#requires -Modules @{ ModuleName="PnP.PowerShell"; ModuleVersion="1.12.0" }

function New-SharedDocumentSummaryReport
{
    [CmdletBinding()]
    param
    (
        # path to SharedContent-SharePoint_<timestamp>.csv or SharedContent-OneDrive_<timestamp>.csv  file.
        [Parameter(Mandatory=$true)]
        [string]
        $Path,

        # optional switch to indicated that only a basic summary report including SiteUrl, SharedDocumentCount, SharedDocumentsViewableByExternalUsersCount will be generated.
        [Parameter(Mandatory=$false)]
        [switch]
        $BasicSummaryReport
    )

    begin
    {
        # build a dictionary to track the number of files by label 
        $documentsBySensitivityLabel = [System.Collections.Generic.Dictionary[Guid, int]]::new()

        # don't fetch service data that is not required for a basic report
        if( -not $BasicSummaryReport.IsPresent )
        {   
            $connection = Get-PnpConnection -ErrorAction Stop

            # default to app-only endpoint
            $graphSensitivityLabelEndpoint = "/beta/security/informationProtection/sensitivityLabels"

            if( $connection.ClientId -eq [PnP.Framework.AuthenticationManager]::CLIENTID_PNPMANAGEMENTSHELL <# 31359c7f-bd7e-475c-86db-fdb8c937548e #> )
            {
                # delegated connection
                $graphSensitivityLabelEndpoint = "/beta/me/security/informationProtection/sensitivityLabels"
            }
            
            # need to use raw api to fetch label formats
            $availableSensitivityLabels = Invoke-PnPGraphMethod -Method Get -Url $graphSensitivityLabelEndpoint -Verbose:$false -ErrorAction Stop | Select-Object -ExpandProperty value

            # pull out file labels
            $availableFileSensitivityLabels = $availableSensitivityLabels | Where-Object -Property contentFormats -Contains "file" | Select-Object Id, Name

            $availableSensitivityLabels | ForEach-Object { Write-Verbose "$(Get-Date) - Available Sensitivity Label: Name:$($_.Name), Id:$($_.Id), Usage:$($_.contentFormats -join ";") "}

            $documentsBySensitivityLabel.Add( [Guid]::Empty, 0 ) # default label that represents an unlabeled file
            
            foreach( $availableFileSensitivityLabel in $availableFileSensitivityLabels )
            {
                $documentsBySensitivityLabel.Add( $availableFileSensitivityLabel.id, 0 )
            }
        }

        # storage for SharedDocumentSummaryModel objects
        $sharedDocumentSummaryModels = [System.Collections.Generic.Dictionary[string, PSCustomObject]]::new( [System.StringComparer]::InvariantCultureIgnoreCase )
        
        # input csv headers
        $headers = New-Object System.Collections.Generic.List[string]
    }
    process
    {
        $fileInfo = [System.IO.FileInfo]::new( $Path )

        if( -not $fileInfo.Exists )
        {
            throw "File not found: $Path"
        }

        # generate summary csv file name
        $exportPath = Join-Path -Path $fileInfo.Directory.FullName -ChildPath "$($fileInfo.BaseName)_BasicSummaryReport$($fileInfo.Extension)"

        $counter = 0

        try
        {
            Write-Verbose "$(Get-Date) - Parsing file: $Path"

            $streamReader = [System.IO.File]::OpenText( $Path )
        
            while( -not $streamReader.EndOfStream )
            {
                if( $counter -eq 0 )
                {
                    # calculdate csv headers
                    $firstrow = $streamReader.ReadLine()

                    $firstrow -split "," | ForEach-Object -Process { $null = $headers.Add( $_.ToString().TrimStart('"').TrimEnd('"').Trim()) }
                }
                
                if( $counter -gt 0 -and $counter % 5000 -eq 0 )
                {
                    Write-Verbose "$(Get-Date) - Processed $counter rows"
                }

                $counter++

                $line = $StreamReader.ReadLine()

                # using this slower option to ensure any column values containing double quotes doesn't cause issues with parsing
                $row = $line | ConvertFrom-Csv -Delimiter "," -Header $headers

                if( [string]::IsNullOrWhiteSpace($row.InformationProtectionLabelId) )
                {
                    # set as empty file label
                    $row.InformationProtectionLabelId = [Guid]::Empty.ToString() 
                }

                if( -not $sharedDocumentSummaryModels.ContainsKey( $row.SPSiteUrl ) )
                {
                    # init a new row for this site
                    Write-Verbose "$(Get-Date) - Discovered site: $($row.SPSiteUrl)"

                    $sharedDocumentSummaryModels.Add( $row.SPSiteUrl, [SharedDocumentSummaryModel]::new() )

                    $sharedDocumentSummaryModels[$row.SPSiteUrl].SharedDocumentByLabelCount = [System.Collections.Generic.Dictionary[Guid, int]]::new($documentsBySensitivityLabel)
                    $sharedDocumentSummaryModels[$row.SPSiteUrl].SiteId                     = $row.SiteId
                    $sharedDocumentSummaryModels[$row.SPSiteUrl].SiteUrl                    = $row.SPSiteUrl
                }

                $viewableByExternalUsers = $false
                
                if( [bool]::TryParse($row.ViewableByExternalUsers, [ref]$viewableByExternalUsers) -and $viewableByExternalUsers )
                {
                    $sharedDocumentSummaryModels[$row.SPSiteUrl].ViewableByExternalUsersCount++
                }

                # update row properties
                $sharedDocumentSummaryModels[$row.SPSiteUrl].SharedDocumentCount++
                $sharedDocumentSummaryModels[$row.SPSiteUrl].SharedDocumentByLabelCount[ $row.InformationProtectionLabelId ]++
            }
        }
        finally
        {
            if( $null -ne $streamReader )
            {
                $streamReader.Close()
                $streamReader.Dispose()
            }
        }

        if( $BasicSummaryReport.IsPresent )
        {
            $basicSummaryResults = foreach( $key in $sharedDocumentSummaryModels.Keys )
            {
                $sharedDocumentSummaryModel = $sharedDocumentSummaryModels[$key]

                $sharedDocumentSiteSummaryModel = [SharedDocumentSiteSummaryModel]::new()
                $sharedDocumentSiteSummaryModel.SiteUrl                                     = $sharedDocumentSummaryModel.SiteUrl
                $sharedDocumentSiteSummaryModel.SharedDocumentCount                         = $sharedDocumentSummaryModel.SharedDocumentCount
                $sharedDocumentSiteSummaryModel.SharedDocumentsViewableByExternalUsersCount = $sharedDocumentSummaryModel.ViewableByExternalUsersCount

                $sharedDocumentSiteSummaryModel
            }
    
            $basicSummaryResults | Sort-Object -Property SharedDocumentCount -Descending | Select-Object SiteUrl, SharedDocumentCount, SharedDocumentsViewableByExternalUsersCount | Export-Csv -Path $exportPath -NoTypeInformation
            
            Write-Host "Basic summary report written to $exportPath"
            
            return
        }

        if( $sharedDocumentSummaryModels[0].Key -match "-my.sharepoint" )
        {
            Write-Verbose "$(Get-Date) - Downloading OneDrive usage details"
            $usageReportUrl = '/beta/reports/getOneDriveUsageAccountDetail(period=''D30'')?$format=application/json&$top=999'
        }
        else
        {
            Write-Verbose "$(Get-Date) - Downloading SharePoint usage details"
            $usageReportUrl = '/beta/reports/getSharePointSiteUsageDetail(period=''D30'')?$format=application/json&$top=999'
        }

        # pull site usage details report
        $siteUsageDetails = Invoke-PnPGraphMethod -Method Get -Url $usageReportUrl -All -Verbose:$false | Select-Object -ExpandProperty value
        
        # enumerate sites found in input .csv
        $results = foreach( $key in $sharedDocumentSummaryModels.Keys )
        {
            $sharedDocumentSummaryModel = $sharedDocumentSummaryModels[$key]

            $sharedDocumentSiteSummaryModel = [SharedDocumentSiteSummaryModel]::new()
            $sharedDocumentSiteSummaryModel.SiteUrl                                     = $sharedDocumentSummaryModel.SiteUrl
            $sharedDocumentSiteSummaryModel.SiteId                                      = $sharedDocumentSummaryModel.SiteId
            $sharedDocumentSiteSummaryModel.SharedDocumentCount                         = $sharedDocumentSummaryModel.SharedDocumentCount
            $sharedDocumentSiteSummaryModel.SharedDocumentsViewableByExternalUsersCount = $sharedDocumentSummaryModel.ViewableByExternalUsersCount
            
            $uniqueLabelNames = @()

            # dynamically add file label names to model
            foreach( $labelId in $sharedDocumentSummaryModel.SharedDocumentByLabelCount.Keys )
            {
                if( $labelId -eq [Guid]::Empty )
                {
                    $sharedDocumentSiteSummaryModel | Add-Member -MemberType NoteProperty -Name "UnlabedFileCount" -Value $sharedDocumentSummaryModel.SharedDocumentByLabelCount[$labelId]
                }
                else
                {
                    if( $label = $availableSensitivityLabels | Where-Object -Property Id -eq $labelId.ToString() )
                    {
                        $labelName = $label.Name -replace " ", "" # remove spaces in label name

                        $columnName = $labelName + "LabelFileCount" # append "LabelFileCount to the label name
                    
                        while( $uniqueLabelNames -contains $columnName )
                        {
                            Write-Verbose "$(Get-Date) - Adding LabelId to name $($label.Name) to make the column name unique."

                            $columnName = $labelName + "_$($labelId)_LabelFileCount" # make the label name unique
                        }
                    }
                    else
                    {
                        $columnName = "$($labelId)_LabelFileCount"
                    }

                    $uniqueLabelNames += $columnName

                    $sharedDocumentSiteSummaryModel | Add-Member -MemberType NoteProperty -Name $columnName -Value $sharedDocumentSummaryModel.SharedDocumentByLabelCount[$labelId]
                }
            }

            Write-Verbose "$(Get-Date) - Looking up site properties for $($sharedDocumentSiteSummaryModel.SiteUrl)"
            
            try
            {
                $siteProperties = Get-PnPTenantSite -Identity $sharedDocumentSiteSummaryModel.SiteUrl -Detailed -ErrorAction Stop -Verbose:$false
                
                $sensitivityLabel = $siteProperties.SensitivityLabel # default to the label guid in case the label is not found

                if( $availableSensitivityLabels | Where-Object -Property Id -eq $siteProperties.SensitivityLabel )
                {
                    $sensitivityLabel = $availableSensitivityLabels | Where-Object -Property Id -eq $siteProperties.SensitivityLabel | Select-Object -ExpandProperty Name
                }

                $sharedDocumentSiteSummaryModel.GroupId                 = $siteProperties.GroupId
                $sharedDocumentSiteSummaryModel.IsGroupConnected        = $siteProperties.GroupId -ne [Guid]::Empty
                $sharedDocumentSiteSummaryModel.IsTeamsConnected        = $false
                $sharedDocumentSiteSummaryModel.LastContentModifiedDate = $siteProperties.LastContentModifiedDate
                $sharedDocumentSiteSummaryModel.OwnerLoginName          = $siteProperties.OwnerLoginName -replace "i:0#\.f\|membership\|", ""
                $sharedDocumentSiteSummaryModel.RootWebTemplate         = $siteProperties.Template
                $sharedDocumentSiteSummaryModel.SiteSensitivityLabel    = $sensitivityLabel
                $sharedDocumentSiteSummaryModel.SiteSharingCapability   = $siteProperties.SiteDefinedSharingCapability
                $sharedDocumentSiteSummaryModel.SiteStorageMB           = $siteProperties.StorageUsageCurrent
                $sharedDocumentSiteSummaryModel.SubSiteCount            = $siteProperties.WebsCount - 1

                # merge site usage with current record
                if( $siteUsageDetail = [Linq.Enumerable]::FirstOrDefault([Linq.Enumerable]::Where( $siteUsageDetails, [Func[Object,bool]]{ param($object) $object.SiteId -eq $sharedDocumentSiteSummaryModel.SiteId } )) )
                {
                    $sharedDocumentSiteSummaryModel.ActiveFileCount  = $siteUsageDetail.activeFileCount 
                    $sharedDocumentSiteSummaryModel.LastActivityDate = $siteUsageDetail.lastActivityDate
                    $sharedDocumentSiteSummaryModel.PageViewCount    = $siteUsageDetail.pageViewCount  
                    $sharedDocumentSiteSummaryModel.SiteFileCount    = $siteUsageDetail.fileCount 
                    $sharedDocumentSiteSummaryModel.VisitedPageCount = $siteUsageDetail.visitedPageCount
                }

                # lookup the group properties
                if( $sharedDocumentSiteSummaryModel.GroupId -ne [Guid]::Empty )
                {
                    try
                    {
                        Write-Verbose "$(Get-Date) - Looking up M365 group properties for site: $($sharedDocumentSiteSummaryModel.SiteUrl)"

                        $m365Group = Get-PnPMicrosoft365Group -Identity $sharedDocumentSiteSummaryModel.GroupId -IncludeOwners -Verbose:$false -ErrorAction Stop

                        if( $m365Group )
                        {
                            Write-Verbose "$(Get-Date) - Looking up M365 group Viva Engage Communities"
                            $yammerCommunity = Get-PnPMicrosoft365GroupYammerCommunity -Identity $m365Group.Id -Verbose:$false -ErrorAction Stop

                            try
                            {
                                Write-Verbose "$(Get-Date) - Looking up M365 group planner plans"
                                $plannerPlans = Get-PnPPlannerPlan -Group $m365Group.Id.ToString() -Verbose:$false -ErrorAction Stop # requires Microsoft Graph > Tasks.Read.All
                            }
                            catch
                            {
                                Write-Warning "Failed to lookup the group $($siteProperties.GroupId) Planner Plan associations. Error: $($_)"
                            }

                            $sharedDocumentSiteSummaryModel.SiteVisibility        = $m365Group.Visibility
                            $sharedDocumentSiteSummaryModel.IsTeamsConnected      = $m365Group.HasTeam
                            $sharedDocumentSiteSummaryModel.OwnerLoginName        = ($m365Group.Owners.UserPrincipalName -join ",") -replace "i:0#\.f\|membership\|", ""
                            $sharedDocumentSiteSummaryModel.IsVivaEngageConnected = $null -ne  $yammerCommunity
                            $sharedDocumentSiteSummaryModel.IsPlannerConnected    = $null -ne  $plannerPlans

                        }
                    }
                    catch
                    {
                        Write-Warning "Failed to lookup group: $($siteProperties.GroupId). Error: $($_)"
                    }
                }
            }
            catch
            {
                Write-Warning "Failed to process site: $($sharedDocumentSiteSummaryModel.SiteUrl). Error: $($_)"
            }

            $sharedDocumentSiteSummaryModel
        }

        if( $results.Count -eq 0 ) { return }

        # generate summary csv file name
        $exportPath = Join-Path -Path $fileInfo.Directory.FullName -ChildPath "$($fileInfo.BaseName)_SummaryReport$($fileInfo.Extension)"

        # output csv column order
        $orderedOutputColumns = [System.Collections.Generic.List[string]]@(
            "SiteUrl", 
            "SharedDocumentCount", 
            "SharedDocumentsViewableByExternalUsersCount", 
            "SiteSensitivityLabel", 
            "SiteVisibility", 
            "OwnerLoginName", 
            "IsGroupConnected", 
            "LastContentModifiedDate", 
            "LastActivityDate", 
            "RootWebTemplate", 
            "SiteFileCount", 
            "SiteSharingCapability", 
            "SiteStorageMB", 
            "SubSiteCount", 
            "IsTeamsConnected",
            "IsVivaEngageConnected",
            "IsPlannerConnected", 
            "ActiveFileCount", 
            "PageViewCount", 
            "VisitedPageCount",
            "GroupId", 
            "SiteId"
        )

        # insert dynamic file label count headers at index 4
        foreach( $noteProperty in $results[0] | Get-Member -MemberType NoteProperty ) 
        {
            $orderedOutputColumns.Insert( 4, $noteProperty.Name )
        }

        # export summary data
        $results | Select-Object $orderedOutputColumns | Sort-Object -Property SharedDocumentCount, SiteUrl -Descending |  Export-Csv -Path $exportPath -NoTypeInformation -Encoding utf8
        
        Write-Host "Summary report written to $exportPath"
    }
    end
    {
    }
}

class SharedDocumentSiteSummaryModel
{
    [int]
    $ActiveFileCount
    
    [Guid]
    $GroupId
    
    [bool]
    $IsGroupConnected
    
    [bool]
    $IsTeamsConnected
    
    [bool]
    $IsVivaEngageConnected

    [bool]
    $IsPlannerConnected

    [Nullable[DateTime]]
    $LastActivityDate
    
    [Nullable[DateTime]]
    $LastContentModifiedDate
    
    [string]
    $OwnerLoginName
    
    [int]
    $PageViewCount
    
    [string]
    $SiteSensitivityLabel
    
    [int]
    $SharedDocumentCount

    [int]
    $SharedDocumentsViewableByExternalUsersCount

    [string]
    $SiteSharingCapability
    
    [int]
    $SiteFileCount
    
    [Guid]
    $SiteId
    
    [string]
    $SiteUrl

    [int]
    $SiteStorageMB
    
    [string]
    $RootWebTemplate
    
    [int]
    $SubSiteCount
    
    [string]
    $SiteVisibility
    
    [int]
    $VisitedPageCount
}

class SharedDocumentSummaryModel
{
    [string]
    $SiteUrl = ""

    [Guid]
    $SiteId = [Guid]::Empty

    [int]
    $SharedDocumentCount = 0

    [int]
    $ViewableByExternalUsersCount = 0

    [System.Collections.Generic.Dictionary[Guid, int]]
    $SharedDocumentByLabelCount = [System.Collections.Generic.Dictionary[Guid, int]]::new()

    SiteCollectionSharedDocumentDetail()
    {
        $this.SharedDocumentByLabelCount.Add( [Guid]::Empty, 0 ) # files with no label will be tagged with empty guid
    }
}


# Permissoins Required
#
#    Delegated Option:   
#        - SharePoint Administrator Role
#        - Microsoft Graph > Delegated > Directory.All
#        - Microsoft Graph > Delegated > Reports.Read.All
#        - Microsoft Graph > Delegated > InformationProtectionPolicy.Read.All (one of the following roles: Global Reader, Organization Management, Security Reader, Compliance Data Administrator, Security Administrator, Compliance Administrator)
#        - Microsoft Graph > Delegated > Tasks.Read.All
#    
#    Application Option: 
#        - SharePoint > Application > Sites.FullControl.All 
#        - Microsoft Graph > Application > Directory.All
#        - Microsoft Graph > Application > Reports.Read.All
#        - Microsoft Graph > Application > InformationProtectionPolicy.Read.All
#        - Microsoft Graph > Application > Tasks.Read.All

<# service prinpcial connection example
Connect-PnPOnline `
        -Url        "https://contoso-admin.sharepoint.com" `
        -ClientId   $env:O365_CLIENTID `
        -Thumbprint $env:O365_THUMBPRINT `
        -Tenant     $env:O365_TENANTID
#>


# update with your tenant admin site URL
Connect-PnPOnline -Url "https://contoso-admin.sharepoint.com" `
                  -ClientId "<YOUR GUID>"
                  -Interactive `
                  -ForceAuthentication

# update path to the location of your SharedContent-SharePoint_<timestamp>.csv file
New-SharedDocumentSummaryReport -Path "C:\temp\SharedDocuments-SharePoint_20240620T1536438280.csv" -Verbose

# update path to the location of your SharedContent-OneDrive_<timestamp>.csv file
New-SharedDocumentSummaryReport -Path "C:\temp\SharedContent-OneDrive__20240620T1536438280.csv" -Verbose