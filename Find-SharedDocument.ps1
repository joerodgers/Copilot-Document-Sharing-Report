#requires -Modules @{ ModuleName="PnP.PowerShell"; ModuleVersion="1.12.0" }

function Find-SharedDocument
{
    [CmdletBinding(DefaultParameterSetName="SharePoint")]
    param
    (
        # optional list of site collection URLs to scan.
        [Parameter(Mandatory=$true,ParameterSetName="SpecificSites")]
        [string[]]
        $SiteUrl,

        # set's search stop to either sharepoint or onedrive
        [Parameter(Mandatory=$true,ParameterSetName="TenantSites")]
        [ValidateSet("SharePoint", "OneDrive")]
        [string]
        $SearchScope,

        # full file path to export search results 
        [Parameter(Mandatory=$true)]
        [string]
        $FilePath,

        # advanced property that allows you to page results at a specifc value between 1 and 500.  Default/max is 500.
        [Parameter(Mandatory=$false)]
        [ValidateRange(1,500)]
        [int]
        $MaxResults = 500,

        # advanced property that allows starting an scan using a specific DocId. Default is 0.
        [Parameter(Mandatory=$false)]
        [Int64]
        $IndexDocId = 0
    )

    begin
    {
        if( $PSCmdlet.ParameterSetName -eq "TenantSites" )
        {
            $context = Get-PnPContext -ErrorAction Stop

            $tenant = [System.Uri]::new($context.Url).Host.Split(".")[0]

            if( $SearchScope -eq "SharePoint" )
            {
                $tenantUrl = "https://$tenant.sharepoint.com"
            }
            else
            {
                $tenantUrl = "https://$tenant-my.sharepoint.com"
            }
            
            $query = "IsDocument:true path:$tenantUrl"
        }
        else
        {
            $paths = $SiteUrl | ForEach-Object { "path:""$($_.ToLower().Trim().TrimEnd('/'))/""" }

            $query = "IsDocument:true ($($paths -join " OR "))"
        }

        $selectProperties = "SPSiteUrl", "SPWebUrl", "Filename", "Path", "Created", "LastModifiedTime", "ViewableByExternalUsers", "SiteId", "InformationProtectionLabelId"

        $columns = @($selectProperties | ForEach-Object { @{ Name="$_"; Expression=[ScriptBlock]::Create("`$_['$_']") }})

        $counter = 1

        $fileCount = 0

        $objects = New-Object System.Collections.Generic.List[object]

        $startTime = [DateTime]::Now
    }
    process
    {
        if( $query.Length -gt 4096 )
        {
            throw "Query length cannot exceed 4,096 characters. The current query is $($query.Length) characters. Please reduce the number of Site URLs and retry execution."
        }

        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        $attempts = 1

        while( $true )
        {
            $pagedQuery = "$Query AND IndexDocId > $IndexDocId"

            try
            {
                Write-Verbose "$(Get-Date) - Executing query: '$pagedQuery'"

                $allResults = Submit-PnPSearchQuery `
                                -Query            $pagedQuery `
                                -SortList         @{ "[DocId]" = "ascending" } `
                                -StartRow         0 `
                                -MaxResults       $MaxResults `
                                -TrimDuplicates   $true `
                                -SelectProperties $SelectProperties `
                                -ErrorAction      Stop
                
                $results = $allResults | Where-Object -Property TableType -eq "RelevantResults"
            }
            catch
            {
                if( $attempts -le 10 )
                {
                    $seconds = ($attempts * 10)

                    Write-Warning "Failed to process page $counter on attempt $attempts, retrying in $seconds seconds."

                    Write-Verbose "$(Get-Date) - Error detail: $($_)"
                    Write-Verbose "$(Get-Date) - Exception detail: $($_.Exception.ToString())"

                    Start-Sleep -Seconds $seconds

                    $attempts++

                    continue
                }

                throw "Failed to process page: $($_)"
            }

            $attempts = 1

            if( $null -ne $results -and $null -ne $results.ResultRows -and $results.RowCount -gt 0 )
            {
                $fileCount += $results.RowCount

                if( $IndexDocId -gt 0 -and $counter % 10 -eq 0 -and $fileCount -lt $totalRows )
                {
                    $average = [Math]::Round( $stopwatch.Elapsed.TotalSeconds / $counter, 2 )
                    
                    $totalseconds = $estimatedbatches * $average

                    Write-Verbose "$(Get-Date) - Estimated Completion: $($startTime.AddSeconds($totalseconds))"
                }

                if( $IndexDocId -eq 0 )
                {
                    $totalRows        = $results.TotalRows
                    $estimatedbatches = [Math]::Max( [Math]::Round( $totalRows/500, 0), 1)

                    Write-Verbose "$(Get-Date) - Estimated Page Count: $estimatedbatches"
                }

                Write-Verbose "$(Get-Date) - Processing Page: $($counter), Page Result Count: $($results.RowCount)"

                $IndexDocId = [Int64]$results.ResultRows[-1]["DocId"]
                
                if( $counter -eq 1 )
                {
                    # need header on the first row
                    $rows = ($results.ResultRows | Select-Object $columns | ConvertTo-Csv -NoTypeInformation) -as [System.Collections.Generic.List[object]]
                }
                else
                {
                    # no header on subsequent exports
                    $rows = ($results.ResultRows | Select-Object $columns | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1) -as [System.Collections.Generic.List[object]]
                }

                $objects.AddRange($rows)

                if( $objects.Count -ge 5000 )
                {
                    $sw = Measure-Command -Expression { $objects | Out-File -FilePath $FilePath -Append -Confirm:$false }

                    Write-Verbose "$(Get-Date) - `t`tFlushed $($objects.Count) rows in $([Math]::Round( $sw.TotalMilliseconds, 0))ms"

                    $objects.Clear()
                }
            }
            else
            {
                break
            }

            $counter++
        }
    
        if( $objects.Count -gt 0 )
        {
            Write-Verbose "$(Get-Date) - Flushing remaining $($objects.Count) rows to disk."

            if( $PSVersionTable.PSVersion.Major -le 5 )
            {
                $objects | Out-File -FilePath $FilePath -Append -Confirm:$false -Encoding ascii # allows for no fuss opening in Excel
            }
            else
            {
                $objects | Out-File -FilePath $FilePath -Append -Confirm:$false
            }
            
            $objects.Clear()
        }
        
        Write-Verbose "$(Get-Date) - Exported $($fileCount) search results in $([Math]::Round( $stopwatch.Elapsed.TotalMinutes, 2)) minutes."
    }
    end
    {
    }
}

# Permissoins Required: RUN AS ACCOUNT WITH NO ACCESS TO SHAREPIONT OR ONEDRIVE CONTENT

$timestamp = Get-Date -Format FileDateTime

# update with your tenant admin site URL
Connect-PnPOnline -Url "https://contoso.sharepoint.com" `
                  -Interactive `
                  -ForceAuthentication

# update path to the location to write the SharedContent-SharePoint_<timestamp>.csv file
Find-SharedDocument -SearchScope SharePoint `
                    -FilePath "C:\temp\SharedDocuments-SharePoint_$timestamp.csv" `
                    -Verbose

# update path to the location to write the SharedContent-OneDrive_<timestamp>.csv file
Find-SharedDocument -SearchScope OneDrive `
                    -FilePath "C:\temp\SharedDocuments-OneDrive_$timestamp.csv" `
                    -Verbose
