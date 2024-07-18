#requires -Modules @{ ModuleName="PnP.PowerShell"; ModuleVersion="1.12.0" }

function Find-SharedDocument
{
    [CmdletBinding(DefaultParameterSetName="SharePoint")]
    param
    (
        [Parameter(Mandatory=$true,ParameterSetName="SpecificSites")]
        [string[]]
        $SiteUrl,

        [Parameter(Mandatory=$true,ParameterSetName="TenantSites")]
        [ValidateSet("SharePoint", "OneDrive")]
        [string]
        $SearchScope,

        [Parameter(Mandatory=$true)]
        [string]
        $FilePath
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

        $selectProperties = "SPSiteUrl", "SPWebUrl", "Filename", "FileExtension", "Path", "Created", "LastModifiedTime", "ViewableByExternalUsers", "ContentClass", "SiteId", "InformationProtectionLabelId"


        $columns = @($selectProperties | ForEach-Object { @{ Name="$_"; Expression=[ScriptBlock]::Create("`$_['$_']") }})

        $counter = 1

        $lastDocumentId = 0

        $fileCount = 0

        $objects = New-Object System.Collections.Generic.List[object]

        $startTime = [DateTime]::Now

        $BatchSize = 500
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
            $pagedQuery = "$Query AND IndexDocId > $lastDocumentId"

            try
            {
                Write-Verbose "$(Get-Date) - Executing query: '$pagedQuery'"

                $allResults = Submit-PnPSearchQuery `
                                -Query            $pagedQuery `
                                -SortList         @{ "[DocId]" = "ascending" } `
                                -StartRow         0 `
                                -MaxResults       500 `
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

                if( $lastDocumentId -gt 0 -and $counter % 10 -eq 0 -and $fileCount -lt $totalRows )
                {
                    $average = [Math]::Round( $stopwatch.Elapsed.TotalSeconds / $counter, 2 )
                    
                    $totalseconds = $estimatedbatches * $average

                    Write-Verbose "$(Get-Date) - Estimated Completion: $($startTime.AddSeconds($totalseconds))"
                }

                if( $lastDocumentId -eq 0 )
                {
                    $totalRows        = $results.TotalRows
                    $estimatedbatches = [Math]::Max( [Math]::Round( $totalRows/500, 0), 1)

                    Write-Verbose "$(Get-Date) - Estimated Page Count: $estimatedbatches"
                }

                Write-Verbose "$(Get-Date) - Processing Page: $($counter), Page Result Count: $($results.RowCount)"

                $lastDocumentId = $results.ResultRows[-1]["DocId"].ToString()
                
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
Connect-PnPOnline `
        -Url "https://contoso.sharepoint.com" `
        -Interactive `
        -ForceAuthentication

# scan spo documents
Find-SharedDocument `
        -SearchScope SharePoint `
        -FilePath "C:\temp\SharedDocuments-SharePoint_$timestamp.csv" `
        -Verbose

# scan onedrive documents
Find-SharedDocument `
        -SearchScope OneDrive `
        -FilePath "C:\temp\SharedDocuments-OneDrive_$timestamp.csv" `
        -Verbose