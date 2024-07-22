function Get-SharedDocument
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [string]
        $FilePath,

        [Parameter(Mandatory=$true)]
        [string]
        $SiteUrl
    )

    begin
    {
        $SiteUrl = $SiteUrl.TrimEnd("/").Trim()

        $headers = New-Object System.Collections.Generic.List[string]
    }
    process
    {
        if( -not (Test-Path -Path $FilePath -PathType Leaf) )
        {
            throw "File not found: $FilePath"
        } 

        try
        {
            $streamReader = [System.IO.File]::OpenText( $FilePath )

            # pull out the header row
            $line = $StreamReader.ReadLine()

            # split by comma
            $chunks = $line.Split(",")

            # find the index of the SPSiteUrl column
            $indexOfUrlColumn = $chunks.IndexOf( '"SPSiteUrl"' )

            if( $indexOfUrlColumn -lt 0 )
            {
                throw """SPSiteUrl"" column not found in $FilePath"
            }

            # build a collection of file headers
            $chunks | ForEach-Object -Process { $null = $headers.Add( $_.ToString().TrimStart('"').TrimEnd('"').Trim()) }

            while( -not $streamReader.EndOfStream )
            {
                $line = $StreamReader.ReadLine()
                
                $chunks = $line.Split(",")

                if( $chunks[$indexOfUrlColumn] -eq """$SiteUrl""" )
                {
                    $line | ConvertFrom-Csv -Delimiter "," -Header $headers
                }
            }
        }
        catch
        {
            Write-Error "Failed to parse input file: $_"
        }
        finally
        {
            if( $null -ne $streamReader )
            {
                $streamReader.Close()
                $streamReader.Dispose()
            }            
        }
    }
    end
    {
    }
}

# list of site collections to extact from csv
$sites = "https://contoso.sharepoint.com/sites/marketing", "https://contoso.sharepoint.com/sites/finance"

# path to csv
$filePath = "C:\temp\SharedContent-SharePoint_20240626T1736337909.csv"

# unique timestamp for output file to prevent accidential data duplication
$timestamp = Get-Date -Format FileDateTime

# enumerate sites
foreach( $site in $sites )
{
    $fi = [System.IO.FileInfo]::new( $filePath )

    # look for all matching sites
    $rows = Get-SharedDocument -FilePath $filePath -SiteUrl $site

    if( $rows )
    {
        # take the input file an append the site name to the file name
        $sitename = $site.Split("/")[-1]

        # remove any invalid chars from the site name
        $sitename = [string]::Concat( $sitename.Split( [System.IO.Path]::GetInvalidFileNameChars()) )
        
        # append the site name to the input file name
        $fileName = "{0}_{1}_{2}{3}" -f $fi.BaseName, $sitename, $timestamp, $fi.Extension

        $path = Join-Path -Path $fi.Directory.FullName -ChildPath $filename

        # save as csv in the same location as input csv
        $rows | Export-Csv -Path $path -NoTypeInformation

        Write-Host "Saved $($rows.Count) rows for $($site) to $path"
    }
    else 
    {
        Write-Host "$($rows.Count) rows found for $($site)."
    }
}
