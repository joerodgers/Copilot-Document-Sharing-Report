## Overview
The collection of PowerShell scripts provided is designed to assist organizations implementing M365 Copilot in identifying documents in SharePoint and OneDrive shared broadly with [_Everyone_](https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/access-management/grant-everyone-claim-to-external-users) or [_Everyone Except External Users (EEEU)_](https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/access-management/grant-everyone-claim-to-external-users). Once the process of pinpointing these extensively shared documents is finished, a subsequent script can be executed to compile a comprehensive summary report that encompasses detailed metadata pertaining to the SharePoint or OneDrive site where they are housed. This equips administrators with the necessary insights to focus on sites with a specific visibility or template type, usage or popularity, or the existance of highly sensitive files.

## Script Reference
||Script Name|Description|Required Permissions|Config|Outputs|
|-|-|-|:-:|:-:|:-:|
|:one:|Find-SharedDocument.ps1|Utilizes a specially crafted search query to locate and report all documents shared with Everyone or EEEU.|[details](https://github.com/joerodgers/Copilot-Document-Sharing-Report/tree/main?tab=readme-ov-file#required-permissions)|[details](https://github.com/joerodgers/Copilot-Document-Sharing-Report/tree/main?tab=readme-ov-file#configuration)|[.csv file](https://github.com/joerodgers/Copilot-Document-Sharing-Report/tree/main?tab=readme-ov-file#outputs)|
|:two:|New-SharedDocumentSummaryReport.ps1|Generates a comprehensive report which encompasses document analysis and detailed metadata pertaining to the SharePoint or OneDrive site where they are housed.|[details](https://github.com/joerodgers/Copilot-Document-Sharing-Report/tree/main?tab=readme-ov-file#required-permissions-1)|[details](https://github.com/joerodgers/Copilot-Document-Sharing-Report/tree/main?tab=readme-ov-file#configuration-1)|[.csv file](https://github.com/joerodgers/Copilot-Document-Sharing-Report/tree/main?tab=readme-ov-file#outputs-1)|
|:three:|Export-SharedDocumentRowsToCsv.ps1|An optional script which can be used to extract all rows for specific sites from the .csv file created by *Find-SharedDocument.ps1*. |None|[details](https://github.com/joerodgers/Copilot-Document-Sharing-Report/tree/main?tab=readme-ov-file#configuration-2)|[.csv file](https://github.com/joerodgers/Copilot-Document-Sharing-Report/tree/main?tab=readme-ov-file#outputs-2)|


Example screenshots of a final summary report:

<p align="center" width="100%">
    <kbd><img src="https://github.com/joerodgers/Copilot-Document-Sharing-Report/blob/main/assets/summary-report-6.png"></kbd>
</p>

<p align="center" width="100%">
    <kbd><img src="https://github.com/joerodgers/Copilot-Document-Sharing-Report/blob/main/assets/summary-report-7.png"></kbd>
</p>

<p align="center" width="100%">
    <kbd><img src="https://github.com/joerodgers/Copilot-Document-Sharing-Report/blob/main/assets/summary-report-8.png"></kbd>
</p>

<p align="center" width="100%">
    <kbd><img src="https://github.com/joerodgers/Copilot-Document-Sharing-Report/blob/main/assets/summary-report-9.png"></kbd>
</p>

## PowerShell Requirements
- Windows PowerShell 5.1 or higher
- [PnP.PowerShell](https://www.powershellgallery.com/packages/PnP.PowerShell) module version 1.12.0 or higher

## Register an Entra ID Application to use with PnP PowerShell
>[!NOTE]  
>As of September 9th, 2024, this has become mandatory step. This article will guide you through how to do so.
- [Register an Entra ID Application to use with PnP PowerShell](https://pnp.github.io/powershell/articles/registerapplication)

## Script Details

### :one: Find-SharedDocument.ps1

Utilizes a specially crafted search query to locate and report all documents shared with Everyone or EEEU

#### Required Permissions
A user account, a domain or cloud account, requires a minimum of read permissions exclusively to the root sites of SharePoint and OneDrive. This access is essential for the user account to perform the search queries needed to locate documents shared across the tenants.

|API | Type | Least Privileged Permission | Justification |
|-|-|-|-|
|SharePoint | Delegated | Sites.Search.All or AllSites.Read | Required to query files hosted in SharePoint Online and OneDrive for Business. |

#### Configuration
Before executing the PowerShell script, update the three lines below with your specific values:

``` PowerShell
# Permissoins Required: RUN AS ACCOUNT WITH NO ACCESS TO SHAREPIONT OR ONEDRIVE CONTENT

$timestamp = Get-Date -Format FileDateTime

# update with your tenant admin site URL
Connect-PnPOnline -Url "https://contoso.sharepoint.com" `
                  -ClientId "<YOUR CLIENT ID>" `
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
```


#### Outputs
Generates a .csv file with following columns:
| Column Name | Description |
|-|-|
|SPSiteUrl                    | Url of the site hosting the document.|
|SPWebUrl                     | Url of the web hosting the document.|
|Filename                     | Name and extension of the documen.|
|Path                         | Full Url to the document.|
|Created                      | Date and time the file was created/added.|
|LastModifiedTime             | Date and time the file was last modified.|
|ViewableByExternalUsers      | Boolean that indicates if the document is viewable by people outside of your organization.|
|SiteId                       | GUID of the SharePoint or OneDrive site.|
|InformationProtectionLabelId | GUID of the Purview Sensitivity Label applied to the document, if one is applied.|


### :two: New-SharedDocumentSummaryReport.ps1

Generates a comprehensive report which encompasses document analysis and detailed metadata pertaining to the SharePoint or OneDrive site where they are housed.

#### Required Permissions

Required permissions for user (delegated) authentication:

|API | Type | Least Privileged Permission | Justification |
|-|-|-|-|
|SharePoint           | Role      | SharePoint Administrator         | Required to retrieve tenant site properties.                         |
|SharePoint           | Delegated | AllSites.FullControl             | Required to retrieve tenant site properties.                         |
|Microsoft&nbsp;Graph | Delegated | User.ReadBasic.All               | Required to retrieve M365 Group owner's email addresses.             |
|Microsoft&nbsp;Graph | Delegated | Group.Read.All                   | Required to retrieve M365 Group properties and associated endpoints. |
|Microsoft&nbsp;Graph | Delegated | Reports.Read.All                 | Required to retrieve details about SharePoint site usage.            |
|Microsoft&nbsp;Graph | Delegated | InformationProtectionPolicy.Read | Required to retrieve labels available for the signed-in user.        |

Required permissions for service principal (application) authentication:

|API | Type | Least Privileged Permission | Justification |
|-|-|-|-|
|SharePoint           | Application | Sites.FullControl.All                | Required to retrieve tenant site properties.                                      |
|Microsoft&nbsp;Graph | Application | Groups.Read.All                      | Required to retrieve M365 Group properties and associated endpoints.              |
|Microsoft&nbsp;Graph | Application | User.ReadBasic.All                   | Required to retrieve M365 Group owner's email addresses.                          |
|Microsoft&nbsp;Graph | Application | Reports.Read.All                     | Required to retrieve details about SharePoint site usage.                         |
|Microsoft&nbsp;Graph | Application | InformationProtectionPolicy.Read.All | Required to retrieve labels available to the organization as a service principal. |
|Microsoft&nbsp;Graph | Application | Tasks.Read.All                       | Required to retrieve Planner associations with M365 Groups.                       |

#### Configuration
Before executing the PowerShell script, update the three lines below with your specific values:

``` PowerShell
# update with your tenant admin site URL
Connect-PnPOnline -Url "https://contoso-admin.sharepoint.com" `
                  -ClientId "<YOUR CLIENT ID>" `
                  -Interactive `
                  -ForceAuthentication

# update path to the location of your SharedContent-SharePoint_<timestamp>.csv file
New-SharedDocumentSummaryReport -Path "C:\temp\SharedContent-SharePoint_20240620T1536438280.csv" -Verbose

# update path to the location of your SharedContent-OneDrive_<timestamp>.csv file
New-SharedDocumentSummaryReport -Path "C:\temp\SharedContent-OneDrive__20240620T1536438280.csv" -Verbose

```

#### Outputs
Generates a .csv file with following columns:
| Column Name | Column Description |
|-|-|
|SiteUrl                                     | Url of the SharePoint or OneDrive Site|
|SharedDocumentCount                         | Total number of shared documents found for the site|
|SharedDocumentsViewableByExternalUsersCount | Total number of shared documents viewable by guests for the site|
|SiteSensitivityLabel                        | Sensitivity Label applied to the site|
|UnlabedFileCount                            | Total number of shared documents without a sensitivity label applied|
|\<DYNAMIC COLUMN\>                          | Total number of shared documents with specific sensitivity label applied|
|SiteVisibility                              | Visibility of the connected M365 Group, otherwise blank.|
|OwnerLoginName                              | The primary owner or M365 Group owners |
|IsGroupConnected                            | True if group connected, otherwise False| 
|IsTeamsConnected                            | True if connected to a Microsoft Team, otherwise false| 
|IsVivaEngageConnected                       | True if connected to a Viva Engage Community, otherwise false | 
|IsPlannerConnected                          | True if connected to one or more Planner Plans, otherwise false.  Only populated with using application authentication. | 
|LastContentModifiedDate                     | Date and time (UTC) when the content of the site was last changed| 
|LastActivityDate                            | The date of the last time file activity was detected or a page was viewed on the site| 
|RootWebTemplate                             | Template of the site collection's root web site| 
|SiteFileCount                               | The number of files on the site|
|SiteSharingCapability                       | The value of the external sharing setting for the site|
|SiteStorageMB                               | The amount of storage currently being used on the site|
|SubSiteCount                                | Total number of sub webs for the site|
|GroupId                                     | GUID of the M365 group, otherwise a empty GUID|
|SiteId                                      | GUID of the SharePoint site|
|ActiveFileCount                             | The number of active files on the site. A file is considered active if it has been saved, synced, modified, or shared within the specified time period|
|PageViewCount                               | The number of times pages were viewed on the site|
|VisitedPageCount                            | The number of unique pages that were visited on the site|

### :three: Export-SharedDocumentRowsToCsv.ps1
An ancillary script which will copy all entries for particular sites from the extensive .csv file generated by Find-SharedDocument.ps1 to new a .csv file. Clients have utilized these condensed .csv files to supply site collection administrators with precise data to aid in focused permission correction.

#### Required Permissions
None, all processing is local.

#### Configuration
Before executing the PowerShell script, update the two lines below with your specific values:

``` PowerShell
# update with the list of site collections you want to extact from the larger .csv file
$sites = "https://contoso.sharepoint.com/sites/marketing", "https://contoso.sharepoint.com/sites/finance"

# update path to the location of your SharedContent-<scope>_<timestamp>.csv file
$filePath = "C:\_temp\SharedContent-SharePoint_20240626T1736337909.csv"
```

#### Outputs
A new CSV file will be generated using for each site URL supplied.  The new .csv files will be saved to the same directory as the input csv file.  The new .csv file names will have the name of the input .csv file, suffixed with the site's URL name.

