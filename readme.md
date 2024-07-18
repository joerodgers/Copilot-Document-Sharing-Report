## Overview
The collection of PowerShell scripts provided is designed to assist organizations implementing M365 Copilot in identifying documents in SharePoint and OneDrive shared broadly with [_Everyone_](https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/access-management/grant-everyone-claim-to-external-users) or [_Everyone Except External Users (EEEU)_](https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/access-management/grant-everyone-claim-to-external-users). Once the process of pinpointing these extensively shared documents is finished, a subsequent script can be executed to compile a comprehensive summary report that encompasses detailed metadata pertaining to the SharePoint or OneDrive site where they are housed. This equips administrators with the necessary insights to focus on sites with a specific visibility or template type, usage or popularity, or the existance of highly sensitive files.

Example screenshots of a final summary report:

<p align="center" width="100%">
    <kbd><img src="https://github.com/joerodgers/Copilot-Document-Sharing-Report/blob/main/assets/summary-report-1.png" width="800"></kbd>
</p>

<p align="center" width="100%">
    <kbd><img src="https://github.com/joerodgers/Copilot-Document-Sharing-Report/blob/main/assets/summary-report-2.png" width="800"></kbd>
</p>

<p align="center" width="100%">
    <kbd><img src="https://github.com/joerodgers/Copilot-Document-Sharing-Report/blob/main/assets/summary-report-3.png" width="800"></kbd>
</p>


## PowerShell Requirements
- Windows PowerShell 5.1 or higher
- [PnP.PowerShell](https://www.powershellgallery.com/packages/PnP.PowerShell) module version 1.12.0 or higher



## Access Requirements

#### [_Find-SharedDocument.ps1_](https://github.com/joerodgers/Copilot-Document-Sharing-Report/blob/main/Find-SharedDocument.ps1)
* A user account, whether on a domain or in the cloud, requires minimum read permissions exclusively to the root sites of SharePoint and OneDrive. This access is essential for the user account to perform the search queries needed to locate documents shared across the tenants.

#### [_New-SharedDocumentSummaryReport.ps1_](https://github.com/joerodgers/Copilot-Document-Sharing-Report/blob/main/New-SharedDocumentSummaryReport.ps1)
* This script can be run under a user (delegated) or service principal (application) context.

  - User Context (Delegated)
    - SharePoint and/or Global Administrator
    - Microsoft Graph > Delegated > Groups.Read.All
    - Microsoft Graph > Delegated > Reports.Read.All
    - Microsoft Graph > Delegated > InformationProtectionPolicy.Read.All
    - Microsoft Graph > Delegated > Tasks.Read.All
    
  - Service Principal Context (Application)
    - SharePoint > Application > Sites.FullControl.All
    - Microsoft Graph > Application > Groups.Read.All  
    - Microsoft Graph > Application > Reports.Read.All
    - Microsoft Graph > Application > InformationProtectionPolicy.Read.All
    - Microsoft Graph > Application > Tasks.Read.All

#### [_Export-SharedDocumentRowsToCsv.ps1_](https://github.com/joerodgers/Copilot-Document-Sharing-Report/blob/main/Export-SharedDocumentRowsToCsv.ps1)
* No M365 service permissions are required, all processing is local.

## Script Execution and Output

#### [_Find-SharedDocument.ps1_](https://github.com/joerodgers/Copilot-Document-Sharing-Report/blob/main/Find-SharedDocument.ps1)

#### Script Execution
  1. Update lines #191, #198 and #204 with your environment specific values.
  2. From a PowerShell prompt, execute _Find-SharedDocument.ps1_.  Enter user credentials when prompted.
  3. Wait patiently, the script will take several hours to complete on large tenants.

#### Script Output
The following document properties will be exported to a .csv file:
  - SPSiteUrl
  - SPWebUrl
  - Filename
  - FileExtension
  - Path
  - Created
  - LastModifiedTime
  - ViewableByExternalUsers
  - ContentClass
  - SiteId
  - InformationProtectionLabelId

#### [_New-SharedDocumentSummaryReport.ps1_](https://github.com/joerodgers/Copilot-Document-Sharing-Report/blob/main/New-SharedDocumentSummaryReport.ps1)
#### Script Execution
1. Update lines #366 #378 and #380 with your environment specific values. If using service principal authentication, update the admin URL, ClientId, Thumbprint and TenantId values.
2. Execute _New-SharedDocumentSummaryReport.ps1_.  If you are using a user context for authentication, enter a SharePoint/Global administrator credential.
3. Wait patiently, the script may also take several hours to complete on large outputs.

#### Script Output
The following document summary and site properties will be exported to a .csv file:

|Column Name| Column Description|
|-|-|
|SiteUrl|Url of the SharePoint or OneDrive Site|
|SharedDocumentCount|Total number of shared documents found for the site|
|SharedDocumentsViewableByExternalUsersCount|Total number of shared documents viewable by guests for the site|
|SiteSensitivityLabel|Sensitivity Label applied to the site|
|UnlabedFileCount|Total number of shared documents without a sensitivity label applied|
|\<DYNAMIC COLUMN\>|Total number of shared documents with specific sensitivity label applied|
|SiteVisibility|Visibility of the connected M365 Group, otherwise blank|
|OwnerLoginName | The primary owner or M365 Group owners |
|IsGroupConnected | True if group connected, otherwise False| 
|IsTeamsConnected | True if connected to a Microsoft Team, otherwise false| 
|IsVivaEngageConnected | True if connected to a Viva Engage Community, otherwise false (added 07-18-24)| 
|IsPlannerConnected | True if connected to one or more Planner Plans, otherwise false (added 07-18-24)| 
|LastContentModifiedDate | Date and time (UTC) when the content of the site was last changed| 
|LastActivityDate | The date of the last time file activity was detected or a page was viewed on the site| 
|RootWebTemplate | Template of the site collection's root web site| 
|SiteFileCount | The number of files on the site|
|SiteSharingCapability | The value of the external sharing setting for the site|
|SiteStorageMB | The amount of storage currently being used on the site|
|SubSiteCount | Total number of sub webs for the site|
|GroupId | GUID of the M365 group, otherwise a empty GUID|
|SiteId | GUID of the SharePoint site|
|ActiveFileCount | The number of active files on the site. A file is considered active if it has been saved, synced, modified, or shared within the specified time period|
|PageViewCount | The number of times pages were viewed on the site|
|VisitedPageCount | The number of unique pages that were visited on the site|


#### [_Export-SharedDocumentRowsToCsv.ps1_](https://github.com/joerodgers/Copilot-Document-Sharing-Report/blob/main/Export-SharedDocumentRowsToCsv.ps1)
#### Script Execution
1. Update lines #81 and #84 with the desired URLs and path to the existing .csv file.
2. Execute Export-SharedDocumentRowsToCsv.ps1

#### Script Output
A new CSV file will be generated using for each site URL supplied.  The new .csv files will be saved to the same directory as the input csv file.  The new .csv file names will have the name of the input .csv file, suffixed with the site's URL name.