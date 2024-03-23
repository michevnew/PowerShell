# Remove user from all groups in Office 365
Use this script to remove a given user, or a list of users, as members from any groups. This task is usually performed when a person is leaving the company.

By default, the script only covers Exchange related groups: Distribution groups and Mail-Enabled Security groups. To include Office 365 (modern) groups, specify the -IncludeOffice365Groups parameter. To also include Azure AD Security groups, specify the -IncludeAADSecurityGroups parameter. The latter will require a recent version of the Azure AD PowerShell module. For download and install instructions refer to the PowerShell Gallery: https://www.powershellgallery.com/packages/AzureAD/

The script will detect and reuse any existing sessions to Exchange Online and/or Azure AD. Basic functionality to facilitate a new connection is also included, but it does not cover all possible scenarios. If you are using an MFA-protected account or any of the non-MT Office 365 instances, make sure you connect manually before running the script. If you need help connecting PowerShell to Exchange Online, follow the steps in this article: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx

When fetching Exchange data, the script will use Invoke-Command to get a minimum set of attributes returned in order to speed up execution. To avoid throttling, some artificial delay is added upon executing the removal process for each user. If needed, adjust lines 104 and 117 accordingly.

The removal action is performed via the Remove-UserFromAllGroups cmdlet. The script can be dot-sourced to expose it for use in other scripts. Additional function is used to handle connectivity checks. Parameters are passed via splatting, here's an example usage:

```powershell
.\Remove_User_All_Groups.ps1 -Identity leaver
```
The following example will remove two users, userA and userB from all types of groups:

```powershell
.\Remove_User_All_Groups.ps1 -Identity UserA,UserB -IncludeAADSecurityGroups -IncludeOffice365Groups
```
You can also pass values for the Identity parameter over the pipeline. To avoid potential errors, it's recommended to first run the script with the -WhatIf parameter in such scenarios:

```powershell
Get-User -Filter {CountryOrRegion -eq "US"} | Remove-UserFromAllGroups -IncludeAADSecurityGroups -IncludeOffice365Groups -WhatIf -Verbose
```
Additional information about the script can be found in this article: https://www.michev.info/Blog/Post/2161/script-to-remove-users-from-all-groups-in-office-365

Questions and feedback are welcome.
