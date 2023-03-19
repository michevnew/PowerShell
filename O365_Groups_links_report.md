# Office 365 Group links (membership) report

The script reports the Links for all Office 365 Groups, more specifically the following link types: Owners, Members and Subscribers. Two additional link types exist currently, Aggregators and EventSubscribers, but those are not yet used in the service.

The script will detect and reuse any existing sessions to Exchange Online. Basic functionality to facilitate a new connection is also included, but it does not cover all possible scenarios. If you are using an MFA-protected account or any of the non-MT Office 365 instances, make sure you connect manually before running the script. If you need help connecting PowerShell to Exchange Online, follow the steps in this article: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx

 When fetching ata from Exchange, the script will use Invoke-Command to get a minimum set of attributes returned in order to speed up execution. If additional attributes are required, they need to be added to the relevant script block first (line 44). Similarly, if additional attributes are required for the Link object, include them in lines 64-67. A small delay is added on line 61 as a simple anti-throttling measure, update it as necesary.
 
The script includes a single cmdlet, Get-O365GroupMembershipInventory, and can be dot-sourced to expose it for use in other scripts. Here's an example usage:
```PowerShell
.\O365_Groups_links_report.ps1 
```
By default the script outputs the result to a CSV file and also stores it in the $varO365GroupMembers global variable. If you want to modify the output, you can do so in the following fashion (the example assumes you have dot-sourced the script to expose the Get-O365GroupMembershipInventory cmdlet):
```PowerShell
Get-O365GroupMembershipInventory -CondensedOutput -OutVariable global:var 
$var | ? {$_.MemberType -eq "Owner"} | Sort Member -Descending | Export-Csv -NoTypeInformation "O365GroupLinks.csv"
```
The default output is one line per each Link entry, which generates a CSV file easy to filter by either Group, User or Link type. If you prefer a smaller output file, with one entry per Group object, use the CondensedOutput parameter: 
```PowerShell
.\O365_Groups_links_report.ps1 -CondensedOutput
```
Additional information about the script can be found in [this article](https://www.michev.info/blog/post/2101/reporting-on-membership-of-office-365-groups).

Questions and feedback are welcome.
