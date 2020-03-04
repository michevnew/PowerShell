# Office 365 MemberOf Inventory Report

Use this script to generate a report of all users of the specified type and their corresponding group membership. In other words, the script will generate a report of the value of the MemberOf parameter for each user. The report is generated via Exchange server-side filtering, more specifically a filter based on the MemberOf attribute. This in turn means that only groups recognized by Exchange will be included: Distribution groups, Mail-Enabled Security groups  and Office 365 (modern) groups.

By default, the script covers only User mailboxes, if you want to include other recipient types use the corresponding parameters. Supported recipient types include User mailboxes, Shared mailboxes, Mail Users, Mail Contacts, Guest Users. To generate the report for All recipient types, use the -IncludeAll switch.

The script will detect and reuse any existing sessions to Exchange Remote PowerShell. Basic functionality to facilitate a new connection is also included, but it does not cover all possible scenarios. If you are using an MFA-protected account or any of the non-MT Office 365 instances, make sure you connect manually before running the script. If you need help connecting PowerShell to Exchange Online, follow the steps in this article: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx

When fetching Exchange data, the script will use Invoke-Command to get a minimum set of attributes returned in order to speed up execution. If you need additional attributes added to the report, make sure to add them inside the Invoke-Command scriptblocks on lines 74-82. Similarly, if you want to change the attribute used to identify the recipient (PrimarySMTPaddress, ExternalEmailAddress or ExternalDirectoryObjectId), or the Group (PrimarySmtpAddress), make sure to add them to the relevant script blocks.

To avoid throttling, some artificial delay can be added by uncommenting line 100. 

The retrieval action is performed via the Get-DGMembershipInventory cmdlet. The script can be dot-sourced to expose the cmdlet for use in other scripts. Additional function is used to handle connectivity checks. Parameters are passed via splatting, here's an example usage: 
```PowerShell
.\DG_MemberOf_inventory.ps1 -IncludeGuestUsers
```

The above example will generate a report for the group membership of all Guest Users in the organization. To generate a report for all supported recipient types, use:
```PowerShell
.\DG_MemberOf_inventory.ps1 -IncludeAll
```

The script will save the report to a CSV file in the working directory and will also store it in a global variable ($varDGMemberOf) for reuse. Additional information about the script can be found in [this article](https://www.michev.info/Blog/Post/2250/generating-a-report-of-users-group-membership-memberof-inventory). 

Questions and feedback are welcome.
