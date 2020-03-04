# Office 365 Send on behalf of Permission Inventory

The script lists all recipients of the specified type(s) that have at least one delegate with Send on behalf of permissions configured. Running the script without any parameter will return entries across all User, Shared, Room, Equipment, Discovery, Team and Group mailboxes, as well as Distribution and Mail-enabled Security Groups.

Server-side filtering is used to make sure only recipients with non-empty GrantSendOnBehalfTo values are returned, which should significantly reduce the amount of time it takes to run the script even in large environments.

The script does not handle connectivity to Exchange Online, due to the variety of methods now available. If any existing session is detected, it will be used to run the script, otherwise an error will be thrown. If you need help connecting PowerShell to Exchange Online, follow the steps in this article: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx

The script includes a single cmdlet, Get-SOBPermissionInventory, and can be dot-sourced to expose it for use in other scripts. Parameters are passed via splatting, here's an example usage:
```PowerShell
.\Mailbox_Permissions_inventory_SOB.ps1 -IncludeUserMailboxes -IncludeSharedMailboxes
```
By default the script outputs the results to the console host and also stores them in the $varPermissions variable. To export the results to CSV file, use the output variable or un-comment the last line of the script:
```PowerShell
Get-SOBPermissionInventory -IncludeAll -OutVariable var  
$var | Export-Csv -NoTypeInformation "accessrights.csv"
```
Additional information about the script can be found at: https://www.cogmotive.com/blog/powershell-scripts/office-365-permissions-inventory-send-on-behalf-of

Questions and feedback are welcome.
