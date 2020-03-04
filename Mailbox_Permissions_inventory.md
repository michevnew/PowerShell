# Office 365 Mailbox Permission Inventory

The script lists all mailboxes of the specified type(s) that have at least one non-default permission entry. Running the script without any parameter will return entries across all User, Shared, Room, Equipment, Discovery, and Team mailboxes.

As the full list of mailboxes needs to be cycled in order to get the permissions, the script will use Invoke-Command in order to get a minimum set of attributes returned. If additional attributes are required, they need to be added to the relevant script block first.
The script does not handle connectivity to Exchange Online, due to the variety of methods now available. If any existing session is detected, it will be used to run the script, otherwise an error will be thrown. If you need help connecting PowerShell to Exchange Online, follow the steps in this article: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx

The script includes a single cmdlet, Get-MailboxPermissionInventory, and can be dot-sourced to expose it for use in other scripts. Parameters are passed via splatting, here's an example usage:

```PowerShell
.\Mailbox_Permissions_inventory.ps1 -IncludeUserMailboxes
```

By default the script outputs the results to the console host and also stores them in the $varPermissions variable. To export the results to CSV file, use the output variable:

```PowerShell
Get-MailboxPermissionInventory -IncludeAll -OutVariable var   
$var | Export-Csv -NoTypeInformation "MailboxPermissions.csv"
```

Alternatively, remove the comment mark in the last line of the script. If the script fails too often due to connectivity issues, you can consider uncommenting line 95 to force the script to write to the CSV file after each iteration.
 
Additional information about the script can be found at: https://www.cogmotive.com/blog/powershell-scripts/office-365-permissions-inventory-full-access

Questions and feedback are welcome.
