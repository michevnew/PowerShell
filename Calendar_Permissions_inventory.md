# Office 365 Calendar Permissions Inventory

UPDATED APR 2020: Microsoft changed the output of the Get-MailboxFolderPermission in the service, it no longer returns the RecucedRecipient object but instead a PermissionSecurityPrincipal object. I've adjusted the script accodrdingly, but this mean it will no longer run as expected against on-opremises installs. You can change it manually: 

$entry.User.RecipientPrincipal.PrimarySmtpAddress.ToString() <-> $entry.User.ADRecipient.PrimarySmtpAddress.ToString()

///end update

The script lists find the default Calendar folder for all mailboxes of the specified type(s) and lists its permissions. Running the script without any parameter will return permissions for User mailboxes only. Use the switches to include Shared, Room or Equipment mailboxes.

As the full list of mailboxes needs to be cycled in order to get the permissions, the script will use Invoke-Command in order to get a minimum set of attributes returned. If additional attributes are required, they need to be added to the relevant script block first.
The script does not handle connectivity to Exchange Online, due to the variety of methods now available. If any existing session is detected, it will be used to run the script, otherwise an error will be thrown. If you need help connecting PowerShell to Exchange Online, follow the steps in this article: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx

The script includes a single cmdlet, Get-CalendarPermissionInventory, and can be dot-sourced to expose it for use in other scripts. Parameters are passed via splatting, here's an example usage:
```PowerShell
.\Calendar_Permissions_inventory.ps1 -IncludeUserMailboxes
```
By default the script outputs the results to a CSV file and also stores them in the $varPermissions variable. If you are dot-sourcing the script and invoking the Get-CalendarPermissionInventory cmdlet instead, output is written only to the console host. Use the -OutVariable parameter to store the output and export to CSV file:
```PowerShell
Get-CalendarPermissionInventory -IncludeAll -OutVariable var    
$var | Export-Csv -NoTypeInformation "CalendarPermissions.csv"
```
If you want to use "condensed" output, limited to one line per mailbox, specify the -CondensedOutput switch. By default, "expanded" output is used, with one line per each permission entry, including the default permissions.
 
If the script fails too often due to connectivity issues, you can consider uncommenting lines 115 and 140 to force the script to write to the CSV file after each iteration. Removing the comment mark from line 71 will add small delay between interations in order to avoid throttling.

Additional information about the script can be found at: https://www.michev.info/blog/post/3676/office-365-permission-inventory-calendar-permissions-2

Questions and feedback are welcome.
