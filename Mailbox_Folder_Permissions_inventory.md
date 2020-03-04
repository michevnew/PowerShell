# Office 365 Mailbox Folder Permissions Inventory

The script enumerates the default and user-created folders for all mailboxes of the specified type(s) and lists their permissions. Running the script without any parameter will return permissions for User mailboxes only. Use the switches to include Shared, Room or Equipment mailboxes.

As the full list of mailboxes and their folders needs to be cycled in order to get the permissions, the script will use Invoke-Command in order to get a minimum set of attributes returned. If additional attributes are required, they need to be added to the relevant script block first.

The script does not handle connectivity to Exchange Online, due to the variety of methods now available. If any existing session is detected, it will be used to run the script, otherwise an error will be thrown. If you need help connecting PowerShell to Exchange Online, follow the steps in this article: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx

The script includes a single cmdlet, Get-MailboxFolderPermissionInventory, and can be dot-sourced to expose it for use in other scripts. Parameters are passed via splatting, here's an example usage: 
```PowerShell
.\Mailbox_Folder_Permissions_inventory.ps1 -IncludeUserMailboxes
```
By default the script outputs the results to a CSV file and also stores them in the $varPermissions variable. If you are dot-sourcing the script and invoking the Get-MailboxPermissionInventory cmdlet instead, output is written only to the console host. Use the -OutVariable parameter to store the output and export to CSV file:  
```PowerShell
Get-MailboxFolderPermissionInventory -IncludeAll -OutVariable var 
$var | Export-Csv -NoTypeInformation "MailboxFolderPermissions.csv"
```
Default permission entries are not included in the report when invoking the script without parameters. Specify the -IncludeDefaultPermissions parameter to add them to the output:
```PowerShell
.\Mailbox_Folder_Permissions_inventory.ps1 -IncludeDefaultPermissions
```
If you want to use "condensed" output, limited to one line per mailbox folder, specify the -CondensedOutput switch. By default, "expanded" output is used, with one line per each permission entry, including the default permissions.

If the script fails too often due to connectivity issues, you can consider uncommenting lines 145 and 170 to force the script to write to the CSV file after each iteration. Removing the comment mark from line 101 will add small delay between interations in order to avoid throttling.

To further reduce the time to execute the script, consider limiting the list of folders to only those you are interested in. This can be achieved by editing the $includedfolders and $excludedfolders arrays:
```PowerShell
    $includedfolders = @("Root","Inbox","Calendar", "Contacts", "DeletedItems", "Drafts", "JunkEmail", "Journal", "Notes", "Outbox", "SentItems", "Tasks", "CommunicatorHistory", "Clutter", "Archive") 
 
    $excludedfolders = @("News Feed","Quick Step Settings","Social Activity Notifications","Suggested Contacts", "SearchDiscoveryHoldsUnindexedItemFolder", "SearchDiscoveryHoldsFolder")
```
To exclude permissions entries for specific user(s), specify the email address(es) via the $ExcludeUsers parameter:
```PowerShell
.\Mailbox_Folder_Permissions_inventory.ps1 -ExcludeUsers admin@domain.com,serviceaccount@domain.com
```
Additional information about the script can be found in [this article](https://www.cogmotive.com/blog/powershell/mailbox-folder-permissions). 

Questions and feedback are welcome.
