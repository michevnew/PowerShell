# Office 365 Mailbox Forwarding Inventory

The script lists all mailboxes of the specified type(s) that have at least one form of forwarding configured. Running the script without any parameter will return entries across all User, Shared, Room, Equipment, Discovery, and Team mailboxes. Both the ForwardingSmtpAddress and the ForwardingAddress attributes are checked.

In addition, you can specify the -CheckInboxRules parameter to cover scenarios where messages are being forwarded or redirected via Inbox rules. Similarly, the -CheckCalendarDelegates parameter will toggle the inclusion of information related to Calendar items forwarding. Lastly, a generic Transport rules information can be returned by specifying the -CheckTransportRules parameter.

As the full list of mailboxes needs to be cycled in order to get Inbox rules and Calendar delegates, the script will use Invoke-Command in order to get a minimum set of attributes returned if the respective parameter is invoked. If additional attributes are required, they need to be added to the relevant script block first.

The script does not handle connectivity to Exchange Online, due to the variety of methods now available. If any existing session is detected, it will be used to run the script, otherwise an error will be thrown. If you need help connecting PowerShell to Exchange Online, follow the steps in this article: [https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx](https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps)

The script includes a single cmdlet, Get-MailboxForwardingInventory, and can be dot-sourced to expose it for use in other scripts. Parameters are passed via splatting, here's an example usage:
```PowerShell
.\Mailbox_Forwarding_inventory.ps1 -CheckInboxRules
```
By default the script outputs the results to the console host and also stores them in the $varForwarding variable. To export the results to CSV file, use the output variable:
```PowerShell
Get-MailboxForwardingInventory -IncludeAll -OutVariable var    
$var | Export-Csv -NoTypeInformation "MailboxForwarding.csv"
```
Alternatively, remove the comment mark in the last line of the script. 
 
Additional information about the script can be found at: [https://www.michev.info/Blog/Post/4438/mailbox-forwarding-inventory-report](https://www.michev.info/Blog/Post/4438/mailbox-forwarding-inventory-report)

Questions and feedback are welcome.
