# Office 365 aliases inventory

The script enumerates email and non-email addresses (aliases) for all recipients of the specified type(s). Running the script without any parameter will return aliases for User mailboxes only. Use the switches to include other recipient types, or use the -IncludeAll to return all aliases across all supported recipient types. See lines 33-48 for the relevant parameters and their description.

SMTP and X500 aliases are included by default. If you also want to include SIP aliases, use the -IncludeSIPAliases parameter. Similarly, to include the SPO aliases, use the -IncludeSPOAliases   parameter. For any mail user/mail contact objects, the ExternalEmailAddress atribute is also included.

The script will use Invoke-Command to get a minimum set of attributes returned in order to speed up execution. If additional attributes are required, they need to be added to the relevant script block first (line 77). Since the UserPrincipalName attribute is not exposed via Get-Recipient, the script uses the WindowsLiveID attribute as a workaround.

The script does not handle connectivity to Exchange Online, due to the variety of methods now available. If any existing session is detected, it will be used to run the script, otherwise an error will be thrown. If you need help connecting PowerShell to Exchange Online, follow the steps in this article: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx  

The script includes a single cmdlet, Get-EmailAddressesInventory, and can be dot-sourced to expose it for use in other scripts. Parameters are passed via splatting, here's an example usage:
```PowerShell
.\O365_aliases_inventory.ps1 -IncludeAll
```
By default the script outputs the results to a CSV file and also stores them in the $varAliases variable. If you are dot-sourcing the script and invoking the Get-EmailAddressesInventory   cmdlet instead, output is written only to the console host. Use the -OutVariable parameter to store the output and export to CSV file:
```PowerShell
Get-EmailAddressesInventory -IncludeAll -OutVariable var  
$var | Export-Csv -NoTypeInformation "EmailAddresses.csv"
```
A "one line per alias" formatting is used by default, you can switch to "condensed" output (one line per recipient) instead by specifying the -CondensedOutput switch.

Additional information about the script can be found in [this article](https://www.michev.info/Blog/Post/3966/office-365-aliases-inventory).

Questions and feedback are welcome.
