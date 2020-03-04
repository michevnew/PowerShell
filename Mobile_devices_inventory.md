# Office 365 Mobile device inventory and statistics

The script enumerates all mobile devices in the tenant and returns additional statistics as well as information about the device owner. A helper function is used to load mailbox details from a CSV file or generate the CSV file if such is not present. A hash table is then used to "match" the mobile devices with their corresponding mailbox.

The script does not handle connectivity to Exchange Online, due to the variety of methods now available. If any existing session is detected, it will be used to run the script, otherwise an error will be thrown. If you need help connecting PowerShell to Exchange Online, follow the steps in this article: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx

The script will use Invoke-Command in order to get a minimum set of attributes returned. If additional attributes are required, they need to be added to the relevant script block(s) first.

A very basic anti-throttling mechanism is included in the script. If you are running it against large environments and/or have modified it to include additional cmdlets, make sure to properly handle throttling and.or session reconnects.

The output of the script will be written to a CSV file in the script directory. Additional CSV file will be generated for mailbox data, if no such file was found when loading the script. Only User mailboxes are accounted for, if your organization assigns mobile devices to other mailbox types, make sure to include them as needed.

Additional information about the script can be found in [this article](https://www.quadrotech-it.com/blog/powershell-script-for-office-365-mobile-device-inventory-and-statistics/).

Questions and feedback are welcome.
