# Office 365 Group Owners Report

The script reports the Owners for all Office 365 Groups. Running the script without any parameter will fetch the data form Azure AD, which means that the following object types will be returned: Security group, Mail-enabled security group, Distribution group, Office 365 group, groups with dynamic membership. If you want to include Exchange-only objects such as Dynamic distribution groups or Room lists, use the -IncludeExchangeManagedBy parameter.
 
You will need a recent version of the Azure AD PowerShell module or the Azure AD Preview PowerShell module. If both modules are installed, the Azure AD Preview module will be used, as it exposes some additional details about Group objects. For download and install instructions refer to the PowerShell Gallery: https://www.powershellgallery.com/packages/AzureAD/
 
The script will detect and reuse any existing sessions to Azure AD and/or Exchange Online. Basic functionality to facilitate   a new connection is also included, but it does not cover all possible scenarios. If you are using an MFA-protected account or any of the non-MT Office 365 instances, make sure you connect manually before running the script. If you need help connecting PowerShell to Exchange Online, follow the steps in this article: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx
 
When fetching Exchange data, the script will use Invoke-Command to get a minimum set of attributes returned in order to speed up execution. If additional attributes are required, they need to be added to the relevant script block first (line 56). It's important to note that the ManagedBy attribute returned by Exchange is the display name, thus it cannot be used to uniquely identify the user.  
 
The script includes two cmdlets, Get-AzureADGroupOwnersInventory and Get-ExchangeObjectsOwnersInventory, and can be dot-sourced to expose them for use in other scripts. Parameters are passed via splatting, here's an example usage:
```PowerShell
.\Groups_Owner_Inventory.ps1 -IncludeExchangeManagedBy
```
By default the script outputs the results to the host window and also stores them in the $varOwners and $varOwnersExchange global variables respectively. If you want to export the output to CSV file directly, remove the comment marks from lines 73 and 79 or use the below example:
```PowerShell
$varOwners | Export-Csv -NoTypeInformation "GroupOwners.csv"
```
Additional information about the script can be found in [this article](https://www.cogmotive.com/blog/powershell/new-script-office-365-aliases-inventory). 

Questions and feedback are welcome.
