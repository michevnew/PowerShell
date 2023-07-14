# Azure AD report on user OU

The script prepares a report of all synchronized users in your Azure AD/Office 365 tenant. For each user, the information about the on-premises OU to which the user is added is exposed. This information is only synced to Azure AD if you are using AAD Connect version 1.1.153.0 or later.

To get the list of synchronized users, the script invokes the Get-AzureADUser cmdlet with a filter "DirSyncEnabled eq true". If you want to include only a subset of the users, edit line 5 of the script.

By default, output will be written to a CSV file in the current directory. If you want to use the output directly in the PowerShell console or modify it before exporting, edit the last line of the script.

Additional information about the script can be found at: [https://www.quadrotech-it.com/blog/reporting-organizational-unit-information-azure-ad-powershell/](https://www.michev.info/blog/post/3344/reporting-on-ou-information-via-azure-ad-powershell)https://www.michev.info/blog/post/3344/reporting-on-ou-information-via-azure-ad-powershell

Questions and feedback are welcome.
