# Azure AD Integrated Applications Inventory

The script lists all Azure AD integrated applications for your tenant. For each application, additional information such as the Publisher and homepage is returned. OAuth permissions required by the application are included in the output, as is information about any users that have authorized the application.

The script requires version 2.0.0.55 of the Azure AD PowerShell module or newer. You can download the latest version here: https://www.powershellgallery.com/packages/AzureAD/

If existing session to Azure AD is detected, the script will try to reuse it. Otherwise, you will be prompted for credentials.
The output closely resembles the report found in products such as Advanced Security Management or Cloud App Security. To export the output to a CSV file, remove the comment mark from the last line of the script.  

Additional information about the script can be found at: https://www.cogmotive.com/blog/cogmotive-reports-news/office-365-permissions-inventory-azure-ad-integrated-applications
https://practical365.com/inventorying-azure-ad-apps-and-their-permissions/

Questions and feedback are welcome.
