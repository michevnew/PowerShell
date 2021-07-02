# Azure AD Admin Roles Inventory

The script lists all Azure AD admin roles assigned in your tenant. Only roles with active assignments are returned, along with details about the corresponding user or service principal object. The script requires the Azure AD PowerShell module. You can download the latest version here: https://www.powershellgallery.com/packages/AzureAD/  

If existing session to Azure AD is detected, the script will try to reuse it. Otherwise, you will be prompted for credentials.
 
The output will list the users and service principals with admin roles assigned sorted by their display name. Multiple role assignments for the same user are concatenated together in the "Roles" field. To export the output to a CSV file, remove the comment mark from the last line of the script.
 
Additional information about the script can be found at: https://www.michev.info/Blog/Post/3350/office-365-permissions-inventory-azure-ad-admin-roles
 
Questions and feedback are welcome.

Dont forget that admin roles are only one of the ways permissions can be granted in Azure AD/Office 365. If you want a comprehensive inventory, make sure to cover any workload-specific controls, as well as application permissions. With regards to this, the script goes hand to hand with the Azure AD Integrated Applications Inventory one: https://gallery.technet.microsoft.com/Azure-AD-Integrated-44658ec2?redir=0
