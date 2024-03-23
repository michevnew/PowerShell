# Azure AD Applications Registrations Inventory

The script lists all Azure AD application registrations (application objects) for your tenant. For each application, additional information such as the verification status is returned. OAuth permissions (both delegate and application) required by the application are included in the output, as is information about any API integrations.

The script calls the Graph API directly, and thus needs to obtain an access token. The method used is via the client credentials flow, meaning you need an app registration. Feel free to replace it with whatever method works best for you. The Directory.Read.All scope is hard requirement for enumerating oauth2PermissionGrants entries, so make sure to include that.

Output will be exported to a CSV file within the current directory. A single entry per application object is used, with permissions concatenated together in a single, delimited string. Format used is as follows:

```
[Microsoft Graph(user@domain.com)]:profile,openid,User.Read;[Microsoft Graph]:Directory.Read.All;[Office 365 Exchange Online]:full_access_as_app,Exchange.ManageAsApp
```
Additional information about the script can be found at: https://www.michev.info/blog/post/3665/azure-ad-application-registration-inventory-script

Questions and feedback are welcome.
