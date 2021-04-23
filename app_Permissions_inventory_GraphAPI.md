# Azure AD Integrated Applications Inventory

The script lists all Azure AD integrated applications (service principals) for your tenant. For each application, additional information such as the Publisher and homepage is returned. OAuth permissions (both delegate and application) granted to the application are included in the output, as is information about any users that have authorized the application.

The script calls the Graph API directly, and thus needs to obtain an access token. The method used is via the client credentials flow, meaning you need an app registration. Feel free to replace it with whatever method works best for you. The Directory.Read.All scope is hard requirement for enumerating oauth2PermissionGrants entriesm, so make sure to include that.

Output will be exported to a CSV file within the current directory. A single entry per service principal is used, with permissions concatenated together in a single, delimited string. Format used is as follows:

```
[Microsoft Graph(user@domain.com)]:profile,openid,User.Read;[Microsoft Graph]:Directory.Read.All;[Office 365 Exchange Online]:full_access_as_app,Exchange.ManageAsApp
```
Additional information about the script can be found at: 

Questions and feedback are welcome.
