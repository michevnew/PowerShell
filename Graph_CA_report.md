This is a small proof-of-concept script that lists Conditional access policies within your Azure AD/Office 365 tenant, along with their details. This is done by querying the Graph API's /identity/conditionalAccess/policies endpoint.

In order to use the script, you will need to first configure some variables, found on top. This is needed in order to obtain an access token for the Graph. The script uses the so-called "client credentials" flow, thus you need to provide the AppID of an Azure AD application you've registered with the tenant. The application needs the following permissions for the script to run as expected:

\# Policy.Read.All to enumerate all Conditional access policies the tenant

\# Directory.Read.All to convert GUIDs to UPNs or display names as needed

After creating the application and granting the permissions, copy the key/secret and use it to configure the $client_secred variable. If you need more help understanding all the concepts mentioned above, start with [this article](https://docs.microsoft.com/en-us/graph/auth/auth-concepts).

If a token is successfuly obtained, the script will query the /beta/identity/conditionalAccess/policies Graph endpoint, fetch the results and transform it to present the details for each Conditional access policies in the console window.

Two small helper functions are added to convert GUIDs to UPN (for user objects) or display name (for Groups and DirectoryRoles). No conversion is done for applicaitons or named locations, add the necesasry functions if needed.

Output is also written to a CSV file in the script directory.

More info about the script can be found here: https://www.michev.info/blog/post/3004/reporting-on-conditional-access-policies-in-microsoft-365
