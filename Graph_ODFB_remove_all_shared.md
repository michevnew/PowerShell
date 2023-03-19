# Remove sharing for user's OneDrive for Business items

The script uses the Graph API to enumerate all shared files within a user's OneDrive for Business drive. For each such item, permissions will be removed where possible. 

if you want to include items in any of the (sub)folders, use the corresponding script parameters as detailed below.

In order to use the script, you will need to first configure some variables, found between lines 141-143. Provide the tenantID and the AppID of an Azure AD application you've registered with the tenant. The application needs the following permissions for the script to run as expected:

\#    User.Read.All to enumerate all users in the tenant

\#    Sites.ReadWrite.All to return all the item sharing details and remove permissions

After creating the application and granting the permissions, copy the key/secret and use it to configure the $client_secred variable. If you need more help understanding all the concepts mentioned above, start with [this article](https://docs.microsoft.com/en-us/graph/auth/auth-concepts).

To run the script against a given user, use the following syntax:

```PowerShell
.\Graph_ODFB_remove_all_shared.ps1 -Verbose -User vasil@michev.info
```
The script has two optional parameters you can use. The -ExpandFolders switch instructs it to enumerate files in any (sub)folders found under the root, and the -Depth parameter controls how deep the expansion is. The default value is $true for ExpandFolders and 2 for Depth. Use this parameter with care, while I've tested the script with few thousand items in multiple nested folders, I cannot guarantee it will work in all scenarios.

```PowerShell
.\Graph_ODFB_remove_all_shared.ps1 -Verbose -User vasil@michev.info -ExpandFolders -depth 2
```

By default, the script will return a filtered list of just the items that have been shared, and will also store the output in a global variable called $varODFBSharedItems in case you want to reuse it. if you want to save it to CSV, uncomment line 192

Additional information about the script can be found at: https://www.michev.info/blog/post/3018/remove-sharing-permissions-on-all-files-in-users-onedrive-for-business

Questions and feedback are welcome.
