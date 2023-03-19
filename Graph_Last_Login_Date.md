This is a small proof-of-concept script that lists users within your Azure AD/Office 365 tenant, along with their Last login date. This is done by querying the Graph API's /users endpoint and including the signInActivity property.

In order to use the script, you will need to first configure some variables, found on top. This is needed in order to obtain an access token for the Graph. The script uses the so-called "client credentials" flow, thus you need to provide the AppID of an Azure AD application you've registered with the tenant. The application needs the following permissions for the script to run as expected:

\# User.Read.All to enumerate all users in the tenant

\# Auditlogs.Read.All to return all the Last login date

\# (optional) Directory.Read.All

After creating the application and granting the permissions, copy the key/secret and use it to configure the $client_secred variable. If you need more help understanding all the concepts mentioned above, start with [this article](https://docs.microsoft.com/en-us/graph/auth/auth-concepts).

If a token is successfuly obtained, the script will query the /beta/users Graph endpoint, fetch the results and transform it to present the Last login date (where available) for each user.

If you want to export the result to CSV file, uncomment the last line.

More info can be found here: https://www.michev.info/blog/post/2968/reporting-on-users-last-logged-in-date-in-office-365
