# Teams and channel email addresses inventory

The script uses the Graph API to enumerate all Teams in the tenant, then enumerates all channels in each Team. For every channel, the information about any email addresses configured is gathered. In addition, information about any email addresses configured on the Team itself is returned.

In order to use the script, you will need to first configure some variables, found between lines 8-11. Provide the tenantID and the AppID of an Azure AD application you've registered with the tenant. The application needs the following permissions for the script to run as expected:

\#    Group.Read.All or Directory.Read.All to read all Groups
\#    Group.Read.All to read Channel info

After creating the application and granting the permissions, copy the key/secret and use it to configure the $client_secred variable. If you need more help understanding all the concepts mentioned above, start with [this article](https://docs.microsoft.com/en-us/graph/auth/auth-concepts).

By default, the script will store the output in a global variable called $varTeamChannels in case you want to reuse it, and will return the output to the console. The unfiltered output will be saved to a CSV file, which you can then format, sort and filter as needed.

Additional information about the script can be found in [this article](https://www.michev.info/blog/post/2676/reporting-on-any-email-addresses-configured-for-teams-and-channels-via-the-graph-api).

Questions and feedback are welcome.
