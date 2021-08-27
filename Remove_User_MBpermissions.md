# Remove user Full Access permissions from all mailboxes

Use this script to remove Full Access permissions for a given user, or a list of users, from all mailboxes within the organization.

By default, the script only covers User mailboxes. To include Shared mailboxes, specify the -IncludeSharedMailboxes parameter. To also include room and equipment mailboxes, specify the -IncludeResourceMailboxes parameter.

The script will detect and reuse any existing sessions to Exchange Remote PowerShell. Basic functionality to facilitate a new connection to Exchange Online is also included, but it does not cover all possible scenarios. If you are using an MFA-protected account or any of the non-MT Office 365 instances, make sure you connect manually before running the script. If you need help connecting PowerShell to Exchange Online, follow the steps in this article: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx

When fetching Exchange data, the script will use Invoke-Command to get a minimum set of attributes returned in order to speed up execution. To avoid throttling, some artificial delay is added upon executiong the removal process for each user. If needed, adjust lines 98 and 128 accordingly.

The removal action is performed via the Remove-MailboxPermission cmdlet. The script can be dot-sourced to expose it for use in other scripts. Additional function is used to handle connectivity checks. Parameters are passed via splatting, here's an example usage:

```
.\Remove_user_MBpermissions.ps1 -Identity vasil
```

The following example will remove permissions for two users, userA and userB from all types of mailboxes:

```
.\Remove_user_MBpermissions.ps1 -Identity userA,userB -IncludeSharedMailboxes -IncludeResourceMailboxes
```

You can also pass values for the Identity parameter over the pipeline. To avoid potential errors, it's recommended to first run the script with the -WhatIf parameter in such scenarios:

```
Get-User -Filter {CountryOrRegion -eq "US"} | Remove-UserMBPermissions -IncludeSharedMailboxes -WhatIf -Verbose
```

Additional information about the script can be found in this article: https://www.michev.info/Blog/Post/3418/

Questions and feedback are welcome.
