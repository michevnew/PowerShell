#Requires -Version 3.0
#Make sure to fill in all the required variables before running the script
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    User.Read.All (for fetching the list of users)
#    Calendars.Read (for fetching the calendar permissions)

param([switch]$IncludeNonDefaultCalendars=$false) #Use the -IncludeNonDefaultCalendars switch to include non-default calendars in the report. This will make the script slower, as it will need to make additional calls for each mailbox.

#For details on what the script does and how to run it, check: https://michev.info/blog/post/8420/report-calendar-permissions-via-the-graph-api

#Get an Access token. Make sure to fill in all the variable values here. Or replace with your own preferred method to obtain token.
$tenantId = "tenant.onmicrosoft.com"
$uri = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'
$clientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$client_secret = "verylongstring"

#Get an Access token for the Graph API
$Scopes = New-Object System.Collections.Generic.List[string]
$Scope = "https://graph.microsoft.com/.default"
$Scopes.Add($Scope)

$body = @{
    grant_type = "client_credentials"
    client_id = $clientId
    client_secret = $client_secret
    scope = $Scopes
}

try {
    $res = Invoke-WebRequest -Method Post -Uri $uri -Verbose:$false -Body $body -UseBasicParsing
    $token = ($res.Content | ConvertFrom-Json).access_token

    $authHeader = @{
       'Authorization'="Bearer $token"
    }}
catch { Write-Host "Failed to obtain token, aborting..." ; return }

#Prepare the list of mailboxes
Write-Verbose "Fetching the list of users..."
$Mailboxes = @()
try {
    $authHeader["consistencyLevel"] = "eventual"
    $uri = "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'member' and mail ne null&`$select=id,Mail,MailNickname,ProxyAddresses&`$top=999&`$count=true"
    do {
        $result = Invoke-WebRequest -Uri $uri -Method GET -Headers $authHeader -UseBasicParsing -ErrorAction Stop -Verbose:$false
        $Mailboxes += ($result.Content | ConvertFrom-Json).value
        $uri = ($result.Content | ConvertFrom-Json).'@odata.nextLink'
    } while ($uri)
    $authHeader.Remove("consistencyLevel")
}
catch { Write-Warning "Failed to get user list, aborting..." ; return }

if (!$Mailboxes -or ($Mailboxes.Count -eq 0)) { Throw "No matching users found, aborting..." }
Write-verbose "Found $($Mailboxes.Count) matching users, processing..."

$output = [System.Collections.Generic.List[Object]]::new() #output variable
#Loop over all mailboxes and get the Calendar permissions
foreach ($mb in $Mailboxes) {
    #Fetch the calendar permissions for each mailbox.
    Write-Verbose "Fetching Calendar permissions for $($mb.mail)..."
    Write-Progress -Activity "Fetching Calendar permissions" -Status "User $($Mailboxes.IndexOf($mb)) of $($Mailboxes.Count): $($mb.mail)" -PercentComplete (($Mailboxes.IndexOf($mb) / $Mailboxes.Count) * 100)

    #If the IncludeNonDefaultCalendars switch is set, fetch the full set of calendars for the mailbox and get their permissions as well.
    if ($IncludeNonDefaultCalendars) {
        try {
            $calendarsResult = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/users/$($mb.id)/calendars?`$top=99" -Method GET -Headers $authHeader -UseBasicParsing -ErrorAction Stop -Verbose:$false
            $calendars = ($calendarsResult.Content | ConvertFrom-Json).value

            if (!$calendars -or ($calendars.Count -eq 0)) { Write-Verbose "No calendars found for $($mb.mail), skipping..." ; continue }

            #Cycle over each calendar and fetch its permissions
            foreach ($calendar in $calendars) {
                Write-Verbose "Fetching Calendar permissions for calendar '$($calendar.name)' ($($calendar.id)) of $($mb.mail)..."
                try {
                    $result = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/users/$($mb.id)/calendars/$($calendar.id)/calendarPermissions" -Method GET -Headers $authHeader -UseBasicParsing -ErrorAction Stop -Verbose:$false

                    #If the request was successful and returned content, parse the JSON and add the permissions to the output list
                    if ($result -and $result.Content) {
                        $permissions = ($result.Content | ConvertFrom-Json).value
                        foreach ($perm in $permissions) {
                            $obj = [PSCustomObject]@{
                                Id = $mb.id
                                Mailbox = $mb.mail
                                CalendarName = $calendar.name
                                CalendarId = $calendar.id
                                User = $perm.emailAddress.name
                                Address = $perm.emailAddress.address
                                Role = $perm.role
                                AllowedRoles = $perm.allowedRoles -join ";"
                                IsInsideOrganization = $perm.isInsideOrganization
                                IsRemovable = $perm.isRemovable
                            }
                            $output.Add($obj)
                        }
                    }

                    Start-Sleep -Milliseconds 100 #add some delay to avoid throttling
                }
                #If the mailbox does not exist, or is not accessible, skip it and continue to the next one.
                catch { Write-Warning "Failed to fetch Calendar permissions for calendar '$($calendar.name)' ($($calendar.id)) of $($mb.mail), skipping..." ; continue }
            }
        }
        #If the mailbox does not exist, or is not accessible, skip it and continue to the next one.
        catch { Write-Warning "Failed to fetch non-default calendars for $($mb.mail), skipping..." ; continue }
    }

    else {
        try {
            $result = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/users/$($mb.id)/calendar/calendarPermissions" -Method GET -Headers $authHeader -UseBasicParsing -ErrorAction Stop -Verbose:$false

            #If the request was successful and returned content, parse the JSON and add the permissions to the output list
            if ($result -and $result.Content) {
                $permissions = ($result.Content | ConvertFrom-Json).value
                foreach ($perm in $permissions) {
                    $obj = [PSCustomObject]@{
                        Id = $mb.id
                        Mailbox = $mb.mail
                        User = $perm.emailAddress.name
                        Address = $perm.emailAddress.address
                        Role = $perm.role
                        AllowedRoles = $perm.allowedRoles -join ";"
                        IsInsideOrganization = $perm.isInsideOrganization
                        IsRemovable = $perm.isRemovable
                    }
                    $output.Add($obj)
                }
            }

            Start-Sleep -Milliseconds 100 #add some delay to avoid throttling
        }
        #If the mailbox does not exist, or is not accessible, skip it and continue to the next one.
        catch { Write-Warning "Failed to fetch Calendar permissions for $($mb.mail), skipping..." ; continue }
    }
}

#Export the output to CSV
$output | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_CalendarPermissions.csv" -NoTypeInformation -Encoding UTF8 -UseCulture

#Export the results to an HTML file for easier viewing
if ($IncludeNonDefaultCalendars) { $columns = @('Id','Mailbox','CalendarName','CalendarId','User','Address','Role','AllowedRoles','IsInsideOrganization','IsRemovable') }
else { $columns = @('Id','Mailbox','User','Address','Role','AllowedRoles','IsInsideOrganization','IsRemovable') }

$rowsBuilder = New-Object System.Text.StringBuilder
foreach ($item in $output) {
    [void]$rowsBuilder.AppendLine('<tr>')
    foreach ($col in $columns) {
        $value = [System.Net.WebUtility]::HtmlEncode([string]$item.$col)
        [void]$rowsBuilder.AppendLine("<td>$value</td>")
    }
    [void]$rowsBuilder.AppendLine('</tr>')
}

$headerHtml = ($columns | ForEach-Object { "<th>$_</th>" }) -join "`n"

$html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>Calendar Permissions</title>
    <style>
        body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
        h2 { margin-top: 0; }
        .table-wrap { overflow: auto; border: 1px solid #ddd; }
        table { border-collapse: collapse; width: 100%; table-layout: fixed; }
        th, td { border: 1px solid #ddd; padding: 6px 8px; text-align: left; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        th { background: #f3f3f3; position: relative; user-select: none; cursor: pointer; }
        th .resize-handle { position: absolute; top: 0; right: 0; width: 6px; height: 100%; cursor: col-resize; }
        th.sort-asc::after { content: ' ▲'; }
        th.sort-desc::after { content: ' ▼'; }
    </style>
</head>
<body>
    <h2>Calendar Permissions</h2>
    <div class="table-wrap">
        <table id="permissionsTable">
            <thead>
                <tr>
                    $headerHtml
                </tr>
            </thead>
            <tbody>
                $($rowsBuilder.ToString())
            </tbody>
        </table>
    </div>

    <script>
        (function () {
            const table = document.getElementById('permissionsTable');
            const headers = table.querySelectorAll('th');

            headers.forEach((th, index) => {
                th.addEventListener('click', function (e) {
                    if (e.target.classList.contains('resize-handle')) return;

                    const tbody = table.tBodies[0];
                    const rows = Array.from(tbody.querySelectorAll('tr'));
                    const isAsc = !th.classList.contains('sort-asc');

                    headers.forEach(h => h.classList.remove('sort-asc', 'sort-desc'));
                    th.classList.add(isAsc ? 'sort-asc' : 'sort-desc');

                    rows.sort((a, b) => {
                        const aText = (a.children[index].innerText || '').trim();
                        const bText = (b.children[index].innerText || '').trim();
                        const aNum = Number(aText);
                        const bNum = Number(bText);

                        let cmp;
                        if (!Number.isNaN(aNum) && !Number.isNaN(bNum)) {
                            cmp = aNum - bNum;
                        } else {
                            cmp = aText.localeCompare(bText, undefined, { numeric: true, sensitivity: 'base' });
                        }
                        return isAsc ? cmp : -cmp;
                    });

                    rows.forEach(r => tbody.appendChild(r));
                });

                const handle = document.createElement('div');
                handle.className = 'resize-handle';
                th.appendChild(handle);

                let startX = 0;
                let startWidth = 0;

                const onMouseMove = (e) => {
                    const newWidth = Math.max(60, startWidth + (e.pageX - startX));
                    th.style.width = newWidth + 'px';
                };

                const onMouseUp = () => {
                    document.removeEventListener('mousemove', onMouseMove);
                    document.removeEventListener('mouseup', onMouseUp);
                };

                handle.addEventListener('mousedown', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    startX = e.pageX;
                    startWidth = th.offsetWidth;
                    document.addEventListener('mousemove', onMouseMove);
                    document.addEventListener('mouseup', onMouseUp);
                });
            });
        })();
    </script>
</body>
</html>
"@

$html | Out-File -FilePath "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_CalendarPermissions.html" -Encoding UTF8