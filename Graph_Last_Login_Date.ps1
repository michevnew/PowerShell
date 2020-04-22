#Set up
$AppId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #clientID of your AAD app, must have User.Read.All, Directory.Read.All, Auditlogs.Read.All permissions
$client_secret = Get-Content .\ReportingAPIsecret.txt | ConvertTo-SecureString
$app_cred = New-Object System.Management.Automation.PsCredential($AppId, $client_secret)
$TenantId = "tenant.onmicrosoft.com" #your tenant

$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $app_cred.GetNetworkCredential().Password
    grant_type    = "client_credentials"
}
 
#simple code to get an access token, add your own handlers as needed
try { $tokenRequest = Invoke-WebRequest -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -ErrorAction Stop }
catch { Write-Host "Unable to obtain access token, aborting..."; return }

$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

#prepare auth header
$authHeader1 = @{
   'Content-Type'='application\json'
   'Authorization'="Bearer $token"
}

#exectue the actual query
$LastLogin = Invoke-WebRequest -Headers $AuthHeader1 -Uri "https://graph.microsoft.com/beta/users?`$select=displayName,userPrincipalName,signInActivity"
$result = ($LastLogin.Content | ConvertFrom-Json).Value
$result  | select DisplayName,UserPrincipalName,@{n="LastLoginDate";e={$_.signInActivity.lastSignInDateTime}}

#$result | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_LastLoginDate.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
