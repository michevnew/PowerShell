#Requires -Version 3.0
# Make sure to fill in all the required variables before running the script
# Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Group.Read.All or Directory.Read.All to read all Groups
#    Group.Read.All to read Channel info

#Variables to configure
$ADALpath = 'C:\Program Files\WindowsPowerShell\Modules\AzureAD\2.0.2.16\Microsoft.IdentityModel.Clients.ActiveDirectory.dll' #path to Microsoft.IdentityModel.Clients.ActiveDirectory.dll
$tenantID = "tenant.onmicrosoft.com" #your tenantID or tenant root domain
$appID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #the GUID of your app.
$client_secret = "verylongsecurestring" #client secret for the app

#==========================================================================
#Main script starts here
#==========================================================================

#Needs the ADAL binaries to obtain token
try { Add-Type -Path $ADALpath -ErrorAction Stop }
catch { Write-Error "Unable to load ADAL binaries, make sure you are using the correct path!" -ErrorAction Stop }

#Obtain access token
$authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList "https://login.windows.net/$tenantID"
$ccred = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential -ArgumentList $appID,$client_secret

$authenticationResult = $authContext.AcquireTokenAsync("https://graph.microsoft.com", $ccred)
if (!$authenticationResult.Result.AccessToken) { Write-Error "Failed to aquire token!"; return }

#Use the access token to set the authentication header
$authHeader = @{'Authorization'=$authenticationResult.Result.CreateAuthorizationHeader()}

#Use the /beta endpoint to fetch a list of all Teams
$uri = "https://graph.microsoft.com/beta/groups?`$filter=resourceProvisioningOptions/Any(x:x eq `'Team`')&`$select=id,displayName,mail,proxyAddresses,resourceBehaviorOptions,resourceProvisioningOptions,visibility"
$result = Invoke-WebRequest -Headers $AuthHeader -Uri $uri -Verbose:$VerbosePreference
$teams = ($result.Content | ConvertFrom-Json).Value

#iterate over each Team and gather channel information
$output = @()
foreach ($team in $teams) {
    $uri = "https://graph.microsoft.com/v1.0/teams/$($team.id)/channels"
    $result = Invoke-WebRequest -Headers $AuthHeader -Uri $uri -Verbose:$VerbosePreference
    $channels = ($result.Content | ConvertFrom-Json).Value

    #$channels | select displayName,email,id,webUrl
    foreach ($channel in $channels) {
        $chaninfo = New-Object psobject
        $chaninfo | Add-Member -MemberType NoteProperty -Name "Id" -Value $channel.id
        $chaninfo | Add-Member -MemberType NoteProperty -Name "Team" -Value $team.displayName
        $chaninfo | Add-Member -MemberType NoteProperty -Name "TeamId" -Value $team.id
        $chaninfo | Add-Member -MemberType NoteProperty -Name "TeamEmailAddresses" -Value $($team.proxyAddresses -join ",")
        $chaninfo | Add-Member -MemberType NoteProperty -Name "Visibility" -Value $team.visibility
        $chaninfo | Add-Member -MemberType NoteProperty -Name "Channel" -Value $channel.displayName
        $chaninfo | Add-Member -MemberType NoteProperty -Name "ChannelEmail" -Value (&{If($channel.email) {$channel.email} Else {"N/A"}})
        $output += $chaninfo
    }

}

#return the output
$global:varTeamChannels = $output | select Team,Visibility,Channel,ChannelEmail,TeamEmailAddresses,TeamId
$output | select * | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_TeamsChannels.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
return $global:varTeamChannels