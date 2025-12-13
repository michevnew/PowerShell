#Variables to configure
$tenantID = "tenant.onmicrosoft.com" #your tenantID or tenant root domain
$appID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #the GUID of your app. For best result, use app with TeamsAppInstallation.ReadForTeam.All and TeamsTab.Read.All scopes granted.
$client_secret = "verylongsecurestring" #client secret for the app

#Prepare token request
$url = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'

$body = @{
    grant_type = "client_credentials"
    client_id = $appID
    client_secret = $client_secret
    scope = "https://graph.microsoft.com/.default"
}

#Obtain the token
try { $tokenRequest = Invoke-WebRequest -Method Post -Uri $url -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -ErrorAction Stop }
catch { Write-Host "Unable to obtain access token, aborting..."; return }

$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

$authHeader1 = @{
   'Content-Type'='application\json'
   'Authorization'="Bearer $token"
}

#Get a list of all Teams
$Teams = @()
$uri = "https://graph.microsoft.com/beta/teams"
do {
    $result = Invoke-WebRequest -Headers $AuthHeader1 -Uri $uri -UseBasicParsing -ErrorAction Stop
    $uri = $result.'@odata.nextLink'
    #If we are getting multiple pages, best add some delay to avoid throttling
    Start-Sleep -Milliseconds 500
    $Teams += ($result.Content | ConvertFrom-Json).Value
} while ($uri)
if (!$Teams -or $Teams.Count -eq 0) { Write-Host "Unable to obtain the list of teams, exiting..."; return }

#Iterate over each Team and prepare the report
$ReportApps = @();$ReportTabs = @();$count = 1
foreach ($team in $Teams) {

    #Progress message
    $ActivityMessage = "Retrieving data for team $($team.displayName). Please wait..."
    $StatusMessage = ("Processing team {0} of {1}: {2}" -f $count, @($Teams).count, $team.id)
    $PercentComplete = ($count / @($Teams).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Simple anti-throttling control
    Start-Sleep -Milliseconds 500

    #get a list of apps for the Team
    $teamApps = Invoke-WebRequest -Headers $authHeader1 -Uri "https://graph.microsoft.com/beta/teams/$($Team.id)/installedApps?`$expand=teamsApp,teamsAppDefinition" -UseBasicParsing -ErrorAction Stop
    $teamApps = ($teamApps.content | ConvertFrom-Json).Value

    $i = 0
    foreach ($app in $teamApps) {
        $i++
        $objApp = New-Object PSObject
        $objApp | Add-Member -MemberType NoteProperty -Name "Number" -Value $i
        $objApp | Add-Member -MemberType NoteProperty -Name "Team" -Value $team.displayName
        $objApp | Add-Member -MemberType NoteProperty -Name "Name" -Value $app.teamsApp.displayName
        $objApp | Add-Member -MemberType NoteProperty -Name "Version" -Value $app.teamsAppDefinition.version
        $objApp | Add-Member -MemberType NoteProperty -Name "AppId" -Value $app.teamsApp.Id
        $objApp | Add-Member -MemberType NoteProperty -Name "AddedVia" -Value $app.teamsApp.distributionMethod
        $objApp | Add-Member -MemberType NoteProperty -Name "Description" -Value $app.teamsAppDefinition.description
        $objApp | Add-Member -MemberType NoteProperty -Name "AADAppID" -Value $app.teamsAppDefinition.azureADAppId
        $objApp | Add-Member -MemberType NoteProperty -Name "AvailableFor" -Value $app.teamsAppDefinition.allowedInstallationScopes

        $ReportApps += $objApp
    }

    #Get a list of channels so we can also cover Tabs
    $TeamChannels = Invoke-WebRequest -Headers $AuthHeader1 -Uri "https://graph.microsoft.com/beta/Teams/$($Team.id)/channels" -UseBasicParsing -ErrorAction Stop
    $TeamChannels = ($TeamChannels.Content | ConvertFrom-Json).value

    #Iterate over each channel, enumerate Tabs
    foreach ($channel in $TeamChannels) {
        $tabs = Invoke-WebRequest -Headers $authHeader1 -Uri "https://graph.microsoft.com/beta/teams/$($Team.id)/channels/$($channel.id)/tabs?`$expand=teamsApp" -UseBasicParsing -ErrorAction Stop
        $tabs = ($tabs.Content | ConvertFrom-Json).value

        $j = 0
        foreach ($tab in $tabs) {
            $j++
            $objTab = New-Object PSObject
            $objTab | Add-Member -MemberType NoteProperty -Name "Number" -Value $j
            $objTab | Add-Member -MemberType NoteProperty -Name "Team" -Value $team.displayName
            $objTab | Add-Member -MemberType NoteProperty -Name "Channel" -Value $channel.displayName
            $objTab | Add-Member -MemberType NoteProperty -Name "Tab" -Value $tab.displayName
            $objTab | Add-Member -MemberType NoteProperty -Name "AppName" -Value $tab.teamsApp.displayName
            $objTab | Add-Member -MemberType NoteProperty -Name "AppId" -Value $tab.teamsApp.Id
            $objTab | Add-Member -MemberType NoteProperty -Name "AddedVia" -Value $tab.teamsApp.distributionMethod
            $objTab | Add-Member -MemberType NoteProperty -Name "DateAdded" -Value $tab.configuration.dateAdded

            $ReportTabs += $objTab
        }
    }
}

#Export the result
$ReportApps | select * -ExcludeProperty Number | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_TeamsAppsReport.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
$ReportTabs | select * -ExcludeProperty Number | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_TeamsTabsReport.csv" -NoTypeInformation -Encoding UTF8 -UseCulture