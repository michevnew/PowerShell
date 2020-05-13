$AppId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$client_secret = Get-Content .\ReportingAPIsecret.txt | ConvertTo-SecureString
$app_cred = New-Object System.Management.Automation.PsCredential($AppId, $client_secret)
$TenantId = "tenant.onmicrosoft.com"

$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $app_cred.GetNetworkCredential().Password
    grant_type    = "client_credentials"
}
 
try { $tokenRequest = Invoke-WebRequest -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -ErrorAction Stop }
catch { Write-Host "Unable to obtain access token, aborting..."; return }

$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

$authHeader1 = @{
   'Content-Type'='application\json'
   'Authorization'="Bearer $token"
}

$Teams = @()
$uri = "https://graph.microsoft.com/beta/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"
do {
    $result = Invoke-WebRequest -Headers $AuthHeader1 -Uri $uri -ErrorAction Stop
    $uri = $result.'@odata.nextLink'
    #If we are getting multiple pages, best add some delay to avoid throttling
    Start-Sleep -Milliseconds 500
    $Teams += ($result.Content | ConvertFrom-Json).Value
} while ($uri)



$ReportApps = @();$ReportTabs = @();$count = 1
foreach ($team in $Teams) {

    #Progress message
    $ActivityMessage = "Retrieving data for team $($team.displayName). Please wait..."
    $StatusMessage = ("Processing team {0} of {1}: {2}" -f $count, @($Teams).count, $team.id)
    $PercentComplete = ($count / @($Teams).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Simple anti-throttling control
    Start-Sleep -Milliseconds 200

    #get a list of apps for the Team
    $teamApps = Invoke-WebRequest -Headers $authHeader1 -Uri "https://graph.microsoft.com/v1.0/teams/$($Team.id)/installedApps?`$expand=teamsApp" -ErrorAction Stop
    $teamApps = ($teamApps.content | ConvertFrom-Json).Value

    $i = 0
    foreach ($app in $teamApps) {
        $i++
        $objApp = New-Object PSObject
        $objApp | Add-Member -MemberType NoteProperty -Name "Number" -Value $i
        $objApp | Add-Member -MemberType NoteProperty -Name "Team" -Value $team.displayName
        $objApp | Add-Member -MemberType NoteProperty -Name "Name" -Value $app.teamsApp.displayName
        $objApp | Add-Member -MemberType NoteProperty -Name "AppId" -Value $app.teamsApp.Id
        $objApp | Add-Member -MemberType NoteProperty -Name "AddedVia" -Value $app.teamsApp.distributionMethod

        $ReportApps += $objApp
    }

    $TeamChannels = Invoke-WebRequest -Headers $AuthHeader1 -Uri "https://graph.microsoft.com/beta/Teams/$($Team.id)/channels" -ErrorAction Stop
    $TeamChannels = ($TeamChannels.Content | ConvertFrom-Json).value
    
    foreach ($channel in $TeamChannels) {
        $tabs = Invoke-WebRequest -Headers $authHeader1 -Uri "https://graph.microsoft.com/beta/teams/$($Team.id)/channels/$($channel.id)/tabs?`$expand=teamsApp" -ErrorAction Stop
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

            $ReportTabs += $objTab
        }
    }
}
$ReportApps | select * -ExcludeProperty Number #| Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_TeamsAppsReport.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
$ReportTabs | select * -ExcludeProperty Number #| Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_TeamsTabsReport.csv" -NoTypeInformation -Encoding UTF8 -UseCulture