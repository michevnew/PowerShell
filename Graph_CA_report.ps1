#Set up
$AppId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #clientID of your AAD app, must have Policy.Read.All permissions. Also needs Directory.Read.All permissions if you want the "readable" object names
$client_secret = Get-Content .\ReportingAPIsecret.txt | ConvertTo-SecureString
$app_cred = New-Object System.Management.Automation.PsCredential($AppId, $client_secret)
$TenantId = "tenant.onmicrosoft.com" #your tenant

#helper functions
function GUIDtoIdentifier ([GUID]$GUID) {

    #Only supports User, Groups,Roles
    #Does NOT support applicaitons, named locations
    $Json = @{
        "ids" = @("$GUID")
    } | ConvertTo-Json

    #$Json | Out-Default
    
    $GObject = Invoke-WebRequest -Headers $AuthHeader1 -Uri "https://graph.microsoft.com/v1.0/directoryObjects/getByIds" -Method Post -Body $Json -ContentType "application/json"
    $result = ($GObject.Content | ConvertFrom-Json).Value

    #$result | Out-Default

    switch ($result.'@odata.type') {
        "#microsoft.graph.user" { return $result.UserPrincipalName }
        "#microsoft.graph.group" { return $result.displayName }
        "#microsoft.graph.directoryRole" { return $result.displayName }
        default { return $GUID.Guid }
    }
}

function ReturnIdentifiers ([string[]]$GUIDs) {
    $id = @()

    foreach ($GUID in $GUIDs) {
        try { [GUID]$GUID | Out-Null ; $id += GUIDtoIdentifier $GUID }
        catch { return $GUID }
    }

    return ($id -join ",")
}

#Get the token
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
$CAs = Invoke-WebRequest -Headers $AuthHeader1 -Uri "https://graph.microsoft.com/beta/identity/conditionalAccess/policies/"
$result = ($CAs.Content | ConvertFrom-Json).Value

$output = @();
foreach ($r in $result) {
    $CAinfo = New-Object PSObject
    $CAinfo | Add-Member -MemberType NoteProperty -Name Id -Value $r.id
    $CAinfo | Add-Member -MemberType NoteProperty -Name Name -Value $r.displayName
    $CAinfo | Add-Member -MemberType NoteProperty -Name State -Value $r.state
    $CAinfo | Add-Member -MemberType NoteProperty -Name Created -Value $r.createdDateTime
    $CAinfo | Add-Member -MemberType NoteProperty -Name Modified -Value $r.modifiedDateTime

    #conditions
    $CAinfo | Add-Member -MemberType NoteProperty -Name cRiskLevel -Value ($r.conditions.signInRiskLevels -join ";")
    $CAinfo | Add-Member -MemberType NoteProperty -Name cClientApp -Value ($r.conditions.clientAppTypes -join ";")
    $CAinfo | Add-Member -MemberType NoteProperty -Name cDeviceState -Value ($r.conditions.deviceStates -join ";")
    #$CAinfo | Add-Member -MemberType NoteProperty -Name cDevices -Value $r.conditions.devices #deprecated, exclude
    $CAinfo | Add-Member -MemberType NoteProperty -Name cApplications -Value ("Included: $($r.conditions.applications.includeApplications -join ',')" + ";Excluded: $($r.conditions.applications.excludeApplications -join ',')" + ";Actions: $($r.conditions.applications.includeUserActions -join ',')")
    #$CAinfo | Add-Member -MemberType NoteProperty -Name cUsers -Value ("Included: $($r.conditions.users.includeUsers -join ',')" + ";Excluded: $($r.conditions.users.excludeUsers -join ',')")
    $CAinfo | Add-Member -MemberType NoteProperty -Name cUsers -Value ("Included: $(ReturnIdentifiers $r.conditions.users.includeUsers)" + ";Excluded: $(ReturnIdentifiers $r.conditions.users.excludeUsers)")
    #$CAinfo | Add-Member -MemberType NoteProperty -Name cGroups -Value ("Included: $($r.conditions.users.includeGroups -join ',')" + ";Excluded: $($r.conditions.users.excludeGroups -join ',')")
    $CAinfo | Add-Member -MemberType NoteProperty -Name cGroups -Value ("Included: $(ReturnIdentifiers $r.conditions.users.includeGroups)" + ";Excluded: $(ReturnIdentifiers $r.conditions.users.excludeGroups)")
    #$CAinfo | Add-Member -MemberType NoteProperty -Name cRoles -Value ("Included: $($r.conditions.users.includeRoles -join ',')" + ";Excluded: $($r.conditions.users.excludeRoles -join ',')")
    $CAinfo | Add-Member -MemberType NoteProperty -Name cRoles -Value ("Included: $(ReturnIdentifiers $r.conditions.users.includeRoles)" + ";Excluded: $(ReturnIdentifiers $r.conditions.users.excludeRoles)")
    $CAinfo | Add-Member -MemberType NoteProperty -Name cPlatforms -Value ("Included: $($r.conditions.platforms.includePlatforms -join ',')" + ";Excluded: $($r.conditions.platforms.excludePlatforms -join ',')")
    $CAinfo | Add-Member -MemberType NoteProperty -Name cLocations -Value ("Included: $($r.conditions.locations.includeLocations -join ',')" + ";Excluded: $($r.conditions.locations.excludeLocations -join ',')")

    #conrtos
    if ($r.grantControls) {
        $CAinfo | Add-Member -MemberType NoteProperty -Name aActions -Value ($r.grantControls.builtInControls -join ";")
        $CAinfo | Add-Member -MemberType NoteProperty -Name aToU -Value ($r.grantControls.termsOfUse -join ";")
        $CAinfo | Add-Member -MemberType NoteProperty -Name aCustom -Value ($r.grantControls.customAuthenticationFactors -join ";")
        $CAinfo | Add-Member -MemberType NoteProperty -Name aOperator -Value $r.grantControls.operator
    }

    #session controls
    if ($r.sessionControls) {
        $CAinfo | Add-Member -MemberType NoteProperty -Name sesRestriction -Value (&{If($r.sessionControls.applicationEnforcedRestrictions.isEnabled) {"Enabled"} Else {"Not enabled"}}) 
        $CAinfo | Add-Member -MemberType NoteProperty -Name sesMCAS -Value $r.sessionControls.cloudAppSecurity
        $CAinfo | Add-Member -MemberType NoteProperty -Name sesBrowser -Value $r.sessionControls.persistentBrowser
        $CAinfo | Add-Member -MemberType NoteProperty -Name sesSignInFrequency -Value (&{If($r.sessionControls.signInFrequency.value) {"Enabled"} Else {"Not enabled"}})
        if ($r.sessionControls.signInFrequency.value) { $CAinfo | Add-Member -MemberType NoteProperty -Name sesSignInFrequencyPeriod -Value "$($r.sessionControls.signInFrequency.value) $($r.sessionControls.signInFrequency.type)" }
    }
    $output += $CAinfo
}

#return output to console
$output | select Name,State,Created,Modified,aActions,aCustom,aOperator,aToU,cApplications,cClientApp,cDeviceState,cGroups,cLocations,cPlatforms,cRiskLevel,cRoles,cUsers,cUsers2,sesBrowser,sesMCAS,sesRestriction,sesSignInFrequency,sesSignInFrequencyPeriod #| ogv
#export to CSV
$output | select Id,Name,State,Created,Modified,aActions,aCustom,aOperator,aToU,cApplications,cClientApp,cDeviceState,cGroups,cLocations,cPlatforms,cRiskLevel,cRoles,cUsers,sesBrowser,sesMCAS,sesRestriction,sesSignInFrequency,sesSignInFrequencyPeriod | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_CApolicies.csv" -NoTypeInformation -Encoding UTF8 -UseCulture