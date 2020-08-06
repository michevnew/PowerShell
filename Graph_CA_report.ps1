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
     $reportLine=[ordered]@{
        'Id' =$r.id
        'Name' =$r.displayName
        'State' =$r.state
        'Created' =$r.createdDateTime
        'Modified' =$r.modifiedDateTime

        #conditions
        'cRiskLevel' =($r.conditions.signInRiskLevels -join ";")
        'cClientApp' =($r.conditions.clientAppTypes -join ";")
        'cDeviceState' =($r.conditions.deviceStates -join ";")
        #'cDevices' =$r.conditions.devices #deprecated, exclude
        'cApplications' =("Included: $($r.conditions.applications.includeApplications -join ',')" + ";Excluded: $($r.conditions.applications.excludeApplications -join ',')" + ";Actions: $($r.conditions.applications.includeUserActions -join ',')")
        #'cUsers' =("Included: $($r.conditions.users.includeUsers -join ',')" + ";Excluded: $($r.conditions.users.excludeUsers -join ',')")
        'cUsers' =("Included: $(ReturnIdentifiers $r.conditions.users.includeUsers)" + ";Excluded: $(ReturnIdentifiers $r.conditions.users.excludeUsers)")
        #'cGroups' =("Included: $($r.conditions.users.includeGroups -join ',')" + ";Excluded: $($r.conditions.users.excludeGroups -join ',')")
        'cGroups' =("Included: $(ReturnIdentifiers $r.conditions.users.includeGroups)" + ";Excluded: $(ReturnIdentifiers $r.conditions.users.excludeGroups)")
        #'cRoles' =("Included: $($r.conditions.users.includeRoles -join ',')" + ";Excluded: $($r.conditions.users.excludeRoles -join ',')")
        'cRoles' =("Included: $(ReturnIdentifiers $r.conditions.users.includeRoles)" + ";Excluded: $(ReturnIdentifiers $r.conditions.users.excludeRoles)")
        'cPlatforms' =("Included: $($r.conditions.platforms.includePlatforms -join ',')" + ";Excluded: $($r.conditions.platforms.excludePlatforms -join ',')")
        'cLocations' =("Included: $($r.conditions.locations.includeLocations -join ',')" + ";Excluded: $($r.conditions.locations.excludeLocations -join ',')")
    }

    #conrtos
    if ($r.grantControls) {
        $reportLine.'aActions' =($r.grantControls.builtInControls -join ";")
        $reportLine.'aToU' =($r.grantControls.termsOfUse -join ";")
        $reportLine.'aCustom' =($r.grantControls.customAuthenticationFactors -join ";")
        $reportLine.'aOperator' =$r.grantControls.operator
    }

    #session controls
    if ($r.sessionControls) {
        $reportLine.'sesRestriction' =(&{If($r.sessionControls.applicationEnforcedRestrictions.isEnabled) {"Enabled"} Else {"Not enabled"}}) 
        $reportLine.'sesMCAS' =$r.sessionControls.cloudAppSecurity
        $reportLine.'sesBrowser' =$r.sessionControls.persistentBrowser
        $reportLine. 'sesSignInFrequency' =(&{If($r.sessionControls.signInFrequency.value) {"Enabled"} Else {"Not enabled"}})
        if ($r.sessionControls.signInFrequency.value) { $reportLine.'sesSignInFrequencyPeriod' ="$($r.sessionControls.signInFrequency.value) $($r.sessionControls.signInFrequency.type)" }
    }
    
    $output += @([pscustomobject]$reportLine)
}

#return output to console
$output | fl
#$output | ogv
#export to CSV
$output | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_CApolicies.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
