#Requires -Version 3.0
#Make sure to fill in all the required variables before running the script
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Application.Read.All (required)
#    AuditLog.Read.All (optional, needed to retrieve Sign-in stats)
#    DirectoryRecommendations.Read.All (optional, needed to retrieve directory recommendations)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/5940/reporting-on-entra-id-application-registrations

[CmdletBinding(SupportsShouldProcess)] #Make sure we can use -Verbose
Param([switch]$IncludeSignInStats,[switch]$IncludeRecommendations)

#==========================================================================
#Helper functions
#==========================================================================

#Lite version of the Parse-JWTtoken function from https://www.michev.info/Blog/Post/2247/parse-jwt-token-in-powershell
function Parse-JWTtoken {

    [cmdletbinding()]
    param([Parameter(Mandatory=$true)][string]$token)

    #Validate as per https://tools.ietf.org/html/rfc7519
    if (!$token.Contains(".") -or !$token.StartsWith("eyJ")) { Write-Error "Invalid token" -ErrorAction Stop }

    #Payload
    $tokenPayload = $token.Split(".")[1].Replace('-', '+').Replace('_', '/')
    #Fix padding as needed, keep adding "=" until string length modulus 4 reaches 0
    while ($tokenPayload.Length % 4) { Write-Verbose "Invalid length for a Base-64 char array or string, adding ""="""; $tokenPayload += "=" }

    #Convert to Byte array
    $tokenByteArray = [System.Convert]::FromBase64String($tokenPayload)
    #Convert to string array
    $tokenArray = [System.Text.Encoding]::ASCII.GetString($tokenByteArray)

    #Convert from JSON to PSObject
    $tokobj = $tokenArray | ConvertFrom-Json

    return $tokobj
}

function parse-AppPermissions {

    Param(
    #App role assignment object
    [Parameter(Mandatory=$true)]$AppRoleAssignments)

    foreach ($AppRoleAssignment in $AppRoleAssignments) {
        $resID = (Get-ServicePrincipalRoleById $AppRoleAssignment.resourceAppId).appDisplayName
        foreach ($entry in $AppRoleAssignment.resourceAccess) {
            if ($entry.Type -eq "Role") {
                $entryValue = ($OAuthScopes[$AppRoleAssignment.resourceAppId].AppRoles | ? {$_.id -eq $entry.id}).Value
                if (!$entryValue) { $entryValue = "Orphaned ($($entry.id))" }
                $OAuthpermA["[" + $resID + "]"] += "," + $entryValue
            }
            elseif ($entry.Type -eq "Scope") {
                $entryValue = ($OAuthScopes[$AppRoleAssignment.resourceAppId].publishedPermissionScopes | ? {$_.id -eq $entry.id}).Value
                if (!$entryValue) { $entryValue = "Orphaned ($($entry.id))" }
                $OAuthpermD["[" + $resID + "]"] += "," + $entryValue
            }
            else { continue }
        }
    }
}

function Get-ServicePrincipalRoleById {

    Param(
    #Service principal object
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$resID)

    #check if we've already collected this SP data
    if (!$OAuthScopes[$resID]) {
        $uri = "https://graph.microsoft.com/beta/servicePrincipals?`$filter=appid eq '$resID'"
        $res = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeader -ErrorAction Stop -Verbose:$false
        $OAuthScopes[$resID] = ($res.Content | ConvertFrom-Json).Value
    }
    return $OAuthScopes[$resID]
}

function parse-Credential {

    Param(
    #Key credential or password credential object
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$cred)

    $credout = @($null,@())
    #Return number of credentials
    $credout[0] = ($cred.count).ToString()
    #Check if any there is an expired credential
    if ((Get-Date) -gt ($cred.endDateTime | Sort-Object -Descending | select -First 1)) { $credout[0] += " (expired)" }
    #Check for credentials with excessive validity
    foreach ($c in $cred) {
        $cstring = $c.keyId
        if ((New-TimeSpan -Start $c.startDateTime -End $c.endDateTime).Days -ge 180) { $excessiveValidity = $true }
        if ((Get-Date) -gt ($c.endDateTime)) { $cstring += "(EXPIRED)" }
        $cstring += "(valid from $($c.startDateTime) to $($c.endDateTime))"
        $credout[1] += $cstring
    }
    if ($excessiveValidity) { $credout[0] += " (excessive validity)" }

    return $credout
}

function parse-SPSignInStats {

    Param(
        #Report object
        [Parameter(Mandatory=$true)]$SPSignInStats)

    foreach ($SPSignInStat in $SPSignInStats) {
        if (!$SPStats[$SPSignInStat.appId]) {
            $SPStats[$SPSignInStat.appId] = @{
                "LastSignIn" = $SPSignInStat.lastSignInActivity.lastSignInDateTime
                "LastDelegateClientSignIn" = $SPSignInStat.delegatedClientSignInActivity.lastSignInDateTime
                "LastDelegateResourceSignIn" = $SPSignInStat.delegatedResourceSignInActivity.lastSignInDateTime
                "LastAppClientSignIn" = $SPSignInStat.applicationAuthenticationClientSignInActivity.lastSignInDateTime
                "LastAppResourceSignIn" = $SPSignInStat.applicationAuthenticationResourceSignInActivity.lastSignInDateTime
            }
        }
    }
    #return $SPStats
}

function parse-AppCredStats {

    Param(
        #Report object
        [Parameter(Mandatory=$true)]$AppCredStats)

    foreach ($AppCredStat in $AppCredStats) {
        if (!$AppCreds[$AppCredStat.appId]) {
            $AppCreds[$AppCredStat.appId] = @{
                "LastSignIn" = $AppCredStat.signInActivity.lastSignInDateTime
                #Add keyId?
                #We can have multiples here?
                #Also credentialOrigin?
            }
        }
    }
    #return $SPStats
}

function parse-Recommendations {

    Param(
        # Report object
        [Parameter(Mandatory=$true)]$dirRec)

    foreach ($dirRec in $dirRecs) {
        #Collect details depending on the recommendation type
        foreach ($impactedResource in $dirRec.impactedResources) {
            #Should contain all the details we need, use for each scenario below
            if ($impactedResource.additionalDetails) { $details = $impactedResource.additionalDetails.value | ConvertFrom-Json }
            #else { continue } #AdditionaDetails can be null, don't skip the rest of the code

            #Parse details depending on the recommendation type, multiple recommendations for the same app supported
            switch ($dirrec.recommendationType) {
                "overprivilegedApps" {
                    $toRemove = $details | ? {$_.recommendation -eq "Remove"} | select overprivileged_permission, grant_type, least_privileged_permission
                    $key = "RemovePermissions"
                    $value = $toRemove
                }
                "adalToMsalMigration" {
                    $key = "StillUsesAdal"
                    $value = $true
                }
                "staleAppCreds" {
                    $key = "UnusedCredentials"
                    $value = $details | select Key_id, key_type, last_active_date
                }
                "applicationCredentialExpiry" {
                    $key = "ExpiredCredentials"
                    $value = $details | select Key_id, key_type, last_active_date
                }
                "staleApps" {
                    $key = "UnusedApps"
                    $value = $details | select last_active_date
                }
                default {
                    # We either cover this recommendation ourselves, or it's not relevant
                    return
                }
            }

            if (!$Recommendations.ContainsKey($impactedResource.Id)) {
                $Recommendations[$impactedResource.Id] = @{$key = $value}
            }
            else {
                # If we already have a recommendation for this app
                $Recommendations[$impactedResource.Id] += @{$key = $value}
            }
        }
    }
}

function prepare-RecommendationOutput {

    param (
        [Parameter(Mandatory=$true)]$rec
    )

    $out = @()
    foreach ($key in $rec.Keys) {
        $value = switch ($key) {
            "RemovePermissions" {
                ($rec[$key] | % {
                    if ($_.least_privileged_permission) { "$($_.overprivileged_permission)($($_.grant_type)) -> $($_.least_privileged_permission)" }
                    else { "$($_.overprivileged_permission)($($_.grant_type))" }
                }) -join ";"
            }
            "StillUsesAdal" { "True" }
            "UnusedCredentials" {
                ($rec[$key] | % {
                    if ($_.last_active_date) { "$($_.key_id)($($_.key_type))(Last active -> $($_.last_active_date))" }
                    else { "$($_.key_id)($($_.key_type))(Last active -> N/A)" }
                }) -join ";"
            }
            "ExpiredCredentials" { ($rec[$key] | % { "$($_.key_id)($($_.key_type))" }) -join ";" }
            "UnusedApps" { (&{if ($rec[$key].last_active_date) { "(Last active -> $($rec[$key].last_active_date))" } else { "(Last active -> N/A)" }}) }
            Default { "" }
        }

        $out += "[$key]:$value"
    }
    return $out -join ";"
}

#==========================================================================
#Main script starts here
#==========================================================================

#Get an Access token. Make sure to fill in all the variable values here. Or replace with your own preferred method to obtain token.
$tenantId = "tenant.onmicrosoft.com"
$uri = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'
$clientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$client_secret = "verylongstring"

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
    Write-Verbose "Obtaining token..."
    $res = Invoke-WebRequest -Method Post -Uri $uri -Body $body -ErrorAction Stop -Verbose:$false
    $token = ($res.Content | ConvertFrom-Json).access_token

    $authHeader = @{
       'Authorization'="Bearer $token"
    }}
catch { Write-Output "Failed to obtain token, aborting..." ; return }

$tokenobj = Parse-JWTtoken $token

#Get the list of application objects within the tenant.
$Apps = @()

Write-Verbose "Retrieving list of applications..."
$uri = "https://graph.microsoft.com/beta/applications?`$top=999"
#once they fix $expand($select)
#$uri = "https://graph.microsoft.com/v1.0/applications?`$top=999&`$expand=owners($select=userPrincipalName)
try {
    do {
        $result = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeader -ErrorAction Stop -Verbose:$false
        $uri = ($result.Content | ConvertFrom-Json).'@odata.nextLink'

        #If we are getting multiple pages, best add some delay to avoid throttling
        Start-Sleep -Milliseconds 100
        $Apps += ($result.Content | ConvertFrom-Json).Value
    } while ($uri)
}
catch {
    Write-Output "Failed to retrieve the list of applications, aborting..."
    Write-Error $_ -ErrorAction Stop
    return
}

#Gather sign-in stats for the service principals, if requested
if ($IncludeSignInStats) {
    Write-Verbose "Retrieving sign-in stats for service principals..."

    if ($tokenobj.roles -notcontains "AuditLog.Read.All") { Write-Warning "The access token does not have the required permissions to retrieve SP sign-in activities, data will not be included in the output..." }
    else {
        $SPSignInStats = @()
        $uri = "https://graph.microsoft.com/beta/reports/servicePrincipalSignInActivities?`$top=999"

        try {
            do {
                $result = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeader -ErrorAction Stop -Verbose:$false
                $uri = ($result.Content | ConvertFrom-Json).'@odata.nextLink'

                #If we are getting multiple pages, best add some delay to avoid throttling
                Start-Sleep -Milliseconds 200
                $SPSignInStats += ($result.Content | ConvertFrom-Json).Value
            } while ($uri)
        }
        catch { Write-Warning "Failed to retrieve the report of service principals sign-ins, data will not be included in the output..." }

        $SPStats = @{} #hash-table to store sign-in stats data
        if ($SPSignInStats) { parse-SPSignInStats $SPSignInStats }

        Write-Verbose "Retrieving application credential usage stats..."
        #This requires Azure AD Premium P2 now, and will require Workload Idenity license when GA :(
        $AppCredStats = @()
        $uri = "https://graph.microsoft.com/beta/reports/appCredentialSignInActivities?`$top=999"

        try {
            do {
                $result = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeader -ErrorAction Stop -Verbose:$false
                $uri = ($result.Content | ConvertFrom-Json).'@odata.nextLink'

                #If we are getting multiple pages, best add some delay to avoid throttling
                Start-Sleep -Milliseconds 200
                $AppCredStats += ($result.Content | ConvertFrom-Json).Value
            } while ($uri)
        }
        catch { Write-Warning "Failed to retrieve the report of application credential usage, data will not be included in the output..." }

        $AppCreds = @{} #hash-table to store sign-in stats data
        if ($AppCredStats) { parse-AppCredStats $AppCredStats }
    }
}

#Gather directory recommendations
if ($IncludeRecommendations) {
    Write-Verbose "Retrieving directory recommendations..."
    $dirRecs = @()
    $uri = "https://graph.microsoft.com/beta/directory/recommendations?`$filter=featureAreas/any(x:x eq 'applications')&`$expand=impactedResources" #are we certain it returns all impacted resources or just 20/100/whatever?

    try {
        $result = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeader -ErrorAction Stop -Verbose:$false
        $dirRecs = ($result.Content | ConvertFrom-Json).Value
    }
    catch { Write-Warning "Failed to retrieve directory recommendations, data will not be included in the output..." }

    $Recommendations = @{}
    if ($dirRecs) { parse-Recommendations $dirRecs }
}

#Prepare variables
$OAuthScopes = @{} #hash-table to store data for app roles and stuff
$output = [System.Collections.Generic.List[Object]]::new() #output variable
$i=0; $count = 1; $PercentComplete = 0;

#Process the list of applications
foreach ($App in $Apps) {
    #Progress message
    $ActivityMessage = "Retrieving data for application $($App.DisplayName). Please wait..."
    $StatusMessage = ("Processing application {0} of {1}: {2}" -f $count, @($Apps).count, $App.id)
    $PercentComplete = ($count / @($Apps).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #simple anti-throttling control
    Start-Sleep -Milliseconds 100
    Write-Verbose "Processing application $($App.id)..."

    #Get owners info. We do not use $expand, as it returns the full set of object properties
    try {
        Write-Verbose "Retrieving owners info..."
        $owners = @()
        $res = Invoke-WebRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/applications/$($App.id)/owners?`$select=id,userPrincipalName&`$top=999" -Headers $authHeader -ErrorAction Stop -Verbose:$false
        $owners += ($res.Content | ConvertFrom-Json).Value.userPrincipalName
    }
    catch { Write-Verbose "Failed to retrieve owners info for application $($App.id) ..." }

    #prepare the output object
    $i++;$objPermissions = [PSCustomObject][ordered]@{
        "Number" = $i
        "Application Name" = (&{if ($App.DisplayName) { $App.DisplayName } else { $null }}) #Apparently DisplayName can be null
        "ApplicationId" = $App.AppId
        "Publisher Domain" = (&{if ($App.PublisherDomain) { $App.PublisherDomain } else { $null }})
        "Verified" = (&{if ($App.verifiedPublisher.verifiedPublisherId) { $App.verifiedPublisher.displayName } else { "Not verified" }})
        "Certification" = (&{if ($App.certification) { $App.certification.certificationDetailsUrl } else { "" }})
        "SignInAudience" = $App.signInAudience
        "ObjectId" = $App.id
        "Created on" = (&{if ($App.createdDateTime) { (Get-Date($App.createdDateTime) -format g) } else { "N/A" }})
        "Owners" = (&{if ($owners) { $owners -join ";" } else { $null }})
        "Permissions (application)" = $null
        "Permissions (delegate)" = $null
        "Permissions (API)" = $null
        "Allow Public client flows" = (&{if ($App.isFallbackPublicClient -eq "true") { "True" } else { "False" }}) #probably need to handle 'null' value as well
        "Key credentials" = (&{if ($App.keyCredentials) { (parse-Credential $App.keyCredentials)[0] } else { "" }})
        "KeyCreds" = (&{if ($App.keyCredentials) { ((parse-Credential $App.keyCredentials)[1]) -join ";" } else { $null }})
        "Next expiry date (key)" = (&{if ($App.keyCredentials) { ($App.keyCredentials.endDateTime | ? {$_ -ge (Get-Date)} | Sort-Object -Descending | select -First 1) } else { "" }})
        "Password credentials" = (&{if ($App.passwordCredentials) { (parse-Credential $App.passwordCredentials)[0] } else { "" }})
        "PasswordCreds" = (&{if ($App.passwordCredentials) { ((parse-Credential $App.passwordCredentials)[1]) -join ";" } else { $null }})
        "Next expiry date (password)" = (&{if ($App.passwordCredentials) { ($App.passwordCredentials.endDateTime | ? {$_ -ge (Get-Date)} | Sort-Object -Descending | select -First 1) } else { "" }})
        "App property lock" = (&{if ($App.servicePrincipalLockConfiguration.isEnabled -and $App.servicePrincipalLockConfiguration.allProperties) { $true } else { $false }})
        "HasBadURIs" = (&{if ($App.web.redirectUris -match "localhost|http://|urn:|\*") { $true } else { $false }})
        "Redirect URIs" = (&{if ($App.web.redirectUris) { $App.web.redirectUris -join ";" } else { $null }})
        #"identifierUris" = (&{if ($App.identifierUris) { $App.identifierUris -join ";" } else { $null }}) #-match "api://($app.AppId)"
    }

    #Include sign-in stats, if requested
    if ($IncludeSignInStats) {
        if ($tokenobj.roles -contains "AuditLog.Read.All") {
            $objPermissions | Add-Member -MemberType NoteProperty -Name "Last sign-in" -Value (&{if ($SPStats[$App.appId].LastSignIn) { (Get-Date($SPStats[$App.appid].LastSignIn) -format g) } else { $null }})
            $objPermissions | Add-Member -MemberType NoteProperty -Name "Last delegate client sign-in" -Value (&{if ($SPStats[$App.appid].LastDelegateClientSignIn) { (Get-Date($SPStats[$App.appid].LastDelegateClientSignIn) -format g) } else { $null }})
            $objPermissions | Add-Member -MemberType NoteProperty -Name "Last delegate resource sign-in" -Value (&{if ($SPStats[$App.appid].LastDelegateResourceSignIn) { (Get-Date($SPStats[$App.appid].LastDelegateResourceSignIn) -format g) } else { $null }})
            $objPermissions | Add-Member -MemberType NoteProperty -Name "Last app client sign-in" -Value (&{if ($SPStats[$App.appid].LastAppClientSignIn) { (Get-Date($SPStats[$App.appid].LastAppClientSignIn) -format g) } else { $null }})
            #This one will always be null, so maybe remove it?
            $objPermissions | Add-Member -MemberType NoteProperty -Name "Last app resource sign-in" -Value (&{if ($SPStats[$App.appid].LastAppResourceSignIn) { (Get-Date($SPStats[$App.appid].LastAppResourceSignIn) -format g) } else { $null }})

            #Add credential usage stats, if available
            $objPermissions | Add-Member -MemberType NoteProperty -Name "Last credential sign-in" -Value (&{if ($AppCreds[$App.appid].LastSignIn) { (Get-Date($AppCreds[$App.appid].LastSignIn) -format g) } else { $null }})
        }
    }

    #Check if the app is leveraging any AADGraph permissions
    if ($App.requiredResourceAccess | ? {$_.resourceAppId -eq "00000002-0000-0000-c000-000000000000"}) {
        $objPermissions | Add-Member -MemberType NoteProperty -Name "UsesAADGraph" -Value $true
    }
    else { $objPermissions | Add-Member -MemberType NoteProperty -Name "UsesAADGraph" -Value $false }

    #Include recommendations, if requested
    if ($IncludeRecommendations) {
        if ($tokenobj.roles -contains "DirectoryRecommendations.Read.All") {
            $objPermissions | Add-Member -MemberType NoteProperty -Name "Recommendations" -Value (&{if ($Recommendations.ContainsKey($App.appid)) { prepare-RecommendationOutput $Recommendations[$App.appid] } else { $null }})
        }
    }

    #Process permissions #Add STATUS of consent per each entry?
    $OAuthpermA = @{};$OAuthpermD = @{};$resID = $null;

    if ($App.requiredResourceAccess) { parse-AppPermissions $App.requiredResourceAccess }
    else { Write-Verbose "No permissions found for application $($App.id), skipping..." }

    #parse-AppPermissions $App.requiredResourceAccess
    $objPermissions.'Permissions (application)' = (($OAuthpermA.GetEnumerator() | % { "$($_.Name):$($_.Value.ToString().TrimStart(','))"}) -join ";")
    $objPermissions.'Permissions (delegate)' = (($OAuthpermD.GetEnumerator() | % { "$($_.Name):$($_.Value.ToString().TrimStart(','))"}) -join ";")
    if ($App.api) { $objPermissions.'Permissions (API)' = ($App.api.oauth2PermissionScopes.value -join ";") }

    $output.Add($objPermissions)
}

#Export the result to CSV file
$output | select * -ExcludeProperty Number | Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_GraphAppRegInventory.csv"
Write-Verbose "Output exported to $($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_GraphAppRegInventory.csv"