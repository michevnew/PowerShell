#Requires -Version 3.0
#Make sure to fill in all the required variables before running the script
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Directory.Read.All

#==========================================================================
#Helper functions
#==========================================================================

function parse-AppPermissions {

    Param(
    #App role assignment object
    [Parameter(Mandatory=$true)]$appRoleAssignments)

    foreach ($appRoleAssignment in $appRoleAssignments) {
        $resID = (Get-ServicePrincipalRoleById $appRoleAssignment.resourceAppId).appDisplayName
        foreach ($entry in $appRoleAssignment.resourceAccess) {
            if ($entry.Type -eq "Role") {
                $entryValue = ($OAuthScopes[$appRoleAssignment.resourceAppId].AppRoles | ? {$_.id -eq $entry.id}).Value
                if (!$entryValue) { $entryValue = "Orphaned ($($entry.id))" }
                $OAuthpermA["[" + $resID + "]"] += "," + $entryValue
            }
            elseif ($entry.Type -eq "Scope") { 
                $entryValue = ($OAuthScopes[$appRoleAssignment.resourceAppId].publishedPermissionScopes | ? {$_.id -eq $entry.id}).Value
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
    #do we need anything other than AppRoles? add a $select statement...
    if (!$OAuthScopes[$resID]) {
        $uri = "https://graph.microsoft.com/beta/servicePrincipals?`$filter=appid eq '$resID'"
        $res = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeader -Verbose:$VerbosePreference
        $OAuthScopes[$resID] = ($res.Content | ConvertFrom-Json).Value
    }
    return $OAuthScopes[$resID]
}

function parse-Credential {

    Param(
    #Key credential or password credential object
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$cred)

    #Return number of credentials
    $credout = ($cred.count).ToString()
    #Check if any there is an expired credential
    if ((Get-Date) -gt ($cred.endDateTime | sort -Descending | select -First 1)) { $credout +=  " (expired)" } 
    else {}
    #Check for credentials with excessive validity
    foreach ($c in $cred) {
        if ((New-TimeSpan -Start $c.startDateTime -End $c.endDateTime).Days -ge 365) { $excessiveValidity = $true }
    }
    if ($excessiveValidity) { $credout +=  " (excessive validity)" }
    return $credout
}


#==========================================================================
#Main script starts here
#==========================================================================

#Get MSAL token. Make sure to fill in all the variable values here. Or replace with your own preferred method to obtain token.
$tenantId = "tenant.onmicrosoft.com"
$url = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'

$Scopes = New-Object System.Collections.Generic.List[string]
$Scope = "https://graph.microsoft.com/.default"
$Scopes.Add($Scope)

$body = @{
    grant_type = "client_credentials"
    client_id = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    client_secret = "verylongstring"
    scope = $Scopes
}

try { 
    $res = Invoke-WebRequest -Method Post -Uri $url -Verbose -Body $body
    $token = ($res.Content | ConvertFrom-Json).access_token
    
    $authHeader = @{
       'Authorization'="Bearer $token"
    }}
catch { Write-Host "Failed to obtain token, aborting..." ; return }

#Get the list of application objects within the tenant.
$Apps = @()

$uri = "https://graph.microsoft.com/beta/applications?`$top=999"
#once they fix $expand($select)
#$uri = "https://graph.microsoft.com/v1.0/applications?`$top=999&`$expand=owners($select=userPrincipalName)
do {
    $result = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeader -Verbose:$VerbosePreference
    $uri = ($result.Content | ConvertFrom-Json).'@odata.nextLink'

    #If we are getting multiple pages, best add some delay to avoid throttling
    Start-Sleep -Milliseconds 500
    $Apps += ($result.Content | ConvertFrom-Json).Value
} while ($uri)


$OAuthScopes = @{} #hash-table to store data for app roles and stuff
$output = [System.Collections.Generic.List[Object]]::new() #output variable
$i=0;

foreach ($App in $Apps) {

    Write-Verbose "Processing application $($App.id)..."

    $OAuthpermA = @{};$OAuthpermD = @{};$resID = $null;

    #prepare the output object
    $i++;$objPermissions = [PSCustomObject][ordered]@{
        "Number" = $i
        "Application Name" = $App.DisplayName
        "ApplicationId" = $App.AppId
        "Publisher" = $App.publisherDomain
        "Verified" = (&{if ($App.verifiedPublisher.verifiedPublisherId) {$App.verifiedPublisher.displayName} else {"Not verified"}})
        "Certification" = $app.certification
        "SignInAudience" = $app.signInAudience
        "ObjectId" = $App.id
        "Created on" = (&{if ($app.createdDateTime) {(Get-Date($App.createdDateTime) -format g)} else {"N/A"}})
        #"Owner" = 
        "Permissions (application)" = $null
        "Permissions (delegate)" = $null
        "Permissions (API)" = $null
        "Allow Public client flows" = (&{if ($app.isFallbackPublicClient -eq "true") {"True"} else {"False"}}) #probably need to handle 'null' value as well
        "Key credentials" = (&{if ($app.keyCredentials) {parse-Credential $app.keyCredentials} else {""}})
        "Key credentials expiry date" = (&{if ($app.keyCredentials) {($app.keyCredentials.endDateTime | sort -Descending | select -First 1)} else {""}})
        "Password credentials" = (&{if ($app.passwordCredentials) {parse-Credential $app.passwordCredentials} else {""}})
        "Password credentials expiry date" = (&{if ($app.passwordCredentials) {($app.passwordCredentials.endDateTime | sort -Descending | select -First 1)} else {""}})
    }

    #Process permissions #Add STATUS of consent per each entry?
    parse-AppPermissions $app.requiredResourceAccess
    $objPermissions.'Permissions (application)' = (($OAuthpermA.GetEnumerator()  | % { "$($_.Name):$($_.Value.ToString().TrimStart(','))"}) -join ";")
    $objPermissions.'Permissions (delegate)' = (($OAuthpermD.GetEnumerator()  | % { "$($_.Name):$($_.Value.ToString().TrimStart(','))"}) -join ";")
    if ($app.api) { $objPermissions.'Permissions (API)' = ($app.api.oauth2PermissionScopes.value -join ";") }

    $output.Add($objPermissions)
}

#Export the result to CSV file
$output | select * -ExcludeProperty Number | Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_GraphAppRegInventory.csv"