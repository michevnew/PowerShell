#Requires -Version 3.0
#Make sure to fill in all the required variables before running the script
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Application.Read.All (to read the service principals)
#    User.Read.All (for "resolving" user IDs to UPNs)
#    Application.ReadWrite.All (required for remediation)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/5894/script-to-review-and-remove-service-principal-credentials

[CmdletBinding(SupportsShouldProcess)] #Make sure we can use -WhatIf and -Verbose
Param([switch]$IncludeBuiltin=$false,[switch]$IncludeConsents=$false,[switch]$Remediate=$false,[switch]$Force=$false)

#==========================================================================
#Helper functions
#==========================================================================
function parse-AppPermissions {

    Param(
    #App role assignment object
    [Parameter(Mandatory=$true)]$appRoleAssignments)

    foreach ($appRoleAssignment in $appRoleAssignments) {
        $resID = $appRoleAssignment.ResourceDisplayName
        $roleID = (Get-ServicePrincipalRoleById $appRoleAssignment.resourceId).appRoles | ? {$_.id -eq $appRoleAssignment.appRoleId} | select -ExpandProperty Value
        if (!$roleID) { $roleID = "Orphaned ($($appRoleAssignment.appRoleId))" }
        $OAuthperm["[" + $resID + "]"] += $("," + $RoleId)
    }
}

function parse-DelegatePermissions {

    Param(
    #oauth2PermissionGrants object
    [Parameter(Mandatory=$true)]$oauth2PermissionGrants)

    foreach ($oauth2PermissionGrant in $oauth2PermissionGrants) {
        $resID = (Get-ServicePrincipalRoleById $oauth2PermissionGrant.ResourceId).appDisplayName
        if ($null -ne $oauth2PermissionGrant.PrincipalId) {
            $userId = "(" + (Get-UserUPNById -objectID $oauth2PermissionGrant.principalId) + ")"
        }
        else { $userId = $null }

        if ($oauth2PermissionGrant.Scope) { $OAuthperm["[" + $resID + $userId + "]"] += ($oauth2PermissionGrant.Scope.Split(" ") -join ",") }
        else { $OAuthperm["[" + $resID + $userId + "]"] += "Orphaned scope" }
    }
}

function Get-ServicePrincipalRoleById {

    Param(
    #Service principal object
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$spID)

    #check if we've already collected this SP data
    #do we need anything other than AppRoles? add a $select statement...
    if (!$SPPerm[$spID]) {
        $res = Invoke-WebRequest -Method Get -Uri "https://graph.microsoft.com/beta/servicePrincipals/$spID" -Headers $authHeader -Verbose:$false
        $SPPerm[$spID] = ($res.Content | ConvertFrom-Json)
    }
    return $SPPerm[$spID]
}

function Get-UserUPNById {

    Param(
    #User objectID
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$objectID)

    #check if we've already collected this User's data
    #currently we store only UPN, store the entire object if needed
    if (!$SPusers[$objectID]) {
        $res = Invoke-WebRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users/$($objectID)?`$select=UserPrincipalName" -Headers $authHeader -Verbose:$false
        $SPusers[$objectID] = ($res.Content | ConvertFrom-Json).UserPrincipalName
    }
    return $SPusers[$objectID]
}

function Remediate-SP {

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(
        #SP object
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$oSP,
        [switch]$Force)

    $SPName = $oSP.'SP Name'; $SPID = $oSP.objectID; $uri = "https://graph.microsoft.com/v1.0/servicePrincipals/$SPID"
    Write-Verbose "Performing remediation actions against service principal ""$SPName"":"
    Write-Verbose ($oSP | select * -ExcludeProperty Number | Out-String)

    #Check for and remove password credentials
    if ($oSP.PasswordCreds) {
        Write-Verbose "Service principal ""$SPName"" has password credentials, removing..."
        if (-not ($Force -or $PSCmdlet.ShouldContinue("Performing operation ""Remove password credentials"" on service principal ""$SPName"" ($SPID)", 'Confirm'))) {
            Write-Verbose "User aborted, skipping operation..."
        }
        else {
            try {
                    Invoke-WebRequest -Method Patch -Uri $uri -Body (@{"passwordCredentials"=@()} | ConvertTo-Json) -Headers $authHeader -ContentType "application/json" -Verbose:$false -ErrorAction Stop | Out-Null
                    Write-Verbose "Successfully removed password credentials for SP ""$SPName"" ($SPID)"
                }
            catch { Write-Error "Failed to remove password credentials for SP ""$SPName"" ($SPID), please try the operation manually..." }
        }
    }
    else { Write-Verbose "No password credentials found for SP ""$SPName"", skipping operation..." }

    #Check for and remove key credentials
    if ($oSP.KeyCreds) {
        Write-Verbose "Service principal ""$SPName"" has key credentials, removing..."
        if (-not ($Force -or $PSCmdlet.ShouldContinue("Performing operation ""Remove key credentials"" on service principal ""$SPName"" ($SPID)", 'Confirm'))) {
            Write-Verbose "User aborted, skipping operation..."
        }
        else {
            try {
                    Invoke-WebRequest -Method Patch -Uri $uri -Body (@{"keyCredentials"=@()} | ConvertTo-Json) -Headers $authHeader -ContentType "application/json" -Verbose:$false -ErrorAction Stop | Out-Null
                    Write-Verbose "Successfully removed key credentials for SP ""$SPName"" ($SPID)"
                }
            catch { Write-Error "Failed to remove key credentials for SP ""$SPName"" ($SPID), please try the operation manually..." }
        }
    }
    else { Write-Verbose "No key credentials found for SP ""$SPName"", skipping operation..." }

    #Check for and null tokenEncryptionKeyId id property
    if ($oSP.TokenKey) {
        Write-Verbose "Service principal ""$SPName"" has token encryption key, removing..."
        if (-not ($Force -or $PSCmdlet.ShouldContinue("Performing operation ""Remove token key"" on service principal ""$SPName"" ($SPID)", 'Confirm'))) {
            Write-Verbose "User aborted, skipping operation..."
        }
        else {
            try {
                    Invoke-WebRequest -Method Patch -Uri $uri -Body (@{"tokenEncryptionKeyId"=$null} | ConvertTo-Json) -Headers $authHeader -ContentType "application/json" -Verbose:$false -ErrorAction Stop | Out-Null
                    Write-Verbose "Successfully removed token encryption key for SP ""$SPName"" ($SPID)"
                }
            catch { Write-Error "Failed to remove key credentials for SP ""$SPName"" ($SPID), please try the operation manually..." }
        }
    }
    else { Write-Verbose "No token encryption key found for SP ""$SPName"", skipping operation..." }
}

#==========================================================================
#Main script starts here
#==========================================================================

#Get an Access token. Make sure to fill in all the variable values here. Or replace with your own preferred method to obtain token.
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
    $res = Invoke-WebRequest -Method Post -Uri $url -Verbose:$false -Body $body
    $token = ($res.Content | ConvertFrom-Json).access_token

    $authHeader = @{
       'Authorization'="Bearer $token"
    }}
catch { Write-Host "Failed to obtain token, aborting..." ; return }

#Invoke-WebRequest does not support -WhatIf, so we suppress it here
$WhatIfPreference = $false

#Get the list of Service principal objects within the tenant.
$SPs = @()

if ($IncludeBuiltin) { $uri = "https://graph.microsoft.com/beta/servicePrincipals?`$top=999" }
else { $uri = "https://graph.microsoft.com/beta/servicePrincipals?`$top=999&`$filter=tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp')" }

do {
    $result = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeader -Verbose:$false
    $uri = ($result.Content | ConvertFrom-Json).'@odata.nextLink'

    #If we are getting multiple pages, best add some delay to avoid throttling
    Start-Sleep -Milliseconds 500
    $SPs += ($result.Content | ConvertFrom-Json).Value
} while ($uri)

#Process the list of service principals
$SPperm = @{} #hash-table to store data for app roles and stuff
$SPusers = @{} #hash-table to store data for users assigned delegate permissions and stuff
$output = [System.Collections.Generic.List[Object]]::new() #output variable
$i=0; $count = 1; $PercentComplete = 0;

foreach ($SP in $SPs) {
    #Progress message
    $ActivityMessage = "Retrieving data for service principal $($SP.DisplayName). Please wait..."
    $StatusMessage = ("Processing service principal {0} of {1}: {2}" -f $count, @($SPs).count, $SP.id)
    $PercentComplete = ($count / @($SPs).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #simple anti-throttling control
    if ($IncludeConsents) { Start-Sleep -Milliseconds 50 }
    else { Start-Sleep -Milliseconds 200 }
    Write-Verbose "Processing service principal $($SP.id)..."

    #Get owners info. We do not use $expand, as it returns the full set of object properties
    $owners = @()
    $res = Invoke-WebRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)/owners?`$select=id,userPrincipalName" -Headers $authHeader -Verbose:$false
    $owners += ($res.Content | ConvertFrom-Json).Value.userPrincipalName

    #prepare the output object
    $i++;$objPermissions = [PSCustomObject][ordered]@{
        "Number" = $i
        "Application Name" = $SP.appDisplayName
        "ApplicationId" = $SP.AppId
        "IsBuiltIn" = $SP.tags -notcontains "WindowsAzureActiveDirectoryIntegratedApp"
        "Publisher" = (&{if ($SP.PublisherName) { $SP.PublisherName } else { $null }})
        "Verified" = (&{if ($SP.verifiedPublisher.verifiedPublisherId) { $SP.verifiedPublisher.displayName } else { "Not verified" }})
        "Homepage" = (&{if ($SP.Homepage) { $SP.Homepage } else { $null }})
        "SP name" = $SP.displayName
        "ObjectId" = $SP.id
        "Created on" = (&{if ($SP.createdDateTime) {(Get-Date($SP.createdDateTime) -format g)} else { $null }})
        "Enabled" = $SP.AccountEnabled
        "PasswordCreds" = (&{if ($SP.passwordCredentials) { $SP.passwordCredentials.keyId -join ";" } else { $null }})
        "KeyCreds" = (&{if ($SP.keyCredentials) { $SP.keyCredentials.keyId -join ";" } else { $null }})
        "TokenKey" = (&{if ($SP.tokenEncryptionKeyId) { $SP.tokenEncryptionKeyId } else { $null }})
        "Last modified" = $null
        "Permissions (application)" = $null
        "Authorized By (application)" = $null
        "Permissions (delegate)" = $null
        "Valid until (delegate)" = $null
        "Authorized By (delegate)" = $null
        "Owners" = (&{if ($owners) { $owners -join ";" } else { $null }})
    }

    if ($IncludeConsents) {
        #Check for appRoleAssignments (application permissions)
        $appRoleAssignments = @()
        $res = Invoke-WebRequest -Method Get -Uri "https://graph.microsoft.com/beta/servicePrincipals/$($sp.id)/appRoleAssignments" -Headers $authHeader -Verbose:$false
        $appRoleAssignments = ($res.Content | ConvertFrom-Json).Value

        $OAuthperm = @{};
        $assignedto = @(); $resID = $null; $userId = $null;

        #process application permissions entries
        if (!$appRoleAssignments) { Write-Verbose "No application permissions to report on for SP $($SP.id), skipping..." }
        else {
            $objPermissions.'Last modified' = (Get-Date($appRoleAssignments.CreationTimestamp | select -Unique | Sort-Object -Descending | select -First 1) -format g)

            parse-AppPermissions $appRoleAssignments
            $objPermissions.'Permissions (application)' = (($OAuthperm.GetEnumerator() | % { "$($_.Name):$($_.Value.ToString().TrimStart(','))"}) -join ";")
            $objPermissions.'Authorized By (application)' = "An administrator (application permissions)"
        }

        #Check for oauth2PermissionGrants (delegate permissions)
        #Use /beta here, as /v1.0 does not return expiryTime
        $oauth2PermissionGrants = @()
        $res = Invoke-WebRequest -Method Get -Uri "https://graph.microsoft.com/beta/servicePrincipals/$($sp.id)/oauth2PermissionGrants" -Headers $authHeader -Verbose:$false
        $oauth2PermissionGrants = ($res.Content | ConvertFrom-Json).Value

        $OAuthperm = @{};
        $assignedto = @(); $resID = $null; $userId = $null;

        #process delegate permissions entries
        if (!$oauth2PermissionGrants) { Write-Verbose "No delegate permissions to report on for SP $($SP.id), skipping..." }
        else {
            parse-DelegatePermissions $oauth2PermissionGrants
            $objPermissions.'Permissions (delegate)' = (($OAuthperm.GetEnumerator() | % { "$($_.Name):$($_.Value.ToString().TrimStart(','))"}) -join ";")
            $objPermissions.'Valid until (delegate)' = (Get-Date($oauth2PermissionGrants.ExpiryTime | select -Unique | Sort-Object -Descending | select -First 1) -format g)

            if (($oauth2PermissionGrants.ConsentType | select -Unique) -eq "AllPrincipals") { $assignedto += "All users (admin consent)" }
            $assignedto += @($OAuthperm.Keys) | % {if ($_ -match "\((.*@.*)\)") {$Matches[1]}}
            $objPermissions.'Authorized By (delegate)' = (($assignedto | select -Unique) -join ",")
        }
    }
    else { Write-Verbose "Not listing any permissions/consents information for SP $($SP.id), use the -IncludeConsents switch to include it..." }

    if ($objPermissions.PasswordCreds -or $objPermissions.KeyCreds -or $objPermissions.TokenKey) { Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "ToBeReviewed" -Value $true }
    else { Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "ToBeReviewed" -Value $false }
    $output.Add($objPermissions)
}

#Export the result to CSV file
$output | select * -ExcludeProperty Number | Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_SPCredsReport.csv" -Confirm:$false
Write-Verbose "Output exported to $($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_SPCredsReport.csv"

#Remediation part
#Filter out any SP with no credentials
$outR = $output | ? {$_.ToBeReviewed}

if (!$outR) { Write-Output "No service principal(s) that need remediation found!" }
else {
    if ($Remediate) {
        Write-Output "Found $($outR.count) service principal(s) that need remediation, proceeding..."
    }
    else {
        Write-Output "Found $($outR.count) service principal(s) that need remediation, use the -Remediate switch to perform the operation..."
        return
    }
}

foreach ($o in $outR) {
    Remediate-SP $o -Confirm:$ConfirmPreference -Verbose:$VerbosePreference -Force:$Force
}

Write-Output "Done!"