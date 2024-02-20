#Requires -Version 3.0
#The script requires the following permissions:
#    Directory.Read.All (hard-requirement for oauth2PermissionGrants, covers everything else needed)
#    CustomSecAttributeAssignment.Read.All (optional, needed to retrieve custom security attributes)
#    AuditLog.Read.All (optional, needed to retrieve Sign-in stats)
#    Reports.Read.All (optional, needed to retrieve Sign-in summary stats)
#    CrossTenantInformation.ReadBasic.All (optional, needed to retrieve owner organization info)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/5922/reporting-on-entra-id-integrated-applications-service-principals-and-their-permissions

[CmdletBinding(SupportsShouldProcess)] #Make sure we can use -WhatIf and -Verbose
Param([switch]$IncludeBuiltin=$false, [switch]$IncludeOwnerOrg=$false, [switch]$IncludeCSA=$false, [switch]$IncludeSignInStats=$false)

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
        $SPPerm[$spID] = Get-MgBetaServicePrincipal -ServicePrincipalId $spID -Verbose:$false -ErrorAction Stop
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
        #$res = Invoke-WebRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users/$($objectID)?`$select=UserPrincipalName" -Headers $authHeader -Verbose:$false
        $SPusers[$objectID] = (Get-MgBetaUser -UserId $objectID -Property UserPrincipalName).UserPrincipalName
    }
    return $SPusers[$objectID]
}

function parse-CustomSecurityAttributes {

    Param(
    #CustomSecurityAttributes object
    [Parameter(Mandatory=$true)]$customSecurityAttributes)

    $out = @();
    foreach ($CSAset in $customSecurityAttributes.GetEnumerator()) {
        $Name = $CSAset.Key; $attr = @()
        foreach ($prop in $CSAset.Value.GetEnumerator()) {
            if ($prop.Key -eq '@odata.type') { continue }
            $key = $prop.Key
            $value = $prop.Value
            $attr += "$($key):$Value"
        }
        $out += "[$Name]$($attr -join "|")"
    }
    return ($out -join ";")
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

function parse-SPSummaryStats {

    Param(
        #Report object
        [Parameter(Mandatory=$true)]$SPSignInSummary)

    foreach ($SPSignInStat in $SPSignInSummary) {
        if (!$SPSummaryStats[$SPSignInStat.Id]) {
            $SPSummaryStats[$SPSignInStat.Id] = @{
                "SignInSuccessCount" = $SPSignInStat.successfulSignInCount
                "SignInFailureCount" = $SPSignInStat.failedSignInCount
            }
        }
    }
    #return $SPSummaryStats
}

function Get-SPOwnerOrg {

    Param(
    #Service principal object
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$ID)

    #check if we've already collected this SP data
    if (!$SPOwnerOrg[$ID]) {
        Write-Verbose "Retrieving owner org info..."
        try {
            $SPOwnerOrg[$ID] = (Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/tenantRelationships/findTenantInformationByTenantId(tenantId=`'$($ID)`')").defaultDomainName
        }
        catch { Write-Verbose "Failed to retrieve owner org info for SP $($SP.id) ..."; return }
    }
    return $SPOwnerOrg[$ID]
}

#==========================================================================
#Main script starts here
#==========================================================================

#Determine the required scopes, based on the parameters passed to the script
$RequiredScopes = switch ($PSBoundParameters.Keys) {
    "IncludeBuiltin" { "Directory.Read.All" }
    "IncludeOwnerOrg" { "CrossTenantInformation.ReadBasic.All" }
    "IncludeCSA" { "CustomSecAttributeAssignment.Read.All" }
    "IncludeSignInStats" { "AuditLog.Read.All", "Reports.Read.All" }
    Default { "Directory.Read.All" }
}

Write-Verbose "Connecting to Graph API..."
Import-Module Microsoft.Graph.Beta.Applications -Verbose:$false -ErrorAction Stop
try {
    Connect-MgGraph -Scopes $RequiredScopes -verbose:$false -ErrorAction Stop -NoWelcome
}
catch { throw $_ }

#Check if we have all the required permissions
$CurrentScopes = (Get-MgContext).Scopes
if ($RequiredScopes | ? {$_ -notin $CurrentScopes }) { Write-Error "The access token does not have the required permissions, rerun the script and consent to the missing scopes!" -ErrorAction Stop }

#Make sure we include Custom security attributes in the report, if requested
if ($IncludeCSA) { $properties = "appDisplayName,appId,appOwnerOrganizationId,displayName,id,servicePrincipalType,createdDateTime,AccountEnabled,passwordCredentials,keyCredentials,tokenEncryptionKeyId,verifiedPublisher,Homepage,PublisherName,tags,customSecurityAttributes" }
else { $properties = "appDisplayName,appId,appOwnerOrganizationId,displayName,id,servicePrincipalType,createdDateTime,AccountEnabled,passwordCredentials,keyCredentials,tokenEncryptionKeyId,verifiedPublisher,Homepage,PublisherName,tags" }

#Get the list of Service principal objects within the tenant.
#Only /beta returns publisherName currently
$SPs = @()

Write-Verbose "Retrieving list of service principals..."
if (!$IncludeBuiltin) { $SPs = Get-MgBetaServicePrincipal -All -Filter "tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp')" -Property $properties -ErrorAction Stop -Verbose:$false }
else { $SPs = Get-MgBetaServicePrincipal -All -Property $properties -ErrorAction Stop -Verbose:$false }

#Gather sign-in stats for the service principals, if requested
if ($IncludeSignInStats) {
    Write-Verbose "Retrieving sign-in stats for service principals..."

    $SPSignInStats = @()
    $SPSignInStats = Get-MgBetaReportServicePrincipalSignInActivity -All -ErrorAction Stop -Verbose:$false

    $SPStats = @{} #hash-table to store sign-in stats data
    if ($SPSignInStats) { parse-SPSignInStats $SPSignInStats }

    Write-Verbose "Retrieving sign-in summary for service principals..."

    $SPSignInSummary = @()
    $SPSignInSummary = Get-MgBetaReportAzureAdApplicationSignInSummary -Period D30 -ErrorAction Stop -Verbose:$false

    $SPSummaryStats = @{} #hash-table to store sign-in stats data
    if ($SPSignInSummary) { parse-SPSummaryStats $SPSignInSummary }
}

#Set up some variables
$SPperm = @{} #hash-table to store data for app roles and stuff
$SPusers = @{} #hash-table to store data for users assigned delegate permissions and stuff
if ($IncludeOwnerOrg) { $SPOwnerOrg = @{} } #hash-table to store data for SP owner organization
$output = [System.Collections.Generic.List[Object]]::new() #output variable
$i=0; $count = 1; $PercentComplete = 0;

#Process the list of service principals
foreach ($SP in $SPs) {
    #Progress message
    $ActivityMessage = "Retrieving data for service principal $($SP.DisplayName). Please wait..."
    $StatusMessage = ("Processing service principal {0} of {1}: {2}" -f $count, @($SPs).count, $SP.id)
    $PercentComplete = ($count / @($SPs).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #simple anti-throttling control
    #Start-Sleep -Milliseconds 500 #The SDK should handle throttling?
    Write-Verbose "Processing service principal $($SP.id)..."

    #Get owners info. We do not use $expand, as it returns the full set of object properties
    Write-Verbose "Retrieving owners info..."
    $owners = @()
    $owners = Get-MgBetaServicePrincipalOwner -ServicePrincipalId $SP.id -Property id,userPrincipalName -All -ErrorAction Stop -Verbose:$false
    if ($owners) { $owners = $owners.userPrincipalName }

    #Include info about the SP owner organization
    if ($IncludeOwnerOrg) {
        if ($SP.appOwnerOrganizationId) { $ownerDomain = Get-SPOwnerOrg $SP.appOwnerOrganizationId }
        else { $ownerDomain = $null }
    }

    #Include information about group/directory role memberships. Cannot use /memberOf/microsoft.graph.directoryRole :(
    Write-Verbose "Retrieving group/directory role memberships..."
    $res = Get-MgBetaServicePrincipalMemberOf -All -ServicePrincipalId $SP.id -ErrorAction Stop -Verbose:$false
    $memberOfGroups = ($res.AdditionalProperties | ? {$_.'@odata.type' -eq "#microsoft.graph.group"}).displayName -join ";" #d is Case-sensitive!
    $memberOfRoles = ($res.AdditionalProperties | ? {$_.'@odata.type' -eq "#microsoft.graph.directoryRole"}).displayName -join ";" #d is Case-sensitive!

    #prepare the output object
    $i++;$objPermissions = [PSCustomObject][ordered]@{
        "Number" = $i
        "Application Name" = (&{if ($SP.appDisplayName) { $SP.appDisplayName } else { $null }}) #Apparently appDisplayName can be null
        "ApplicationId" = $SP.AppId
        "IsBuiltIn" = $SP.tags -notcontains "WindowsAzureActiveDirectoryIntegratedApp"
        "Publisher" = (&{if ($SP.PublisherName) { $SP.PublisherName } else { $null }})
        "Owned by org" = (&{if ($ownerDomain) { "$($SP.appOwnerOrganizationId) ($ownerDomain)" } else { $SP.appOwnerOrganizationId }}) #Apparently appOwnerOrganizationId can be null?
        "Verified" = (&{if ($SP.verifiedPublisher.verifiedPublisherId) { $SP.verifiedPublisher.displayName } else { "Not verified" }})
        "Homepage" = (&{if ($SP.Homepage) { $SP.Homepage } else { $null }})
        "SP name" = $SP.displayName
        "ObjectId" = $SP.id
        "Type" = $SP.servicePrincipalType
        "Created on" = (&{if ($SP.AdditionalProperties.createdDateTime) {(Get-Date($SP.AdditionalProperties.createdDateTime) -format g)} else { "N/A" }})
        "Enabled" = $SP.AccountEnabled
        "Owners" = (&{if ($owners) { $owners -join "," } else { $null }})
        "Member of (groups)" = $memberOfGroups
        "Member of (roles)" = $memberOfRoles
        "PasswordCreds" = (&{if ($SP.passwordCredentials) { $SP.passwordCredentials.keyId -join ";" } else { $null }})
        "KeyCreds" = (&{if ($SP.keyCredentials) { $SP.keyCredentials.keyId -join ";" } else { $null }})
        "TokenKey" = (&{if ($SP.tokenEncryptionKeyId) { $SP.tokenEncryptionKeyId } else { $null }})
        "Permissions (application)" = $null
        "Authorized By (application)" = $null
        "Last modified (application)" = $null
        "Permissions (delegate)" = $null
        "Authorized By (delegate)" = $null
        "Valid until (delegate)" = $null
    }

    #Include sign-in stats, if requested
    if ($IncludeSignInStats) {
        $objPermissions | Add-Member -MemberType NoteProperty -Name "Last sign-in" -Value (&{if ($SPStats[$SP.appId].LastSignIn) { (Get-Date($SPStats[$SP.appid].LastSignIn) -format g) } else { $null }})
        $objPermissions | Add-Member -MemberType NoteProperty -Name "Last delegate client sign-in" -Value (&{if ($SPStats[$SP.appid].LastDelegateClientSignIn) { (Get-Date($SPStats[$SP.appid].LastDelegateClientSignIn) -format g) } else { $null }})
        $objPermissions | Add-Member -MemberType NoteProperty -Name "Last delegate resource sign-in" -Value (&{if ($SPStats[$SP.appid].LastDelegateResourceSignIn) { (Get-Date($SPStats[$SP.appid].LastDelegateResourceSignIn) -format g) } else { $null }})
        $objPermissions | Add-Member -MemberType NoteProperty -Name "Last app client sign-in" -Value (&{if ($SPStats[$SP.appid].LastAppClientSignIn) { (Get-Date($SPStats[$SP.appid].LastAppClientSignIn) -format g) } else { $null }})
        $objPermissions | Add-Member -MemberType NoteProperty -Name "Last app resource sign-in" -Value (&{if ($SPStats[$SP.appid].LastAppResourceSignIn) { (Get-Date($SPStats[$SP.appid].LastAppResourceSignIn) -format g) } else { $null }})

        $objPermissions | Add-Member -MemberType NoteProperty -Name "Sign-in success count (30 days)" -Value (&{if ($SPSummaryStats[$SP.appid].SignInSuccessCount) { $SPSummaryStats[$SP.appid].SignInSuccessCount } else { $null }})
        $objPermissions | Add-Member -MemberType NoteProperty -Name "Sign-in failure count (30 days)" -Value (&{if ($SPSummaryStats[$SP.appid].SignInFailureCount) { $SPSummaryStats[$SP.appid].SignInFailureCount } else { $null }})
    }

    #Include Custom security attributes, if requested
    if ($IncludeCSA) {
        $objPermissions | Add-Member -MemberType NoteProperty -Name "CustomSecurityAttributes" -Value (&{if ($SP.customSecurityAttributes) { parse-CustomSecurityAttributes $SP.customSecurityAttributes.AdditionalProperties } else { $null }})
    }

    #Check for appRoleAssignments (application permissions)
    Write-Verbose "Retrieving application permissions..."
    try {
        $appRoleAssignments = @()
        $appRoleAssignments = Get-MgBetaServicePrincipalAppRoleAssignment -All -ServicePrincipalId $SP.id -ErrorAction Stop -Verbose:$false

        $OAuthperm = @{};
        $assignedto = @();$resID = $null; $userId = $null;

        #process application permissions entries
        if (!$appRoleAssignments) { Write-Verbose "No application permissions to report on for SP $($SP.id), skipping..." }
        else {
            $objPermissions.'Last modified (application)' = (Get-Date($appRoleAssignments.CreationTimestamp | select -Unique | sort -Descending | select -First 1) -format g)

            parse-AppPermissions $appRoleAssignments
            $objPermissions.'Permissions (application)' = (($OAuthperm.GetEnumerator()  | % { "$($_.Name):$($_.Value.ToString().TrimStart(','))"}) -join ";")
            $objPermissions.'Authorized By (application)' = "An administrator (application permissions)"
        }
    }
    catch { Write-Verbose "Failed to retrieve application permissions for SP $($SP.id) ..." }

    #Check for oauth2PermissionGrants (delegate permissions)
    #Use /beta here, as /v1.0 does not return expiryTime
    Write-Verbose "Retrieving delegate permissions..."
    try {
        $oauth2PermissionGrants = @()
        $oauth2PermissionGrants = Get-MgBetaServicePrincipalOAuth2PermissionGrant -All -ServicePrincipalId $SP.id -ErrorAction Stop -Verbose:$false

        $OAuthperm = @{};
        $assignedto = @();$resID = $null; $userId = $null;

        #process delegate permissions entries
        if (!$oauth2PermissionGrants) { Write-Verbose "No delegate permissions to report on for SP $($SP.id), skipping..." }
        else {
            parse-DelegatePermissions $oauth2PermissionGrants
            $objPermissions.'Permissions (delegate)' = (($OAuthperm.GetEnumerator() | % { "$($_.Name):$($_.Value.ToString().TrimStart(','))"}) -join ";")
            $objPermissions.'Valid until (delegate)' = (Get-Date($oauth2PermissionGrants.ExpiryTime | select -Unique | sort -Descending | select -First 1) -format g)

            if (($oauth2PermissionGrants.ConsentType | select -Unique) -eq "AllPrincipals") { $assignedto += "All users (admin consent)" }
            $assignedto +=  @($OAuthperm.Keys) | % {if ($_ -match "\((.*@.*)\)") {$Matches[1]}}
            $objPermissions.'Authorized By (delegate)' = (($assignedto | select -Unique) -join ",")
        }
    }
    catch { Write-Verbose "Failed to retrieve delegate permissions for SP $($SP.id) ..." }

    $output.Add($objPermissions)
}

#Export the result to CSV file
$output | select * -ExcludeProperty Number | Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_GraphAppInventory.csv"
Write-Verbose "Output exported to $($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_GraphAppInventory.csv"