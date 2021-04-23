#Requires -Version 3.0
#Make sure to fill in all the required variables before running the script
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Directory.Read.All (hard-requirement for oauth2PermissionGrants, covers everything else needed)

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
        $OAuthperm["[" + $resID + $userId + "]"] += ($oauth2PermissionGrant.Scope.Split(" ") -join ",")
    }
}

function Get-ServicePrincipalRoleById {

    Param(
    #Service principal object
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$spID)

    #check if we've already collected this SP data
    #do we need anything other than AppRoles? add a $select statement...
    if (!$SPPerm[$spID]) {
        $res = Invoke-WebRequest -Method Get -Uri "https://graph.microsoft.com/beta/servicePrincipals/$spID" -Headers $authHeader -Verbose:$VerbosePreference
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
        $res = Invoke-WebRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users/$($objectID)?`$select=UserPrincipalName" -Headers $authHeader -Verbose:$VerbosePreference
        $SPusers[$objectID] = ($res.Content | ConvertFrom-Json).UserPrincipalName
    }
    return $SPusers[$objectID]
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

#Get the list of Service principal objects within the tenant.
#Filter out any "built-in" service principals. Remove the filter if you want to include them.
#Only /beta returns publisherName currently
$SPs = @()

$uri = "https://graph.microsoft.com/beta/servicePrincipals?`$top=999&`$filter=tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp')"
#using the list endpoint returns empty appRoles?!?! Do per-SP query later on...
do {
    $result = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeader -Verbose:$VerbosePreference
    $uri = ($result.Content | ConvertFrom-Json).'@odata.nextLink'

    #If we are getting multiple pages, best add some delay to avoid throttling
    Start-Sleep -Milliseconds 500
    $SPs += ($result.Content | ConvertFrom-Json).Value
} while ($uri)

$SPperm = @{} #hash-table to store data for app roles and stuff
$SPusers = @{} #hash-table to store data for users assigned delegate permissions and stuff
$Output = @() #output variable
$i=0; $count = 1; $PercentComplete = 0;

foreach ($SP in $SPs) {
    #Progress message
    $ActivityMessage = "Retrieving data for service principal $($SP.DisplayName). Please wait..."
    $StatusMessage = ("Processing service principal {0} of {1}: {2}" -f $count, @($SPs).count, $SP.id)
    $PercentComplete = ($count / @($SPs).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #simple anti-throttling control
    Start-Sleep -Milliseconds 500
    Write-Verbose "Processing service principal $($SP.id)..."

    #Check for appRoleAssignments (application permissions)
    $appRoleAssignments = @()
    $res = Invoke-WebRequest -Method Get -Uri "https://graph.microsoft.com/beta/servicePrincipals/$($sp.id)/appRoleAssignments" -Headers $authHeader -Verbose:$VerbosePreference
    $appRoleAssignments = ($res.Content | ConvertFrom-Json).Value

    $OAuthperm = @{};
    $assignedto = @();$resID = $null; $userId = $null;
    $objPermissions = New-Object PSObject

    #prepare the output object
    $i++;Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Number" -Value $i
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Application Name" -Value $SP.appDisplayName
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "ApplicationId" -Value $SP.AppId
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Publisher" -Value $SP.PublisherName
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Verified" -Value (&{if ($SP.verifiedPublisher.verifiedPublisherId) {$SP.verifiedPublisher.displayName} else {"Not verified"}})
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Homepage" -Value $SP.Homepage
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "SP name" -Value $SP.displayName
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "ObjectId" -Value $SP.id
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Created on" -Value $SP.createdDateTime
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Enabled" -Value $SP.AccountEnabled

    #process application permissions entries
    if (!$appRoleAssignments) { Write-Verbose "No application permissions to report on for SP $($SP.id), skipping..." }
    else {
        Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Last modified" -Value ($appRoleAssignments.CreationTimestamp | select -Unique | sort -Descending | select -First 1)
    
        parse-AppPermissions $appRoleAssignments
        Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Permissions (application)" -Value (($OAuthperm.GetEnumerator() | % { "$($_.Name):$($_.Value.ToString().TrimStart(','))"}) -join ";")
        Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Authorized By (application)" -Value "An administrator (application permissions)"
    }
    

    #Check for oauth2PermissionGrants (delegate permissions)
    #Use /beta here, as /v1.0 does not return expiryTime
    $oauth2PermissionGrants = @()
    $res = Invoke-WebRequest -Method Get -Uri "https://graph.microsoft.com/beta/servicePrincipals/$($sp.id)/oauth2PermissionGrants" -Headers $authHeader -Verbose:$VerbosePreference
    $oauth2PermissionGrants = ($res.Content | ConvertFrom-Json).Value
    
    $OAuthperm = @{};
    $assignedto = @();$resID = $null; $userId = $null;

    #process delegate permissions entries
    if (!$oauth2PermissionGrants) { Write-Verbose "No delegate permissions to report on for SP $($SP.id), skipping..." }
    else {
        parse-DelegatePermissions $oauth2PermissionGrants
        Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Permissions (delegate)" -Value (($OAuthperm.GetEnumerator() | % { "$($_.Name):$($_.Value.ToString().TrimStart(','))"}) -join ";")

        Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Valid until (delegate)" -Value ($oauth2PermissionGrants.ExpiryTime | select -Unique | sort -Descending | select -First 1)
        
        if (($oauth2PermissionGrants.ConsentType | select -Unique) -eq "AllPrincipals") { $assignedto += "All users (admin consent)" }
        $assignedto +=  @($OAuthperm.Keys) | % {if ($_ -match "\((.*@.*)\)") {$Matches[1]}}
        Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Authorized By (delegate)" -Value (($assignedto | select -Unique) -join ",")
    }

    $Output += $objPermissions
}

#Export the result to CSV file
$Output | select 'Application name', 'ApplicationId', 'Publisher', 'Verified', 'Homepage', 'SP name', 'ObjectId', 'Created on', 'Enabled', 'Last modified', 'Permissions (application)', 'Authorized By (application)',`
'Permissions (delegate)', 'Valid until (delegate)', 'Authorized By (delegate)' | Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_GraphAppInventory.csv"