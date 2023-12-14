#Requires -Version 3.0
#Make sure to fill in all the required variables before running the script
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Directory.Read.All (to cover assigned directory roles)
#    RoleManagement.Read.Directory (should cover both assigned and PIM-eligible roles)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/3958/generate-a-report-of-azure-ad-role-assignments-via-the-graph-api-or-powershell

[CmdletBinding()] #Make sure we can use -WhatIf and -Verbose
Param([switch]$IncludePIMEligibleAssignments) #Indicate whether to include PIM elibigle role assignments in the output. NOTE: Currently the RoleManagement.Read.Directory application permissions seems to be required!

#region Authentication
#We use the client credentials flow as an example. For production use, REPLACE the code below with your preferred auth method. NEVER STORE CREDENTIALS IN PLAIN TEXT!!!

#Variables to configure
$tenantID = "tenant.onmicrosoft.com" #your tenantID or tenant root domain
$appID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #the GUID of your app. For best result, use app with Directory.Read.All scope granted. For PIM use RoleManagement.Read.Directory
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
Write-Verbose "Authenticating..."
try { $tokenRequest = Invoke-WebRequest -Method Post -Uri $url -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -Verbose:$false -ErrorAction Stop }
catch { Write-Host "Unable to obtain access token, aborting..."; return }

$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

$authHeader = @{
   'Content-Type'='application\json'
   'Authorization'="Bearer $token"
}
#endregion Authentication

#region Roles
Write-Verbose "Collecting role assignments..."
#Use the /roleManagement/directory/roleAssignments endpoint to collect a list of all role assignments.
$roles = @()
$uri = 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?$expand=principal' #$expand=* is BROKEN

try {
    do {
        $result = Invoke-WebRequest -Uri $uri -Verbose:$false -ErrorAction Stop -Headers $authHeader
        $uri = $($result | ConvertFrom-Json).'@odata.nextLink'
        #If we are getting multiple pages, best add some delay to avoid throttling
        Start-Sleep -Milliseconds 500
        $roles += ($result | ConvertFrom-Json).Value
    } while ($uri)
}
catch { Write-Error "Unable to obtain role assignments, make sure the required permissions have been granted..."; return }

#fix to also fetch the roleDefinition
$uri = 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?$expand=roleDefinition' #$expand=* is BROKEN

$roles1 = @()
try {
    do {
        $result = Invoke-WebRequest -Uri $uri -Verbose:$false -ErrorAction Stop -Headers $authHeader
        $uri = $($result | ConvertFrom-Json).'@odata.nextLink'
        #If we are getting multiple pages, best add some delay to avoid throttling
        Start-Sleep -Milliseconds 500
        $roles1 += ($result | ConvertFrom-Json).Value
    } while ($uri)
}
catch { Write-Error "Unable to obtain role assignments, make sure the required permissions have been granted..."; return }

foreach ($role in $roles) { Add-Member -InputObject $role -MemberType NoteProperty -Name roleDefinition -Value ($roles1 | ? {$_.id -eq $role.id}).roleDefinition }

#process PIM eligible role assignments
if ($IncludePIMEligibleAssignments) {
    Write-Verbose "Collecting PIM eligible role assignments..."
    #$expand=* is BROKEN in /v1.0, so we need to use the beta endpoint
    $uri = 'https://graph.microsoft.com/beta/roleManagement/directory/roleEligibilitySchedules?$select=id,principalId,directoryScopeId,roleDefinitionId,status&$expand=*'

    try {
        do {
            $result = Invoke-WebRequest -Uri $uri -Verbose:$false -ErrorAction Stop -Headers $authHeader
            $uri = $($result | ConvertFrom-Json).'@odata.nextLink'
            #If we are getting multiple pages, best add some delay to avoid throttling
            Start-Sleep -Milliseconds 500
            $roles += ($result | ConvertFrom-Json).Value
        } while ($uri)
    }
    catch { Write-Host "Unable to obtain PIM eligible role assignments, make sure the required permissions have been granted..."; return }
}

if (!$roles) { Write-Verbose "No valid role assignments found, verify the required permissions have been granted?" }

Write-Verbose "A total of $($roles.count) role assignments were found, of which $(($roles | ? {$_.directoryScopeId -eq "/"}).Count) are tenant-wide and $(($roles | ? {$_.directoryScopeId -ne "/"}).Count) are AU-scoped. $(($roles | ? {!$_.status}).Count) roles are permanently assigned, you might want to address that!"
#endregion Roles

#region Output
Write-Verbose "Preparing the output..."
$report = @()
foreach ($role in $roles) {
    $reportLine=[ordered]@{
        "Principal" = switch ($role.principal.'@odata.type') {
            '#microsoft.graph.user' {$role.principal.userPrincipalName}
            '#microsoft.graph.servicePrincipal' {$role.principalId}
            '#microsoft.graph.group' {$role.principal.id}
        }
        "PrincipalDisplayName" = $role.principal.displayName
        "PrincipalType" = $role.principal.'@odata.type'.Split(".")[-1]
        "AssignedRole" = $role.roleDefinition.displayName
        "AssignedRoleScope" = $role.directoryScopeId
        "AssignmentType" = (&{if ($role.status -eq "Provisioned") {"Eligible"} else {"Permanent"}})
        "IsBuiltIn" = $role.roleDefinition.isBuiltIn
        "RoleTemplate" = $role.roleDefinition.templateId
    }
    $report += @([pscustomobject]$reportLine)
}
#endregion Output

#format and export
$report | sort PrincipalDisplayName | Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_AzureADRoleInventory.csv"
Write-Verbose "Report saved to ""$($pwd)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_AzureADRoleInventory.csv"""