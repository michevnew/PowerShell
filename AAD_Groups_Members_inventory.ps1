#Requires -Version 3.0
#Make sure to fill in all the required variables before running the script
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Directory.Read.All

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/4357/report-on-azure-ad-group-members-via-the-graph-api

param([string[]]$GroupList,[switch]$TransitiveMembership=$false)
[CmdletBinding()] #Make sure we can use -Verbose

#region Authentication
#We use the client credentials flow as an example. For production use, REPLACE the code below wiht your preferred auth method. NEVER STORE CREDENTIALS IN PLAIN TEXT!!!

#Variables to configure
$tenantID = "tenant.onmicrosoft.com" #your tenantID or tenant root domain
$appID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #the GUID of your app. For best result, use app with Directory.Read.All scope granted.
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
try {
    $tokenRequest = Invoke-WebRequest -Method Post -Uri $url -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -ErrorAction Stop
    $token = ($tokenRequest.Content | ConvertFrom-Json).access_token

    $authHeader = @{
       'Content-Type'='application\json'
       'Authorization'="Bearer $token"
    }
}
catch { Write-Error "Unable to obtain access token, aborting..." -ErrorAction Stop; return }
#endregion Authentication


#region Groups
$Groups = @()

#If a list of groups was provided via the -GroupList parameter, only run against a set of groups
if ($GroupList) {
    Write-Verbose "Running the script against the provided list of groups..."
    foreach ($group in $GroupList) {
        try {
            $uri = "https://graph.microsoft.com/v1.0/groups/$($group)?`$select=id,displayName,groupTypes,securityEnabled,mailEnabled,membershipRule,isAssignableToRole,mail,assignedLicenses&`$expand=owners(`$select=userPrincipalName)"
            $res = Invoke-WebRequest -Method Get -Headers $authHeader -Uri $uri -ErrorAction Stop -Verbose:$VerbosePreference
            $gres = ($res.Content | ConvertFrom-Json)

            $Groups += $gres
        }
        catch {
            Write-Verbose "No match found for provided group entry $group, skipping..."
            continue
        }
    }
    Write-Verbose "The following list of groups will be used: $($Groups.displayName -join ",")"
}
else {
    #Get the list of all user objects within the tenant.
    Write-Verbose "Running the script against all groups in the tenant..."

    $uri = "https://graph.microsoft.com/v1.0/groups?`$top=999&`$select=id,displayName,groupTypes,securityEnabled,mailEnabled,membershipRule,isAssignableToRole,mail,assignedLicenses&`$expand=owners(`$select=userPrincipalName)"
    do {
        $result = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeader -Verbose:$VerbosePreference
        $uri = ($result.Content | ConvertFrom-Json).'@odata.nextLink'

        #If we are getting multiple pages, best add some delay to avoid throttling
        Start-Sleep -Milliseconds 500
        $Groups += ($result.Content | ConvertFrom-Json).Value
    } while ($uri)
}
#endregion Groups

#region GroupMembership
#Cycle over each group and fetch group membership
$output = [System.Collections.Generic.List[Object]]::new() #output variable for expanded CSV (one line per member)
$output2 = [System.Collections.Generic.List[Object]]::new() #output variable for summary CSV (one line per group)
$count = 1; $PercentComplete = 0;
foreach ($g in $Groups) {
    #Progress message
    $ActivityMessage = "Retrieving data for group $($g.displayName). Please wait..."
    $StatusMessage = ("Processing group object {0} of {1}: {2}" -f $count, @($Groups).count, $g.id)
    $PercentComplete = ($count / @($Groups).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Simple anti-throttling control
    Start-Sleep -Milliseconds 200

    #Set generic group properties to avoid re-evaluating them
    $g | Add-Member -MemberType NoteProperty -Name GroupType -Value (&{
        if ($g.groupTypes -eq "Unified" -and $g.securityEnabled) { "Microsoft 365 (security-enabled)" }
        elseif ($g.groupTypes -eq "Unified" -and !$g.securityEnabled) { "Microsoft 365" }
        elseif (!($g.groupTypes -eq "Unified") -and $g.securityEnabled -and $g.mailEnabled) { "Mail-enabled Security" }
        elseif (!($g.groupTypes -eq "Unified") -and $g.securityEnabled) { "Azure AD Security" }
        elseif (!($g.groupTypes -eq "Unified") -and $g.mailEnabled) { "Distribution" }
        else { "N/A" }
    })

    #Prepare the query depending on the type of membership we are interested in
    if ($TransitiveMembership) { $QueryType = "transitiveMembers" } else { $QueryType = "members" }

    #Obtain the list of members, taking into account the desired query type and pagination
    Write-Verbose "Processing single group entry $($g.displayName) with $QueryType query..."
    $gMembers = @()

    #We use /beta here, as /v1.0 does not return service principal objects yet
    $uri = "https://graph.microsoft.com/beta/groups/$($g.id)/$($QueryType)?`$top=999&`$select=id,displayName,mailEnabled,securityEnabled,membershipRule,mail,isAssignableToRole,groupTypes,userPrincipalName,userType,deviceId"
    do {
        $result = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeader -Verbose:$VerbosePreference -ErrorAction Stop
        $uri = ($result.Content | ConvertFrom-Json).'@odata.nextLink'
        $gMembers += ($result.Content | ConvertFrom-Json).Value
    } while ($uri)


    #prepare the output for the expanded CSV
    $uInfo = [PSCustomObject][ordered]@{
        "Id" = $g.id
        "DisplayName" = $g.displayName
        "GroupType" = $g.GroupType
        "Owners" = (&{if ($g.owners) { $($g.Owners.UserPrincipalName -join ",") } else { "N/A" }})
        "PrimarySmtpAddress" = (&{if ($g.mail) { $g.mail } else { "N/A" }})
        "RoleAssignable" = (&{if ($g.isAssignableToRole) { $true } else { $false }})
        "AssignedLicenses" = (&{if ($g.assignedLicenses) { ($g.assignedLicenses.SkuId -join ",") } else { $false }})
        "MembershipType" = (&{if ($g.membershipRule) { "Dynamic" } else { "Assigned" }})
        "MembershipRule" = (&{if ($g.membershipRule) { $g.membershipRule } else { "N/A" }})
        "MemberId" = $null
        "MemberDisplayName" = $null
        "MemberType" = $null
        "MemberMail" = $null
        "MemberIdentifier" = $null
    }
    if (!$gMembers) { $output.Add($uInfo) } #add the "empty" value

    #For each member returned, include the relevant details
    $j = 0;$usermembers = @();$groupmembers = @();$devicemembers = @();$contactmembers = @();$SPmembers = @();
    foreach ($Member in $gMembers) {
        $j++ #cheap member count that accounts for unhandled member types
        $uInfo.MemberId = $Member.Id
        $uInfo.MemberDisplayName = $Member.displayName
        if ($Member.userType -eq "Guest") { $uInfo.MemberType = "Guest" }
        else { $uInfo.MemberType = $Member.'@odata.type'.Split(".")[-1] }
        $uInfo.MemberMail = (&{if ($Member.mail) { $Member.mail } else { "N/A" }})

        #add to the lists used by the summary CSV file
        switch ($Member.'@odata.type'.Split(".")[-1]) {
            "user" { $uInfo.MemberIdentifier = $Member.userPrincipalName; $usermembers += $Member.UserPrincipalName }
            "group" { $uInfo.MemberIdentifier = $Member.id; $groupmembers += $Member.id }
            "device" { $uInfo.MemberIdentifier = $Member.deviceId; $devicemembers += $Member.deviceId }
            "orgContact" { $uInfo.MemberIdentifier = $Member.Mail; $contactmembers += $Member.Mail }
            "servicePrincipal" { $uInfo.MemberIdentifier = $Member.id; $SPmembers += $Member.id }
            default { Write-Verbose "Unhandled scenario" }
        }

        $output.Add($uInfo.psobject.Copy()) #!
    }

    #prepare the output for summary CSV
    $uInfo2 = [PSCustomObject][ordered]@{
        "Id" = $g.id
        "DisplayName" = $g.displayName
        "GroupType" = $g.GroupType
        "Owners" = (&{if ($g.owners) { $($g.Owners.UserPrincipalName -join ",") } else { "N/A" }})
        "HasNestedGroups" = &{If ($groupmembers) { $groupmembers.Count } else {$false} }
        "PrimarySmtpAddress" = (&{if ($g.mail) { $g.mail } else { "N/A" }})
        "RoleAssignable" = (&{if ($g.isAssignableToRole) { $true } else { $false }})
        "MembershipType" = (&{if ($g.membershipRule) { "Dynamic" } else { "Assigned" }})
        "MembershipRule" = (&{if ($g.membershipRule) { $g.membershipRule } else { "N/A" }})
        "AssignedLicenses" = (&{if ($g.assignedLicenses) { ($g.assignedLicenses.SkuId -join ",") } else { $false }})
        "MemberCountTotal" = $j
        "UserMemberCount" = $usermembers.count
        "GroupMemberCount" = $groupmembers.count
        "DeviceMemberCount" = $devicemembers.count
        "ContactMemberCount" = $contactmembers.count
        "SPMemberCount" = $SPmembers.count
        "UserMembers" = &{If ($usermembers) { $usermembers -join ","}}
        "GroupMembers" = &{If ($groupmembers) { $groupmembers -join ","}}
        "DeviceMembers" = &{If ($devicemembers) { $devicemembers -join ","}}
        "ContactMembers" = &{If ($contactmembers) { $contactmembers -join ","}}
        "SPmembers" = &{If ($SPmembers) { $SPmembers -join ","}}
    }

    $output2.Add($uInfo2)
}
#endregion GroupMembership

#Finally, export to CSV
Write-Verbose "Writing output to CSV files..."
$output | select * | Export-CSV -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_AADGroupMembersExpanded.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
$output2 | select * | Export-CSV -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_AADGroupMembers.csv" -NoTypeInformation -Encoding UTF8 -UseCulture