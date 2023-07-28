#Requires -Version 3.0
#Make sure to fill in all the required variables before running the script
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Directory.Read.All

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/4235/reporting-on-users-group-membership-in-azure-ad

param([string[]]$UserList,[switch]$TransitiveMembership=$false)

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


#region Users
$Users = @()

#If a list of users was provided via the -UserList parameter, only run against a set of users
if ($UserList) {
    Write-Verbose "Running the script against the provided list of users..."
    foreach ($user in $UserList) {
        try {
            $uri = "https://graph.microsoft.com/v1.0/users/$($user)?`$select=id,userPrincipalName"
            $res = Invoke-WebRequest -Headers $authHeader -Uri $uri -ErrorAction Stop
            $ures = ($res.Content | ConvertFrom-Json) | select Id,userPrincipalName

            $Users += $ures
        }
        catch {
            Write-Verbose "No match found for provided user entry $user, skipping..."
            continue
        }
    }
}
else {
    #Get the list of all user objects within the tenant.
    Write-Verbose "Running the script against all users in the tenant..."

    $uri = "https://graph.microsoft.com/v1.0/users?`$top=999&`$select=id,userPrincipalName"
    do {
        $result = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeader -Verbose:$VerbosePreference
        $uri = ($result.Content | ConvertFrom-Json).'@odata.nextLink'

        #If we are getting multiple pages, best add some delay to avoid throttling
        Start-Sleep -Milliseconds 500
        $Users += ($result.Content | ConvertFrom-Json).Value
    } while ($uri)
}
#endregion Users

#region GroupMembership
#Cycle over each user and fetch group membership
$output = [System.Collections.Generic.List[Object]]::new() #output variable
$count = 1; $PercentComplete = 0;
foreach ($u in $Users) {
    #Progress message
    $ActivityMessage = "Retrieving data for user $($u.userPrincipalName). Please wait..."
    $StatusMessage = ("Processing user object {0} of {1}: {2}" -f $count, @($Users).count, $u.id)
    $PercentComplete = ($count / @($Users).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Simple anti-throttling control
    Start-Sleep -Milliseconds 100
    Write-Verbose "Processing user object $($u.userPrincipalName)..."

    #Prepare the query depending on the type of membership we are interested in
    if ($TransitiveMembership) { $QueryType = "transitiveMemberOf" } else { $QueryType = "memberOf" }

    $uri = "https://graph.microsoft.com/v1.0/users/$($u.id)/$QueryType/microsoft.graph.group?`$select=id,displayName,mailEnabled,securityEnabled,membershipRule,mail,isAssignableToRole,groupTypes"
    $res = Invoke-WebRequest -Headers $authHeader -Uri $uri -ErrorAction Stop
    $uGroups = ($res.Content | ConvertFrom-Json).Value

    #If no group objects returned for the user, still write to output
    if (!$uGroups) {
        #prepare the output
        $uInfo = [PSCustomObject][ordered]@{
            "Id" = $u.id
            "UPN" = $u.userPrincipalName
            "Group" = "N/A"
            "GroupName" = $null
            "Mail" = $null
            "RoleAssignable" = $null
            "GroupType" = $null
            "MembershipType" = $null
            "GroupRule" = $null
        }

        $output.Add($uInfo)
        continue
    }

    #For each group returned, output the relevant details
    foreach ($Group in $uGroups) {
        #prepare the output
        $uInfo = [PSCustomObject][ordered]@{
            "Id" = $u.id
            "UPN" = $u.userPrincipalName
            "Group" = $Group.Id
            "GroupName" = $Group.displayName
            "Mail" = (&{if ($Group.mail) { $Group.mail } else { "N/A" }})
            "RoleAssignable" = (&{if ($Group.isAssignableToRole) { $true } else { $false }})
            "GroupType" = (&{
                if ($Group.groupTypes -eq "Unified" -and $Group.securityEnabled) { "Microsoft 365 (security-enabled)" }
                elseif ($Group.groupTypes -eq "Unified" -and !$Group.securityEnabled) { "Microsoft 365" }
                elseif (!($Group.groupTypes -eq "Unified") -and $Group.securityEnabled -and $Group.mailEnabled) { "Mail-enabled Security" }
                elseif (!($Group.groupTypes -eq "Unified") -and $Group.securityEnabled) { "Azure AD Security" }
                elseif (!($Group.groupTypes -eq "Unified") -and $Group.mailEnabled) { "Distribution" }
                else { "N/A" }
            }) #triple-check this
            "MembershipType" = (&{if ($Group.membershipRule) { "Dynamic" } else { "Assigned" }})
            "GroupRule" = (&{if ($Group.membershipRule) { $Group.membershipRule } else { "N/A" }})
        }

        $output.Add($uInfo)
    }
}
#endregion GroupMembership

#Finally, export to CSV
$output | select * #| Export-CSV -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_AADGroupMembership.csv" -NoTypeInformation -Encoding UTF8 -UseCulture