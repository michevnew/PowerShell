#Requires -Version 7.1
#The script requires the following permissions:
#    Device.Read.All (required for /memberOf)
#    Group.Read.All (required)
#    Directory.Read.All (optional, if you don't care about least privilege and don't want to add multiple permissions)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/6444/reporting-on-group-membership-for-entra-id-devices-including-assigned-licenses

param([string[]]$DeviceList)

#Set the authentication details
$tenantID = "tenant.onmicrosoft.com" #your tenantID or tenant root domain
$appID = "12345678-1234-1234-1234-1234567890AB" #the GUID of your app. For best result, use app with Directory.Read.All scope granted
$client_secret = "XXXXXXXXXXXXXXXXXXX" #client secret for the app

$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $client_secret
    grant_type    = "client_credentials"
}

#Get a token
$authenticationResult = Invoke-WebRequest -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -ErrorAction Stop -Verbose:$false
$token = ($authenticationResult.Content | ConvertFrom-Json).access_token
$authHeader = @{'Authorization'="Bearer $token"}

$GraphDevices = @()
#If a list of devices was provided via the -DeviceList parameter, only run against a set of devices
if ($DeviceList) {
    Write-Verbose "Running the script against the provided list of devices..."
    foreach ($device in $DeviceList) {
        try {
            #Make sure the device entry is valid, check against both Id and DeviceId
            $filter = "deviceId eq `'$device`' or id eq `'$device`'"
            $uri = "https://graph.microsoft.com/v1.0/devices?`$filter=$filter&`$select=id,deviceId,displayName"
            $dres = Invoke-WebRequest -Headers $authHeader -Uri $uri -UseBasicParsing -ErrorAction Stop -Verbose:$false
            $dres = (${dres}?.Content | ConvertFrom-Json).Value

            $GraphDevices += $dres
        }
        catch {
            Write-Verbose "No match found for provided device entry $device, skipping..."
            continue
        }
    }
}
else {
    #Get a list of all devices, make sure to handle multiple pages
    Write-Verbose "Running the script against all devices in the tenant..."

    $uri = "https://graph.microsoft.com/v1.0/devices?`$select=id,deviceId,displayName"
    do {
        $result = Invoke-WebRequest -Headers $authHeader -Uri $uri -UseBasicParsing -ErrorAction Stop -Verbose:$false
        $uri = $result.'@odata.nextLink'
        #If we are getting multiple pages, best add some delay to avoid throttling
        Start-Sleep -Milliseconds 500
        $GraphDevices += (${result}?.Content | ConvertFrom-Json).Value
    } while ($uri)
}

#Cycle over each device and fetch group membership
$output = [System.Collections.Generic.List[Object]]::new() #output variable
$count = 1; $PercentComplete = 0;
foreach ($device in $GraphDevices) {
    #Simple progress indicator
    $ActivityMessage = "Retrieving data for user $($device.displayName). Please wait..."
    $StatusMessage = ("Processing user {0} of {1}: {2}" -f $count, @($GraphDevices).count, $device.id)
    $PercentComplete = ($count / @($GraphDevices).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Get the list of groups for the device
    Write-Verbose "Fetching transitive group membership for device $($device.displayName)..."
    $uri = "https://graph.microsoft.com/v1.0/devices/$($device.id)/transitivememberof/microsoft.graph.group?`$select=Id,displayName,groupTypes,securityEnabled,mailEnabled,onPremisesSyncEnabled,Visibility,assignedLicenses"
    $DeviceGroups = Invoke-WebRequest -Headers $authHeader -Uri $uri -UseBasicParsing -ErrorAction Stop -Verbose:$false
    $rGroups = ($DeviceGroups.Content | ConvertFrom-Json).Value

    #If no groups returned for the device, still write to output
    if (!$rGroups) {
        #prepare the output
        $dInfo = [PSCustomObject][ordered]@{
            "Id" = $device.id
            "DeviceId" = $device.deviceId
            "Display Name" = $device.DisplayName
            "Group" = "N/A"
            "GroupName" = $null
            "GroupType" = $null
            "GroupSynced" = $null
            "GroupMembershipType" = $null
            "GroupRule" = $null
            "HiddenMembership" = $null
            "GroupLicenses" = $null
        }

        $output.Add($dInfo)
        continue
    }

    #For each group returned, output the relevant details
    foreach ($group in $rGroups) {
        #prepare the output
        $dInfo = [PSCustomObject][ordered]@{
            "Id" = $device.id
            "DeviceId" = $device.deviceId
            "Display Name" = $device.DisplayName
            "Group" = $group.Id
            "GroupName" = $group.DisplayName
            "GroupType" = (&{
                if ($Group.groupTypes -eq "Unified" -and $Group.securityEnabled) { "Microsoft 365 (security-enabled)" }
                elseif ($Group.groupTypes -eq "Unified" -and !$Group.securityEnabled) { "Microsoft 365" }
                elseif (!($Group.groupTypes -eq "Unified") -and $Group.securityEnabled -and $Group.mailEnabled) { "Mail-enabled Security" }
                elseif (!($Group.groupTypes -eq "Unified") -and $Group.securityEnabled) { "Entra ID Security" }
                elseif (!($Group.groupTypes -eq "Unified") -and $Group.mailEnabled) { "Distribution" }
                else { "N/A" }
            })
            "GroupSynced" = ($group.onPremisesSyncEnabled) ? "Yes" : "No"
            "GroupMembershipType" = ($group.groupTypes -eq "Dynamic") ? "Dynamic" : "Assigned"
            "GroupMembershipRule" = ($group.groupTypes -eq "Dynamic") ? ($group.membershipRule) : $null
            "HiddenMembership" = ($group.Visibility) ? "Yes" : "No"
            "GroupLicenses" = ($group.assignedLicenses) ? ($group.assignedLicenses.SkuId -join ",") : $null
        }

        $output.Add($dInfo)
    }
}

#Finally, export to CSV
$output | select * | Export-CSV -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_DeviceMemberOf.csv" -NoTypeInformation -Encoding UTF8 -UseCulture