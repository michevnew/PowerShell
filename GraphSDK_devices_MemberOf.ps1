#Requires -Version 7.1
#The script requires the following permissions:
#    Device.Read.All (required for /memberOf)
#    Group.Read.All (required)
#    Directory.Read.All (optional, if you don't care about least privilege and don't want to add multiple permissions)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/6444/reporting-on-group-membership-for-entra-id-devices-including-assigned-licenses

param([string[]]$DeviceList)

#Connect to the Graph with required permissions
Write-Verbose "Connecting to Graph API..."
try {
    Connect-MGgraph -Scopes "Device.Read.All,Group.Read.All" -NoWelcome -ErrorAction Stop -Verbose:$false
}
catch { throw $_ }

$Devices = @()
#If a list of devices was provided via the -DeviceList parameter, only run against a set of devices
if ($DeviceList) {
    Write-Verbose "Running the script against the provided list of devices..."
    foreach ($device in $DeviceList) {
        try {
            #Make sure the device entry is valid, check against both Id and DeviceId
            $filter = "deviceId eq `'$device`' or id eq `'$device`'"
            $dres = Get-MgDevice -Filter $filter -Property id,deviceId,displayName -ErrorAction Stop -Verbose:$false

            $Devices += $dres
        }
        catch {
            Write-Verbose "No match found for provided device entry $device, skipping..."
            continue
        }
    }
}
else {
    #Get the list of all Entra ID device objects within the tenant.
    Write-Verbose "Running the script against all devices in the tenant..."

    $Devices = Get-MgDevice -All -Property id,deviceId,displayName -ErrorAction Stop
}

#Cycle over each device and fetch group membership
$output = [System.Collections.Generic.List[Object]]::new() #output variable
$count = 1; $PercentComplete = 0;
foreach ($d in $Devices) {
    #Progress message
    $ActivityMessage = "Retrieving data for device $($d.displayName). Please wait..."
    $StatusMessage = ("Processing device object {0} of {1}: {2}" -f $count, @($Devices).count, $u.id)
    $PercentComplete = ($count / @($Devices).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Get the list of groups for the device
    Write-Verbose "Fetching transitive group membership for device $($d.displayName)..."
    $rGroups = Get-MgDeviceTransitiveMemberOfAsGroup -DeviceId $d.id -Property Id,displayName,groupTypes,securityEnabled,mailEnabled,onPremisesSyncEnabled,Visibility,assignedLicenses -ErrorAction Stop

    #If no groups returned for the device, still write to output
    if (!$rGroups) {
        #prepare the output
        $dInfo = [PSCustomObject][ordered]@{
            "Id" = $d.id
            "DeviceId" = $d.deviceId
            "Display Name" = $d.DisplayName
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
            "Id" = $d.id
            "DeviceId" = $d.deviceId
            "Display Name" = $d.DisplayName
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