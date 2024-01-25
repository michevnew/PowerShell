#Requires -Version 3.0
#The script requires the following permissions:
#    BitLockerKey.Read.All (required)
#    Device.Read.All (optional, needed to retrieve device details)
#    User.ReadBasic.All (optional, needed to retrieve device owner's UPN)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/5950/reporting-on-bitlocker-recovery-keys-and-associated-devices

[CmdletBinding(SupportsShouldProcess)] #Make sure we can use -Verbose
Param([switch]$IncludeDeviceInfo,[switch]$IncludeDeviceOwner,[switch]$DeviceReport)

#==========================================================================
#Helper functions
#==========================================================================

function DriveType {
    Param($Drive)
    switch ($Drive) {
        1 { "operatingSystemVolume" }
        2 { "fixedDataVolume" }
        3 { "removableDataVolume" }
        4 { "unknownFutureValue" }
        Default { "Unknown" }
    }
}

#==========================================================================
#Main script starts here
#==========================================================================

#Handle parameter dependencies, device report requires device info and so on. We set $PSBoundParameters here, as we use it to determine the required scopes later on
if ($PSBoundParameters.ContainsKey("DeviceReport") -and $PSBoundParameters["DeviceReport"]) {
    $PSBoundParameters["IncludeDeviceInfo"] = $true
    $PSBoundParameters["IncludeDeviceOwner"] = $true
}
if ($PSBoundParameters.ContainsKey("IncludeDeviceOwner") -and $PSBoundParameters["IncludeDeviceOwner"]) {
    $PSBoundParameters["IncludeDeviceInfo"] = $true
}

#Determine the required scopes, based on the parameters passed to the script
$RequiredScopes = switch ($PSBoundParameters.Keys) {
    "IncludeDeviceInfo" { if ($PSBoundParameters["IncludeDeviceInfo"]) {"Device.Read.All" } }
    "IncludeDeviceOwner" { if ($PSBoundParameters["IncludeDeviceOwner"]) {"User.ReadBasic.All" } } #Otherwise we only get the UserId
    Default { "BitLockerKey.Read.All" }
}

Write-Verbose "Connecting to Graph API..."
Import-Module Microsoft.Graph.Identity.SignIns -Verbose:$false -ErrorAction Stop
try {
    Connect-MgGraph -Scopes $RequiredScopes -verbose:$false -ErrorAction Stop -NoWelcome
}
catch { throw $_ }

#Check if we have all the required permissions
$CurrentScopes = (Get-MgContext).Scopes
if ($RequiredScopes | ? {$_ -notin $CurrentScopes }) { Write-Error "The access token does not have the required permissions, rerun the script and consent to the missing scopes!" -ErrorAction Stop }

#If requested, retrieve the device details
if ($PSBoundParameters["IncludeDeviceInfo"]) {
    Write-Verbose "Retrieving device details..."

    $Devices = @()
    if ($PSBoundParameters["IncludeDeviceOwner"]) {
        Write-Verbose "Retrieving device owner..."
        $Devices = Get-MgDevice -All -ExpandProperty registeredOwners -ErrorAction Stop -Verbose:$false
    }
    else { $Devices = Get-MgDevice -All -ErrorAction Stop -Verbose:$false }

    if ($Devices) { Write-Verbose "Retrieved $($Devices.Count) devices" }
    else { Write-Verbose "No devices found"; continue }

    #Prepare the device object to be used later on
    if ($PSBoundParameters["DeviceReport"]) {
        $Devices | Add-Member -MemberType NoteProperty -Name "BitLockerKeyId" -Value $null
        $Devices | Add-Member -MemberType NoteProperty -Name "BitLockerRecoveryKey" -Value $null
        $Devices | Add-Member -MemberType NoteProperty -Name "BitLockerDriveType" -Value $null
        $Devices | Add-Member -MemberType NoteProperty -Name "BitLockerBackedUp" -Value $null
    }
    $Devices | % { Add-Member -InputObject $_ -MemberType NoteProperty -Name "DeviceOwner" -Value (&{if ($_.registeredOwners) { $_.registeredOwners[0].AdditionalProperties.userPrincipalName } else { "N/A" }}) }    
}

#Get the list of application objects within the tenant.
$Keys = @()

#Get the list of BitLocker keys
Write-Verbose "Retrieving BitLocker keys..."
$Keys = Get-MgInformationProtectionBitlockerRecoveryKey -All -ErrorAction Stop -Verbose:$false

#Cycle through the keys and retrieve the key
Write-Verbose "Retrieving BitLocker Recovery keys..."
foreach ($Key in $Keys) {
    $RecoveryKey = Get-MgInformationProtectionBitlockerRecoveryKey -BitlockerRecoveryKeyId $Key.Id -Property key -ErrorAction Stop -Verbose:$false
    $Key.Key = (&{if ($RecoveryKey.Key) { $RecoveryKey.Key } else { "N/A" }})
    $Key | Add-Member -MemberType NoteProperty -Name "BitLockerKeyId" -Value $Key.Id
    $Key | Add-Member -MemberType NoteProperty -Name "BitLockerRecoveryKey" -Value $Key.Key
    $Key | Add-Member -MemberType NoteProperty -Name "BitLockerDriveType" -Value (DriveType $Key.VolumeType)
    $Key | Add-Member -MemberType NoteProperty -Name "BitLockerBackedUp" -Value (&{if ($Key.CreatedDateTime) { Get-Date($Key.CreatedDateTime) -Format g } else { "N/A" }})

    #If requested, include the device details
    if ($PSBoundParameters["IncludeDeviceInfo"]) {

        $Device = $Devices | ? { $Key.DeviceId -eq $_.DeviceId }
        if (!$Device) { Write-Warning "Device with ID $($Key.DeviceId) not found!"; continue }

        #If building a device report, add the BitLocker key details to the device object
        if ($PSBoundParameters["DeviceReport"]) {
            $Device.BitLockerKeyId = $Key.Id
            $Device.BitLockerRecoveryKey = $Key.Key
            $Device.BitLockerDriveType = (DriveType $Key.VolumeType)
            $Device.BitLockerBackedUp = (&{if ($Key.CreatedDateTime) { Get-Date($Key.CreatedDateTime) -Format g } else { "N/A" }})
        }

        $Key | Add-Member -MemberType NoteProperty -Name "DeviceName" -Value $Device.DisplayName
        $Key | Add-Member -MemberType NoteProperty -Name "DeviceGUID" -Value $Device.Id #key actually used by the stupid module...
        $Key | Add-Member -MemberType NoteProperty -Name "DeviceOS" -Value $Device.OperatingSystem
        $Key | Add-Member -MemberType NoteProperty -Name "DeviceTrustType" -Value $Device.TrustType
        $Key | Add-Member -MemberType NoteProperty -Name "DeviceMDM" -Value $Device.AdditionalProperties.managementType #can be null! ALWAYS null when using a filter...
        $Key | Add-Member -MemberType NoteProperty -Name "DeviceCompliant" -Value $Device.isCompliant #can be null!
        $Key | Add-Member -MemberType NoteProperty -Name "DeviceRegistered" -Value (&{if ($Device.registrationDateTime) { Get-Date($Device.registrationDateTime) -Format g } else { "N/A" }})
        $Key | Add-Member -MemberType NoteProperty -Name "DeviceLastActivity" -Value (&{if ($Device.approximateLastSignInDateTime) { Get-Date($Device.approximateLastSignInDateTime) -Format g } else { "N/A" }})

        #If requested, include the device owner
        if ($PSBoundParameters["IncludeDeviceOwner"]) {
            $Key | Add-Member -MemberType NoteProperty -Name "DeviceOwner" -Value (&{if ($Device.registeredOwners) { $Device.registeredOwners[0].AdditionalProperties.userPrincipalName } else { "N/A" }})
        }
    }
}

#Export the result to CSV file
if ($PSBoundParameters["DeviceReport"]) {
    #BitLocker keys are at the front, followed by the device details. Cleaned up few "internal" properties, adjust the list below as needed
    $ExcludeProps = @("AdditionalProperties","AlternativeSecurityIds","complianceExpirationDateTime","deviceMetadata","deviceVersion","memberOf","PhysicalIds","SystemLabels","transitiveMemberOf","RegisteredOwners","RegisteredUsers")
    $Devices | select * -ExcludeProperty $ExcludeProps | Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_BitLockerKeys.csv"
}
else {
    $Keys | select * -ExcludeProperty Id,VolumeType,AdditionalProperties,CreatedDateTime,Key | Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_BitLockerKeys.csv"
}
Write-Verbose "Output exported to $($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_BitLockerKeys.csv"