#Helper function for loading the mailbox data. If no existing CSV file is found or it is outdated, the function will generate a new file (might take some time)...
function Load-MailboxMatchInputFile {
    $importCSV = Get-ChildItem -Path $PSScriptRoot -Filter "*MailboxReport.csv" | sort LastWriteTime -Descending | select -First 1 #| select -ExpandProperty FullName

    if (!$importCSV -or $importCSV.LastWriteTime -le (Get-Date).AddDays(-30)) {
        #No CSV file detected or it's too old, generate new mailbox report
        Write-Host "No Mailbox report file detected, or it's too old. Generating new report file..." -ForegroundColor Yellow

        try { $session = Get-PSSession -InstanceId (Get-AcceptedDomain | select -First 1).RunspaceId.Guid -ErrorAction Stop }
        catch { Write-Error "No active Exchange Online session detected, please connect to ExO first: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop }

        #ONLY User mailboxes are included, add other types as needed. Add other properties as needed.
        $toImport = Invoke-Command -Session $session -ScriptBlock { Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Select-Object DisplayName,Alias,UserPrincipalName,DistinguishedName,PrimarySmtpAddress } -HideComputerName | select * -ExcludeProperty RunspaceId
        $toImport | Export-Csv -Path "$PSScriptRoot\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxReport.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
    }
    else { $toImport = Import-Csv $importCSV.FullName }
    $global:mailboxes = arrayToHash -array $toImport -key "DistinguishedName"
}


#Helper function to load the mailbox report as hashtable
function arrayToHash {
    Param($array,[string]$key)

    $hash = @{}
    foreach ($a in $array) {
        if (!$key) { $i++; $hash[$i] = $a }
        else { $hash[$a.($key)] = $a }
        }
    return $hash
}


#Helper function to find a matching mailbox based on the DN
function Get-MailboxMatch {
    [CmdletBinding()]

    Param([parameter(Position=0, Mandatory=$true)][String]$dn)

    if (!$mailboxes) { Load-MailboxMatchInputFile }
    return $mailboxes[$dn]
}



#Main script starts here. Check for connectivity to ExO first.
try { $session = Get-PSSession -InstanceId (Get-AcceptedDomain | select -First 1).RunspaceId.Guid -ErrorAction Stop }
catch { Write-Error "No active Exchange Online session detected, please connect to ExO first: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop }

#Gather the inventory of all mobile devices
#Make sure to add any additional properties you need to the below list
#Make sure to add any Filters as needed!
$MobileDevices = Invoke-Command -Session $session -ScriptBlock { Get-MobileDevice -ResultSize Unlimited | Select-Object FriendlyName,UserDisplayName,DeviceId,DeviceOS,DeviceType,DeviceUserAgent,DeviceModel,DistinguishedName,FirstSyncTime,DeviceAccessState,DeviceAccessStateReason,DeviceAccessControlRule,ClientType } -HideComputerName | select * -ExcludeProperty RunspaceId
if (!$MobileDevices) { Write-Host "No mobile devices found, make sure you are using the correct credentials and/or adjust any filters." -ForegroundColor Yellow; return }

#Load mailbox data and prepare the hashtable for lookups
Load-MailboxMatchInputFile

#Loop over each device to prepare the output
$count = 1; $PercentComplete = 0;
foreach ($device in $MobileDevices) {
    #Progress message. This also adds some delay, so consider removing it...
    $ActivityMessage = "Retrieving data for device $($device.DeviceId). Please wait..."
    $StatusMessage = ("Processing mailbox {0} of {1}: {2}" -f $count, @($MobileDevices).count, $device.FriendlyName)
    $PercentComplete = ($count / @($MobileDevices).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Check for session state and add some artificial delay every 100 iterations or so...
    if ($count /100 -is [int]) {
        #Consider adding the session reconnect logic from the robust cmdlet execution script here!
        $session = Get-PSSession -InstanceId (Get-AcceptedDomain | select -First 1).RunspaceId.Guid -ErrorAction Stop
        Start-Sleep -Seconds 1
        }

    #Get additional properties from Get-MobileDeviceStatistics. If you dont need those, comment the below lines to greatly speed up the script. If you need additional properties, add them to the list below.
    $devicestats = Invoke-Command -Session $session -ScriptBlock { Get-MobileDeviceStatistics $using:device.DistinguishedName | Select-Object LastSuccessSync,Status,DevicePolicyApplied,DevicePolicyApplicationStatus } -HideComputerName -ErrorAction SilentlyContinue

    if ($devicestats) {
        $device | Add-Member -MemberType NoteProperty -Name LastSuccessSync -Value (&{If($devicestats.LastSuccessSync) {$devicestats.LastSuccessSync.ToString()} Else {"Never"}})
        $device | Add-Member -MemberType NoteProperty -Name Status -Value $devicestats.Status.Value
        $device | Add-Member -MemberType NoteProperty -Name DevicePolicyApplied -Value $devicestats.DevicePolicyApplied.Name
        $device | Add-Member -MemberType NoteProperty -Name DevicePolicyApplicationStatus -Value $devicestats.DevicePolicyApplicationStatus.Value
        #NumberOfFoldersSynced always seems to return 0 on Outlook/REST?!
    }

    #Find the mailbox owner and add relevant properties
    $mailbox = New-Object psobject
    Get-MailboxMatch ($device.DistinguishedName.Split(",")[2..10] -join ",") -OutVariable mailbox | Out-Null
    if ($mailbox.UserPrincipalName) {
        $device | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $mailbox.UserPrincipalName
        $device | Add-Member -MemberType NoteProperty -Name PrimarySmtpAddress -Value $mailbox.PrimarySmtpAddress.ToString()
    }

    #If you need properties from additional cmdlets, such as Get-User or Get-MsolUser, add them here. This will greatly increase the script runtime though, so use only if needed!
    #$device | Add-Member -MemberType NoteProperty -Name License -Value ((Get-MsolUser -UserPrincipalName $mailbox.UserPrincipalName).Licenses.ServiceStatus | ? {$_.ServicePlan.ServiceName -like "Exchange_?_*" -and $_.ServicePlan.TargetClass -eq "User"}).ServicePlan.ServiceName
}

#Export the output to a CSV file
$MobileDevices #| Export-Csv -Path "$PSScriptRoot\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MobileDeviceReport.csv" -NoTypeInformation -Encoding UTF8 -UseCulture