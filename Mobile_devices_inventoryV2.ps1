#Requires -Version 3.0
#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.0.0" }
#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/5639/mobile-device-inventory-and-statistics-report-2023-updated-version

#Helper function for ExO connectivity
function Check-Connectivity {
    [cmdletbinding()]
    [OutputType([bool])]
    param()

    #Make sure we are connected to Exchange Online PowerShell
    Write-Verbose "Checking connectivity to Exchange Online PowerShell..."

    #Check via Get-ConnectionInformation first
    if (Get-ConnectionInformation) { return $true }

    #Double-check and try to eastablish a session
    try { Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null }
    catch {
        try { Connect-ExchangeOnline -CommandName Get-EXOMailbox, Get-MobileDevice, Get-EXOMobileDeviceStatistics -SkipLoadingFormatData -ShowBanner:$false } #custom for this script
        catch { Write-Error "No active Exchange Online session detected. To connect to ExO: https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps"; return $false }
    }

    return $true
}

#Helper function for loading the mailbox data. If no existing CSV file is found or it is outdated, the function will generate a new file (might take some time)...
function Load-MailboxMatchInputFile {
    $importCSV = Get-ChildItem -Path $PSScriptRoot -Filter "*MailboxReport.csv" | sort LastWriteTime -Descending | select -First 1 #| select -ExpandProperty FullName

    if (!$importCSV -or $importCSV.LastWriteTime -le (Get-Date).AddDays(-30)) {
        #No CSV file detected or it's too old, generate new mailbox report
        Write-Host "No Mailbox report file detected, or it's too old. Generating new report file..." -ForegroundColor Yellow

        if (!(Check-Connectivity)) { return }

        #ONLY User mailboxes are included, add other types as needed. Add other properties as needed.
        $toImport = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Select-Object DisplayName,Alias,UserPrincipalName,DistinguishedName,PrimarySmtpAddress
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
if (!(Check-Connectivity)) { return }

#Gather the inventory of all mobile devices. We use Get-MobileDevice to avoid looping over all mailboxes.
#Make sure to add any additional properties you need to the below list, or any Filters as needed!
$MobileDevices = Get-MobileDevice -ResultSize Unlimited | Select-Object FriendlyName,UserDisplayName,DeviceId,DeviceOS,DeviceType,DeviceUserAgent,DeviceModel,DistinguishedName,Identity,OrganizationId,FirstSyncTime,DeviceAccessState,DeviceAccessStateReason,DeviceAccessControlRule,ClientType
if (!$MobileDevices) { Write-Host "No mobile devices found, make sure you are using the correct credentials and/or adjust any filters." -ForegroundColor Yellow; return }

#Load mailbox data and prepare the hashtable for lookups
Load-MailboxMatchInputFile

#Loop over each device to prepare the output
$count = 1; $PercentComplete = 0;
foreach ($device in $MobileDevices) {
    #Progress message. This also adds some delay, so consider removing it...
    $ActivityMessage = "Retrieving data for device $($device.DeviceId). Please wait..."
    $StatusMessage = ("Processing device {0} of {1}: {2}" -f $count, @($MobileDevices).count, $device.FriendlyName)
    $PercentComplete = ($count / @($MobileDevices).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Add some artificial delay every 100 iterations or so...
    if ($count /100 -is [int]) {
        Start-Sleep -Seconds 1
    }

    #Get additional properties from Get-EXOMobileDeviceStatistics. If you dont need those, comment the below lines to greatly speed up the script. If you need additional properties, add them to the list below.
    $deviceid = $device.OrganizationId.Split("-")[0].TrimEnd() + "/"+ $device.Identity.Replace("\","/")
    #You can use the Guid propery instead, however Get-EXOMobileDeviceStatistics is veeeeery slow with it :/
    $devicestats = Get-EXOMobileDeviceStatistics -Identity $deviceid | Select-Object LastSuccessSync,Status,DevicePolicyApplied,DevicePolicyApplicationStatus

    if ($devicestats) {
        $device | Add-Member -MemberType NoteProperty -Name LastSuccessSync -Value (&{If($devicestats.LastSuccessSync) {$devicestats.LastSuccessSync} Else {"Never"}})
        $device | Add-Member -MemberType NoteProperty -Name Status -Value $devicestats.Status
        $device | Add-Member -MemberType NoteProperty -Name DevicePolicyApplied -Value $devicestats.DevicePolicyApplied
        $device | Add-Member -MemberType NoteProperty -Name DevicePolicyApplicationStatus -Value $devicestats.DevicePolicyApplicationStatus
        #NumberOfFoldersSynced always seems to return 0 on Outlook/REST?!
    }

    #Find the mailbox owner and add relevant properties
    $mailbox = New-Object psobject
    Get-MailboxMatch ($device.DistinguishedName.Split(",")[2..10] -join ",") -OutVariable mailbox | Out-Null
    if ($mailbox.UserPrincipalName) {
        $device | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $mailbox.UserPrincipalName
        $device | Add-Member -MemberType NoteProperty -Name PrimarySmtpAddress -Value $mailbox.PrimarySmtpAddress
    }

    #If you need properties from additional cmdlets, such as Get-MgUser or Get-MsolUser, add them here. This will greatly increase the script runtime though, so use only if needed!
    #$device | Add-Member -MemberType NoteProperty -Name License -Value ((Get-MsolUser -UserPrincipalName $mailbox.UserPrincipalName).Licenses.ServiceStatus | ? {$_.ServicePlan.ServiceName -like "Exchange_?_*" -and $_.ServicePlan.TargetClass -eq "User"}).ServicePlan.ServiceName
}

#Export the output to a CSV file
$MobileDevices | Export-Csv -Path "$PSScriptRoot\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MobileDeviceReport.csv" -NoTypeInformation -Encoding UTF8 -UseCulture