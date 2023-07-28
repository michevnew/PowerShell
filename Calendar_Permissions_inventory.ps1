param([switch]$IncludeAll,[switch]$IncludeUserMailboxes,[switch]$IncludeSharedMailboxes,[switch]$IncludeRoomMailboxes,[switch]$CondensedOutput)

function Get-CalendarPermissionInventory {
<#
.Synopsis
    Lists all permissions on the default Calendar folder for all mailboxes of the selected type(s).
.DESCRIPTION
    The Get-CalendarPermissionInventory cmdlet finds the default Calendar folder for all mailboxes of the selected type(s) and lists its permissions.
    Running the cmdlet without parameters will return entries for User mailboxes only. Specifying particular mailbox type(s) can be done with the corresponding switch parameter.
    To use condensed output (one line per Calendar folder/mailbox), use the CondensedOutput switch.
    To specify a variable in which to hold the cmdlet output, use the OutVariable parameter.

.EXAMPLE
    Get-CalendarPermissionInventory -IncludeUserMailboxes

    This command will return a list of permissions for the default Calendar folder for all user mailboxes.

.EXAMPLE
    Get-CalendarPermissionInventory -IncludeAll -OutVariable global:var
    $var | Export-Csv -NoTypeInformation "accessrights.csv"

    To export the results to a CSV file, use the OutVariable parameter.
.INPUTS
    None.
.OUTPUTS
    Array with information about the mailbox, delegate and type of permissions.
#>

    [CmdletBinding()]

    Param
    (
    #Specify whether to include user mailboxes in the result
    [Switch]$IncludeUserMailboxes,
    #Specify whether to include Shared mailboxes in the result
    [Switch]$IncludeSharedMailboxes,
    #Specify whether to include Room and Equipment mailboxes in the result
    [Switch]$IncludeRoomMailboxes,
    #Specify whether to return all mailbox types
    [Switch]$IncludeAll,
    #Specify whether to write the output in condensed format
    [Switch]$CondensedOutput)

    #Initialize the variable used to designate mailbox types, based on the input parameters
    $included = @()
    if($IncludeSharedMailboxes) { $included += "SharedMailbox"}
    if($IncludeRoomMailboxes) { $included += "RoomMailbox"; $included += "EquipmentMailbox"}
    #if no parameters specified, return only User mailboxes
    if($IncludeUserMailboxes -or !$included) { $included += "UserMailbox"}

    #Confirm connectivity to Exchange Online
    try { $session = Get-PSSession -InstanceId (Get-OrganizationConfig).RunspaceId.Guid -ErrorAction Stop  }
    catch { Write-Error "No active Exchange Online session detected, please connect to ExO first: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop }

    #Get the list of mailboxes, depending on the parameters specified when invoking the script
    if ($IncludeAll) {
        $MBList = Invoke-Command -Session $session -ScriptBlock { Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails }
    }
    else {
        $MBList = Invoke-Command -Session $session -ScriptBlock { Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $Using:included | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails }
    }

    #If no mailboxes are returned from the above cmdlet, stop the script and inform the user
    if (!$MBList) { Write-Error "No mailboxes of the specifyied types were found, specify different criteria." -ErrorAction Stop}

    #Once we have the mailbox list, cycle over each mailbox to gather Calendar permissions inventory
    $arrPermissions = @()
    $count = 1; $PercentComplete = 0;
    foreach ($MB in $MBList) {
        #Start-Sleep -Milliseconds 200 #uncomment if getting throttled
        #Progress message
        $ActivityMessage = "Retrieving data for mailbox $($MB.Identity). Please wait..."
        $StatusMessage = ("Processing {0} of {1}: {2}" -f $count, @($MBList).count, $MB.PrimarySmtpAddress.ToString())
        $PercentComplete = ($count / @($MBList).count * 100)
        Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
        $count++

        #Get the default Calendar folder for each mailbox, depending on localization, etc.
        $calendarfolder = (Get-MailboxFolderStatistics $MB.PrimarySmtpAddress.ToString() -FolderScope Calendar | ? {$_.FolderType -eq "Calendar"}).Identity.ToString().Replace("\",":\").Replace([char]63743,"/")
        #Get the Calendar folder permissions
        $MBrights = Get-MailboxFolderPermission -Identity $calendarfolder #| ? {$_.User.DisplayName -notin @("Default","Anonymous","Owner@local","Member@local") #filter out default permissions
        #No non-default permissions found, continue to next mailbox
        if (!$MBrights) { continue }

        if ($CondensedOutput) {
            #Prepare the output object
            $objPermissions = New-Object PSObject
            $i++;Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Number" -Value $i
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox address" -Value $MB.PrimarySmtpAddress.ToString()
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox type" -Value $MB.RecipientTypeDetails.ToString()
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Calendar folder" -Value $calendarfolder
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Default level" -Value (($MBrights | ? {$_.User.DisplayName -eq "Default"}).AccessRights -join ";")
            #Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Anonymous level" -Value (($MBrights | ? {$_.User.DisplayName -eq "Anonymous"}).AccessRights -join ";")

            $internal = "";$external = "";$orphaned = ""
            foreach ($entry in $MBrights) {
                if ($entry.User.UserType.ToString() -eq "Internal") {
                    $internal = ("$($entry.User.RecipientPrincipal.PrimarySmtpAddress.ToString()):$($entry.AccessRights)" + ";" + $internal)
                }
                elseif ($entry.User.UserType.ToString() -eq "External") {
                    $external = ("$($entry.User.RecipientPrincipal.PrimarySmtpAddress.Replace("ExchangePublishedUser.",$null)):$($entry.AccessRights)" + ";" + $external)
                }
                elseif ($entry.User.UserType.ToString() -eq "Unknown") {
                    $orphaned = ("$($entry.User.DisplayName):$($entry.AccessRights)" + ";" + $orphaned)
                }
                else { continue }
            }
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Internal levels" -Value $internal.Trim(";")
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "External levels" -Value $external.Trim(";")
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Orphaned levels" -Value $orphaned.Trim(";")

            $arrPermissions += $objPermissions
            #Uncomment if the script is failing due to connectivity issues, the line below will write the output to a CSV file for each individual Calendar folder
            #$objPermissions | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd'))_CalendarPermissions.csv" -Append -NoTypeInformation -Encoding UTF8 -UseCulture
        }
        else {
            #Write each permission entry on separate line
            foreach ($entry in $MBrights) {
                #Prepare the output object
                $objPermissions = New-Object PSObject
                $i++;Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Number" -Value $i
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox address" -Value $MB.PrimarySmtpAddress.ToString()
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox type" -Value $MB.RecipientTypeDetails.ToString()
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Calendar folder" -Value $calendarfolder

                $varUser = "";$varType = "";
                if ($entry.User.UserType.ToString() -eq "Internal") { $varUser = $entry.User.RecipientPrincipal.PrimarySmtpAddress.ToString(); $varType = "Internal" }
                elseif ($entry.User.UserType.ToString() -eq "Default") { $varUser = $entry.User.DisplayName; $varType = "Default" }
                elseif ($entry.User.UserType.ToString() -eq "External") { $varUser = $entry.User.RecipientPrincipal.PrimarySmtpAddress.Replace("ExchangePublishedUser.",$null); $varType = "External" }
                elseif ($entry.User.UserType.ToString() -eq "Unknown") { $varUser = $entry.User.DisplayName; $varType = "Orphaned" }
                else { continue }

                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "User" -Value $varUser
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Permissions" -Value $($entry.AccessRights -join ";")
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Permission Type" -Value $varType

                $arrPermissions += $objPermissions
                #Uncomment if the script is failing due to connectivity issues, the line below will write the output to a CSV file for each individual permissions entry
                #$objPermissions | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd'))_CalendarPermissions.csv" -Append -NoTypeInformation -Encoding UTF8 -UseCulture
            }}
    }
    #Output the result to the console host. Rearrange/sort as needed.
    $arrPermissions | select * -ExcludeProperty Number
}

#Invoke the Get-CalendarPermissionInventory function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
Get-CalendarPermissionInventory @PSBoundParameters -OutVariable global:varPermissions  #| Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_CalendarPermissions.csv" -NoTypeInformation -Encoding UTF8 -UseCulture