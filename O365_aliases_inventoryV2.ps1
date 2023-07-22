#Requires -Version 3.0
#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.0.0" }
#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/5711/report-on-all-microsoft-365-email-addresses
param([switch]$IncludeAll,[switch]$IncludeUserMailboxes,[switch]$IncludeSharedMailboxes,[switch]$IncludeRoomMailboxes,[switch]$IncludeGroupMailboxes,[switch]$IncludeDGs,[switch]$IncludeMailUsers,[switch]$IncludeMailContacts,[switch]$CondensedOutput,[switch]$IncludeSIPAliases,[switch]$IncludeSPOAliases)

#Helper function for ExO connectivity
function Check-Connectivity {
    [cmdletbinding()]param()

    #Make sure we are connected to Exchange Online PowerShell
    Write-Verbose "Checking connectivity to Exchange Online PowerShell..."

    #Check via Get-ConnectionInformation first
    if (Get-ConnectionInformation) { return $true }

    #Double-check and try to eastablish a session
    try { Get-EXORecipient -ResultSize 1 -ErrorAction Stop | Out-Null }
    catch {
        try { Connect-ExchangeOnline -CommandName Get-EXORecipient, Get-Recipient -SkipLoadingFormatData -ShowBanner:$false } #custom for this script
        catch { Write-Error "No active Exchange Online session detected. To connect to ExO: https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps"; return $false }
    }

    return $true
}

function Get-EmailAddressesInventory {
<#
.Synopsis
    Lists all aliases for all recipients of the selected type(s).
.DESCRIPTION
    The Get-EmailAddressesInventory cmdlet finds all recipients of the selected type(s) and lists their email and/or non-email aliases.
    Running the cmdlet without parameters will return entries for User mailboxes only. Specifying particular recipient type(s) can be done with the corresponding switch parameter.
    To use condensed output (one line per recipient), use the -CondensedOutput switch.
    To specify a variable in which to hold the cmdlet output, use the -OutVariable parameter.

.EXAMPLE
    Get-EmailAddressesInventory -IncludeUserMailboxes

    This command will return a list of email aliases for all user mailboxes.

.EXAMPLE
    Get-EmailAddressesInventory -IncludeAll -OutVariable global:var
    $var | Export-Csv -NoTypeInformation "accessrights.csv"

    To export the results to a CSV file, use the OutVariable parameter.
.INPUTS
    None.
.OUTPUTS
    Array with information about the recipient type and its aliases.
#>

    [CmdletBinding()]
    
    Param
    (
    #Specify whether to include User mailboxes in the output
    [Switch]$IncludeUserMailboxes,    
    #Specify whether to include Shared mailboxes in the output
    [Switch]$IncludeSharedMailboxes,
    #Specify whether to include Room, Equipment and Scheduling mailboxes in the output
    [Switch]$IncludeRoomMailboxes,
    #Specify whether to include Group mailboxes in the output
    [Switch]$IncludeGroupMailboxes,
    #Specify whether to include Distribution Groups, Dynamic Distribution Groups, Room Lists and Mail-enabled Security Groups in the output
    [Switch]$IncludeDGs,
    #Specify whether to include Mail Users, Guest Mail Users and SharedWIth Mail Users in the output
    [switch]$IncludeMailUsers,
    #Specify whether to include Mail Contacts in the output
    [switch]$IncludeMailContacts,
    #Specify whether to return all supported recipient types in the output
    [Switch]$IncludeAll,
    #Specify whether to write the output in condensed format
    [Switch]$CondensedOutput,
    #Specify whether to include SIP/EUM aliases in the output
    [switch]$IncludeSIPAliases,
    #Specify whether to include SPO aliases in the output
    [switch]$IncludeSPOAliases)

    
    #Initialize the variable used to designate recipient types, based on the input parameters
    $included = @()
    if($IncludeSharedMailboxes) { $included += "SharedMailbox"}
    if($IncludeRoomMailboxes) { $included += "RoomMailbox","EquipmentMailbox","SchedulingMailbox"}
    if($IncludeMailUsers) { $included += "MailUser","GuestMailUser","SharedWithMailUser"}
    if($IncludeMailContacts) { $included += "MailContact"}
    if($IncludeGroupMailboxes) { $included += "GroupMailbox"}
    if($IncludeDGs) { $included += 'DynamicDistributionGroup', 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', 'RoomList'}
    
    #If no parameters specified, return only User mailboxes
    if($IncludeUserMailboxes -or !$included) { $included += "UserMailbox" }

    #Use the -IncludeAll parameter if you want to cover all recipient types. Full list below:
    if($IncludeAll) { 
        $Included = @(
            'UserMailbox',
            'SharedMailbox',
            'RoomMailbox',
            'EquipmentMailbox',
            'SchedulingMailbox',
            'TeamMailbox',
            'DiscoveryMailbox',
            'MailUser',
            'GuestMailUser',
            'SharedWithMailUser',
            'MailContact',
            'DynamicDistributionGroup',
            'MailUniversalDistributionGroup',
            'MailUniversalSecurityGroup',
            'RoomList',
            'PublicFolder',
            'PublicFolderMailbox',
            'GroupMailbox')
    }

    #Confirm connectivity to Exchange Online
    if (!(Check-Connectivity)) { return }

    #Get a minimal set of properties for the selected recipients. Make sure to add any additional properties you want included in the report to the list here!
    $MBList = Get-EXORecipient -ResultSize Unlimited -RecipientTypeDetails $included -Properties WindowsLiveID,ExternalEmailAddress | Select-Object -Property Displayname,PrimarySMTPAddress,WindowsLiveID,EmailAddresses,ExternalEmailAddress,RecipientTypeDetails

    #If no recipients are returned from the above cmdlet, stop the script and inform the user
    if (!$MBList) { Write-Error "No recipients of the specifyied types were found, specify different criteria." -ErrorAction Stop }

    #Once we have the recipient list, cycle over each recipient to prepare the output object
    $arrAliases = @()
    foreach ($MB in $MBList) {
        #If we want condensed output, one line per recipient
        if ($CondensedOutput) {
            $objAliases = New-Object PSObject
            $i++;Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Number" -Value $i
            Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Recipient" -Value $MB.DisplayName
            Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Primary SMTP address" -Value $MB.PrimarySMTPAddress
            #we use WindowsLiveID as a workaround to get the UPN, as Get-ExORecipient does not return the UPN property
            if ($MB.WindowsLiveID) { Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "UPN" -Value $MB.WindowsLiveID }
            Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Recipient type" -Value $MB.RecipientTypeDetails
            Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Email Aliases" -Value (($MB.EmailAddresses | ? {$_.Split(":")[0] -eq "SMTP" -or $_.Split(":")[0] -eq "X500"}) -join ";")

            #Handle SIP/SPO aliases and external email address depending on the parameters provided
            if ($IncludeSIPAliases -or $IncludeAll) { Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "SIP Aliases" -Value (($MB.EmailAddresses | ? {$_.Split(":")[0] -eq "SIP" -or $_.Split(":")[0] -eq "EUM"}) -join ";") }
            if ($IncludeSPOAliases -or $IncludeAll) { Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "SPO Aliases" -Value (($MB.EmailAddresses | ? {$_.Split(":")[0] -eq "SPO"}) -join ";") }
            if ($IncludeMailUsers -or $IncludeMailContacts -or $IncludeAll) { Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "External email address" -Value $MB.ExternalEmailAddress }

            $arrAliases += $objAliases
        }

        #Otherwise, write each permission entry on separate line
        else {
            foreach ($entry in $MB.EmailAddresses) {
                $objAliases = New-Object PSObject
                $i++;Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Number" -Value $i
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Recipient" -Value $MB.DisplayName
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Primary SMTP address" -Value $MB.PrimarySMTPAddress
                #we use WindowsLiveID as a workaround to get the UPN, as Get-ExORecipient does not return the UPN property
                if ($MB.WindowsLiveID) { Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "UPN" -Value $MB.WindowsLiveID }
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Recipient type" -Value $MB.RecipientTypeDetails
                #Handle SIP/SPO aliases depending on the parameters provided
                if (($entry.Split(":")[0] -eq "SIP" -or $entry.Split(":")[0] -eq "EUM") -and !($IncludeSIPAliases -or $IncludeAll)) { continue }
                if ($entry.Split(":")[0] -eq "SPO" -and !($IncludeSPOAliases -or $IncludeAll)) { continue }
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Aliases" -Value $entry
                
                $arrAliases += $objAliases
            }
            #Handle External email address for Mail User/Mail Contact objects
            if (($IncludeMailUsers -or $IncludeMailContacts -or $IncludeAll) -and $MB.ExternalEmailAddress) {
                if ($MB.ExternalEmailAddress.Split(":")[1] -eq $MB.PrimarySMTPAddress) { continue }
                $objAliases = New-Object PSObject
                $i++;Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Number" -Value $i
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Recipient" -Value $MB.DisplayName
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Primary SMTP address" -Value $MB.PrimarySMTPAddress
                if ($MB.WindowsLiveID) { Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "UPN" -Value $MB.WindowsLiveID }
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Recipient type" -Value $MB.RecipientTypeDetails
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Aliases" -Value $MB.ExternalEmailAddress
                $arrAliases += $objAliases
            }
        }
    }
    #Output the result to the console host. Rearrange/sort as needed.
    $arrAliases | select * -ExcludeProperty Number | sort Aliases -Unique
}

#Invoke the Get-EmailAddressesInventory function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
Get-EmailAddressesInventory @PSBoundParameters -OutVariable global:varAliases | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_EmailAddresses.csv" -NoTypeInformation -Encoding UTF8 -UseCulture

<# All possible RecipientTypeDetails values
None, UserMailbox, LinkedMailbox, SharedMailbox, LegacyMailbox, RoomMailbox, EquipmentMailbox, MailContact, MailUser, MailUniversalDistributionGroup, MailNonUniversalGroup, MailUniversalSecurityGroup,
DynamicDistributionGroup, PublicFolder, SystemAttendantMailbox, SystemMailbox, MailForestContact, User, Contact, UniversalDistributionGroup, UniversalSecurityGroup, NonUniversalGroup, DisabledUser, MicrosoftExchange,
ArbitrationMailbox, MailboxPlan, LinkedUser, RoomList, DiscoveryMailbox, RoleGroup, RemoteUserMailbox, Computer, RemoteRoomMailbox, RemoteEquipmentMailbox, RemoteSharedMailbox, PublicFolderMailbox, TeamMailbox,
RemoteTeamMailbox, MonitoringMailbox, GroupMailbox, LinkedRoomMailbox, AuditLogMailbox, RemoteGroupMailbox, SchedulingMailbox, GuestMailUser, AuxAuditLogMailbox, SupervisoryReviewPolicyMailbox, ExchangeSecurityGroup,
SubstrateGroup, SubstrateADGroup, WorkspaceMailbox, SharedWithMailUser, ServicePrinciple, BlobShard, DeskMailbox, SubstrateTenantRecipient, AllUniqueRecipientTypes
#>

<# All *supported* RecipientTypeDetails values
'RoomMailbox', 'LinkedRoomMailbox', 'EquipmentMailbox', 'SchedulingMailbox', 'LegacyMailbox', 'LinkedMailbox', 'UserMailbox', 'MailContact', 'DynamicDistributionGroup', 'MailForestContact', 'MailNonUniversalGroup',
'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', 'RoomList', 'MailUser', 'GuestMailUser', 'GroupMailbox', 'DiscoveryMailbox', 'PublicFolder', 'TeamMailbox', 'SharedMailbox', 'RemoteUserMailbox',
'RemoteRoomMailbox', 'RemoteEquipmentMailbox', 'RemoteTeamMailbox', 'RemoteSharedMailbox','PublicFolderMailbox', 'SharedWithMailUser'
#>