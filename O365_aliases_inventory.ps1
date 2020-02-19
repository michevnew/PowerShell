param([switch]$IncludeAll,[switch]$IncludeUserMailboxes,[switch]$IncludeSharedMailboxes,[switch]$IncludeRoomMailboxes,[switch]$IncludeGroupMailboxes,[switch]$IncludeDGs,[switch]$IncludeMailUsers,[switch]$IncludeMailContacts,[switch]$CondensedOutput,[switch]$IncludeSIPAliases,[switch]$IncludeSPOAliases)

function Get-EmailAddressesInventory {
<#
.Synopsis
    Lists all aliases for all recipients of the selected type(s).
.DESCRIPTION
    The Get-EmailAddressesInventory cmdlet finds all recipients of the selected type(s) and lists their email and/or non-email aliases.
    Running the cmdlet without parameters will return entries for User mailboxes only. Specifying particular recipient type(s) can be done with the corresponding switch parameter.
    To use condensed output (one line per recipient), use the CondensedOutput switch.
    To specify a variable in which to hold the cmdlet output, use the OutVariable parameter.

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
    #Specify whether to include User mailboxes in the result
    [Switch]$IncludeUserMailboxes,    
    #Specify whether to include Shared mailboxes in the result
    [Switch]$IncludeSharedMailboxes,
    #Specify whether to include Room and Equipment mailboxes in the result
    [Switch]$IncludeRoomMailboxes,
    #Specify whether to include Group mailboxes in the result
    [Switch]$IncludeGroupMailboxes,
    #Specify whether to include Distribution Groups, Dynamic Distribution Groups, Room Lists and Mail-enabled Security Groups in the result
    [Switch]$IncludeDGs,
    #Specify whether to include Mail Users and Guest Mail Users in the result
    [switch]$IncludeMailUsers,
    #Specify whether to include Mail Contacts in the result
    [switch]$IncludeMailContacts,
    #Specify whether to return all recipient types in the result
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
    if($IncludeRoomMailboxes) { $included += "RoomMailbox","EquipmentMailbox"}
    if($IncludeMailUsers) { $included += "MailUser","GuestMailUser"}
    if($IncludeMailContacts) { $included += "MailContact"}
    if($IncludeGroupMailboxes) { $included += "GroupMailbox"}
    if($IncludeDGs) { $included += 'DynamicDistributionGroup', 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', 'RoomList'}
    
    #If no parameters specified, return only User mailboxes
    if($IncludeUserMailboxes -or !$included) { $included += "UserMailbox"}
    #Use the -IncludeAll parameter if you want to cover all recipient types, full list below
    #'UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox','TeamMailbox','DiscoveryMailbox','MailUser','MailContact', 'DynamicDistributionGroup', 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', 'RoomList','GuestMailUser','PublicFolder','GroupMailbox'
    if($IncludeAll) {$Included = @('UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox','TeamMailbox','DiscoveryMailbox','MailUser','MailContact', 'DynamicDistributionGroup', 'MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', 'RoomList','GuestMailUser','PublicFolder','GroupMailbox') }

    #Confirm connectivity to Exchange Online
    try { $session = Get-PSSession -InstanceId (Get-OrganizationConfig).RunspaceId.Guid -ErrorAction Stop  }
    catch { Write-Error "No active Exchange Online session detected, please connect to ExO first: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop }

    #Get a minimal set of properties for the selected recipients. Make sure to add any additional properties you want included in the report to the list here!
    $MBList = Invoke-Command -Session $session -ScriptBlock { Get-Recipient -ResultSize Unlimited -RecipientTypeDetails $Using:included | Select-Object -Property Displayname,PrimarySMTPAddress,WindowsLiveID,EmailAddresses,ExternalEmailAddress,RecipientTypeDetails }

    #If no recipients are returned from the above cmdlet, stop the script and inform the user
    if (!$MBList) { Write-Error "No recipients of the specifyied types were found, specify different criteria." -ErrorAction Stop}

    #Once we have the recipient list, cycle over each recipient to prepare the output object
    $arrAliases = @()
    foreach ($MB in $MBList) {
        #If we want condensed output, one line per recipient
        if ($CondensedOutput) {
            $objAliases = New-Object PSObject
            $i++;Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Number" -Value $i
            Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Recipient" -Value $MB.DisplayName
            Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Primary SMTP address" -Value $MB.PrimarySMTPAddress
            #we use WindowsLiveID as a workaround to get the UPN, as Get-Recipient does not return the UPN property
            if ($MB.WindowsLiveID) { Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "UPN" -Value $MB.WindowsLiveID.Address }
            Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Recipient type" -Value $MB.RecipientTypeDetails.Value
            Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Email Aliases" -Value (($MB.EmailAddresses | ? {$_.Prefix -eq "SMTP" -or $_.Prefix -eq "X500"}).ProxyAddressString -join ";")

            #Handle SIP/SPO aliases and external email address depending on the parameters provided
            if ($IncludeSIPAliases -or $IncludeAll) { Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "SIP Aliases" -Value (($MB.EmailAddresses | ? {$_.Prefix -eq "SIP" -or $_.Prefix -eq "EUM"}).ProxyAddressString -join ";") }
            if ($IncludeSPOAliases -or $IncludeAll) { Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "SPO Aliases" -Value (($MB.EmailAddresses | ? {$_.Prefix -eq "SPO"}).ProxyAddressString -join ";") }
            if ($IncludeMailUsers -or $IncludeMailContacts -or $IncludeAll) { Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "External email address" -Value $MB.ExternalEmailAddress.ProxyAddressString }

            $arrAliases += $objAliases
        }
        #Otherwise, write each permission entry on separate line
        else {
            foreach ($entry in $MB.EmailAddresses) {
                $objAliases = New-Object PSObject
                $i++;Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Number" -Value $i
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Recipient" -Value $MB.DisplayName
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Primary SMTP address" -Value $MB.PrimarySMTPAddress
                #we use WindowsLiveID as a workaround to get the UPN, as Get-Recipient does not return the UPN property
                if ($MB.WindowsLiveID) { Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "UPN" -Value $MB.WindowsLiveID.Address }
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Recipient type" -Value $MB.RecipientTypeDetails.Value
                #Handle SIP/SPO aliases depending on the parameters provided
                if (($entry.Prefix -eq "SIP" -or $entry.Prefix -eq "EUM") -and !($IncludeSIPAliases -or $IncludeAll)) { continue }
                if ($entry.Prefix -eq "SPO" -and !($IncludeSPOAliases -or $IncludeAll)) { continue }
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Aliases" -Value $entry.ProxyAddressString
                
                $arrAliases += $objAliases
            }
            #Handle External email address for Mail User/Mail Contact objects
            if (($IncludeMailUsers -or $IncludeMailContacts -or $IncludeAll) -and $MB.ExternalEmailAddress) {
                if ($MB.ExternalEmailAddress.AddressString -eq $MB.PrimarySMTPAddress) { continue }
                $objAliases = New-Object PSObject
                $i++;Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Number" -Value $i
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Recipient" -Value $MB.DisplayName
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Primary SMTP address" -Value $MB.PrimarySMTPAddress
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Recipient type" -Value $MB.RecipientTypeDetails.Value
                Add-Member -InputObject $objAliases -MemberType NoteProperty -Name "Aliases" -Value $MB.ExternalEmailAddress.ProxyAddressString
                $arrAliases += $objAliases
            }
        }
    }
    #Output the result to the console host. Rearrange/sort as needed.
    $arrAliases | select * -ExcludeProperty Number
}

#Invoke the Get-EmailAddressesInventory function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
Get-EmailAddressesInventory @PSBoundParameters -OutVariable global:varAliases | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_EmailAddresses.csv" -NoTypeInformation -Encoding UTF8 -UseCulture

<# None, UserMailbox, LinkedMailbox, SharedMailbox, LegacyMailbox, RoomMailbox, EquipmentMailbox, MailContact, MailUser,
MailUniversalDistributionGroup, MailNonUniversalGroup, MailUniversalSecurityGroup, DynamicDistributionGroup, PublicFolder, SystemAttendantMailbox, SystemMailbox,
MailForestContact, User, Contact, UniversalDistributionGroup, UniversalSecurityGroup, NonUniversalGroup, DisabledUser, MicrosoftExchange, ArbitrationMailbox, MailboxPlan,
LinkedUser, RoomList, DiscoveryMailbox, RoleGroup, RemoteUserMailbox, Computer, RemoteRoomMailbox, RemoteEquipmentMailbox, RemoteSharedMailbox, PublicFolderMailbox, TeamMailbox,
RemoteTeamMailbox, MonitoringMailbox, GroupMailbox, LinkedRoomMailbox, AuditLogMailbox, RemoteGroupMailbox, SchedulingMailbox, GuestMailUser, AuxAuditLogMailbox,
SupervisoryReviewPolicyMailbox, ExchangeSecurityGroup, AllUniqueRecipientTypes""
#>