#Requires -Version 3.0
param([switch]$CondensedOutput)

#Helper function for fetching data from Exchange Online
function Get-O365GroupMembershipInventory {
<#
.Synopsis
    Lists all Office 365 Groups and their corresponding Links (members)
.DESCRIPTION
    The Get-O365GroupMembershipInventory cmdlet enumerates all Office 365 Groups and lists their Links. This includes Members, Owners and Subscriber type of links (Aggregators and EventSubscribers are still not being used in the service)
    Output will be written to a CSV file and also exposed via the $varO365GroupMembers global variable for reuse.
    To use condensed output (one line per Group), use the -CondensedOutput switch.

.EXAMPLE
    Get-O365GroupMembershipInventory

    This command will return a list of all Office 365 Groups and lists their Links.

.EXAMPLE
    Get-O365GroupMembershipInventory -CondensedOutput -OutVariable global:var
    $var | Export-Csv -NoTypeInformation "O365GroupLinks.csv"

    To modify the output before exporting to CSV file, use the OutVariable parameter.
.INPUTS
    None.
.OUTPUTS
    Array with information about the Group and any Links.
#>

    [CmdletBinding()]
    Param(
    #Specify whether to write the output in condensed format
    [Switch]$CondensedOutput)

    #Confirm connectivity to Exchange Online
    try { $session = Get-PSSession -InstanceId (Get-AcceptedDomain | select -First 1).RunspaceId.Guid -ErrorAction Stop }
    catch {
        try {
            $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential (Get-Credential) -Authentication Basic -AllowRedirection -ErrorAction Stop
            Import-PSSession $session -ErrorAction Stop | Out-Null
            }
        catch { Write-Error "No active Exchange Online session detected, please connect to ExO first: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop; return }
    }

    #Get a list of all recipients that support ManagedBy/Owner attribute
    $O365Groups = Invoke-Command -Session $session -ScriptBlock { Get-UnifiedGroup -ResultSize Unlimited | Select-Object -Property Displayname,PrimarySMTPAddress,ExternalDirectoryObjectId } -HideComputerName 

    #If no objects are returned from the above cmdlet, stop the script and inform the user
    if (!$O365Groups) { Write-Error "No Office 365 groups found" -ErrorAction Stop }

    #Once we have the O365 Groups list, cycle over each group to gather membership
    $arrMembers = @()
    $count = 1; $PercentComplete = 0;
    foreach ($o in $O365Groups) {
        #Progress message
        $ActivityMessage = "Retrieving data for mailbox $($o.DisplayName). Please wait..."
        $StatusMessage = ("Processing mailbox {0} of {1}: {2}" -f $count, @($O365Groups).count, $o.PrimarySmtpAddress.ToString())
        $PercentComplete = ($count / @($O365Groups).count * 100)
        Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
        $count++

        #Add some artificial delay to combat throttling
        Start-Sleep -Milliseconds 500

        #Gather the LINKS for each Group
        $oMembers = Invoke-Command -Session $session -ScriptBlock { Get-UnifiedGroupLinks -Identity $using:o.ExternalDirectoryObjectId -LinkType Members -ResultSize Unlimited | Select-Object -Property WindowsLiveID, RecipientTypeDetails} -HideComputerName
        $oGuests = $oMembers | ? {$_.RecipientTypeDetails.ToString() -eq "GuestMailUser"}
        $oOwners = Invoke-Command -Session $session -ScriptBlock { Get-UnifiedGroupLinks -Identity $using:o.ExternalDirectoryObjectId -LinkType Owners -ResultSize Unlimited | Select-Object -Property WindowsLiveID, RecipientTypeDetails} -HideComputerName
        $oSubscribers = Invoke-Command -Session $session -ScriptBlock { Get-UnifiedGroupLinks -Identity $using:o.ExternalDirectoryObjectId -LinkType Subscribers -ResultSize Unlimited | Select-Object -Property WindowsLiveID, RecipientTypeDetails} -HideComputerName

        #If NOT using the $condensedoutput switch, each individual Link will be listed on a single line in the output
        if (!$CondensedOutput) {
            #Make sure to add Aggregators and EventSubscribers once they start working
            foreach ($oMember in $oMembers) {
                #Prepare the output object
                $objMember = New-Object PSObject
                $i++;Add-Member -InputObject $objMember -MemberType NoteProperty -Name "Number" -Value $i
                Add-Member -InputObject $objMember -MemberType NoteProperty -Name "DisplayName" -Value $o.DisplayName
                Add-Member -InputObject $objMember -MemberType NoteProperty -Name "PrimarySMTPAddress" -Value $o.PrimarySMTPAddress.ToString()
                Add-Member -InputObject $objMember -MemberType NoteProperty -Name "Member" -Value $oMember.WindowsLiveID.ToString()
                Add-Member -InputObject $objMember -MemberType NoteProperty -Name "MemberType" -Value "Member"
                $arrMembers += $objMember
            }
            foreach ($oMember in $oOwners) {
                #Prepare the output object
                $objMember = New-Object PSObject
                $i++;Add-Member -InputObject $objMember -MemberType NoteProperty -Name "Number" -Value $i
                Add-Member -InputObject $objMember -MemberType NoteProperty -Name "DisplayName" -Value $o.DisplayName
                Add-Member -InputObject $objMember -MemberType NoteProperty -Name "PrimarySMTPAddress" -Value $o.PrimarySMTPAddress.ToString()
                Add-Member -InputObject $objMember -MemberType NoteProperty -Name "Member" -Value $oMember.WindowsLiveID.ToString()
                Add-Member -InputObject $objMember -MemberType NoteProperty -Name "MemberType" -Value "Owner"
                $arrMembers += $objMember
            }
            foreach ($oMember in $oSubscribers) {
                #Prepare the output object
                $objMember = New-Object PSObject
                $i++;Add-Member -InputObject $objMember -MemberType NoteProperty -Name "Number" -Value $i
                Add-Member -InputObject $objMember -MemberType NoteProperty -Name "DisplayName" -Value $o.DisplayName
                Add-Member -InputObject $objMember -MemberType NoteProperty -Name "PrimarySMTPAddress" -Value $o.PrimarySMTPAddress.ToString()
                Add-Member -InputObject $objMember -MemberType NoteProperty -Name "Member" -Value $oMember.WindowsLiveID.ToString()
                Add-Member -InputObject $objMember -MemberType NoteProperty -Name "MemberType" -Value "Subscriber"
                $arrMembers += $objMember
            }
            foreach ($oMember in $oGuests) {
                #Prepare the output object
                $objMember = New-Object PSObject
                $i++;Add-Member -InputObject $objMember -MemberType NoteProperty -Name "Number" -Value $i
                Add-Member -InputObject $objMember -MemberType NoteProperty -Name "DisplayName" -Value $o.DisplayName
                Add-Member -InputObject $objMember -MemberType NoteProperty -Name "PrimarySMTPAddress" -Value $o.PrimarySMTPAddress.ToString()
                Add-Member -InputObject $objMember -MemberType NoteProperty -Name "Member" -Value $oMember.WindowsLiveID.ToString()
                Add-Member -InputObject $objMember -MemberType NoteProperty -Name "MemberType" -Value "Guest"
                $arrMembers += $objMember
            }
        }

        else {
            #If using condensed output, use single line per Group
            #Make sure to add Aggregators and EventSubscribers once they start working
            $o | Add-Member "Owners" $($oOwners.WindowsLiveID -join ";")
            $o | Add-Member "Members" $($oMembers.WindowsLiveID -join ";")
            $o | Add-Member "Subscribers" $($oSubscribers.WindowsLiveID -join ";")
            $o | Add-Member "Guests" (&{If ($oGuests) {$($oGuests.WindowsLiveID -join ",")} else {""}})
            $arrMembers += $o
        }}

    #Return the output
    $arrMembers | select * -ExcludeProperty Number,PSComputerName,RunspaceId,PSShowComputerName,ExternalDirectoryObjectId
}

#Get the Office 365 Group membership report
Get-O365GroupMembershipInventory @PSBoundParameters -OutVariable global:varO365GroupMembers #| Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_O365GroupMembers.csv" -NoTypeInformation
#Write-Host "Office 365 Groups membership report data is stored in the `$varO365GroupMembers global variable" -ForegroundColor Cyan