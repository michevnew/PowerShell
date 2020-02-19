param([switch]$IncludeAll,[switch]$IncludeUserMailboxes,[switch]$IncludeSharedMailboxes,[switch]$IncludeMailUsers,[switch]$IncludeMailContacts,[switch]$IncludeGuestUsers)

function Check-Connectivity {
    #Make sure we are connected to Exchange Remote PowerShell
    Write-Verbose "Checking connectivity to Exchange Remote PowerShell..."
    if (!$session -or ($session.State -ne "Opened")) {
        try { $script:session = Get-PSSession -InstanceId (Get-AcceptedDomain | select -First 1).RunspaceId.Guid -ErrorAction Stop  }
        catch {
            try { 
                $script:session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential (Get-Credential) -Authentication Basic -AllowRedirection -ErrorAction Stop 
                Import-PSSession $session -ErrorAction Stop | Out-Null 
            }  
            catch { Write-Error "No active Exchange Remote PowerShell session detected, please connect first. To connect to ExO: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop }
        }
    }

    return $true
}


function Get-DGMembershipInventory {
<#
.Synopsis
    Lists all groups for which a given user is a member of, for all of the selected user types.
.DESCRIPTION
    The Get-DGMembershipInventory cmdlet prepares a list of all users of the selected type(s) and for each user, returns all groups he is a member of. Running the cmdlet without parameters will return entries for all User mailboxes.
    Specifying particular user type(s) can be done with the corresponding parameter. To specify a variable in which to hold the cmdlet output, use the OutVariable parameter.

.EXAMPLE
    Get-DGMembershipInventory -IncludeUserMailboxes

    This command will return a list of groups for each user mailboxes in the company

.EXAMPLE
    Get-DGMembershipInventory -IncludeAll -OutVariable global:var
    $var | Export-Csv -NoTypeInformation "DGMemberOf.csv"

    To export the results to a CSV file, use the OutVariable parameter.
.INPUTS
    None.
.OUTPUTS
    Array with information about the user and a list of groups he's a member of
#>

    [CmdletBinding()]
    
    Param
    (
    #Specify whether to include user mailboxes in the result
    [Switch]$IncludeUserMailboxes,    
    #Specify whether to include Shared mailboxes in the result
    [Switch]$IncludeSharedMailboxes,
    #Specify whether to include Mail users in the result
    [Switch]$IncludeMailUsers,
    #Specify whether to include Mail contacts in the result
    [Switch]$IncludeMailContacts,
    #Specify whether to include Guest (Mail) users in the result
    [Switch]$IncludeGuestUsers,
    #Specify whether to include every type of recipient in the result
    [Switch]$IncludeAll)

    #Initialize the variable used to designate recipient types, based on the script parameters
    $included = @()
    if($IncludeUserMailboxes) { $included += "UserMailbox"}
    if($IncludeSharedMailboxes) { $included += "SharedMailbox"}
    if($IncludeMailUsers) { $included += "MailUser"}
    if($IncludeMailContacts) { $included += "MailContact"}
    if($IncludeGuestUsers) { $included += "GuestMailUser"}
    
    #Check if we are connected to Exchange PowerShell
    if (!(Check-Connectivity)) { return }

    #Get the list of users, depending on the parameters specified when invoking the script
    if ($IncludeAll) {
        $MBList = Invoke-Command -Session $session -ScriptBlock { Get-Recipient -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,MailUser,MailContact,GuestMailUser | Select-Object -Property PrimarySmtpAddress,DistinguishedName,ExternalEmailAddress,ExternalDirectoryObjectId } -HideComputerName
    }
    elseif (!$included -or ($included -eq "UserMailbox" -and $Included.Length -eq 1)) {
        $MBList = Invoke-Command -Session $session -ScriptBlock { Get-Recipient gosho -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Select-Object -Property PrimarySmtpAddress,DistinguishedName,ExternalEmailAddress,ExternalDirectoryObjectId } -HideComputerName
    }
    else {
        $MBList = Invoke-Command -Session $session -ScriptBlock { Get-Recipient -ResultSize Unlimited -RecipientTypeDetails $Using:included | Select-Object -Property PrimarySmtpAddress,DistinguishedName,ExternalEmailAddress,ExternalDirectoryObjectId } -HideComputerName
    }
    
    #If no users are returned from the above cmdlet, stop the script and inform the user
    if (!$MBList) { Write-Error "No users of the specifyied types were found, specify different criteria." -ErrorAction Stop }

    #prepare the output
    $arrMembers = @(); $count = 1; $PercentComplete = 0;

    #cycle over each object from the list
    foreach ($mailbox in $MBList) { 
        #display a simple progress message
        $ActivityMessage = "Retrieving data for mailbox $($mailbox.PrimarySmtpAddress). Please wait..."
        $StatusMessage = ("Processing {0} of {1}: {2}" -f $count, @($MBList).count, $mailbox.DistinguishedName)
        $PercentComplete = ($count / @($MBList).count * 100)
        Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
        $count++

        #Add some delay to avoid throttling
        #Start-Sleep -Milliseconds 80 #Add some delay to avoid throttling...

        #use server-side filtering to obtain the list of groups a given user is a member of
        $dn =  $mailbox.DistinguishedName
        $dnnew = "'" + "$($dn.Replace("'","''"))" + "'" #handle ' in DN for Invoke-Command (only '' needed)
        $cmd =  'Invoke-Command -Session $session -ScriptBlock { Get-Recipient -Filter "Members -eq $using:dnnew" | Select-Object PrimarySmtpAddress } -HideComputerName'
        $list = Invoke-Expression $cmd

        #save the output
        $objMembers = New-Object PSObject
        #$i++;Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Number" -Value $i
        #Add-Member -InputObject $objMembers -MemberType NoteProperty -Name "User" -Value $mailbox.UserPrincipalName #Not returned via Get-Recipient
        if ($mailbox.PrimarySmtpAddress.Address) { Add-Member -InputObject $objMembers -MemberType NoteProperty -Name "Email" -Value $mailbox.PrimarySmtpAddress.ToString() } 
        elseif ($mailbox.ExternalEmailAddress) { Add-Member -InputObject $objMembers -MemberType NoteProperty -Name "Email" -Value $mailbox.ExternalEmailAddress.AddressString.ToString() }
        else { Add-Member -InputObject $objMembers -MemberType NoteProperty -Name "Email" -Value $mailbox.ExternalDirectoryObjectId.ToString() } 
        Add-Member -InputObject $objMembers -MemberType NoteProperty -Name "Groups" -Value $($list.PrimarySmtpAddress -join ",")
        $arrMembers += $objMembers
    }

    #return the output
    $arrMembers | select Email,Groups
}

#Invoke the Get-DGMembershipInventory function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
Get-DGMembershipInventory @PSBoundParameters -OutVariable global:varDGMemberOf #| Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_DGMemberOfReport.csv" -NoTypeInformation