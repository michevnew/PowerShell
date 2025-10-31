#For details on what the script does and how to run it, check: https://michev.info/blog/post/7028/powershell-cmdlets-to-add-or-remove-resource-delegates-in-bulk

function Add-RoomDelegate {
<#
.Synopsis
    Adds a (resource) delegate to the specified Room mailbox(es) while preserving the rest of the delegates.
.DESCRIPTION
    The Add-RoomDelegate cmdlet processes all specified mailbox(es) and adds the specified user(s) to the list of (resource) delegates.

.EXAMPLE
    Add-RoomDelegate -Mailbox room@domain.com -User userA@domain.com,userB@domain.com

    This cmdlet will add userA and userB as delegates for the room@domain.com mailbox.

.EXAMPLE
    Add-RoomDelegate -Mailbox (Get-Mailbox -RecipientTypeDetails RoomMailbox) -User userA@domain.com

    This cmdlet will add userA as (resource) delegate for all Room mailboxes in the company.

.INPUTS
   MailboxIdParameter
   UserIdParameter
   String
.OUTPUTS
   None
#>

    [CmdletBinding(SupportsShouldProcess)]
    Param(
    <#The Mailbox parameter specifies the identity of one or more room mailboxes.

This parameter accepts the following values:
* Alias: JPhillips
* Display Name: Jeff Phillips
* Distinguished Name (DN): CN=JPhillips,CN=Users,DC=Atlanta,DC=Corp,DC=contoso,DC=com
* ExternalDirectoryObjectId: 584b2b38-888c-4d58-9d15-5af57d0354c2
* GUID: fb456636-fe7d-4d58-9d15-5af57d0354c2
* Legacy Exchange DN: /o=Contoso/ou=AdministrativeGroup/cn=Recipients/cn=JPhillips
* SMTP Address: Jeff.Phillips@contoso.com
* User Principal Name: JPhillips@contoso.com
        #>
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String[]][Alias("Identity")]$Mailbox,

    <#The User parameter specifies the identity of one or more user or group objects to designate as resource delegates.

This parameter accepts the following values:
* Alias: JPhillips
* Display Name: Jeff Phillips
* Distinguished Name (DN): CN=JPhillips,CN=Users,DC=Atlanta,DC=Corp,DC=contoso,DC=com
* ExternalDirectoryObjectId: 584b2b38-888c-4d58-9d15-5af57d0354c2
* GUID: fb456636-fe7d-4d58-9d15-5af57d0354c2
* Legacy Exchange DN: /o=Contoso/ou=AdministrativeGroup/cnRecipients/cn=JPhillips
* SMTP Address: Jeff.Phillips@contoso.com
* User Principal Name: JPhillips@contoso.com
        #>
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String[]][Alias("Delegate")]$User)

#region BEGIN
    #NO validation on ExO side, which is why we check each input value here and provide you with proper experience

    # Make sure we're connected and have access to all the required cmdlets
    try { Get-Command Get-Mailbox,Get-Recipient,Get-CalendarProcessing,Set-CalendarProcessing -ErrorAction Stop -Verbose:$false | Out-Null}
    catch { Throw "The required cmdlets were not found in the current session. Please connect to Exchange Online or Exchange Server with sufficient permissions and try again." }

    #Prepare the list of mailboxes
    Write-Verbose "Parsing the Mailbox parameter..."
    $SMTPAddresses = @{}
    foreach ($mb in $Mailbox) {
        Start-Sleep -Milliseconds 80 #Add some delay to avoid throttling...
        #Make sure a matching mailbox is found and return its Primary SMTP Address
        $SMTPAddress = Get-Mailbox $mb -RecipientTypeDetails RoomMailbox,EquipmentMailbox -ErrorAction SilentlyContinue -Verbose:$false | Select-Object -ExpandProperty PrimarySmtpAddress
        if (!$SMTPAddress) { if (!$Quiet) { Write-Warning "Resource mailbox with identifier $mb not found, skipping..." }; continue }
        elseif (($SMTPAddress.count -gt 1) -or ($SMTPAddresses[$mb]) -or ($SMTPAddresses.ContainsValue($SMTPAddress))) { Write-Warning "Multiple mailboxes matching the identifier $mb found, skipping..."; continue }
        else { $SMTPAddresses[$mb] = $SMTPAddress.ToString() }
    }
    if (!$SMTPAddresses -or ($SMTPAddresses.Count -eq 0)) { Throw "No matching mailboxes found, check the parameter values." }
    Write-Verbose "The following list of mailboxes will be used: ""$($SMTPAddresses.Values -join ", ")"""

    #Prepare the list of delegates
    Write-Verbose "Parsing the User parameter..."
    $GUIDs = @{}
    foreach ($us in $User) {
        Start-Sleep -Milliseconds 80 #Add some delay to avoid throttling...
        #Make sure a matching recipient object is found and return its Primary SMTP Address
        try { $GUID = Get-Recipient $us -RecipientTypeDetails UserMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup,GroupMailbox,DynamicDistributionGroup -ErrorAction Stop -Verbose:$false | select Name, PrimarySmtpAddress }
        catch { if (!$Quiet) { Write-Warning "Recipient with identifier $us not found, skipping..." }; continue }

        #Normalize to GUID as key, Name as value
        if (($GUID.count -gt 1) -or ($GUIDs[$us]) -or ($GUIDs.ContainsValue($GUID.Name))) { Write-Warning "Multiple recipients matching the identifier $us found, skipping..."; continue }
        else { $GUIDs[$GUID.PrimarySmtpAddress.ToString()] = $GUID.Name } #As Name is returned by ResourceDelegates
    }
    if (!$GUIDs -or ($GUIDs.Count -eq 0)) { Throw "No matching recipients found, check the parameter values." }
    Write-Verbose "The following list of delegates will be used: ""$($GUIDs.Values -join ", ")"""
#endregion BEGIN

#region PROCESS
    #Iterate over each mailbox and add the specified delegates
    foreach ($RMB in $SMTPAddresses.Values) {#should be unique, if needed select/sort
        Write-Verbose "Processing mailbox ""$RMB""..."
        Start-Sleep -Milliseconds 200 #Add some delay to avoid throttling...

        #Get the current set of delegates. Use ArrayList instead or array, because why not!
        try {
            $delegates = [System.Collections.ArrayList]@((Get-CalendarProcessing $RMB -ErrorAction Stop -Verbose:$false).ResourceDelegates)
            $delegatesold = $delegates.Count
        }
        catch {
            if ($_.Exception.Message -match "couldn't be found") { Write-Host "ERROR: mailbox ""$RMB"" not found, this should not happen..." -ForegroundColor Red ; continue }
            else {$_ | fl * -Force; continue} #catch-all for any unhandled errors
        }

        #As service-side validation is broken, we might as well add the full list in one go... should they fix it, we can always revert to adding one by one and generate output for each
        foreach ($u in $GUIDs.GetEnumerator()) {
            if (!$delegates.Contains($u.Value)) { $delegates.Add($u.Value) | Out-Null } #DO NOT use Sort or select to filter out Unique values, it will convert the ArrayList!!!!
        }
        if ($delegates.Count -eq $delegatesold) { if (!$Quiet) { Write-Host "No changes in resource delegate needed, skipping mailbox $RMB" -ForegroundColor Yellow }; continue }
        try {
            Write-Verbose "Updated list of resource delegates for mailbox ""$RMB"": $($delegates -join ", ")"
            Set-CalendarProcessing -Identity $RMB -ResourceDelegates $delegates -WhatIf:$WhatIfPreference -Confirm:$false -ErrorAction Stop
        }
        catch [System.Exception] {
            if ($_.Exception.Message -match "couldn't be found") { Write-Host "ERROR: mailbox ""MB"" not found, this should not happen..." -ForegroundColor Red }
            elseif ($_.Exception.Message -match "There are multiple recipients matching the identity") { Write-Host "ERROR: Multiple recipients matching the identifier ""MB"" found, removing from the list..." -ForegroundColor Red }
            elseif ($_.Exception.Message -match "ResourceDelegates can only be enabled on resource mailboxes.") { Write-Host "ERROR: Mailbox ""MB"" is not a resource mailbox, this should not happen..." -ForegroundColor Red }
            else {$_ | fl * -Force; continue} #catch-all for any unhandled errors
        }
        catch {$_ | fl * -Force; continue} #catch-all for any unhandled errors
    }
#endregion PROCESS

    Write-Verbose "Finish..."
}

function Remove-RoomDelegate {
<#
.Synopsis
    Removes a (resource) delegate from the specified Room mailbox(es) while preserving the rest of the delegates.
.DESCRIPTION
    The Remove-RoomDelegate cmdlet processes all specified mailbox(es) and removes the specified user(s) from the list of (resource) delegates.

.EXAMPLE
    Remove-RoomDelegate -Mailbox room@domain.com -User userA@domain.com,userB@domain.com

    This cmdlet will remove userA and userB as delegates for the room@domain.com mailbox.

.EXAMPLE
    Remove-RoomDelegate -Mailbox (Get-Mailbox -RecipientTypeDetails RoomMailbox) -User userA@domain.com

    This cmdlet will remove userA as (resource) delegate for all Room mailboxes in the company.

.INPUTS
   MailboxIdParameter
   UserIdParameter
   String
.OUTPUTS
   None
#>

    [CmdletBinding(SupportsShouldProcess)]
    Param(
    <#The Mailbox parameter specifies the identity of one or more room mailboxes.

This parameter accepts the following values:
* Alias: JPhillips
* Display Name: Jeff Phillips
* Distinguished Name (DN): CN=JPhillips,CN=Users,DC=Atlanta,DC=Corp,DC=contoso,DC=com
* ExternalDirectoryObjectId: 584b2b38-888c-4d58-9d15-5af57d0354c2
* GUID: fb456636-fe7d-4d58-9d15-5af57d0354c2
* Legacy Exchange DN: /o=Contoso/ou=AdministrativeGroup/cn=Recipients/cn=JPhillips
* SMTP Address: Jeff.Phillips@contoso.com
* User Principal Name: JPhillips@contoso.com
        #>
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String[]][Alias("Identity")]$Mailbox,

<#The User parameter specifies the identity of one or more user or group objects to designate as resource delegates.

This parameter accepts the following values:
* Alias: JPhillips
* Display Name: Jeff Phillips
* Distinguished Name (DN): CN=JPhillips,CN=Users,DC=Atlanta,DC=Corp,DC=contoso,DC=com
* ExternalDirectoryObjectId: 584b2b38-888c-4d58-9d15-5af57d0354c2
* GUID: fb456636-fe7d-4d58-9d15-5af57d0354c2
* Legacy Exchange DN: /o=Contoso/ou=AdministrativeGroup/cn=Recipients/cn=JPhillips
* SMTP Address: Jeff.Phillips@contoso.com
* User Principal Name: JPhillips@contoso.com
        #>
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [String[]][Alias("Delegate")]$User)

#region BEGIN
    #NO validation on ExO side, which is why we check each input value here and provide you with proper experience

    # Make sure we're connected and have access to all the required cmdlets
    try { Get-Command Get-Mailbox,Get-Recipient,Get-CalendarProcessing,Set-CalendarProcessing -ErrorAction Stop -Verbose:$false | Out-Null}
    catch { Throw "The required cmdlets were not found in the current session. Please connect to Exchange Online or Exchange Server with sufficient permissions and try again." }

    #Prepare the list of mailboxes
    Write-Verbose "Parsing the Mailbox parameter..."
    $SMTPAddresses = @{}
    foreach ($mb in $Mailbox) {
        Start-Sleep -Milliseconds 80 #Add some delay to avoid throttling...
        #Make sure a matching mailbox is found and return its Primary SMTP Address
        $SMTPAddress = Get-Mailbox $mb -RecipientTypeDetails RoomMailbox,EquipmentMailbox -ErrorAction SilentlyContinue -Verbose:$false | Select-Object -ExpandProperty PrimarySmtpAddress
        if (!$SMTPAddress) { if (!$Quiet) { Write-Warning "Resource mailbox with identifier $mb not found, skipping..." }; continue }
        elseif (($SMTPAddress.count -gt 1) -or ($SMTPAddresses[$mb]) -or ($SMTPAddresses.ContainsValue($SMTPAddress))) { Write-Warning "Multiple mailboxes matching the identifier $mb found, skipping..."; continue }
        else { $SMTPAddresses[$mb] = $SMTPAddress.ToString() }
    }
    if (!$SMTPAddresses -or ($SMTPAddresses.Count -eq 0)) { Throw "No matching mailboxes found, check the parameter values." }
    Write-Verbose "The following list of mailboxes will be used: ""$($SMTPAddresses.Values -join ", ")"""

    #Prepare the list of delegates
    Write-Verbose "Parsing the User parameter..."
    $GUIDs = @{}
        foreach ($us in $User) {
        Start-Sleep -Milliseconds 80 #Add some delay to avoid throttling...
        #Make sure a matching recipient object is found and return its Primary SMTP Address
        try { $GUID = Get-Recipient $us -RecipientTypeDetails UserMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup,GroupMailbox,DynamicDistributionGroup -ErrorAction Stop -Verbose:$false | select Name, PrimarySMTPAddress }
        catch { if (!$Quiet) { Write-Warning "Recipient with identifier $us not found, skipping..." }; continue }

        #Normalize to GUID as key, Name as value
        if (($GUID.count -gt 1) -or ($GUIDs[$us]) -or ($GUIDs.ContainsValue($GUID.Name))) { Write-Warning "Multiple recipients matching the identifier $us found, skipping..."; continue }
        else { $GUIDs[$GUID.PrimarySmtpAddress.ToString()] = $GUID.Name } #As Name is returned by ResourceDelegates
    }
    if (!$GUIDs -or ($GUIDs.Count -eq 0)) { Throw "No matching recipients found, check the parameter values." }
    Write-Verbose "The following list of delegates will be used: ""$($GUIDs.Values -join ", ")"""
#endregion BEGIN

#region PROCESS
    #Iterate over each mailbox and remove the specified delegates
    foreach ($RMB in $SMTPAddresses.Values) {#should be unique, if needed select/sort
        Write-Verbose "Processing mailbox ""$RMB""..."
        Start-Sleep -Milliseconds 200 #Add some delay to avoid throttling...

        #Get the current set of delegates. Use ArrayList instead or array, because why not!
        try {
            $delegates = [System.Collections.ArrayList]@((Get-CalendarProcessing $RMB -ErrorAction Stop -Verbose:$false).ResourceDelegates)
            $delegatesold = $delegates.Count
        }
        catch {
            if ($_.Exception.Message -match "couldn't be found") { Write-Host "ERROR: mailbox ""$RMB"" not found, this should not happen..." -ForegroundColor Red ; continue }
            else {$_ | fl * -Force; continue} #catch-all for any unhandled errors
        }

        #As service-side validation is broken, we might as well remove the full list in one go... should they fix it, we can always revert to removing one by one and generate output for each
        foreach ($u in $GUIDs.GetEnumerator()) {
            if ($delegates.Contains($u.Value)) { $delegates.Remove($u.Value) | Out-Null } #DO NOT use Sort or select to filter out Unique values, it will convert the ArrayList!!!!
        }
        if ($delegates.Count -eq $delegatesold) { if (!$Quiet) { Write-Host "No changes in resource delegate needed, skipping mailbox $RMB" -ForegroundColor Yellow }; continue }
        try {
            Write-Verbose "Updated list of resource delegates for mailbox ""$RMB"": $($delegates -join ", ")"
            Set-CalendarProcessing -Identity $RMB -ResourceDelegates $delegates -WhatIf:$WhatIfPreference -Confirm:$false -ErrorAction Stop
        }
        catch [System.Exception] {
            if ($_.Exception.Message -match "couldn't be found") { Write-Host "ERROR: mailbox ""MB"" not found, this should not happen..." -ForegroundColor Red }
            else {$_ | fl * -Force; continue} #catch-all for any unhandled errors
        }
        catch {$_ | fl * -Force; continue} #catch-all for any unhandled errors
    }
#endregion PROCESS

    Write-Verbose "Finish..."
}