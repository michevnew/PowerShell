#Requires -Version 3.0
[CmdletBinding()] #Make sure we can use -Verbose
Param([switch]$IncludeAll,[switch]$IncludeDGs,[switch]$IncludeDynamicDGs,[switch]$IncludeO365Groups,[switch]$RecursiveOutput,[switch]$RecursiveOutputListGroups)

function Check-Connectivity {
    #Make sure we are connected to Exchange Remote PowerShell
    Write-Verbose "Checking connectivity to Exchange Remote PowerShell..."
    if (!$session -or ($session.State -ne "Opened")) {
        try { $script:session = Get-PSSession -InstanceId (Get-AcceptedDomain | select -First 1).RunspaceId.Guid -ErrorAction Stop  }
        catch {
            try {
                #Failing to detect an active session, try connecting to ExO via Basic auth...
                $script:session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential (Get-Credential) -Authentication Basic -AllowRedirection -ErrorAction Stop 
                Import-PSSession $session -ErrorAction Stop | Out-Null 
            }  
            catch { Write-Error "No active Exchange Remote PowerShell session detected, please connect first. To connect to ExO: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop }
        }
    }

    #As the function is called every once in a while, use it to trigger some artifical delay in order to prevent throttling
    Start-Sleep -Milliseconds 300
    return $true
}


function Get-GroupMemberRecursive {
<#
.Synopsis
    List all members of a given group, including nested groups
.DESCRIPTION
    The Get-GroupMemberRecursive cmdlet lists all members of the specified group, and can be used to also expand the members of any nested groups

.EXAMPLE
    Get-GroupMemberRecursive group@domain.com

    This command will return a list of direct members of the group@domain.com group

.EXAMPLE
    Get-DistributionGroup new | Get-GroupMemberRecursive

    The command accepts pipeline input (unlike Get-DistributionGroup new | Get-DistributionGroupMember)!

.EXAMPLE
    Get-GroupMemberRecursive -Identity group@domain.com -OutVariable var
    $var | Export-Csv -NoTypeInformation "accessrights.csv"

    To export the results to a CSV file, use the OutVariable parameter.
.INPUTS
    Distribution group identifier
.OUTPUTS
    Array with basic information about the group and list of all members.
#>    
    
    [CmdletBinding()]

    Param(
    #Use the Identity parameter to provide an unique identifier for the group object.
    [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,ValueFromRemainingArguments=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$Identity,
    #Specify whether to recursively expand membership of any nested groups. Default value is $false, meaning the primary SMTP address of the group is returned instead of any members it might contain.
    [switch]$RecursiveOutput,
    #Specify whether to include an entry for any nested groups in the output object, or just their expanded member objects.
    [Switch]$RecursiveOutputListGroups)

    if (!(Check-Connectivity)) { return }

    Write-Verbose "Processing single group entry $Identity with RecursiveOutput set to $RecursiveOutput"
    #Get the group object. If additional properties of the group are required, make sure to add them to the script block!
    $DG = Invoke-Command -Session $session -ScriptBlock { Get-Recipient -Identity $using:Identity -RecipientTypeDetails MailUniversalDistributionGroup,DynamicDistributionGroup,MailUniversalSecurityGroup,GroupMailbox -ErrorAction SilentlyContinue | Select-Object -Property Name,PrimarySmtpAddress,Guid,RecipientTypeDetails } -HideComputerName
    if (!$DG) { Throw "Group $Identity not found" }
    
    #Prepare the output object.
    $members = New-Object System.Collections.ArrayList
    #Use the hash table to prevent infinite looping in Get-Membership. This is the only reason we're using a separate funciton.
    $processed = @{}; $processed[$Identity] = $dg.Guid.Guid
    #This variable is used to feed info on the presence of nested Groups.
    $script:HasNestedGroups = $false
    Write-Verbose "Checking whether nested groups were detected: $HasNestedGroups"

    #Do the actual "membership" part.
    Get-Membership -Group $DG -RecursiveOutput:$RecursiveOutput -RecursiveOutputListGroups:$RecursiveOutputListGroups

    #Make sure we return an unique-valued identifier for each member.
    $members = $members | select @{n="Identifier";e={if ($_.PrimarySmtpAddress) { $_.PrimarySmtpAddress } else {$_.UserPrincipalName}}}
    #$members | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MembershipReport_$($DG.Name).csv" -NoTypeInformation -Encoding UTF8 -UseCulture 
    return $members
}


function Get-Membership {

    #DO NOT CALL DIRECTLY!
    [CmdletBinding()]

    Param(
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$Group,
    [switch]$RecursiveOutput,
    [Switch]$RecursiveOutputListGroups)

    #We expect a valid Recipient type!
    if (!$group.RecipientTypeDetails) { return }

    #Process membership depending on the recipient type.
    Write-Verbose "Processing group $($group.Name) of type $($group.RecipientTypeDetails.Value) ..."
    if ($group.RecipientTypeDetails.Value -eq "GroupMailbox") { $list = Invoke-Command -Session $session -ScriptBlock { Get-UnifiedGroupLinks -Identity ($using:group).PrimarySmtpAddress.ToString() -ResultSize Unlimited -LinkType Member } -HideComputerName | Select-Object -Property Name,WindowsLiveID,UserPrincipalName,PrimarySmtpAddress,Guid,RecipientTypeDetails,ExternalEmailAddress,ExternalDirectoryObjectId }
    elseif ($group.RecipientTypeDetails.Value -eq "DynamicDistributionGroup") { 
        $filter = (Get-DynamicDistributionGroup $group.PrimarySmtpAddress.ToString()).RecipientFilter
        $list = Invoke-Command -Session $session -ScriptBlock { Get-Recipient -RecipientPreviewFilter $using:filter -ResultSize Unlimited | Select-Object -Property Name,WindowsLiveID,UserPrincipalName,PrimarySmtpAddress,Guid,RecipientTypeDetails,ExternalEmailAddress,ExternalDirectoryObjectId } -HideComputerName
    }
    elseif ($group.RecipientTypeDetails.Value -eq "RoomList") { Write-Verbose "Skipping group $($group.Name) of type RoomList"; continue } #Just in case
    elseif ($group.RecipientTypeDetails.Value -eq "ExchangeSecurityGroup") { Write-Verbose "Skipping group $($group.Name) of type ExchangeSecurityGroup as those groups cannot be handled by Exchange cmdlets..."; continue }
    else { $list = Invoke-Command -Session $session -ScriptBlock { Get-DistributionGroupMember ($using:group).PrimarySmtpAddress.ToString() -ResultSize Unlimited | Select-Object -Property Name,WindowsLiveID,UserPrincipalName,PrimarySmtpAddress,Guid,RecipientTypeDetails,ExternalEmailAddress,ExternalDirectoryObjectId } -HideComputerName }

    #Loop over each member and process them accordingly...
    Write-Verbose "A total of $(&{If ($list) { ($list | measure).count} else {0}}) entries found, processing..."
    foreach ($l in $list) {
        Write-Verbose "Processig $($l.Name) ..."
        #Check whether we have already processed this object and if so, skip it.
        if ($l.Guid.Guid -eq $group.Guid.Guid -or $processed.ContainsValue($l.Guid.Guid)) { Write-Verbose "Recusrion detected, aborting..."; continue } 
        
        #If the object is not yet processed, and is of type Group, toggle the variable to signal presence of nested groups.
        if ($l.RecipientTypeDetails.Value -match "Group") {
            $script:HasNestedGroups = $true
            Write-Verbose "Signaling that nested groups were detected: $HasNestedGroups"
        }
        #If the object is not yet processed, and is of type Group and the function was called with the $RecursiveOutput switch, call the Get-Membership function again to expand its membership...
        if ($l.RecipientTypeDetails.Value -match "Group" -and $RecursiveOutput) {
            Write-Verbose "Processing group $($l.Name) of type $($l.RecipientTypeDetails.Value) because RecursiveOutput is set to $true"

            #If using the $RecursiveOutputListGroups switch, add an entry to the output object.
            if ($RecursiveOutputListGroups) {
                $obj = New-Object PSObject
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "Name" -Value $l.DisplayName
                #We are using the same object schema as with other recipient types, thus the UserPrincipalName property. But populate it with the Group GUID instead.
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "UserPrincipalName" -Value (&{If ($l.RecipientTypeDetails.Value -ne "ExchangeSecurityGroup") { $l.Guid.Guid } else { $l.ExternalDirectoryObjectId }})
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "PrimarySmtpAddress" -Value $l.PrimarySmtpAddress.ToString()
                $members.Add($obj) > $Null
            }

            #Recursively process the nested group(s), while keeping track of the objects we've already proessed.
            $processed[$l.PrimarySmtpAddress.ToString()] = $l.Guid.Guid
            Get-Membership -Group $l -RecursiveOutput:$RecursiveOutput -RecursiveOutputListGroups:$RecursiveOutputListGroups
        }
        #Otherwise return the flattened list of members...
        else {
            #Prepare the output object.
            $obj = New-Object PSObject
            # Use UserPrincipalName for Users, MailUsers; use WindowsLiveID for GuestMailUsers; use GUID for Mail Contacts.
            Add-Member -InputObject $obj -MemberType NoteProperty -Name "UserPrincipalName" -Value (&{If($l.UserPrincipalName) { $l.UserPrincipalName } Else { &{If($l.WindowsLiveID.Length) {$l.WindowsLiveID.ToString()} else { $l.Guid.Guid } }}})
            # Use PrimarySmtpAddress where exists, ExternalEmailAddress for Mail Contacts and GuestMailUsers, return empty string for User objects.
            Add-Member -InputObject $obj -MemberType NoteProperty -Name "PrimarySmtpAddress" -Value (&{If($l.PrimarySmtpAddress.Length) { $l.PrimarySmtpAddress.ToString() } Else { &{If ($l.ExternalEmailAddress) { $l.ExternalEmailAddress.ToString().Replace("SMTP:","") } else { "" } }}})
            $members.Add($obj) > $Null
        }
    Write-Verbose "End Processig $($l.Name) ..."
    }
}

function Get-GroupMembershipReport {
<#
.Synopsis
    Lists members of all groups of the selected type(s).
.DESCRIPTION
    The Get-GroupMembershipReport cmdlet enumerates all group objects of the selected type(s) and lists their membership.
    Running the cmdlet without parameters will return direct members of all Distribution groups and Mail-enabled Security Groups in the organization. To include other group type(s), use the corresponding switch parameter or -IncludeAll.
    Membership of nested groups is NOT returned by default, you need to specify the -RecursiveOutput switch when running the cmdlet/script.
    To specify a variable in which to hold the cmdlet output, use the -OutVariable parameter.

.EXAMPLE
    Get-GroupMembershipReport -IncludeDGs

    This command will return a list of direct members for all Distribution groups in the tenant.

.EXAMPLE
    Get-GroupMembershipReport -IncludeO365Groups -RecursiveOutput

    This command will return a list of direct and indirect members for all Office 365 Groups in the tenant.

.EXAMPLE
    Get-GroupMembershipReport -IncludeAll -OutVariable global:var
    $var | Export-Csv -NoTypeInformation "members.csv"

    To export the results to a CSV file, use the OutVariable parameter.
.INPUTS
    None.
.OUTPUTS
    Array with information about the group, its managers and any members found.
#>

    [CmdletBinding(DefaultParameterSetName='None')]
    
    Param
    (
    #Specify whether to include "regular" DGs in the result.
    [Switch]$IncludeDGs,
    #Specify whether to include dynamic DGs in the result.
    [Switch]$IncludeDynamicDGs,
    #Specify whether to include Office 365 Groups in the result.
    [Switch]$IncludeO365Groups,
    #Specify whether to include all groups in the result.
    [Switch]$IncludeAll,
    #Specify whether to recursively expand membership of any nested groups. Default value is $false, meaning the primary SMTP address of the group is returned instead of any members it might contain.
    [Switch]$RecursiveOutput,
    #Specify whether to include an entry for any nested groups in the output object, or just their expanded member objects.
    [Switch]$RecursiveOutputListGroups)

    #Initialize the parameters
    if (!$RecursiveOutput -and $RecursiveOutputListGroups) {
        $RecursiveOutputListGroups = $false
        Write-Verbose "The parameter -RecursiveOutputListGroups can only be used when the -RecursiveOutput is specified as well, ignoring..." 
    }

    #Initialize the variable used to designate group types, based on the input parameters.
    $included = @()
    if ($IncludeDynamicDGs) { $included += "DynamicDistributionGroup" }
    if ($IncludeO365Groups) { $included += "GroupMailbox" }
    #If no parameters specified, return only "standard" DGs
    if ($IncludeDGs -or !$included) { $included += "MailUniversalDistributionGroup";$included += "MailUniversalSecurityGroup" }

    #Check if we are connected to Exchange PowerShell.
    if (!(Check-Connectivity)) { return }

    #Get the list of groups, depending on the parameters specified when invoking the script. If you want to include other object types or additional properties, make sure to add them to the script blocks below!
    if ($IncludeAll) {
        $Groups = Invoke-Command -Session $session -ScriptBlock { Get-Recipient -ResultSize Unlimited -RecipientTypeDetails MailUniversalDistributionGroup,DynamicDistributionGroup,MailUniversalSecurityGroup,GroupMailbox -ErrorAction SilentlyContinue | Select-Object -Property Name,PrimarySmtpAddress,RecipientTypeDetails,ManagedBy } -HideComputerName
    }
    else {
        $Groups = Invoke-Command -Session $session -ScriptBlock { Get-Recipient -ResultSize Unlimited -RecipientTypeDetails $Using:included |  Select-Object -Property Name,PrimarySmtpAddress,RecipientTypeDetails,ManagedBy } -HideComputerName
    }
    
    #If no groups are returned from the above cmdlet, stop the script and inform the user.
    if (!$Groups) { Write-Error "No groups of the specifyied types were found, specify different criteria." -ErrorAction Stop }

    #Once we have the group list, cycle over each group to gather a list of direct or recursive members.
    $arrGroupData = @()
    $PercentComplete = 0; $count = 1;
    foreach ($GName in $Groups) {
        #Progress message
        $ActivityMessage = "Processing group $($GName.Name). Please wait..."
        $StatusMessage = ("Processing {0} of {1}: {2}" -f $count, @($Groups).count, $GName.PrimarySmtpAddress.ToString())
        $PercentComplete = ($count / @($Groups).count * 100)
        Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
        $count++    

        #Get the list of members.
	    $users = Get-GroupMemberRecursive -Identity $GName.PrimarySmtpAddress.ToString() -RecursiveOutput:$RecursiveOutput -RecursiveOutputListGroups:$RecursiveOutputListGroups
        #Filter out any duplicates and sort
        $users = $users | sort Identifier -Unique

        #Prepare the output object.
	    $objProperties = New-Object PSObject
   	    Add-Member -InputObject $objProperties -MemberType NoteProperty -Name "PrimarySmtpAddress" -Value $GName.PrimarySmtpAddress.ToString()
        Add-Member -InputObject $objProperties -MemberType NoteProperty -Name "GroupType" -Value $GName.RecipientTypeDetails.Value
        Add-Member -InputObject $objProperties -MemberType NoteProperty -Name "HasNestedGroups" -Value $HasNestedGroups
        Add-Member -InputObject $objProperties -MemberType NoteProperty -Name "ManagedBy" -Value $($GName.ManagedBy.Name -join ",") # maybe change that to UPNs, care for multiple values, etc?
        Add-Member -InputObject $objProperties -MemberType NoteProperty -Name "MemberCount" -Value $(&{If ($users) { ($users | measure).count} else {0}})
        Add-Member -InputObject $objProperties -MemberType NoteProperty -Name "Members" -Value ($users.Identifier -join ",")

	    $arrGroupData += $objProperties
    }
    
    #Output the result to the console host. Rearrange/sort as needed.
    $arrGroupData | sort PrimarySmtpAddress
}

#Invoke the Get-GroupMembershipReport function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
Get-GroupMembershipReport @PSBoundParameters -OutVariable global:varGroupMembership #| Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_DGMembershipReport.csv" -NoTypeInformation -Encoding UTF8 -UseCulture