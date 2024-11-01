#Requires -Version 3.0
#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.0.0" }
[CmdletBinding()] #Make sure we can use -Verbose
Param([switch]$IncludeAll,[switch]$IncludeDGs,[switch]$IncludeDynamicDGs,[switch]$IncludeO365Groups,[switch]$RecursiveOutput,[switch]$RecursiveOutputListGroups,[string[]]$GroupList)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/4394/report-on-recursive-group-membership-via-exchange-powershell

function Check-Connectivity {
    #Make sure we are connected to Exchange Online PowerShell
    Write-Verbose "Checking connectivity to Exchange Online PowerShell..."

    #Check via Get-ConnectionInformation first
    if (Get-ConnectionInformation) { return $true } #REMOVE ALL OTHER CHECKS?

    #Confirm connectivity to Exchange Online
    try { Get-EXORecipient -ResultSize 1 -ErrorAction Stop | Out-Null }
    catch {
        try { Connect-ExchangeOnline -CommandName Get-EXORecipient, Get-DistributionGroupMember, Get-DynamicDistributionGroup, Get-Recipient, Get-UnifiedGroupLinks -SkipLoadingFormatData } #custom for this script
        catch { Write-Error "No active Exchange Online session detected. To connect to ExO: https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps"; return $false }
    }

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

    #if (!(Check-Connectivity)) { return } #overkill if using the V3 module
    Start-Sleep -Milliseconds 222

    Write-Verbose "Processing single group entry $Identity with RecursiveOutput set to $RecursiveOutput"
    #Get the group object. If additional properties of the group are required, make sure to add them to the script block!
    try { $DG = Get-EXORecipient -Identity $Identity -RecipientTypeDetails MailUniversalDistributionGroup,DynamicDistributionGroup,MailUniversalSecurityGroup,GroupMailbox -ErrorAction Stop -Properties Guid | Select-Object -Property Name,PrimarySmtpAddress,Guid,RecipientTypeDetails }
    catch { Throw "Group $Identity not found" }

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
    Write-Verbose "Processing group $($group.Name) of type $($group.RecipientTypeDetails) ..."
    #MemberOfGroup filter is not supported for Get-EXORecipient, so we have to use non-REST cmdlets :/
    if ($group.RecipientTypeDetails -eq "GroupMailbox") { $list = Get-UnifiedGroupLinks -Identity $Group.PrimarySmtpAddress -ResultSize Unlimited -LinkType Member | Select-Object -Property Name,WindowsLiveID,UserPrincipalName,PrimarySmtpAddress,Guid,RecipientTypeDetails,ExternalEmailAddress,ExternalDirectoryObjectId }
    elseif ($group.RecipientTypeDetails -eq "DynamicDistributionGroup") {
        $filter = (Get-DynamicDistributionGroup $group.PrimarySmtpAddress).RecipientFilter #Get-DynamicDistributionGroupMember instead?
        #here's another place where Get-EXORecipient fails... both with -RecipientPreviewFilter and -Filter. Use Get-Recipient instead.
        $list = Get-Recipient -RecipientPreviewFilter $filter -ResultSize Unlimited | Select-Object -Property Name,WindowsLiveID,UserPrincipalName,PrimarySmtpAddress,Guid,RecipientTypeDetails,ExternalEmailAddress,ExternalDirectoryObjectId
    }
    elseif ($group.RecipientTypeDetails -eq "RoomList") { Write-Verbose "Skipping group $($group.Name) of type RoomList"; continue } #Just in case
    elseif ($group.RecipientTypeDetails -eq "ExchangeSecurityGroup") { Write-Verbose "Skipping group $($group.Name) of type ExchangeSecurityGroup as those groups cannot be handled by Exchange cmdlets..."; continue }
    else { $list = Get-DistributionGroupMember $Group.PrimarySmtpAddress -ResultSize Unlimited | Select-Object -Property Name,WindowsLiveID,UserPrincipalName,PrimarySmtpAddress,Guid,RecipientTypeDetails,ExternalEmailAddress,ExternalDirectoryObjectId }

    #Loop over each member and process them accordingly...
    Write-Verbose "A total of $(&{If ($list) { ($list | measure).count} else {0}}) entries found, processing..."
    foreach ($l in $list) {
        Write-Verbose "Processig $($l.Name) ..."
        #Check whether we have already processed this object and if so, skip it.
        if ($l.Guid.Guid -eq $group.Guid.Guid -or $processed.ContainsValue($l.Guid.Guid)) { Write-Verbose "Recursion detected, aborting..."; continue }

        #If the object is not yet processed, and is of type Group, toggle the variable to signal presence of nested groups.
        if ($l.RecipientTypeDetails -match "Group") {
            $script:HasNestedGroups = $true
            Write-Verbose "Signaling that nested groups were detected: $HasNestedGroups"
        }
        #If the object is not yet processed, and is of type Group and the function was called with the $RecursiveOutput switch, call the Get-Membership function again to expand its membership...
        if ($l.RecipientTypeDetails -match "Group" -and $RecursiveOutput) {
            Write-Verbose "Processing group $($l.Name) of type $($l.RecipientTypeDetails) because RecursiveOutput is set to $true"

            #If using the $RecursiveOutputListGroups switch, add an entry to the output object.
            if ($RecursiveOutputListGroups) {
                $obj = New-Object PSObject
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "Name" -Value $l.DisplayName
                #We are using the same object schema as with other recipient types, thus the UserPrincipalName property. But populate it with the Group GUID instead.
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "UserPrincipalName" -Value (&{If ($l.RecipientTypeDetails -ne "ExchangeSecurityGroup") { $l.Guid.Guid } else { $l.ExternalDirectoryObjectId }})
                Add-Member -InputObject $obj -MemberType NoteProperty -Name "PrimarySmtpAddress" -Value $l.PrimarySmtpAddress
                $members.Add($obj) > $Null
            }

            #Recursively process the nested group(s), while keeping track of the objects we've already proessed.
            $processed[$l.PrimarySmtpAddress] = $l.Guid.Guid
            Get-Membership -Group $l -RecursiveOutput:$RecursiveOutput -RecursiveOutputListGroups:$RecursiveOutputListGroups
        }
        #Otherwise return the flattened list of members...
        else {
            #Prepare the output object.
            $obj = New-Object PSObject
            # Use UserPrincipalName for Users, MailUsers; use WindowsLiveID for GuestMailUsers; use GUID for Mail Contacts.
            Add-Member -InputObject $obj -MemberType NoteProperty -Name "UserPrincipalName" -Value (&{If($l.UserPrincipalName) { $l.UserPrincipalName } Else { &{If($l.WindowsLiveID.Length) {$l.WindowsLiveID} else { $l.Guid.Guid } }}})
            #Override for Service principal objects, as GUID is not suitable
            if ($l.RecipientTypeDetails -eq "ServicePrinciple") { $obj.UserPrincipalName = $l.ExternalDirectoryObjectId }
            # Use PrimarySmtpAddress where exists, ExternalEmailAddress for Mail Contacts and GuestMailUsers, return empty string for User objects.
            Add-Member -InputObject $obj -MemberType NoteProperty -Name "PrimarySmtpAddress" -Value (&{If($l.PrimarySmtpAddress.Length) { $l.PrimarySmtpAddress } Else { &{If ($l.ExternalEmailAddress) { $l.ExternalEmailAddress.Replace("SMTP:","") } else { "" } }}})
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
    Get-GroupMembershipReport -GroupList (Import-Csv .\Groups.csv).PrimarySmtpAddress

    Generate the report for a subset of the groups in the tenant, imported via an CSV file or array of email addresses.

.EXAMPLE
    Get-GroupMembershipReport -IncludeAll -OutVariable global:var
    $var | Export-Csv -NoTypeInformation "members.csv"

    To export the results to a CSV file, use the OutVariable parameter.
.INPUTS
    None.
.OUTPUTS
    Array with information about the group, its managers and any members found.
#>

    [CmdletBinding(DefaultParameterSetName='ByGroupType')]

    Param
    (
    #Specify whether to include "regular" DGs in the result.
    [Parameter(ParameterSetName = 'ByGroupType')][Switch]$IncludeDGs,
    #Specify whether to include dynamic DGs in the result.
    [Parameter(ParameterSetName = 'ByGroupType')][Switch]$IncludeDynamicDGs,
    #Specify whether to include Office 365 Groups in the result.
    [Parameter(ParameterSetName = 'ByGroupType')][Switch]$IncludeO365Groups,
    #Specify whether to include all groups in the result.
    [Parameter(ParameterSetName = 'ByGroupType')][Switch]$IncludeAll,
    #Specify whether to recursively expand membership of any nested groups. Default value is $false, meaning the primary SMTP address of the group is returned instead of any members it might contain.
    [Parameter(ParameterSetName = 'ByGroup')]
    [Parameter(ParameterSetName = 'ByGroupType')]
    [Switch]$RecursiveOutput,
    #Specify whether to include an entry for any nested groups in the output object, or just their expanded member objects.
    [Parameter(ParameterSetName = 'ByGroup')]
    [Parameter(ParameterSetName = 'ByGroupType')]
    [Switch]$RecursiveOutputListGroups,
    #Specify the list of groups to cover, by passing an array value
    [Parameter(ParameterSetName = 'ByGroup')][string[]]$GroupList)

    #Initialize the parameters
    if (!$RecursiveOutput -and $RecursiveOutputListGroups) {
        $RecursiveOutputListGroups = $false
        Write-Verbose "The parameter -RecursiveOutputListGroups can only be used when the -RecursiveOutput is specified as well, ignoring..."
    }

    #Check if we are connected to Exchange Online PowerShell.
    if (!(Check-Connectivity)) { return }

    #region GroupList
    $Groups = @()

    #If running the script against a list of groups
    if ($GroupList) {
        Write-Verbose "Running the script against the provided list of groups..."
        foreach ($Group in $GroupList) {
            #Filter this out if you want to pass other group identifiers
            #try { $null = [mailaddress]($Group) }
            #catch { Write-Verbose "Entry $group does not contain a valid SMTP value, removing..."; continue }

            try { $gres = Get-EXORecipient $Group -RecipientTypeDetails MailUniversalDistributionGroup,DynamicDistributionGroup,MailUniversalSecurityGroup,GroupMailbox -ErrorAction Stop -Properties ManagedBy | Select-Object -Property Name,PrimarySmtpAddress,RecipientTypeDetails,ManagedBy }
            catch { Write-Verbose "Entry $group does not match a valid group recipient in your tenant, removing..."; continue }

            $Groups += $gres
        }
    }

    #If running the script against specific group types
    else {
        #Initialize the variable used to designate group types, based on the input parameters.
        $included = @()
        if ($IncludeDynamicDGs) { $included += "DynamicDistributionGroup" }
        if ($IncludeO365Groups) { $included += "GroupMailbox" }
        #If no parameters specified, return only "standard" DGs
        if ($IncludeDGs -or !$included) { $included += "MailUniversalDistributionGroup";$included += "MailUniversalSecurityGroup" }

        #Get the list of groups, depending on the parameters specified when invoking the script. If you want to include other object types or additional properties, make sure to add them to the script blocks below!
        if ($IncludeAll) {
            $Groups = Get-EXORecipient -ResultSize Unlimited -RecipientTypeDetails MailUniversalDistributionGroup,DynamicDistributionGroup,MailUniversalSecurityGroup,GroupMailbox -ErrorAction SilentlyContinue -Properties ManagedBy | Select-Object -Property Name,PrimarySmtpAddress,RecipientTypeDetails,ManagedBy
        }
        else {
            $Groups = Get-EXORecipient -ResultSize Unlimited -RecipientTypeDetails $included -ErrorAction SilentlyContinue -Properties ManagedBy | Select-Object -Property Name,PrimarySmtpAddress,RecipientTypeDetails,ManagedBy
        }
    }

    #If no groups are returned from the above cmdlet, stop the script and inform the user.
    if (!$Groups) { Write-Error "No groups of the specifyied types were found, specify different criteria." -ErrorAction Stop }
    #endregion GroupList

    #Filter out any potential duplicates
    $Groups = ($Groups | Sort-Object -Unique -Property PrimarySmtpAddress)

    #Once we have the group list, cycle over each group to gather a list of direct or recursive members.
    $arrGroupData = @()
    $PercentComplete = 0; $count = 1;
    foreach ($GName in $Groups) {
        #Progress message
        $ActivityMessage = "Processing group $($GName.Name). Please wait..."
        $StatusMessage = ("Processing {0} of {1}: {2}" -f $count, @($Groups).count, $GName.PrimarySmtpAddress)
        $PercentComplete = ($count / @($Groups).count * 100)
        Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
        $count++

        #Get the list of members.
	    $members = Get-GroupMemberRecursive -Identity $GName.PrimarySmtpAddress -RecursiveOutput:$RecursiveOutput -RecursiveOutputListGroups:$RecursiveOutputListGroups
        #Filter out any duplicates and sort
        $members = $members | Sort-Object Identifier -Unique

        #Prepare the output object.
	    $objProperties = New-Object PSObject
        Add-Member -InputObject $objProperties -MemberType NoteProperty -Name "Name" -Value $GName.Name
   	    Add-Member -InputObject $objProperties -MemberType NoteProperty -Name "PrimarySmtpAddress" -Value $GName.PrimarySmtpAddress
        Add-Member -InputObject $objProperties -MemberType NoteProperty -Name "GroupType" -Value $GName.RecipientTypeDetails
        Add-Member -InputObject $objProperties -MemberType NoteProperty -Name "HasNestedGroups" -Value $HasNestedGroups
        Add-Member -InputObject $objProperties -MemberType NoteProperty -Name "ManagedBy" -Value $($GName.ManagedBy -join ",") # maybe change that to UPNs, care for multiple values, etc?
        Add-Member -InputObject $objProperties -MemberType NoteProperty -Name "MemberCount" -Value $(&{If ($members) { ($members | measure).count} else {0}})
        Add-Member -InputObject $objProperties -MemberType NoteProperty -Name "Members" -Value ($members.Identifier -join ",")

	    $arrGroupData += $objProperties
    }

    #Output the result to the console host. Rearrange/sort as needed.
    $arrGroupData | Sort-Object PrimarySmtpAddress
}

#Invoke the Get-GroupMembershipReport function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
Get-GroupMembershipReport @PSBoundParameters -OutVariable global:varGroupMembership | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_DGMembershipReport.csv" -NoTypeInformation -Encoding UTF8 -UseCulture