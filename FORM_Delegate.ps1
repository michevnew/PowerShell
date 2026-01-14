param([Parameter(Mandatory)][Alias("Manager")][string]$Owner,[Parameter(Mandatory)][string]$Delegate,[switch]$ReadOnly=$false)

#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.7.2" } -Version 7.1
#For details on what the script does and how to run it, check: https://michev.info/blog/post/7306/free-tool-to-man…utlook-delegates


#Reset WhatIfPreference
$script:WhatIfPreference = $false

# Clean up before startup
Remove-Variable params,ExOSession,cUser,cFolders,cPermissions,cPermissionsNew,cFlags,cFlagsValue,FormCancelled -ErrorAction SilentlyContinue
$script:params = $PSBoundParameters

# Connect to ExO with required scopes
Write-Verbose "Connecting to Exchange Online..."
if (-not ($script:ExOSession = Get-ConnectionInformation) -and ((Get-Command *-MailboxFolderPermission).count -lt 4)) {
    # No existing session, connect. Load only the cmdlets we need.No need for Set-Mailbox, the delegate flag toggles it automatically
    # Remember that we can also leverage the REST cmdlets for GET operations, no need to load these
    Connect-ExchangeOnline -DisableWAM -CommandName *-MailboxFolderPermission,Get-MailboxFolderStatistics,Get-MailboxFolder -ShowBanner:$false -Verbose:$false -ErrorAction Stop
}
else {
    # Seems we have an existing session, reuse it
    Write-Verbose "Reusing existing Exchange Online session..."
}

# Cater to self-service scenarios, we need to know the current user
if (!(Get-Command Get-MailboxFolderStatistics -ErrorAction SilentlyContinue)) {
    Write-Verbose "Running as an end user/limited permissions account"
}
$script:cUser = ${ExOSession}?.UserPrincipalName

# Load Windows Forms assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Make sure we use the system theme for controls
[System.Windows.Forms.Application]::EnableVisualStyles()

# Main form object
$form1 = New-Object System.Windows.Forms.Form
$form1.ClientSize = New-Object System.Drawing.Size(424,309)
$form1.Name = "Form1"
$form1.Text = "Delegate Permissions: $Delegate on $Owner"
$form1.StartPosition = "CenterScreen"
$form1.FormBorderStyle = 'FixedDialog'
$form1.MaximizeBox = $false
$form1.MinimizeBox = $false
$form1.Font = New-Object System.Drawing.Font("Segoe UI SemiBold",8,[System.Drawing.FontStyle]::Regular)
$form1.KeyPreview = $True

# Main group box
$groupBox1 = New-Object System.Windows.Forms.GroupBox
$groupBox1.Name = "groupBox1"
$groupBox1.Location = New-Object System.Drawing.Size(6,8)
$groupBox1.Size = New-Object System.Drawing.Size(412,220)
$groupBox1.TabStop = $False #skips control in the tab order (tab-pressing)
$groupBox1.Text = "This delegate has the following permissions"
$form1.Controls.Add($groupBox1)

#Region Calendar Controls
# First Icon
$pictureBox1 = New-Object System.Windows.Forms.PictureBox
$pictureBox1.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\Calendar3.ico")
$pictureBox1.Location = New-Object System.Drawing.Size(10,19)
$pictureBox1.Size = New-Object System.Drawing.Size(24,24)
$pictureBox1.Name = "pictureBox1"
$pictureBox1.SizeMode = 2 #controls icon shrink/resize/etc
$pictureBox1.TabStop = $False
$groupBox1.Controls.Add($pictureBox1)

# First label
$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Size(50,29)
$label1.Size = New-Object System.Drawing.Size(50,18)
$label1.Name = "label1"
$label1.UseMnemonic = $true # Use & for ALT-shortcuts; Doesnt seem to work when combobox.DropDownStyle = DropDownList so use Add_keydown event
$label1.Text = "&Calendar"
$groupBox1.Controls.Add($label1)

# First combobox
$comboBox1 = New-Object System.Windows.Forms.ComboBox
$comboBox1.FormattingEnabled = $True
$comboBox1.Location = New-Object System.Drawing.Size(110,25)
$comboBox1.Size = New-Object System.Drawing.Size(245,21)
$comboBox1.Name = "comboBox1"
$comboBox1.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList #disable editing
$comboBox1.Items.AddRange(("None","Reviewer (can read items)","Author (can read and create items)","Editor (can read, create, and modify items)"))
$comboBox1.SelectedIndex = 0 #select based on index
$groupBox1.Controls.Add($comboBox1)

# First checkbox
$checkBox1 = New-Object System.Windows.Forms.CheckBox
$checkbox1.Location = New-Object System.Drawing.Size(54,49)
$checkbox1.Size = New-Object System.Drawing.Size(355,24)
$checkBox1.Name = "checkBox1"
$checkBox1.Text = "&Delegate receives copies of meeting-related messages sent to me"
$checkBox1.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft #text position relative to the control
$checkBox1.UseVisualStyleBackColor = $True
$groupBox1.Controls.Add($checkBox1)
#EndRegion Calendar Controls

#Region Tasks Controls
# Second Icon
$pictureBox2 = New-Object System.Windows.Forms.PictureBox
$pictureBox2.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\Tasks2.ico")
$pictureBox2.Location = New-Object System.Drawing.Size(10,76)
$pictureBox2.Size = New-Object System.Drawing.Size(24,24)
$pictureBox2.Name = "pictureBox2"
$pictureBox2.SizeMode = 2
$pictureBox2.TabStop = $False
$groupBox1.Controls.Add($pictureBox2)

# Second label
$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Size(50,85)
$label2.Size = New-Object System.Drawing.Size(50,18)
$label2.UseMnemonic = $true # Use & for ALT-shortcuts
$label2.Name = "label2"
$label2.Text = "&Tasks"
$groupBox1.Controls.Add($label2)

# Second combobox
$comboBox2 = New-Object System.Windows.Forms.ComboBox
$comboBox2.DataBindings.DefaultDataSourceUpdateMode = 0
$comboBox2.FormattingEnabled = $True
$comboBox2.Location = New-Object System.Drawing.Size(110,80)
$comboBox2.Size = New-Object System.Drawing.Size(245,21)
$comboBox2.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList #disable editing
$comboBox2.Name = "comboBox2"
$comboBox2.Items.AddRange(("None","Reviewer (can read items)","Author (can read and create items)","Editor (can read, create, and modify items)"))
$comboBox2.SelectedIndex = 0 #select based on index
$comboBox2.add_SelectedIndexChanged($handler_comboBox2_SelectedIndexChanged)
$groupBox1.Controls.Add($comboBox2)
#EndRegion Tasks Controls

#Region Inbox Controls
# Third Icon
$pictureBox3 = New-Object System.Windows.Forms.PictureBox
$pictureBox3.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\inbox.ico")
$pictureBox3.Location = New-Object System.Drawing.Size(10,111)
$pictureBox3.Size = New-Object System.Drawing.Size(24,24)
$pictureBox3.Name = "pictureBox3"
$pictureBox3.SizeMode = 2
$pictureBox3.TabStop = $False
$groupBox1.Controls.Add($pictureBox3)

# Third label
$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Size(50,120)
$label3.Size = New-Object System.Drawing.Size(50,18)
$label3.Name = "label3"
$label3.UseMnemonic = $true # Use & for ALT-shortcuts
$label3.Text = "&Inbox"
$groupBox1.Controls.Add($label3)

# Third combobox
$comboBox3 = New-Object System.Windows.Forms.ComboBox
$comboBox3.DataBindings.DefaultDataSourceUpdateMode = 0
$comboBox3.FormattingEnabled = $True
$comboBox3.Location = New-Object System.Drawing.Size(110,115)
$comboBox3.Size = New-Object System.Drawing.Size(245,21)
$comboBox3.Name = "comboBox3"
$comboBox3.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList #disable editing
$comboBox3.Items.AddRange(("None","Reviewer (can read items)","Author (can read and create items)","Editor (can read, create, and modify items)"))
$comboBox3.SelectedIndex = 0 #select based on index
$groupBox1.Controls.Add($comboBox3)
#EndRegion Inbox Controls

#Region Contacts Controls
$pictureBox4 = New-Object System.Windows.Forms.PictureBox
$pictureBox4.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\contacts2.ico")
$pictureBox4.Location = New-Object System.Drawing.Size(10,146)
$pictureBox4.Size = New-Object System.Drawing.Size(24,24)
$pictureBox4.Name = "pictureBox4"
$pictureBox4.SizeMode = 2
$pictureBox4.TabStop = $False
$groupBox1.Controls.Add($pictureBox4)

# Fourth label
$label4 = New-Object System.Windows.Forms.Label
$label4.Location = New-Object System.Drawing.Size(50,155)
$label4.Size = New-Object System.Drawing.Size(50,18)
$label4.Name = "label4"
$label4.UseMnemonic = $true # Use & for ALT-shortcuts
$label4.Text = "C&ontacts"
$groupBox1.Controls.Add($label4)

# Fourth combobox
$comboBox4 = New-Object System.Windows.Forms.ComboBox
$comboBox4.DataBindings.DefaultDataSourceUpdateMode = 0
$comboBox4.FormattingEnabled = $True
$comboBox4.Location = New-Object System.Drawing.Size(110,150)
$comboBox4.Size = New-Object System.Drawing.Size(245,21)
$comboBox4.Name = "comboBox4"
$comboBox4.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList #disable editing
$comboBox4.Items.AddRange(("None","Reviewer (can read items)","Author (can read and create items)","Editor (can read, create, and modify items)"))
$comboBox4.SelectedIndex = 0 #select based on index
$groupBox1.Controls.Add($comboBox4)
#EndRegion Contacts Controls

#Region Notes Controls
# Fifth Icon
$pictureBox5 = New-Object System.Windows.Forms.PictureBox
$pictureBox5.Image = [System.Drawing.Image]::FromFile("$PSScriptRoot\notes2.ico")
$pictureBox5.Location = New-Object System.Drawing.Size(10,181)
$pictureBox5.Size = New-Object System.Drawing.Size(24,24)
$pictureBox5.Name = "pictureBox5"
$pictureBox5.SizeMode = 2
$pictureBox5.TabStop = $False
$groupBox1.Controls.Add($pictureBox5)

# Fifth label
$label5 = New-Object System.Windows.Forms.Label
$label5.Location = New-Object System.Drawing.Size(50,190)
$label5.Size = New-Object System.Drawing.Size(50,18)
$label5.Name = "label5"
$label5.UseMnemonic = $true # Use & for ALT-shortcuts
$label5.Text = "&Notes"
$groupBox1.Controls.Add($label5)

# Fifth combobox
$comboBox5 = New-Object System.Windows.Forms.ComboBox
$comboBox5.DataBindings.DefaultDataSourceUpdateMode = 0
$comboBox5.FormattingEnabled = $True
$comboBox5.Location = New-Object System.Drawing.Size(110,185)
$comboBox5.Size = New-Object System.Drawing.Size(245,21)
$comboBox5.Name = "comboBox5"
$comboBox5.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList #disable editing
$comboBox5.Items.AddRange(("None","Reviewer (can read items)","Author (can read and create items)","Editor (can read, create, and modify items)"))
$comboBox5.SelectedIndex = 0 #select based on index
$groupBox1.Controls.Add($comboBox5)
#EndRegion Notes Controls

#Region Remaining Non-Groupbox Controls
#Last two checkboxes
$checkBox2 = New-Object System.Windows.Forms.CheckBox
$checkBox2.Location = New-Object System.Drawing.Size(6,228)
$checkBox2.Size = New-Object System.Drawing.Size(400,24)
$checkBox2.Name = "checkBox2"
$checkBox2.UseMnemonic = $true # Use & for ALT-shortcuts
$checkBox2.Text = "Automatically &send a message to delegate summarizing these permissions"
$checkBox2.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft #text position relative to the control
$checkBox2.UseVisualStyleBackColor = $True
$form1.Controls.Add($checkBox2)

$checkBox3 = New-Object System.Windows.Forms.CheckBox
$checkBox3.Location = New-Object System.Drawing.Size(6,248)
$checkBox3.Size = New-Object System.Drawing.Size(200,24)
$checkBox3.Name = "checkBox3"
$checkBox3.UseMnemonic = $true # Use & for ALT-shortcuts
$checkBox3.Text = "Delegate can see my &private items"
$checkBox3.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft #text position relative to the control
$checkBox3.UseVisualStyleBackColor = $True
$form1.Controls.Add($checkBox3)

$checkBox4 = New-Object System.Windows.Forms.CheckBox
$checkBox4.Location = New-Object System.Drawing.Size(206,248)
$checkBox4.Size = New-Object System.Drawing.Size(400,24)
$checkBox4.Name = "checkBox4"
$checkBox4.UseMnemonic = $true # Use & for ALT-shortcuts
$checkBox4.Text = "Delegate can &manage categories"
$checkBox4.TextAlign = [System.Drawing.ContentAlignment]::BottomLeft #text position relative to the control
$checkBox4.UseVisualStyleBackColor = $True
$form1.Controls.Add($checkBox4)
#Endregion Remaining Non-Groupbox Controls

#Region OK and Cancel buttons
# OK button
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(132,279)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form1.AcceptButton = $OKButton
$form1.Controls.Add($OKButton)

# Cancel button
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(216,279)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form1.CancelButton = $CancelButton
$form1.Controls.Add($CancelButton)
#Endregion OK and Cancel buttons

#region Event Handlers
# Load event handler
$form1.add_Load({
    if ($cPermissions.ContainsKey("Calendar")) {
        if ($cPermissions["Calendar"] -notin @("Author","Editor","Reviewer","None")) {
            $comboBox1.Items.Add("Custom")
            $comboBox1.SelectedItem = "Custom"
        }
        else { $comboBox1.SelectedIndex = Convert-ComboTextToInt($cPermissions["Calendar"]) }
    }
    if ($cPermissions.ContainsKey("Tasks")) {
            if ($cPermissions["Tasks"] -notin @("Author","Editor","Reviewer","None")) {
            $comboBox2.Items.Add("Custom")
            $comboBox2.SelectedItem = "Custom"
        }
        else { $comboBox2.SelectedIndex = Convert-ComboTextToInt($cPermissions["Tasks"]) }
    }
    if ($cPermissions.ContainsKey("Inbox")) {
        if ($cPermissions["Inbox"] -notin @("Author","Editor","Reviewer","None")) {
            $comboBox3.Items.Add("Custom")
            $comboBox3.SelectedItem = "Custom"
        }
        else { $comboBox3.SelectedIndex = Convert-ComboTextToInt($cPermissions["Inbox"]) }
    }
    if ($cPermissions.ContainsKey("Contacts")) {
        if ($cPermissions["Contacts"] -notin @("Author","Editor","Reviewer","None")) {
            $comboBox4.Items.Add("Custom")
            $comboBox4.SelectedItem = "Custom"
        }
        else { $comboBox4.SelectedIndex = Convert-ComboTextToInt($cPermissions["Contacts"]) }
    }
    if ($cPermissions.ContainsKey("Notes")) {
        if ($cPermissions["Notes"] -notin @("Author","Editor","Reviewer","None")) {
            $comboBox5.Items.Add("Custom")
            $comboBox5.SelectedItem = "Custom"
        }
        else { $comboBox5.SelectedIndex = Convert-ComboTextToInt($cPermissions["Notes"]) }
    }
    if ($comboBox1.SelectedItem -notmatch 'Editor') {
        $checkBox1.Enabled = $false; $checkBox1.Checked = $false
        $checkBox3.Enabled = $false; $checkBox3.Checked = $false
        $checkBox4.Enabled = $false; $checkBox4.Checked = $false
    }
    else {
        if (!$checkBox1.Checked) {
            $checkBox3.Enabled = $false; $checkBox3.Checked = $false
            $checkBox4.Enabled = $false; $checkBox4.Checked = $false
        }
    }
    if ($cPermissions.ContainsKey("CalendarExt")) {
        $checkbox1.checked = $cPermissions["CalendarExt"].Contains("Delegate")
        $checkbox3.checked = $cPermissions["CalendarExt"].Contains("CanViewPrivateItems")
        $checkbox4.checked = $cPermissions["CalendarExt"].Contains("CanManageCategories")

    }
    if ($WhatIfPreference) { $OKButton.Enabled = $false }
})

# KeyDown event handler
$form1.Add_KeyDown({
    if (($_.KeyCode -eq "Enter") -and ($form1.ActiveControl -eq $OKButton)) { Return-FormValues }
    if ($_.KeyCode -eq "Escape") { CancelForm }
    if ($_.KeyCode -eq "C" -and $_.modifiers -match "ALT") { $comboBox1.DroppedDown = $True; $form1.ActiveControl = $comboBox1 }
    if ($_.KeyCode -eq "T" -and $_.modifiers -match "ALT") { $comboBox2.DroppedDown = $True; $form1.ActiveControl = $comboBox2 }
    if ($_.KeyCode -eq "I" -and $_.modifiers -match "ALT") { $comboBox3.DroppedDown = $True; $form1.ActiveControl = $comboBox3 }
    if ($_.KeyCode -eq "O" -and $_.modifiers -match "ALT") { $comboBox4.DroppedDown = $True; $form1.ActiveControl = $comboBox4 }
    if ($_.KeyCode -eq "N" -and $_.modifiers -match "ALT") { $comboBox5.DroppedDown = $True; $form1.ActiveControl = $comboBox5 }
    if ($_.KeyCode -eq "D" -and $_.modifiers -match "ALT") {
        if ($checkBox1.Enabled) {
            if ($checkBox1.Checked) { $checkBox1.Checked = $false } else { $checkBox1.Checked = $true }
            $_.SuppressKeyPress = $true
        }
    }
    if ($_.KeyCode -eq "S" -and $_.modifiers -match "ALT") {
        if ($checkBox2.Enabled) {
            if ($checkBox2.Checked) { $checkBox2.Checked = $false } else { $checkBox2.Checked = $true }
            $_.SuppressKeyPress = $true
        }
    }
    if ($_.KeyCode -eq "P" -and $_.modifiers -match "ALT") {
        if ($checkBox3.Enabled) {
            if ($checkBox3.Checked) { $checkBox3.Checked = $false } else { $checkBox3.Checked = $true }
        }
        $_.SuppressKeyPress = $true
    }
    if ($_.KeyCode -eq "M" -and $_.modifiers -match "ALT") {
        if ($checkBox4.Enabled) {
            if ($checkBox4.Checked) { $checkBox4.Checked = $false } else { $checkBox4.Checked = $true }
        }
        $_.SuppressKeyPress = $true
    }
})

# ComboBox1 SelectedIndexChanged event handler
$comboBox1.add_SelectedIndexChanged({
    if ($comboBox1.SelectedIndex -ne '3') {
        $checkBox1.Enabled = $false; $checkBox1.Checked = $false
        #CanViewPrivateItems is only allowed when the Delegate flag is set as well
        $checkBox3.Enabled = $false; $checkBox3.Checked = $false
        #Similarly, CanManageCategories is only allowed when Delegate flag is set
        $checkBox4.Enabled = $false; $checkBox4.Checked = $false
    }
    else { $checkBox1.Enabled = $true }
})

#Checkbox1 state change handler
$checkBox1.add_CheckedChanged({
    if ($checkBox1.Checked -eq $false) {
        #CanViewPrivateItems is only allowed when the Delegate flag is set as well
        $checkBox3.Enabled = $false; $checkBox3.Checked = $false
        #Similarly, CanManageCategories is only allowed when Delegate flag is set
        $checkBox4.Enabled = $false; $checkBox4.Checked = $false
    }
    else {
        $checkBox3.Enabled = $true
        $checkBox4.Enabled = $true
    }
})

#Cancel button Click event handler
$CancelButton.add_Click({ CancelForm })

#OK button click event handler
$OKButton.Add_Click({ Return-FormValues })

#Handle closing the form via the X button
$form1.add_FormClosing({
    if ($form1.DialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        $script:FormCancelled = $true
    }
})
#endregion Event Handlers

#region Helper Functions
# Fetch the set of folders for the given mailbox, we cannot hardcode them due to localization
function Get-DefaultFolders {
    #Validation should be done prior to calling this function, only do mandatory flag here
    param(
        [Parameter(Mandatory=$true)][Alias("Manager")][string]$Owner
        )

        # Fetch default folders
        Write-Verbose "Fetching default folders for mailbox $Owner ..."
        try {
            if (!(Get-Command Get-MailboxFolderStatistics -ErrorAction SilentlyContinue)) { # Do we run in end-user scenario?
                Write-Verbose "Get-MailboxFolderStatistics cmdlet not available in the current session. Trying Get-MailboxFolder instead, should work for end-user scenarios."
                if (($Owner -ne $cUser) -and ($Owner -ne $cUser.Split("@")[0])) { # Not the current user
                    Throw "Get-MailboxFolder cannot be used to retrieve other users' folder information. Please run the script as an administrator or with sufficient privileges."
                }
                $defaultFolders = Get-MailboxFolder -Identity $Owner -Recurse -ResultSize Unlimited -ErrorAction Stop -Verbose:$false | Where-Object {$_.DefaultFolderType -in @("Calendar","Inbox","Contacts","Tasks","Notes")} | Select-Object @{n="Name";e={$_.DefaultFolderType}},Identity
            }
            else {
                # No server-side filter, no property selection :(
                # Get-ExOMailboxFolderStatistics is freaking slow when run against non-existent values, we get no benefit from using it
                $defaultFolders = Get-MailboxFolderStatistics -Identity $Owner -ErrorAction Stop -Verbose:$false -WarningAction SilentlyContinue | Where-Object {$_.FolderType -in @("Calendar","Inbox","Contacts","Tasks","Notes")} | Select-Object @{n="Name";e={$_.FolderType}},@{n="Identity";e={($_.Identity.ToString().replace('\',':\'))}}
            }}
        catch {
            if ($_.Exception.Message.ToString() -match "Get-MailboxFolder cannot be used") {
                Write-Warning $_.Exception.Message.ToString()
                Throw $_.Exception.Message.ToString().Split(".")[1]
            }
            if ($_.Exception.Message.ToString() -match "Couldn't find|doesn't exist") {
                Throw "Mailbox $Owner not found, this should not happen..."
            }
            Write-Warning "Failed to retrieve default folders for $Owner. Using hardcoded names instead, beware of localization issues."
            # Return hardcoded names as fallback
            $defaultFolders =  @{
                Calendar = "$($owner):\Calendar"
                Inbox    = "$($owner):\Inbox"
                Contacts = "$($owner):\Contacts"
                Tasks    = "$($owner):\Tasks"
                Notes    = "$($owner):\Notes"
            }
            return $defaultFolders
        }

    $defaultFoldersHash = @{}
    $defaultFolders | Foreach { $defaultFoldersHash[$_.Name] = $_.Identity }
    return $defaultFoldersHash
}

# Fetch the standard set of folder permissions
function Get-StandardFolderPermissions {
    #Validation should be done prior to calling this function, only do mandatory flag here
    param(
        [Parameter(Mandatory=$true)][Alias("Manager")][string]$Owner,
		[Parameter(Mandatory=$true)][string]$Delegate
    )

    # Get the set of folders
    $script:cFolders = Get-DefaultFolders -Owner $Owner
    $folderPermissions = @{}

    # Fetch permissions for each folder
    foreach ($folder in $cFolders.GetEnumerator()) {
        Write-Verbose "Fetching permissions for folder $($folder.Value) ..."
        try {
            #No point using Get-ExOMailboxFolderPermission, it's terribly slow when erroring out
            $varPermissions = Get-MailboxFolderPermission -Identity $folder.Value.ToString() -User $Delegate -ErrorAction Stop -Verbose:$false
            if ($folder.Key -eq "Calendar") {
                if ($varPermissions.SharingPermissionFlags) {
                    $folderPermissions["CalendarExt"] = @($varPermissions.SharingPermissionFlags -split ",").Trim() ?? @()
                }
                else { $folderPermissions["CalendarExt"] = @() } # Covered below, remove?
            }
            $folderPermissions[$($folder.Name)] = ($varPermissions.AccessRights)[0] ?? "None"
        }
        catch {
            if ($_.CategoryInfo.Reason -eq "ManagementObjectNotFoundException") { throw "Folder ""$($folder.Value)"" not found, this should not happen..." }
            # Gotta love the lack of attention to detail in error handling :(
            if (($_.CategoryInfo.Reason -eq "UserNotFoundInPermissionEntryException") -or ($_.CategoryInfo.Reason -eq "Exception" -and $_.Exception.Data['RemoteException'].TypeName.Split(".")[-1] -eq "UserNotFoundInPermissionEntryException")) {
                $folderPermissions[$($folder.Name)] = "None"
                continue
            }
            else { $_ | fl * -Force; continue }
        }
    }

    if (!$folderPermissions.ContainsKey("CalendarExt")) {
        $folderPermissions["CalendarExt"] = @()
    }
    $folderPermissions["Owner"] = $Owner
    $folderPermissions["Delegate"] = $Delegate
    return $folderPermissions
}

# Get the Form values and close the form
function Return-FormValues {
    $script:cPermissionsNew = @{}
    $cPermissionsNew["Calendar"] = $combobox1.SelectedItem.Split(" ")[0]
    $cPermissionsNew["Tasks"] = $combobox2.SelectedItem.Split(" ")[0]
    $cPermissionsNew["Inbox"] = $combobox3.SelectedItem.Split(" ")[0]
    $cPermissionsNew["Contacts"] = $combobox4.SelectedItem.Split(" ")[0]
    $cPermissionsNew["Notes"] = $combobox5.SelectedItem.Split(" ")[0]
    $cPermissionsNew["CalendarExt"] = @()
    #SharingPermissionFlags parameter only applies to calendar folders and can only be used when the AccessRights parameter value is Editor.
    if ($checkbox1.Checked) { $cPermissionsNew["CalendarExt"] += "Delegate" }
    #CanViewPrivateItems is only allowed when the Delegate flag is set as well
    if ($checkbox3.Checked -and ($cPermissionsNew["CalendarExt"].Contains("Delegate"))) { $cPermissionsNew["CalendarExt"] += "CanViewPrivateItems" }
    #Similarly, CanManageCategories is only allowed when Delegate flag is set
    if ($checkbox4.Checked -and ($cPermissionsNew["CalendarExt"].Contains("Delegate"))) { $cPermissionsNew["CalendarExt"] += "CanManageCategories" }
    #SendNotificationToUser parameter only applies to calendar folders and can only be used with the following AccessRights parameter values: AvailabilityOnly, LimitedDetails, Reviewer, Editor
    if ($checkBox2.Checked) { $cPermissionsNew["SendSummary"] = $true }
    else { $cPermissionsNew["SendSummary"] = $false }

    $form1.Close() | Out-Null
}

# Close the form and mark it as cancelled
function CancelForm {
    $form1.Close() | Out-Null
    $script:FormCancelled = $true
}

# Process any permission changes as needed
function Apply-PermissionChanges {
    # Validaitons should be done prior to calling this function

    # Check for changes in SharingPermissionFlags as they need special handling
    if (Compare-Object -ReferenceObject $cPermissions["CalendarExt"] -DifferenceObject $cPermissionsNew["CalendarExt"]) {
        Write-Verbose "Detected change in permission for SharingPermissionFlags, to be addressed as part of Calendar folder processing..."
        $script:cFlags = $true
        #Empty array is not accepted, use "None" instead
        if ($cPermissionsNew["CalendarExt"].Count -eq 0) { $script:cFlagsValue = "None" } else { $script:cFlagsValue = $cPermissionsNew["CalendarExt"] }
    }
    else { $script:cFlags = $false }

    #Process changes
    foreach ($perm in $cPermissionsNew.GetEnumerator()) {
        # Skip SendNotificationToUser, we don't have an initial value for it, but we respect the new value below
        # Also skip CalendarExt here, we handle SharingPermissionFlags separately
        if (($perm.Key -eq "SendSummary") -or ($perm.Key -eq "CalendarExt")) { continue }

        # Skip custom permissions
        if ($perm.Value -eq "Custom") {
            Write-Verbose "Custom permissions are not currently supported, skipping folder $($cFolders[$perm.Key])..."
            continue
        }

        # Check if there is a change
        if ($cPermissions.ContainsKey($perm.Key)) {
            # Old and new permissions match, but we still need to account for SharingPermissionFlags changes
            if ($cPermissions[$perm.Key] -eq $perm.Value) {
                # Only act on Calendar folder, with Editor access right and modified SharingPermissionFlags
                if (($perm.Key -eq "Calendar") -and ($perm.Value -eq "Editor") -and ($cFlags)) {
                    Write-Verbose "Updating sharing permission flags on folder $($cFolders[$perm.Key])..."
                    try {
                        Set-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -SendNotificationToUser:$($cPermissionsNew["SendSummary"]) -SharingPermissionFlags $cFlagsValue -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                    }
                    catch {
                        Write-Verbose "Failed to update sharing permission flags on folder $($cFolders[$perm.Key]): $($_.Exception.Message)"
                        $_ | fl * -Force
                    }
                }
                else { Write-Verbose "No changes detected for folder $($cFolders[$perm.Key]), skipping..." }
                continue
            }
            # Change necessary
            else {
                # Add
                if (($null -eq $cPermissions[$perm.Key]) -or ($cPermissions[$perm.Key] -eq "None")) {
                    Write-Verbose "Adding permission on folder $($cFolders[$perm.Key])..."
                    try {
                        # Calendar needs special handling
                        if ($perm.Key -eq "Calendar") {
                            # Changes to the two SharingPermissionFlags should be processed for Calendar folder with Editor access right and modified SharingPermissionFlags
                            if ($perm.Value -eq "Editor") {
                                Add-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -SendNotificationToUser:$($cPermissionsNew["SendSummary"]) -SharingPermissionFlags $cFlagsValue -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                            }
                            # Else just process SendNotificationToUser. We don't support LimitedDetails and AvailabilityOnly currently, but include them regardless
                            elseif ($perm.Value -in @("AvailabilityOnly","LimitedDetails","Reviewer","Author")) {
                                Add-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -SendNotificationToUser:$($cPermissionsNew["SendSummary"]) -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                            }
                            # Else just add the access right
                            else { Add-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference }
                        }
                        # For non-calendar folders
                        else {
                            Add-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                        }
                    }
                    catch {
                        Write-Verbose "Failed to add permission on folder $($cFolders[$perm.Key]): $($_.Exception.Message)"
                        $_ | fl * -Force
                    }
                    continue
                }

                # Remove
                if ($perm.Value -eq "None") {
                    Write-Verbose "Removing permission on folder $($cFolders[$perm.Key])..."
                    try {
                        # Process SendNotificationToUser only for Calendar folder
                        if ($perm.Key -eq "Calendar") {
                            Remove-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -SendNotificationToUser:$($cPermissionsNew["SendSummary"]) -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                        }
                        else {
                            Remove-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                        }
                    }
                    catch {
                        Write-Verbose "Failed to remove permission on folder $($cFolders[$perm.Key]): $($_.Exception.Message)"
                        $_ | fl * -Force
                    }
                    continue
                }

                # Set
                Write-Verbose "Updating permission on folder $($cFolders[$perm.Key]) to $($perm.Value)..."
                try {
                    # Calendar needs special handling
                    if ($perm.Key -eq "Calendar") {
                        # Changes to the two SharingPermissionFlags should be processed for Calendar folder with Editor access right
                        if ($perm.Value -eq "Editor") {
                            Set-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -SendNotificationToUser:$($cPermissionsNew["SendSummary"]) -SharingPermissionFlags $cFlagsValue -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                        }
                        # Process remaining access rights that support SendNotificationToUser. We don't support LimitedDetails and AvailabilityOnly currently, but include them regardless
                        elseif ($perm.Value -in @("AvailabilityOnly","LimitedDetails","Reviewer","Author")) {
                            Set-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -SendNotificationToUser:$($cPermissionsNew["SendSummary"]) -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                        }
                        # Else just set the access right
                        else { Set-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference }
                    }
                    # For non
                    else {
                        Set-MailboxFolderPermission -Identity $cFolders[$perm.Key] -User $Delegate -AccessRights $perm.Value -Confirm:$false -ErrorAction Stop -Verbose:$false -WhatIf:$WhatIfPreference
                    }
                }
                catch {
                    Write-Verbose "Failed to update permission on folder $($cFolders[$perm.Key]): $($_.Exception.Message)"
                    $_ | fl * -Force
                }
                continue
            }
        }
    }
}

# Convert combobox text to integer index
function Convert-ComboTextToInt($role) {
    switch -wildcard ($role) {
		'Reviewer*' {"1"}
		'Author*' {"2"}
		'Editor*' {"3"}
	    'None' {"0"}
    }
}

function GenerateForm {
    # Validate all the values
	[CmdletBinding()]
	param(
        [Parameter(Mandatory)][Alias("Manager")][string]$Owner,
        [Parameter(Mandatory)][string]$Delegate,
        [switch]$ReadOnly=$false
	)

    # Verify that both Owner and Delegate mailboxes exist
    try {
        Get-ExORecipient -Identity $Owner -RecipientType UserMailbox -ErrorAction Stop -Verbose:$false | Out-Null
        Get-ExORecipient -Identity $Delegate -RecipientTypeDetails UserMailbox -ErrorAction Stop -Verbose:$false | Out-Null
    }
    catch { throw $_ }

    # Fetch current permissions
    $script:cPermissions = Get-StandardFolderPermissions -Owner $Owner -Delegate $Delegate

    # Set whatIfPreference based on ReadOnly switch
    $script:WhatIfPreference = $ReadOnly.IsPresent

    # Show the form
    $form1.Text = "Delegate Permissions for $($cPermissions["Delegate"]) on $($cPermissions["Owner"])"
    $form1.ShowDialog() | Out-Null
}
#endregion Helper Functions

#Call the form
GenerateForm @PSBoundParameters
if ($FormCancelled) {
    Write-Verbose "Form cancelled by user, exiting..."
    return
}
if (!($cPermissionsNew) -or !($cPermissions)) {
    Throw "Cannot apply permission changes, missing current or new permissions data..."
}
else { Apply-PermissionChanges }

Write-Verbose "Finished"