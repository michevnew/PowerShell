#Requires -Modules @{ ModuleName="Microsoft.Graph.Groups"; ModuleVersion="2.25.0" }
#For details on what the script does and how to run it, check: https://michev.info/blog/post/7084/free-tool-to-manage-directory-settings-in-entra-id

# Clean up before startup
Remove-Variable settings,templates,groupId,settingsG,err,newList -ErrorAction SilentlyContinue #$err should be local to the event handler, but just in case

# Connect to Microsoft Graph with required scopes
Write-Verbose "Connecting to Microsoft Graph..."
#Policy.ReadWrite.Authorization needed for consent settings
Connect-MgGraph -Scopes GroupSettings.ReadWrite.All,Directory.Read.All,Directory.ReadWrite.All,Policy.ReadWrite.Authorization -NoWelcome -Verbose:$false -ErrorAction Stop

# Retrieve templates and settings
Write-Verbose "Retrieving directory setting templates..."
$templates = Get-MgGroupSettingTemplateGroupSettingTemplate -Verbose:$false -ErrorAction Stop | sort DisplayName
if (!$templates) {
    Write-Error "No directory setting templates found, check your permissions. The script will now exit."
    return
}

Write-Verbose "Retrieving directory settings..."
$settings = Get-MgGroupSetting -Verbose:$false -ErrorAction Stop | sort DisplayName
if (!$settings) {
    Write-Warning "No directory settings found."
}

# Load Windows Forms assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Main dialog properties
$form = New-Object System.Windows.Forms.Form
$form.Text = "Directory Settings"
$form.Size = New-Object System.Drawing.Size(600, 381)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = 'FixedSingle'

# Dropdown text and control
$labelT = New-Object System.Windows.Forms.Label
$labelT.Location = New-Object System.Drawing.Point(10, 20)
$labelT.Size = New-Object System.Drawing.Size(190, 20)
$labelT.Text = "Select Directory Setting Template:"
$form.Controls.Add($labelT)

$dropdownT = New-Object System.Windows.Forms.ComboBox
$dropdownT.Location = New-Object System.Drawing.Point(210, 20)
$dropdownT.Size = New-Object System.Drawing.Size(250, 20)
$dropdownT.Items.AddRange($templates.DisplayName)
#$dropdownT.SelectedIndex = 0
$dropdownT.DropDownStyle = 'DropDownList'
$form.Controls.Add($dropdownT)

# Data grid control for working with the setting values
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 65)
$dataGridView.Size = New-Object System.Drawing.Size(450, 255)
$datagridView.AutoSizeColumnsMode = 'Fill'
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.ReadOnly = $false
$dataGridView.SelectionMode = 'CellSelect'
$form.Controls.Add($dataGridView)

# Status bar (maybe force an old style one instead of this strip crap?)
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.SizingGrip = $false
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusLabel.BorderStyle = "Adjust"
$statusLabel.Spring = $true
$statusLabel.TextAlign = "MiddleLeft"
$statusBar.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusBar)

# Select grpoup button (might replace with an actual button as this one only pops up on hover)
$statusButton = New-Object System.Windows.Forms.ToolStripButton
$statusButton.Text = "Select group"
$statusButton.Visible = $false
$statusButton.Alignment = [System.Windows.Forms.ToolStripItemAlignment]::Right
$statusBar.Items.Add($statusButton) | Out-Null

#region Buttons
$AddButton = New-Object System.Windows.Forms.Button
$AddButton.Location = New-Object System.Drawing.Point(480, 70)
$AddButton.Size = New-Object System.Drawing.Size(80, 23)
$AddButton.Text = "Add"
$form.Controls.Add($AddButton)

$RemButton = New-Object System.Windows.Forms.Button
$RemButton.Location = New-Object System.Drawing.Point(480, 100)
$RemButton.Size = New-Object System.Drawing.Size(80, 23)
$RemButton.Text = "Remove"
$form.Controls.Add($RemButton)

$UpdButton = New-Object System.Windows.Forms.Button
$UpdButton.Location = New-Object System.Drawing.Point(480, 130)
$UpdButton.Size = New-Object System.Drawing.Size(80, 23)
$UpdButton.Text = "Update"
$form.Controls.Add($UpdButton)

$RefButton = New-Object System.Windows.Forms.Button
$RefButton.Location = New-Object System.Drawing.Point(480, 160)
$RefButton.Size = New-Object System.Drawing.Size(80, 23)
$RefButton.Text = "Refresh"
$form.Controls.Add($RefButton)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(480, 260)
$okButton.Size = New-Object System.Drawing.Size(80, 23)
$okButton.Text = "OK"
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(480, 290)
$cancelButton.Size = New-Object System.Drawing.Size(80, 23)
$cancelButton.Text = "Cancel"
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)
#endregion

#region Events
# Dropdown selection changed event handler
$dropdownT.Add_SelectedIndexChanged({
    $temp = $templates | ? {$_.DisplayName -eq $dropdownT.SelectedItem} | select -ExpandProperty Id
    $objSettings = $settings | ? {$_.TemplateId -eq $temp}
    if ($objSettings) {
        $labelSettings = "Directory Setting found."
    }
    else {
        if ($settingsG -and ($settingsG.TemplateId -eq $temp)) {
            $labelSettings = "Directory Setting found for group $groupId."
        }
        else {
            $labelSettings = "No existing setting for template ""$($dropdownT.SelectedItem)""."
        }
    }

    # Clear existing data
    $dataGridView.DataSource = $null
    $dataGridView.Rows.Clear()
    $dataGridView.Columns.Clear()
    $statusButton.Visible = $false

    # Add columns
    $dataGridView.Columns.Add("Name", "Name")
    $dataGridView.Columns.Add("Value", "Value")
    $dataGridView.Columns[0].ReadOnly = $true

    if ($temp -eq "08d542b9-071f-4e16-94b0-74abb372e3d9" -or $temp -eq "7e0abea2-5c20-405f-9658-bfc9a523fd49") {
        # If we have a group ID, try loading the settings
        if ($groupId) {
            # Settings object exists and was successfully retrieved
            if ($settingsG.TemplateId -eq $temp) { #Stupid, but saves us from doing the same thing twice
                $AddButton.Enabled = $false
                $RemButton.Enabled = $true
                $UpdButton.Enabled = $true

                # Add rows for each setting value
                foreach ($value in $settingsG.Values) {
                    $dataGridView.Rows.Add($value.Name, $value.Value)
                }

                $StatusLabel.Text = "Settings for group $groupId loaded."
            }
            else { # No settings object for the selected group
                $AddButton.Enabled = $true
                $RemButton.Enabled = $false
                $UpdButton.Enabled = $false

                $statusLabel.Text = "No existing settings for group $groupId."
            }
        }
        # No group ID selected yet
        else {
            # Clear existing data
            $dataGridView.DataSource = $null
            $dataGridView.Rows.Clear()
            $dataGridView.Columns.Clear()

            $AddButton.Enabled = $false
            $RemButton.Enabled = $false
            $UpdButton.Enabled = $false

            $statusLabel.Text = "No group selected. Press the button on the right to select a group."
        }
        $statusButton.Visible = $true
    }
    else {
        $StatusLabel.Text = $labelSettings

        if ($objSettings) {
            $AddButton.Enabled = $false
            $RemButton.Enabled = $true
            $UpdButton.Enabled = $true

            # Add rows for each setting value
            foreach ($value in $objSettings.Values) {
                $dataGridView.Rows.Add($value.Name, $value.Value)
            }
        }
        else {
            $AddButton.Enabled = $true
            $RemButton.Enabled = $false
            $UpdButton.Enabled = $false
        }
    }
    # Special handling for the Password rule settings template
    if ($temp -eq "5cf42378-d67d-4f36-ba46-e8b86229381d") {
        # Find the BannedPasswordList row and add a button cell to the last column
        foreach ($row in $dataGridView.Rows) {
            if ($row.Cells[0].Value -eq "BannedPasswordList") {
                $buttonCell = New-Object System.Windows.Forms.DataGridViewButtonCell
                $buttonCell.FlatStyle = 'Flat'

                if ($row.Cells[1].Value -and $row.Cells[1].Value -ne "") {
                    $script:newlist = $row.Cells[1].Value -split "`t" # This overwrites the previous value on template selection change, should we avoid that?
                }
                else { $script:newlist = @() }

                $buttonCell.Value = "$($newList.Count) Banned Passwords"
                $row.Cells[1] = $buttonCell
    }}}
})

# BannedPasswordList button event handler
$dataGridView.Add_CellContentClick({
    param($sender, $e)

    # Only handle clicks on the BannedPasswordList button cell
    if ($e.ColumnIndex -ne 1) { return }
    if ($sender.Rows[$e.RowIndex].Cells[0].Value -ne "BannedPasswordList") { return }

    # Create a multiline input dialog for editing the banned password list
    $inputForm = New-Object System.Windows.Forms.Form
    $inputForm.Text = "Edit Banned Password List (one per line)"
    $inputForm.Size = New-Object System.Drawing.Size(300, 240)
    $inputForm.StartPosition = "CenterScreen"
    $inputForm.FormBorderStyle = 'FixedDialog'
    $inputForm.MaximizeBox = $false
    $inputForm.MinimizeBox = $false

    $inputTextBox = New-Object System.Windows.Forms.TextBox
    $inputTextBox.Location = New-Object System.Drawing.Point(10, 10)
    $inputTextBox.Size = New-Object System.Drawing.Size(260, 150)
    $inputTextBox.Multiline = $true
    $inputTextBox.ScrollBars = 'Vertical'
    $inputTextBox.Text = ($newList -join "`r`n")
    $inputForm.Controls.Add($inputTextBox)

    $inputOkButton = New-Object System.Windows.Forms.Button
    $inputOkButton.Location = New-Object System.Drawing.Point(110, 170)
    $inputOkButton.Size = New-Object System.Drawing.Size(75, 23)
    $inputOkButton.Text = "OK"
    $inputOkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $inputForm.AcceptButton = $null # Set the AcceptButton to None to prevent Enter fromg closing the dialog
    # Uncomment the line below to enable Enter key submission
    #$inputForm.AcceptButton = $inputOkButton
    $inputForm.Controls.Add($inputOkButton)
    $inputOkButton.Select() # Prevent auto-selecting of all the text in the textbox

    $inputCancelButton = New-Object System.Windows.Forms.Button
    $inputCancelButton.Location = New-Object System.Drawing.Point(195, 170)
    $inputCancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $inputCancelButton.Text = "Cancel"
    $inputCancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $inputForm.CancelButton = $inputCancelButton
    $inputForm.Controls.Add($inputCancelButton)

    $inputResult = $inputForm.ShowDialog()

    if ($inputResult -eq [System.Windows.Forms.DialogResult]::OK) {
        # Trim and split the input text into an array, removing empty lines and duplicates. Remove any entry that does not meet the length requirements
        $script:newList = $inputTextBox.Text -split "`r`n" | Where-Object { $_ -ne "" } | Where-Object { ($_.Length -ge 4) -and ($_.Length -le 16) } | ForEach-Object { $_.Trim() } | Sort-Object -Unique
        $sender.Rows[$e.RowIndex].Cells[1].Value = "$($newList.Count) Banned Passwords"
    }
    $inputForm.Dispose()
})

# Basic validation for the setting object values based on the template object constraints.
$dataGridView.Add_CellValidating({
    param($sender, $e)

    $newValue = $e.FormattedValue

    # Only validate the Value column
    if ($e.ColumnIndex -eq "1") {
        $settingName = $sender.Rows[$e.RowIndex].Cells[0].Value

        # Get the current template to check value constraints
        $template = $templates | ? {$_.DisplayName -eq $dropdownT.SelectedItem}

        if ($template) {
            $templateValue = $template.Values | ? {$_.Name -eq $settingName}

            if ($templateValue) {
                # Check if it's a boolean setting
                if ($templateValue.Type -eq "System.Boolean") {
                    if ($newValue -notmatch "^(true|false)$") {
                        $e.Cancel = $true
                        [System.Windows.Forms.MessageBox]::Show("Value must be 'true' or 'false' for boolean settings.", "Invalid Value", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                        return
                    }
                }

                # Validate integer values
                if ($templateValue.Type -eq "System.Int32") {
                    if (-not [int]::TryParse($newValue, [ref]$null)) {
                        $e.Cancel = $true
                        [System.Windows.Forms.MessageBox]::Show("Value must be a valid integer.", "Invalid Value", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                        return
                    }

                    # Set the allowed range for Lockout Duration and Threshold
                    if ($settingName -eq "LockoutDurationInSeconds") {
                            if ([int]$newValue -lt 5 -or [int]$newValue -gt 18000) {
                            $e.Cancel = $true
                            [System.Windows.Forms.MessageBox]::Show("LockoutDurationInSeconds must be between 5 and 18000.", "Invalid Value", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                            return
                        }
                    }
                    if ($settingName -eq "LockoutThreshold") {
                        if ([int]$newValue -lt 1 -or [int]$newValue -gt 50) {
                            $e.Cancel = $true
                            [System.Windows.Forms.MessageBox]::Show("LockoutThreshold must be between 1 and 50.", "Invalid Value", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                            return
                        }
                    }
                }

                # Valudate GUID values
                if ($templateValue.Type -eq "System.Guid") {
                    if ($newValue -eq "") { return } # Allow empty string
                    try {
                        $Result = [System.Guid]::Parse($newValue)
                    }
                    catch {
                        $e.Cancel = $true
                        [System.Windows.Forms.MessageBox]::Show("Value must be a valid GUID.", "Invalid Value", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                        return
                    }
                    Remove-Variable Result -ErrorAction SilentlyContinue
                }

                # What do we want to validate for strings? We should allow empty values, unless default one is defined (covered below)
                if ($templateValue.Type -eq "System.String") {
                    if ($newValue -eq "") { return } # Allow empty string

                    # Valudate URLs
                    if ($settingName -in @("UsageGuidelinesUrl","GuestUsageGuidelinesUrl","CustomConditionalAccessPolicyUrl")) {
                        if (-not [System.Uri]::IsWellFormedUriString($newValue, [System.UriKind]::Absolute)) {
                            $e.Cancel = $true
                            [System.Windows.Forms.MessageBox]::Show("Value must be a valid URL.", "Invalid Value", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                            return
                        }
                    }

                    # Validate BannedPasswordCheckOnPremisesMode
                    if ($settingName -eq "BannedPasswordCheckOnPremisesMode") {
                        if ($newValue -notin @("Audit","Enforce")) {
                            $e.Cancel = $true
                            [System.Windows.Forms.MessageBox]::Show("BannedPasswordCheckOnPremisesMode must be one of the following values: Audit, Enforce.", "Invalid Value", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                            return
                        }
                    }
                }

                # Check for empty values if required
                if ([string]::IsNullOrWhiteSpace($newValue) -and $templateValue.DefaultValue) {
                    $e.Cancel = $true
                    [System.Windows.Forms.MessageBox]::Show("Value cannot be empty for this setting.", "Invalid Value", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    return
                }
            }
        }
        else {
            $e.Cancel = $true
            [System.Windows.Forms.MessageBox]::Show("Invalid template... this should not happen!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }
})

# Group selection button event handler/custom input dialog
$statusButton.Add_Click({
    # Create a simple input dialog for Group ID
    $inputForm = New-Object System.Windows.Forms.Form
    $inputForm.Text = "Enter Group ID"
    $inputForm.Size = New-Object System.Drawing.Size(400, 110)
    $inputForm.StartPosition = "CenterScreen"
    $inputForm.FormBorderStyle = 'FixedDialog'
    $inputForm.MaximizeBox = $false
    $inputForm.MinimizeBox = $false

    $inputTextBox = New-Object System.Windows.Forms.TextBox
    $inputTextBox.Location = New-Object System.Drawing.Point(10, 10)
    $inputTextBox.Size = New-Object System.Drawing.Size(360, 20)
    $inputForm.Controls.Add($inputTextBox)

    $inputOkButton = New-Object System.Windows.Forms.Button
    $inputOkButton.Location = New-Object System.Drawing.Point(210, 40)
    $inputOkButton.Size = New-Object System.Drawing.Size(75, 23)
    $inputOkButton.Text = "OK"
    $inputOkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $inputForm.AcceptButton = $inputOkButton
    $inputForm.Controls.Add($inputOkButton)

    $inputCancelButton = New-Object System.Windows.Forms.Button
    $inputCancelButton.Location = New-Object System.Drawing.Point(295, 40)
    $inputCancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $inputCancelButton.Text = "Cancel"
    $inputCancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $inputForm.CancelButton = $inputCancelButton
    $inputForm.Controls.Add($inputCancelButton)

    $inputResult = $inputForm.ShowDialog()

    if ($inputResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:groupId = $inputTextBox.Text

        try {
            $guid = [System.Guid]::Parse($groupId)
            $statusLabel.Text = "Selected group: $($guid.ToString())"
            Write-Verbose "Group ID selected: $($guid.ToString())"

            # If valid GUID, get the group-specific settings for the selected template
            $script:settingsG = Get-MgGroupSetting -GroupId $groupId -Verbose:$false -ErrorAction Stop
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Invalid GUID format. Please enter a valid GUID.", "Invalid Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $script:groupId = $null
        }
        Remove-Variable guid -ErrorAction SilentlyContinue
    }
    $inputForm.Dispose()
    # Stipud, but the only thing that seems to does the trick...
    $oldIndex = $dropdownT.SelectedIndex
    $dropdownT.SelectedIndex = -1
    $dropdownT.SelectedIndex = $oldIndex
})

# Add button event handler
$AddButton.Add_Click({
    $temp = $templates | ? {$_.DisplayName -eq $dropdownT.SelectedItem} | select -ExpandProperty Id
    $objSettings = $settings | ? {$_.TemplateId -eq $temp}
    $template = $templates | ? {$_.Id -eq $temp}

    # Disable the Add button to prevent multiple clicks, enable the Update button to facilitate saving the new setting object
    $AddButton.Enabled = $false
    $UpdButton.Enabled = $true

    # Just in case, check if a setting object already exists for the selected template and abort is needed
    if ($objSettings -or ($settingsG -and ($settingsG.TemplateId -eq $temp))) {
        Write-Warning "A setting already exists for the selected template. Only one setting object is allowed per template."
        [System.Windows.Forms.MessageBox]::Show("A setting already exists for the selected template. Only one setting object is allowed per template.", "Setting object already exists!", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    if ($template) {
        # Clear existing data from the DataGridView
        $dataGridView.DataSource = $null
        $dataGridView.Rows.Clear()
        $dataGridView.Columns.Clear()
        $statusButton.Visible = $false

        # Fill in the DataGridView with default values from the template
        $dataGridView.Columns.Add("Name", "Name")
        $dataGridView.Columns.Add("Value", "Value")
        $dataGridView.Columns[0].ReadOnly = $true
        foreach ($value in $template.Values) {
            $dataGridView.Rows.Add($value.Name, $value.DefaultValue)
        }

        $statusLabel.Text = "Ready to create new setting. Press the ""Update"" button to save."
    }
    else {
        Write-Warning "Invalid template selected... this should not happen!"
        $objSettings | Out-Default
        [System.Windows.Forms.MessageBox]::Show("Invalid template... this should not happen!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
})

# Update button event handler
$UpdButton.Add_Click({
    $temp = $templates | ? {$_.DisplayName -eq $dropdownT.SelectedItem} | select -ExpandProperty Id
    $objSettings = $settings | ? {$_.TemplateId -eq $temp}
    #$template = $templates | ? {$_.Id -eq $temp}

    # Collect values from DataGridView
    $values = @()
    foreach ($row in $dataGridView.Rows) {
        if ($row.IsNewRow) { continue }

        $settingName = $row.Cells[0].Value
        $settingValue = $row.Cells[1].Value

        # Validate BannedPasswordList value from the button cell
        if ($settingName -eq "BannedPasswordList") {
            $settingValue = $newList -join "`t"
        }

        # Must be a hashtable with Name and Value keys
        $values += @{
            "Name"  = $settingName
            "Value" = $settingValue
        }
    }

    # Group-specific template check
    if ($temp -eq "08d542b9-071f-4e16-94b0-74abb372e3d9" -or $temp -eq "7e0abea2-5c20-405f-9658-bfc9a523fd49") {
        # Just in case, check if a group ID is selected
        if (!$groupId) {
            $errorMsg = "No group selected for group-specific setting template. Please select a group first."
            Write-Warning $errorMsg
            [System.Windows.Forms.MessageBox]::Show("No group selected for group-specific setting template. Please select a group first.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = $errorMsg
            return
        }

        # Check if a setting object exists for the selected template and act accordingly
        if ($settingsG -and ($settingsG.TemplateId -eq $temp)) {
            # Update existing setting object
            Update-MgGroupSetting -GroupSettingId $settingsG.Id -Values $values -GroupId $groupId -ErrorVariable err -ErrorAction SilentlyContinue -Verbose:$false

            if ($err -or $err.count) {
                $errorMsg = "Setting object creation failed! Check your connection and permissions and restart the script."
                Write-Error $err.ErrorDetails.Message -ErrorAction Continue
                [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $statusLabel.Text = $errorMsg
                return
            }

            $statusLabel.Text = "Setting object updated successfully group $groupId"
            Write-Verbose "Setting object updated successfully for group $groupId and template: $($dropdownT.SelectedItem)"

            # Refresh the settings cache
            $RefButton.PerformClick()
        }
        else {
            # Create new setting object
            New-MgGroupSetting -TemplateId $temp -Values $values -GroupId $groupId -ErrorVariable err -ErrorAction SilentlyContinue -Verbose:$false

            if ($err -or $err.count) {
                $errorMsg = "Setting object creation failed! Check your connection and permissions and restart the script."
                Write-Error $err.ErrorDetails.Message -ErrorAction Continue
                [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $statusLabel.Text = $errorMsg
                return
            }

            $statusLabel.Text = "Setting object created successfully for group $groupId"
            Write-Verbose "Setting object created successfully for group $groupId and template: $($dropdownT.SelectedItem)"

            # Refresh the settings cache
            $RefButton.PerformClick()
        }
    }
    # Non-group-specific template
    else {
        # Check if a setting object exists for the selected template and act accordingly
        if ($objSettings) {
            # Update existing setting
            Update-MgGroupSetting -GroupSettingId $objSettings.Id -Values $values -ErrorVariable err -ErrorAction SilentlyContinue -Verbose:$false

            if ($err -or $err.count) {
                $errorMsg = "Setting update failed! Check your connection and permissions and restart the script."
                Write-Error $err.ErrorDetails.Message -ErrorAction Continue
                [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $statusLabel.Text = $errorMsg
                return
            }

            $statusLabel.Text = "Setting updated successfully"
            Write-Verbose "Setting updated successfully for template: $($dropdownT.SelectedItem)"

            # Refresh the settings cache
            $RefButton.PerformClick()
        }
        else {
            # Create new setting
            New-MgGroupSetting -TemplateId $temp -Values $values -ErrorVariable err -ErrorAction SilentlyContinue -Verbose:$false

            if ($err -or $err.count) {
                $errorMsg = "Setting creation failed! Check your connection and permissions and restart the script."
                Write-Error $err.ErrorDetails.Message -ErrorAction Continue
                [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $statusLabel.Text = $errorMsg
                return
            }

            $statusLabel.Text = "Setting created successfully"
            Write-Verbose "Setting created successfully for template: $($dropdownT.SelectedItem)"

            # Refresh the settings cache
            $RefButton.PerformClick()
        }
    }
})

# Remove button event handles
$RemButton.Add_Click({
    $temp = $templates | ? {$_.DisplayName -eq $dropdownT.SelectedItem} | select -ExpandProperty Id
    $objSettings = $settings | ? {$_.TemplateId -eq $temp}

    # Just in case, check if a setting object exists for the selected template and abort if not
    if (!$objSettings -and !($settingsG -and ($settingsG.TemplateId -eq $temp))) {
        Write-Warning "No setting exists for this template."
        [System.Windows.Forms.MessageBox]::Show("No setting exists for this template.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    # Ask for confirmation before removal
    $confirmResult = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove this directory setting?", "Confirm Removal", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

    if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        Write-Verbose "Removing directory setting with ID: $($objSettings.Id)"

        # Check whether it's a group-specific setting and act accordingly
        if ($settingsG -and ($settingsG.TemplateId -eq $temp) -and $groupId) {
            Remove-MgGroupSetting -GroupSettingId $settingsG.Id -GroupId $groupId -ErrorAction SilentlyContinue -ErrorVariable err -Verbose:$false
        }
        else {
            Remove-MgGroupSetting -GroupSettingId $objSettings.Id -ErrorAction SilentlyContinue -ErrorVariable err -Verbose:$false
        }

        # Check for errors
        if ($err -or $err.count) {
            $errorMsg = "Removal failed! Check your connection and permissions and restart the script."
            Write-Error $err.ErrorDetails.Message -ErrorAction Continue
            [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = $errorMsg
            return
        }

        $statusLabel.Text = "Setting removed successfully."
        Write-Verbose "Directory setting object for template ""$($dropdownT.SelectedItem)"" removed successfully"

        # Refresh the settings cache
        $RefButton.PerformClick()

        # Clear the DataGridView
        $dataGridView.DataSource = $null
        $dataGridView.Rows.Clear()
        $dataGridView.Columns.Clear()

        $AddButton.Enabled = $true
        $RemButton.Enabled = $false
        $UpdButton.Enabled = $false
    }
})

# Refresh button event handler
$RefButton.Add_Click({

    $statusLabel.Text = "Refreshing..."
    $oldIndex = $dropdownT.SelectedIndex

    # Refresh templates and settings
    Write-Verbose "Retrieving directory setting templates..."
    $script:templates = Get-MgGroupSettingTemplateGroupSettingTemplate -Verbose:$false -ErrorAction SilentlyContinue -ErrorVariable err | sort DisplayName
    if ($err -or $err.count -or !$templates) {
        $errorMsg = "Refresh failed! Check your connection and permissions and restart the script."
        Write-Error $err.ErrorDetails.Message -ErrorAction Continue
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = $errorMsg
        return
    }

    Write-Verbose "Retrieving directory settings..."
    $script:settings = Get-MgGroupSetting -Verbose:$false -ErrorAction SilentlyContinue -ErrorVariable err | sort DisplayName
    if ($err -or $err.count) {
        $errorMsg = "Refresh failed! Check your connection and permissions and restart the script."
        Write-Error $err.ErrorDetails.Message -ErrorAction Continue
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = $errorMsg
        return
    }

    # Also refresh group-specific settings if a group is selected
    if ($groupId) {
        Write-Verbose "Retrieving group-specific directory settings for group ID: $groupId"
        $script:settingsG = Get-MgGroupSetting -GroupId $groupId -Verbose:$false -ErrorAction SilentlyContinue -ErrorVariable err
        if ($err -or $err.count) {
            $errorMsg = "Refresh failed! Check your connection and permissions and restart the script."
            Write-Error $err.ErrorDetails.Message -ErrorAction Continue
            [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = $errorMsg
            return
        }
    }

    # Update dropdown items
    $dropdownT.Items.Clear()
    $dropdownT.Items.AddRange($templates.DisplayName)

    if ($dropdownT.Items.Count -gt 0) {
        $dropdownT.SelectedIndex = $oldIndex
    }

    $statusLabel.Text = "Refresh completed"
    Write-Verbose "Templates and settings refreshed successfully"
})
#endregion

$dropdownT.SelectedIndex = 0

$form.ShowDialog() | Out-Null

$form.Dispose()