# 0.8.0
# Will Crabtree
# 

Import-Module ActiveDirectory -WarningAction SilentlyContinue

#region Likely values to be updated

# Used for Campus UI dropdown and useful Active Directory OU to computer name conversion
$campusOUConversionList = @(
    ('WC', 'West'), 
    ('29', '29th St.'), 
    ('ER', 'El Rio'), 
    ('EP', 'El Pueblo'), 
    ('DV', 'Desert Vista'), 
    ('DO', 'District'), 
    ('DC', 'Downtown'), 
    ('DM', 'DM'), 
    ('DP', 'Douglas Prison'), 
    ('EC', 'East'), 
    ('MS', 'Maintenance and Security'), 
    ('NW', 'Northwest'), 
    ('WP', 'Wilmot Prison'), 
    ('AT', 'Downtown'),
    ('PCC', 'West')
)

# Used for Suffix UI dropdowns and useful Active Directory OU to computer name conversion
$userSuffixConversionList = @(
    ('A', 'Administrator'), 
    ('B', 'Borrowed'), 
    ('S', 'Staff'), 
    ('F', 'Faculty'), 
    ('I', 'Instructor'), 
    ('C', 'Class'), 
    ('L', 'Lab'), 
    ('P', 'Public'), 
    ('M', 'Meeting/conference'), 
    ('D', 'DPS')
)

$deviceTypeSuffixConversionList = @(
    ('C', 'Computer'), 
    ('N', 'Notebook'),
    ('K', 'Kiosk'),
    ('T', 'Tablet'),
    ('V', 'Virtual Machine')
)
#endregion

#region GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

#region Computer Info Window
$CompInfoWindow = New-Object System.Windows.Forms.Form    
$CompInfoWindow.Text = 'Computer Information'
$CompInfoWindow.AutoScaleDimensions = '7,15'
$CompInfoWindow.StartPosition = 'CenterScreen'
$CompInfoWindow.ControlBox = $false
$CompInfoWindow.TopMost = $true
$CompInfoWindow.MinimumSize = '400,300'
$CompInfoWindow.Padding = New-Object System.Windows.Forms.Padding(10)

$CompInfoLabel = New-Object system.Windows.Forms.Label
$CompInfoLabel.Text = 'Computer Name'
$CompInfoLabel.Font = 'Segoe UI, 10pt,style=bold'
$CompInfoLabel.Dock = 'Fill'
$CompInfoLabel.Anchor = 'Bottom'
$CompInfoLabel.AutoSize = $true

$CampusLabel = New-Object system.Windows.Forms.Label
$CampusLabel.Text = 'Campus'
$CampusLabel.Font = 'Segoe UI, 8pt,style=bold'
$CampusLabel.AutoSize = $true
$CampusLabel.Anchor = 'Right'
$CampusDropdown = New-Object System.Windows.Forms.ComboBox
$CampusDropdown.Dock = 'Top'
$CampusDropdown.Anchor = 'Left'
$CampusDropdown.AutoCompleteMode = 'SuggestAppend'
$CampusDropdown.AutoCompleteSource = 'ListItems'
$CampusDropdown.DropDownStyle = 'DropDown'
$CampusDropdown.MinimumSize = '50,50'
$CampusDropdown.TabIndex = 1
$CampusDropdown.Items.AddRange(($campusOUConversionList | ForEach-Object { $_[0] }))

$RoomNumberLabel = New-Object system.Windows.Forms.Label
$RoomNumberLabel.Text = 'Bldg/Room'
$RoomNumberLabel.Font = 'Segoe UI, 8pt,style=bold'
$RoomNumberLabel.Dock = 'Top'
$RoomNumberLabel.Anchor = 'Right'
$RoomNumberLabel.AutoSize = $true
$RoomNumberTextbox = New-Object System.Windows.Forms.TextBox
$RoomNumberTextbox.Dock = 'Top'
$RoomNumberTextbox.Anchor = 'Left'
$RoomNumberTextbox.MinimumSize = '50,20'
$RoomNumberTextbox.TabIndex = 2

$PccNumberLabel = New-Object system.Windows.Forms.Label
$PccNumberLabel.Text = 'PCC#'
$PccNumberLabel.Font = 'Segoe UI, 8pt,style=bold'
$PccNumberLabel.Dock = 'Top'
$PccNumberLabel.Anchor = 'Right'
$PccNumberLabel.AutoSize = $true
$PccNumberTextBox = New-Object System.Windows.Forms.TextBox
$PccNumberTextBox.Dock = 'Top'
$PccNumberTextBox.Anchor = 'Left'
$PccNumberTextBox.MinimumSize = '50,20'
$PccNumberTextBox.TabIndex = 3

$UserSuffixLabel = New-Object system.Windows.Forms.Label
$UserSuffixLabel.Text = 'User Suffix'
$UserSuffixLabel.Font = 'Segoe UI, 8pt,style=bold'
$UserSuffixLabel.Dock = 'Top'
$UserSuffixLabel.Anchor = 'Right'
$UserSuffixLabel.AutoSize = $true
$UserSuffixDropdown = New-Object System.Windows.Forms.ComboBox
$UserSuffixDropdown.Dock = 'Top'
$UserSuffixDropdown.Anchor = 'Left'
$UserSuffixDropdown.AutoCompleteMode = 'SuggestAppend'
$UserSuffixDropdown.AutoCompleteSource = 'ListItems'
$UserSuffixDropdown.DropDownStyle = 'DropDown'
$UserSuffixDropdown.MinimumSize = '50,50'
$UserSuffixDropdown.TabIndex = 4
$UserSuffixDropdown.Items.AddRange(($userSuffixConversionList | ForEach-Object { $_[1] }))

$HardwareSuffixLabel = New-Object system.Windows.Forms.Label
$HardwareSuffixLabel.Text = 'Hardware Suffix'
$HardwareSuffixLabel.Font = 'Segoe UI, 8pt,style=bold'
$HardwareSuffixLabel.Dock = 'Top'
$HardwareSuffixLabel.Anchor = 'Right'
$HardwareSuffixLabel.AutoSize = $true
$HardwareSuffixDropdown = New-Object System.Windows.Forms.ComboBox
$HardwareSuffixDropdown.Dock = 'Top'
$HardwareSuffixDropdown.Anchor = 'Left'
$HardwareSuffixDropdown.AutoCompleteMode = 'SuggestAppend'
$HardwareSuffixDropdown.AutoCompleteSource = 'ListItems'
$HardwareSuffixDropdown.DropDownStyle = 'DropDown'
$HardwareSuffixDropdown.MinimumSize = '50,50'
$HardwareSuffixDropdown.TabIndex = 5
$HardwareSuffixDropdown.Items.AddRange(($deviceTypeSuffixConversionList | ForEach-Object { $_[1] }))

$ActiveDirTreeViewLabel = New-Object system.Windows.Forms.Label
$ActiveDirTreeViewLabel.Text = 'Select an OU'
$ActiveDirTreeViewLabel.Font = 'Segoe UI, 10pt,style=bold'
$ActiveDirTreeViewLabel.Dock = 'Fill'
$ActiveDirTreeViewLabel.Anchor = 'Bottom'
$ActiveDirTreeViewLabel.AutoSize = $true
$ActiveDirTreeView = New-Object System.Windows.Forms.TreeView
$ActiveDirTreeView.Dock = 'Fill'
$ActiveDirTreeView.TabIndex = 6

$SubmitInfoButton = New-Object System.Windows.Forms.Button
$SubmitInfoButton.Text = '&Submit'
$SubmitInfoButton.Dock = 'Bottom'
$SubmitInfoButton.AutoSize = $true
$SubmitInfoButton.BackColor = 'LightGray'
$SubmitInfoButton.TabIndex = 7
$CompInfoWindow.AcceptButton = $SubmitInfoButton
#endregion

#region Login Window
$LoginWindow = New-Object System.Windows.Forms.Form
$LoginWindow.Text = 'Active Directory Sign In'
$LoginWindow.ControlBox = $false
$LoginWindow.TopMost = $true
$LoginWindow.Size = New-Object System.Drawing.Size(300, 200)
$LoginWindow.StartPosition = 'CenterScreen'
$LoginWindow.AutoSizeMode = 'GrowAndShrink'
$LoginWindow.MinimumSize = New-Object System.Drawing.Size(200, 150)  # Minimum form size
$LoginWindow.Padding = New-Object System.Windows.Forms.Padding(10)

$UsernameLabel = New-Object System.Windows.Forms.Label
$UsernameLabel.Text = 'Username:'
$UsernameLabel.Anchor = 'None'
$UsernameTextBox = New-Object System.Windows.Forms.TextBox
$UsernameTextBox.Dock = 'Fill'
$UsernameTextBox.MinimumSize = New-Object System.Drawing.Size(100, 0)
$UsernameTextBox.TabIndex = 1
$UsernameTextBox.Anchor = 'Left, Right'

$PasswordLabel = New-Object System.Windows.Forms.Label
$PasswordLabel.Text = 'Password:'
$PasswordLabel.Anchor = 'None'
$PasswordTextBox = New-Object System.Windows.Forms.TextBox
$PasswordTextBox.Dock = 'Fill'
$PasswordTextBox.UseSystemPasswordChar = $true
$PasswordTextBox.MinimumSize = New-Object System.Drawing.Size(100, 0)
$PasswordTextBox.TabIndex = 2
$PasswordTextBox.Anchor = 'Left, Right'

$EduDomainOption = New-Object System.Windows.Forms.RadioButton
$EduDomainOption.Text = 'EDU'
$EduDomainOption.Anchor = 'None'
$EduDomainOption.TabStop = $true
$EduDomainOption.Checked = $true

$PccDomainOption = New-Object System.Windows.Forms.RadioButton
$PccDomainOption.Text = 'PCC'
$PccDomainOption.Anchor = 'None'
$PccDomainOption.TabStop = $true

$LoginButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$LoginButtonPanel.Dock = 'Fill'
$LoginButtonPanel.AutoSize = $true
$LoginButtonPanel.AutoSizeMode = 'GrowAndShrink'
$LoginButtonPanel.Anchor = 'None'
$LoginButtonPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight

$SignInButton = New-Object System.Windows.Forms.Button
$SignInButton.Text = '&Sign In'
$SignInButton.Anchor = 'None'
$SignInButton.AutoSize = $true
$SignInButton.MinimumSize = New-Object System.Drawing.Size(100, 0)
$SignInButton.TabIndex = 3
$LoginWindow.AcceptButton = $SignInButton
$LoginWindow.AcceptButton.DialogResult = 'OK'

$CancelLoginButton = New-Object System.Windows.Forms.Button
$CancelLoginButton.Text = '&Cancel'
$CancelLoginButton.Anchor = 'None'
$CancelLoginButton.AutoSize = $true
$CancelLoginButton.MinimumSize = New-Object System.Drawing.Size(100, 0)
$CancelLoginButton.TabIndex = 4
$LoginWindow.CancelButton = $CancelLoginButton
$LoginWindow.CancelButton.DialogResult = 'Cancel'
#endregion

#region Duplicate Computer Delete
$DuplicateSystemWindow = New-Object System.Windows.Forms.Form 
$DuplicateSystemWindow.Text = "Duplicate PCC Number"
$DuplicateSystemWindow.ControlBox = $false
$DuplicateSystemWindow.AutoSize = $true
$DuplicateSystemWindow.Topmost = $true
$DuplicateSystemWindow.AutoSizeMode = 'GrowAndShrink'
$DuplicateSystemWindow.StartPosition = 'CenterScreen'
$DuplicateSystemWindow.MinimumSize = New-Object System.Drawing.Size(300, 200)
$DuplicateSystemWindow.Padding = New-Object System.Windows.Forms.Padding(10)

$DuplicateSystemLabel = New-Object System.Windows.Forms.Label
$DuplicateSystemLabel.Text = "Select computers to delete:"
$DuplicateSystemLabel.Dock = 'Fill'
$DuplicateSystemLabel.AutoSize = $true
$DuplicateSystemLabel.Anchor = 'Left, Right'

$DuplicateSystemList = New-Object System.Windows.Forms.CheckedListBox 
$DuplicateSystemList.Dock = 'Fill'
$DuplicateSystemList.AutoSize = $true
$DuplicateSystemList.MinimumSize = New-Object System.Drawing.Size(100, 0)
$DuplicateSystemList.Anchor = 'Left, Right, Top, Bottom'

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Text = 'OK'
$OKButton.Anchor = 'None'
$OKButton.AutoSize = $true
$OKButton.MinimumSize = New-Object System.Drawing.Size(50, 0)
$OKButton.DialogResult = 'OK'
$DuplicateSystemWindow.AcceptButton = $OKButton
#endregion
#endregion


#region UI Layouts
#region Login Layout
$Login_LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$Login_LayoutPanel.RowCount = 4
$Login_LayoutPanel.ColumnCount = 2
$Login_LayoutPanel.Dock = 'Fill'
$Login_LayoutPanel.AutoSize = $true
$Login_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
$Login_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null

$LoginWindow.Controls.Add($Login_LayoutPanel)
$Login_LayoutPanel.Controls.Add($UsernameLabel, 0, 0)
$Login_LayoutPanel.Controls.Add($UsernameTextBox, 1, 0)
$Login_LayoutPanel.Controls.Add($PasswordLabel, 0, 1)
$Login_LayoutPanel.Controls.Add($PasswordTextBox, 1, 1)
$Login_LayoutPanel.Controls.Add($EduDomainOption, 0, 2)
$Login_LayoutPanel.Controls.Add($PccDomainOption, 1, 2)
$Login_LayoutPanel.Controls.Add($LoginButtonPanel, 0, 3)
$Login_LayoutPanel.SetColumnSpan($LoginButtonPanel, 2)

$LoginButtonPanel.Controls.Add($SignInButton)
$LoginButtonPanel.Controls.Add($CancelLoginButton)
#endregion

#region ComputerInfo UI Layout
$ComputerInfo_LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$ComputerInfo_LayoutPanel.Dock = "Fill"
$ComputerInfo_LayoutPanel.ColumnCount = 4
$ComputerInfo_LayoutPanel.RowCount = 7
#$ComputerInfo_LayoutPanel.CellBorderStyle = 3
[void]$ComputerInfo_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 20)))
[void]$ComputerInfo_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 30)))
[void]$ComputerInfo_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 5)))
[void]$ComputerInfo_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))
[void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))

$ComputerInfo_LayoutPanel.Controls.Add($CompInfoLabel, 0, 0)
$ComputerInfo_LayoutPanel.SetColumnSpan($CompInfoLabel, 2)

$ComputerInfo_LayoutPanel.Controls.Add($CampusLabel, 0, 1)
$ComputerInfo_LayoutPanel.Controls.Add($CampusDropdown, 1, 1)

$ComputerInfo_LayoutPanel.Controls.Add($RoomNumberLabel, 0, 2)
$ComputerInfo_LayoutPanel.Controls.Add($RoomNumberTextbox, 1, 2)

$ComputerInfo_LayoutPanel.Controls.Add($PccNumberLabel, 0, 3)
$ComputerInfo_LayoutPanel.Controls.Add($PccNumberTextBox, 1, 3)

$ComputerInfo_LayoutPanel.Controls.Add($UserSuffixLabel, 0, 4)
$ComputerInfo_LayoutPanel.Controls.Add($UserSuffixDropdown, 1, 4)
$ComputerInfo_LayoutPanel.Controls.Add($HardwareSuffixLabel, 0, 5)
$ComputerInfo_LayoutPanel.Controls.Add($HardwareSuffixDropdown, 1, 5)

$ComputerInfo_LayoutPanel.Controls.Add($ActiveDirTreeViewLabel, 3, 0)
$ComputerInfo_LayoutPanel.Controls.Add($ActiveDirTreeView, 3, 1)
$ComputerInfo_LayoutPanel.SetrowSpan($ActiveDirTreeView, 5)

$ComputerInfo_LayoutPanel.Controls.Add($SubmitInfoButton, 0, 6)
$ComputerInfo_LayoutPanel.SetcolumnSpan($SubmitInfoButton, 4)

$CompInfoWindow.controls.Add($ComputerInfo_LayoutPanel)
#endregion
#region Duplicate Computer UI Layout
$DuplicateComp_LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$DuplicateComp_LayoutPanel.RowCount = 3
$DuplicateComp_LayoutPanel.ColumnCount = 1
$DuplicateComp_LayoutPanel.Dock = 'Fill'
$DuplicateComp_LayoutPanel.AutoSize = $true
[void]$DuplicateComp_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$DuplicateComp_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))

$DuplicateSystemWindow.Controls.Add($DuplicateComp_LayoutPanel)
$DuplicateComp_LayoutPanel.Controls.Add($DuplicateSystemLabel, 0, 0)
$DuplicateComp_LayoutPanel.Controls.Add($DuplicateSystemList, 0, 1)
$DuplicateComp_LayoutPanel.Controls.Add($OKButton, 0, 2)
#endregion
#endregion
#endregion
#region Functions

function Show-ADLoginWindow {
    # Show the login window and log the domain for later
    [void]$LoginWindow.ShowDialog()

    # Check if user clicks 'OK'
    if ($LoginWindow.DialogResult -eq 'OK') {
        # Convert password to secure string for security reasons
        $Password = ConvertTo-SecureString $PasswordTextBox.Text -AsPlainText -Force
        $Credentials = New-Object System.Management.Automation.PSCredential ($UsernameTextBox.text, $Password)
    
        # Determine which radio button is checked to determine which domain to query
        try {
            if ($PccDomainOption.Checked) {
                $ADDomain = Get-ADDomain -Credential $Credentials -Server $($PccDomainOption.text + '-domain.pima.edu')
            }
            else {
                $ADDomain = Get-ADDomain -Credential $Credentials -Server $($EduDomainOption.text + '-domain.pima.edu')
            }
        }
        # If an authentication error occurs, recursively call the function again
        catch [System.Security.Authentication.AuthenticationException] {
            Show-ADLoginWindow
            return
        }
        
        # Check if the domain name matches with either PCC or EDU to determine if login is successful
        if (($ADDomain.Name -match $PccDomainOption.text) -or ($ADDomain.Name -match $EduDomainOption.text)) {
            [void]$CompInfoWindow.ShowDialog()
            break
        }
        # If the login fails, prompt the user with a choice to retry or cancel
        else {
            $RelogChoice = [System.Windows.Forms.MessageBox]::Show("Login Failed, please relaunch.", 'Warning', 'RetryCancel', 'Warning')
            switch ($RelogChoice) {
                'Retry' { Show-ADLoginWindow }
                'Cancel' {
                    [System.Windows.Forms.MessageBox]::Show("Login was cancelled, rebooting the computer.", 'Warning', 'Ok', 'Warning')
                    Start-Sleep -Seconds 5
                    Restart-Computer -Force -WhatIf 
                }
            }        
        }
    }
    # If the user cancels the operation, show a warning message and reboot the computer after 5 seconds
    elseif ($LoginWindow.DialogResult -eq 'Cancel') {
        [System.Windows.Forms.MessageBox]::Show("Login was cancelled, rebooting the computer.", 'Warning', 'Ok', 'Warning')
        Start-Sleep -Seconds 5
        Restart-Computer -Force -WhatIf
    }  
}

Function Confirm-ComputerName {
    # Clearing any previous errors
    $ErrorProvider.SetError($CompInfoLabel, '')

    # Setting the button color to LightGray
    $PccCheckButton.BackColor = 'LightGray'

    try {
        # Checking if a campus from the approved list is selected
        if (!($CampusDropdown.Items -contains $CampusDropdown.Text)) {
            throw 'Select a proper campus'
        }

        # Checking the entered building/room text against a regular expression
        if (!($RoomNumberTextbox.Text -match '^[a-z]{1}\d{3}$|^[a-z]{2}\d{2}$|^[a-z]{2}\d{3}$|^[a-z]{3}$')) {
            throw 'Enter a proper building/room'
        }

        # Checking the entered PCC number is a 6 digit number
        if (!($PccNumberTextBox.Text -match '^\d{6}$')) {
            throw 'Enter a proper PCC Number'
        }
        elseif ($PccNumberTextBox.Text -match '^\d{6}$') {

            # Getting the Active Directory computer matching the PCC number
            $PCCSearch = Get-ADComputer -Filter ('Name -like "*' + $PccNumberTextBox.Text + '*"') -Server $ADDomain.Forest

            # Defining the regular expression pattern
            $regexPattern = "$($PccNumberTextBox.Text)..$"
        
            # Filtering the matching computers
            $matchingComputers = $PCCSearch | Where-Object { $_.Name -match $regexPattern }

            if ($matchingComputers) {

                # Adding each matching computer to the list box
                foreach ($computer in $matchingComputers.Name) {
                    [void]$DuplicateSystemList.Items.Add($computer, $false)
                }
            
                # Removing selected computers
                if ($DuplicateSystemWindow.ShowDialog() -eq 'OK') {
                    foreach ($selectedComputer in $listBox.CheckedItems) {
                        try {
                            Remove-ADComputer -Identity $selectedComputer -Confirm:$false -WhatIf
                        }
                        catch {
                            <#Do this if a terminating exception happens#>
                        }
                    }
                }
            }
        }

        # Checking if a valid suffix is selected
        if (!($UserSuffixDropdown.Items -contains $UserSuffixDropdown.Text) -or !($HardwareSuffixDropdown.Items -contains $HardwareSuffixDropdown.Text)) {
            throw 'Enter a proper suffix'
        }

        # Building the computer name
        switch ($CampusDropdown.Text) {
            'DC' { 
                $Global:ComputerName = $CampusDropdown.Text + $RoomNumberTextbox.Text + $PccNumberTextBox.Text + $userSuffixConversionList[$UserSuffixDropdown.SelectedIndex][0] + $deviceTypeSuffixConversionList[$HardwareSuffixDropdown.SelectedIndex][0]
            }
            Default {
                $Global:ComputerName = $CampusDropdown.Text + '-' + $RoomNumberTextbox.Text + $PccNumberTextBox.Text + $userSuffixConversionList[$UserSuffixDropdown.SelectedIndex][0] + $deviceTypeSuffixConversionList[$HardwareSuffixDropdown.SelectedIndex][0]
            }
        }

        # Checking the length of the computer name
        if ($ComputerName.Length -gt 15) {
            throw 'Name too long'
        }

        # Setting the button color to Green if all inputs are valid
        $PccCheckButton.BackColor = 'Green'
    }
    catch {
        # Catching the error and setting it in the form
        $ErrorProvider.SetError($CompInfoLabel, $_)

        # Stopping further execution
        return
    }
}

Function Get-ADTreeNode ($ParentNode, $CurrentOrganizationalUnit) {
    # Add a new child node to the parent node in the AD tree. 
    # The 'DistinguishedName' property of the CurrentOU object is used as the key, and 'Name' property as the value.
    $ChildNode = $ParentNode.Nodes.Add($CurrentOrganizationalUnit.DistinguishedName.toString(), $CurrentOrganizationalUnit.Name)
    
    # Add a placeholder node that will be replaced when the user expands the node.
    $ChildNode.Nodes.Add("Loading...")
}

#endregion
#region GUI Actions
$ActiveDirTreeView.Add_BeforeExpand({
        param($sender, $e)
    
        # Assign the node being expanded to the variable $node.
        $node = $e.Node

        # Check if the node contains a placeholder (an item with the text "Loading...").
        # If it does, it means the real child nodes have not been loaded yet.
        if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Text -eq "Loading...") {
            # Remove the placeholder.
            $node.Nodes.Clear()

            # Query Active Directory to get the organizational units (OUs) one level below the current node,
            # and add them as child nodes of the current node.
            Get-ADOrganizationalUnit -Filter * -SearchScope OneLevel -SearchBase $node.Name -Server $ADDomain.Forest | ForEach-Object { Get-ADTreeNode $node $_ }
        }
    })

# Define actions to be taken when clicking on the Username and Password textboxes - Clears the content.
$UsernameTextBox.Add_Click( { $UsernameTextBox.Clear() })
$PasswordTextBox.Add_Click( { $PasswordTextBox.Clear() })

# Define actions to be taken when changing the selection in the Campus dropdown.
$CampusDropdown.Add_SelectedIndexChanged({
        # Determine the search base for Get-ADOrganizationalUnit based on the selected radio button
        $searchBase = if ($EduDomainOption.Checked) {
            "OU=EDU_Computers,DC=edu-domain,DC=pima,DC=edu"
        }
        elseif ($PccDomainOption.Checked) {
            if ($CampusDropdown.Items -contains $CampusDropdown.Text) {
                "OU=$($campusOUConversionList[$CampusDropdown.SelectedIndex][1]),OU=PCC,DC=PCC-Domain,DC=pima,DC=edu"
            }
            else {
                "OU=PCC,DC=PCC-Domain,DC=pima,DC=edu"
            }
        }

        # Set the label text and clear the treeview for population
        $ActiveDirTreeViewLabel.Text = "Loading OU's..."
        $ActiveDirTreeView.Nodes.Clear()

        # Populate the treeview with the OUs found
        Get-ADOrganizationalUnit -Filter * -SearchScope OneLevel -SearchBase $searchBase -Server $ADDomain.Forest | ForEach-Object { Get-ADTreeNode $ActiveDirTreeView $_ }
        $ActiveDirTreeViewLabel.Text = 'Select an OU'
    })

# Retrieve the PCC number set in the BIOS, if available, and enter it into the PCC number field
$PCCNumber = (Get-CimInstance -Query "Select * from Win32_SystemEnclosure").SMBiosAssetTag
if ($PCCNumber -match '^\d{6}$') {
    $PccNumberTextBox.Text = $PCCNumber
    $PccNumberTextBox.ReadOnly = $true
    $PccNumberLabel.Text = 'PCC# : Loaded from BIOS'
}

$SubmitInfoButton.Add_Click( { 
        Confirm-ComputerName

        # Verify a target OU is selected
        if ($null -eq $ActiveDirTreeView.SelectedNode) {
            $ErrorProvider.SetError($ActiveDirTreeViewLabel, 'Select an OU')
        }
        else {
            $ErrorProvider.SetError($ActiveDirTreeViewLabel, '')
        }          
        
        # Submit data to Task Sequence if there are no errors on UI elements
        if (-not($ErrorProvider.GetError($CompInfoLabel) -or $ErrorProvider.GetError($ActiveDirTreeViewLabel))) {
            [System.Windows.Forms.MessageBox]::Show("Submitted Data:`n`nComputer Name: $($ComputerName.ToUpper())`n`nOU: $("LDAP://$($ActiveDirTreeView.SelectedNode.Name)")`n`nDomain: $($ADDomain.Forest)", 'Warning', 'Ok', 'Warning')

            # Set the task sequence environment variables and close the form
            <#
            $TSEnvironment = New-Object -COMObject Microsoft.SMS.TSEnvironment 
            $TSEnvironment.Value("OSDComputerName") = "$($ComputerName.ToUpper())"
            $TSEnvironment.Value("OSDDomainOUName") = "$("LDAP://$($ActiveDirTreeView.SelectedNode.Name)")"
            $TSEnvironment.Value("OSDDomainName") = "$($ADDomain.Forest)"     
            #>     
            # Close the form
            [void]$CompInfoWindow.Close()
        }
    })
#endregion

# Launches main login window function which the gets AD creds needed for the rest of the script
Show-ADLoginWindow

# Enable to view forms for testing
#[void]$CompInfoWindow.ShowDialog()
#[void]$DuplicateSystemList.Items.Add("test", $false)
#[void]$DuplicateSystemWindow.ShowDialog()