# 0.7.0
# Will Crabtree
# 

import-module activedirectory

#region GUI

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

$screen = [System.Windows.Forms.Screen]::AllScreens
$font = New-Object System.Drawing.Font("Segoe UI", 8)
#region Computer Info Window
$ComputerInfo_Form = New-Object System.Windows.Forms.Form    
$ComputerInfo_Form.AutoScaleDimensions = '7,15'
$ComputerInfo_Form.AutoScaleMode = 'Font'
$ComputerInfo_Form.StartPosition = 'CenterScreen'
$ComputerInfo_Form.Width = $($screen[0].bounds.Width / 3)
$ComputerInfo_Form.Height = $($screen[0].bounds.Height / 3)
#$ComputerInfo_Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSHome + '\powershell.exe')
$ComputerInfo_Form.Text = 'Active Directory Information'
$ComputerInfo_Form.ControlBox = $false
$ComputerInfo_Form.TopMost = $true
$ComputerInfo_Form.MinimumSize = '400,300'

$ComputerForm_Label = New-Object system.Windows.Forms.Label
$ComputerForm_Label.Text = 'Computer Name'
$ComputerForm_Label.Font = 'Segoe UI, 10pt,style=bold'
$ComputerForm_Label.Dock = 'Fill'
$ComputerForm_Label.Anchor = 'Bottom'
$ComputerForm_Label.AutoSize = $true

$Campus_Label = New-Object system.Windows.Forms.Label
$Campus_Label.Text = 'Campus'
$Campus_Label.AutoSize = $true
$Campus_Dropdown = New-Object System.Windows.Forms.ComboBox
$Campus_Dropdown.DropDownStyle = 'DropDown'
$Campus_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Campus_Dropdown.AutoCompleteSource = 'ListItems'
$Campus_Dropdown.TabIndex = 1
$Campus_Dropdown.Dock = 'Top'
$Campus_Dropdown.MinimumSize = '50,50'

$BuildingRoom_Label = New-Object system.Windows.Forms.Label
$BuildingRoom_Label.Text = 'Bldg/Room'
$BuildingRoom_Label.AutoSize = $true
$BuildingRoom_Textbox = New-Object System.Windows.Forms.TextBox
$BuildingRoom_Textbox.Dock = 'Top'
$BuildingRoom_Textbox.TabIndex = 2
$BuildingRoom_Textbox.MinimumSize = '50,20'

$pccNumber_Label = New-Object system.Windows.Forms.Label
$pccNumber_Label.Text = 'PCC#'
$pccNumber_Label.AutoSize = $true
$pccNumber_Textbox = New-Object System.Windows.Forms.TextBox
$pccNumber_Textbox.TabIndex = 3
$pccNumber_Textbox.Dock = 'Top'
$pccNumber_Textbox.MinimumSize = '50,20'

$userSuffix_Label = New-Object system.Windows.Forms.Label
$userSuffix_Label.Text = 'User Type Suffix'
$userSuffix_Label.AutoSize = $true
$userSuffix_Dropdown = New-Object System.Windows.Forms.ComboBox
$userSuffix_Dropdown.DropDownStyle = 'DropDown'
$userSuffix_Dropdown.AutoCompleteMode = 'SuggestAppend'
$userSuffix_Dropdown.AutoCompleteSource = 'ListItems'
$userSuffix_Dropdown.TabIndex = 4
$userSuffix_Dropdown.Dock = 'Top'
$userSuffix_Dropdown.MinimumSize = '50,50'

$hardwareSuffix_Label = New-Object system.Windows.Forms.Label
$hardwareSuffix_Label.Text = 'Hardware Type Suffix'
$hardwareSuffix_Label.AutoSize = $true
$hardwareSuffix_Dropdown = New-Object System.Windows.Forms.ComboBox
$hardwareSuffix_Dropdown.DropDownStyle = 'DropDown'
$hardwareSuffix_Dropdown.AutoCompleteMode = 'SuggestAppend'
$hardwareSuffix_Dropdown.AutoCompleteSource = 'ListItems'
$hardwareSuffix_Dropdown.TabIndex = 5
$hardwareSuffix_Dropdown.Dock = 'Top'
$hardwareSuffix_Dropdown.MinimumSize = '50,50'

$CheckPCC_Button = New-Object System.Windows.Forms.Button
$CheckPCC_Button.Text = 'Check Name'
$CheckPCC_Button.TabIndex = 6
$CheckPCC_Button.Dock = 'Bottom'
$CheckPCC_Button.AutoSize = $true
$CheckPCC_Button.BackColor = 'LightGray'

$adTree_Label = New-Object system.Windows.Forms.Label
$adTree_Label.Text = 'Select an OU'
$adTree_Label.Font = 'Segoe UI, 10pt,style=bold'
$adTree_Label.Dock = 'Fill'
$adTree_Label.Anchor = 'Bottom'
$adTree_Label.AutoSize = $true
$adTree = New-Object System.Windows.Forms.TreeView
$adTree.Dock = 'Fill'

$Submit_Button = New-Object System.Windows.Forms.Button
$Submit_Button.Text = 'Submit'
$Submit_Button.TabIndex = 7
$Submit_Button.Dock = 'Bottom'
$Submit_Button.AutoSize = $true
$ComputerInfo_Form.AcceptButton = $Submit_Button

#endregion
#region Login Window
$Login_Form = New-Object system.Windows.Forms.Form
$Login_Form.FormBorderStyle = "FixedDialog"
$Login_Form.TopMost = $true
$Login_Form.AutoScaleDimensions = '7,15'
$Login_Form.AutoScaleMode = 'Font'
$Login_Form.StartPosition = 'CenterScreen'
$Login_Form.Width = $($screen[0].bounds.Width / 5)
$Login_Form.Height = $($screen[0].bounds.Height / 5)
$Login_Form.ControlBox = $false
$Login_Form.AutoSize = $true
$Login_Form.MinimumSize = '100,100'

$Username_TextBox = New-Object system.Windows.Forms.TextBox
$Username_TextBox.multiline = $false
$Username_TextBox.Text = 'Username'
$Username_TextBox.TabIndex = 1
$Username_TextBox.Dock = 'Bottom'
$Username_TextBox.AutoSize = $true

$Password_TextBox = New-Object system.Windows.Forms.TextBox
$Password_TextBox.multiline = $false
$Password_TextBox.Text = "PimaRocks"
$Password_TextBox.TabIndex = 2
$Password_TextBox.PasswordChar = '*'
$Password_TextBox.Dock = 'Bottom'
$Password_TextBox.AutoSize = $true

$DomainSelection_Label = New-Object system.Windows.Forms.Label
$DomainSelection_Label.Text = 'Domain'
$DomainSelection_Label.Font = 'Segoe UI, 10pt,style=bold'
$DomainSelection_Label.Dock = 'Bottom'
$DomainSelection_Label.Anchor = 'Bottom'
$DomainSelection_Label.AutoSize = $true

$EDU_RadioButton = New-Object System.Windows.Forms.RadioButton
$EDU_RadioButton.Text = 'EDU'
$EDU_RadioButton.TabStop = $true
$EDU_RadioButton.Dock = 'Fill'
$EDU_RadioButton.AutoSize = $true
$EDU_RadioButton.CheckAlign = 'MiddleRight'
$EDU_RadioButton.TextAlign = 'MiddleRight'
$EDU_RadioButton.Checked = $true

$PCC_RadioButton = New-Object System.Windows.Forms.RadioButton
$PCC_RadioButton.Text = 'PCC'
$PCC_RadioButton.TabStop = $true
$PCC_RadioButton.Dock = 'Fill'
$PCC_RadioButton.AutoSize = $true

$OK_Button_Login = New-Object system.Windows.Forms.Button
$OK_Button_Login.Text = "Login"
$OK_Button_Login.Dock = 'Fill'
$OK_Button_Login.TabIndex = 3
$OK_Button_Login.FlatStyle = 1
$OK_Button_Login.FlatAppearance.BorderSize = 0
$Login_Form.AcceptButton = $OK_Button_Login
$Login_Form.AcceptButton.DialogResult = 'OK'

$Cancel_Button_Login = New-Object system.Windows.Forms.Button
$Cancel_Button_Login.Text = "Cancel"
$Cancel_Button_Login.TabIndex = 4
$Cancel_Button_Login.Dock = 'Fill'

$Login_Form.CancelButton = $Cancel_Button_Login
$Login_Form.CancelButton.DialogResult = 'Cancel'

#endregion
#region UI Layouts
#region Login Layout
$Domain_LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$Domain_LayoutPanel.Dock = "Fill"
$Domain_LayoutPanel.ColumnCount = 3
$Domain_LayoutPanel.RowCount = 2
[void]$Domain_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$Domain_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$Domain_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$Domain_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$Domain_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))

$Domain_LayoutPanel.Controls.Add($DomainSelection_Label, 1, 0)
$Domain_LayoutPanel.Controls.Add($EDU_RadioButton, 0, 1)
$Domain_LayoutPanel.Controls.Add($PCC_RadioButton, 2, 1)

$Login_LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$Login_LayoutPanel.Dock = "Fill"
$Login_LayoutPanel.ColumnCount = 2
$Login_LayoutPanel.RowCount = 3
$Login_LayoutPanel.CellBorderStyle = 3
[void]$Login_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$Login_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))

[void]$Login_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 30)))
[void]$Login_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$Login_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 30)))

$Login_LayoutPanel.Controls.Add($Username_TextBox, 0, 0)
$Login_LayoutPanel.Controls.Add($Password_TextBox, 1, 0)
$Login_LayoutPanel.Controls.Add($Domain_LayoutPanel, 0, 1)
$Login_LayoutPanel.SetColumnSpan($Domain_LayoutPanel, 2)
$Login_LayoutPanel.Controls.Add($OK_Button_Login, 0, 2)
$Login_LayoutPanel.Controls.Add($Cancel_Button_Login, 1, 2)

$Login_Form.controls.Add($Login_LayoutPanel)

#endregion

#region ComputerInfo UI Layout
$ComputerInfo_LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$ComputerInfo_LayoutPanel.Dock = "Fill"
$ComputerInfo_LayoutPanel.ColumnCount = 4
$ComputerInfo_LayoutPanel.RowCount = 6
$ComputerInfo_LayoutPanel.CellBorderStyle = 3
[void]$ComputerInfo_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 20)))
[void]$ComputerInfo_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 30)))
[void]$ComputerInfo_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$ComputerInfo_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))
[void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))

$ComputerInfo_LayoutPanel.Controls.Add($ComputerForm_Label, 0, 0)
$ComputerInfo_LayoutPanel.SetColumnSpan($ComputerForm_Label, 2)

$ComputerInfo_LayoutPanel.Controls.Add($Campus_Label, 0, 1)
$ComputerInfo_LayoutPanel.Controls.Add($Campus_Dropdown, 1, 1)

$ComputerInfo_LayoutPanel.Controls.Add($BuildingRoom_Label, 0, 2)
$ComputerInfo_LayoutPanel.Controls.Add($BuildingRoom_Textbox, 1, 2)

$ComputerInfo_LayoutPanel.Controls.Add($pccNumber_Label, 0, 3)
$ComputerInfo_LayoutPanel.Controls.Add($pccNumber_Textbox, 1, 3)

$ComputerInfo_LayoutPanel.Controls.Add($userSuffix_Label, 0, 4)
$ComputerInfo_LayoutPanel.Controls.Add($userSuffix_Dropdown, 1, 4)
$ComputerInfo_LayoutPanel.Controls.Add($hardwareSuffix_Label, 0, 5)
$ComputerInfo_LayoutPanel.Controls.Add($hardwareSuffix_Dropdown, 1, 5)

$ComputerInfo_LayoutPanel.Controls.Add($CheckPCC_Button, 0, 6)
$ComputerInfo_LayoutPanel.SetcolumnSpan($CheckPCC_Button, 2)

$ComputerInfo_LayoutPanel.Controls.Add($adTree_Label, 3, 0)
$ComputerInfo_LayoutPanel.Controls.Add($adTree, 3, 1)
$ComputerInfo_LayoutPanel.SetrowSpan($adTree, 5)

$ComputerInfo_LayoutPanel.Controls.Add($Submit_Button, 3, 6)
$ComputerInfo_Form.controls.Add($ComputerInfo_LayoutPanel)
#endregion
#endregion
#endregion
#region Functions

function Show-ADLoginWindow {
    # Show the login window and log the domain for later
    [void]$Login_Form.ShowDialog()

    # Check if user clicks 'OK'
    if ($Login_Form.DialogResult -eq 'OK') {
        # Convert password to secure string for security reasons
        $Password = ConvertTo-SecureString $Password_TextBox.Text -AsPlainText -Force
        $Credentials = New-Object System.Management.Automation.PSCredential ($Username_TextBox.text, $Password)
    
        # Determine which radio button is checked to determine which domain to query
        try {
            if ($PCC_RadioButton.Checked) {
                $ADDomain = Get-ADDomain -Credential $Credentials -Server $($PCC_RadioButton.text + '-domain.pima.edu')
            }
            else {
                $ADDomain = Get-ADDomain -Credential $Credentials -Server $($EDU_RadioButton.text + '-domain.pima.edu')
            }
        }
        # If login fails, recursively call the function again
        catch [System.Security.Authentication.AuthenticationException] {
            Show-ADLoginWindow
            return
        }
        
        # Check if the domain name matches with either PCC or EDU to determine if login is successful
        if ($ADDomain.Name -in $PCC_RadioButton.text, $EDU_RadioButton.text) {
            [void]$ComputerInfo_Form.ShowDialog()
            break
        }
        # If login fails, ask user if they want to retry or cancel
        else {
            $RelogChoice = [System.Windows.Forms.MessageBox]::Show("Login Failed, please relaunch.", 'Warning', 'RetryCancel', 'Warning')
            switch ($RelogChoice) {
                'Retry' { Show-ADLoginWindow }
                'Cancel' { exit }
            }        
        }
    }
    # If user clicks 'Cancel', prompt a warning message and reboot the computer after 5 seconds
    elseif ($Login_Form.DialogResult -eq 'Cancel') {
        [System.Windows.Forms.MessageBox]::Show("Login was cancelled, rebooting the computer.", 'Warning', 'Ok', 'Warning')
        Start-Sleep -Seconds 5
        Restart-Computer -Force -WhatIf
    }  
}

Function Confirm-ComputerName {

    # Verify each text box against regex to verify they are valid entries and if not verified, set error on that text box
    $CheckPCC_Button.BackColor = 'LightGray'

    try {
        # Verify a Campus from approved list is selected
        if (!($Campus_Dropdown.Items -contains $Campus_Dropdown.Text)) {
            throw 'Select a proper campus'
        }

        # Verify the building/room text
        if (!($BuildingRoom_Textbox.Text -match '^[a-z]{1}\d{3}$|^[a-z]{2}\d{2}$|^[a-z]{2}\d{3}$|^[a-z]{3}$')) {
            throw 'Enter a proper building/room'
        }

        # Verify a 6 digit number for PCC number
        if (!($pccNumber_Textbox.Text -match '^\d{6}$')) {
            throw 'Enter a proper PCC Number'
        }

        # Verify suffix
        if (!($userSuffix_Dropdown.Items -contains $userSuffix_Dropdown.Text) -or !($hardwareSuffix_Dropdown.Items -contains $hardwareSuffix_Dropdown.Text)) {
            throw 'Enter a proper suffix'
        }

        # Build $ComputerName which will be used to name the computer
        switch ($Campus_Dropdown.Text) {
            'DC' { 
                $Global:ComputerName = $Campus_Dropdown.Text + $BuildingRoom_Textbox.Text + $pccNumber_Textbox.Text + $userSuffixList[$userSuffix_Dropdown.SelectedIndex][0] + $hardwareSuffixList[$hardwareSuffix_Dropdown.SelectedIndex][0]
            }
            Default {
                $Global:ComputerName = $Campus_Dropdown.Text + '-' + $BuildingRoom_Textbox.Text + $pccNumber_Textbox.Text + $userSuffixList[$userSuffix_Dropdown.SelectedIndex][0] + $hardwareSuffixList[$hardwareSuffix_Dropdown.SelectedIndex][0]
            }
        }

        # Check computer name length
        if ($ComputerName.Length -gt 15) {
            throw 'Name too long'
        }

        $CheckPCC_Button.BackColor = 'Green'
    }
    catch {
        $ErrorProvider.SetError($ComputerForm_Label, $_)
        return
    }

    # Search AD to see if there is a computer with the supplied PCC number and notify the technician
    # Wishlist: Currently just shows a warning, deal with those results in some way?
    $PCCSearch = Get-ADComputer -Filter ('Name -Like "*' + $pccNumber_Textbox.Text + '*"') -Server $ADDomain.Forest
    if ($null -ne $PCCSearch) {
        $duplicateComputerList = ($PCCSearch.Name -join "`n")
        [System.Windows.Forms.MessageBox]::Show("The following system(s) matches the entered PCC Number:`n$duplicateComputerList`nYou may need to remove these entries!", 'Warning', 'Ok', 'Warning')
    }
}

Function Get-ADTreeNode ($Node, $CurrentOU) {
    # Used to populate Active Directory tree in the UI
    # Add the current OU to the tree as a child node
    $NodeSub = $Node.Nodes.Add($CurrentOU.DistinguishedName.toString(), $CurrentOU.Name)
    # Get all child OUs of the CurrentOU
    $ChildOUs = Get-ADOrganizationalUnit -Filter * -SearchScope OneLevel -SearchBase $CurrentOU -Server $ADDomain.Forest
    # Recursively add child OUs to the tree
    $ChildOUs | ForEach-Object { 
        Get-ADTreeNode $NodeSub $_ $ADDomain.Forest 
    }
}

#endregion
#region UI Actions

$Username_TextBox.Add_Click( { 
        $Username_TextBox.Clear()
    })

$Password_TextBox.Add_Click( { 
        $Password_TextBox.Clear()
    })

# Populates the AD tree based on the campus and domain selected
$Campus_Dropdown.Add_SelectedIndexChanged({

        # Set the label text and hide the treeview while it's being populated
        $adTree_Label.Text = "Loading OU's..."
        $adTree.Visible = $false
        $adTree.Nodes.Clear()

        # Determine the search base for Get-ADOrganizationalUnit based on the selected radio button
        $searchBase = if ($EDU_RadioButton.Checked) {
            "OU=EDU_Computers,DC=edu-domain,DC=pima,DC=edu"
        }
        elseif ($PCC_RadioButton.Checked) {
            if ($Campus_Dropdown.Items -contains $Campus_Dropdown.Text) {
                "OU=$($CampusList[$Campus_Dropdown.SelectedIndex][1]),OU=PCC,DC=PCC-Domain,DC=pima,DC=edu"
            }
            else {
                "OU=PCC,DC=PCC-Domain,DC=pima,DC=edu"
            }
        }

        # Populate the treeview with the OUs found
        Get-ADOrganizationalUnit -Filter * -SearchScope OneLevel -SearchBase $searchBase -Server $ADDomain.Forest | ForEach-Object {
            Get-ADTreeNode $adTree $_
        }

        # Reset the label text and show the treeview
        $adTree_Label.Text = 'Select an OU'
        $adTree.Visible = $true
    })

# If the machine has a PCC number set in the BIOS, pull that and enter it into the PCC number field
$PCCNumber = (Get-CimInstance -Query "Select * from Win32_SystemEnclosure").SMBiosAssetTag
if ($PCCNumber -match '^\d{6}$') {
    $pccNumber_Textbox.Text = $PCCNumber
    $pccNumber_Textbox.ReadOnly = $true
    $pccNumber_Label.Text = 'PCC# : Loaded from BIOS'
}
# Confirms entered computer name values are correct
$CheckPCC_Button.Add_Click( { 
        Confirm-ComputerName
    })


$Submit_Button.Add_Click( { 
        Confirm-ComputerName

        # Verify a target OU is selected
        if ($null -eq $adTree.SelectedNode) {
            $ErrorProvider.SetError($adTree_Label, 'Select an OU')
        }
        else {
            $ErrorProvider.SetError($adTree_Label, '')
        }          
        
        # Submit data to Task Sequence if there are no errors on UI elements
        if (-not($ErrorProvider.GetError($ComputerForm_Label) -or $ErrorProvider.GetError($adTree_Label))) {
            [System.Windows.Forms.MessageBox]::Show("Submitted Data:`n`nComputer Name: $($ComputerName.ToUpper())`n`nOU: $("LDAP://$($adTree.SelectedNode.Name)")`n`nDomain: $($ADDomain.Forest)", 'Warning', 'Ok', 'Warning')

            # Set the task sequence environment variables and close the form
            <#
            $TSEnvironment = New-Object -COMObject Microsoft.SMS.TSEnvironment 
            $TSEnvironment.Value("OSDComputerName") = "$($ComputerName.ToUpper())"
            $TSEnvironment.Value("OSDDomainOUName") = "$("LDAP://$($adTree.SelectedNode.Name)")"
            $TSEnvironment.Value("OSDDomainName") = "$($ADDomain.Forest)"     
            #>          
            [void]$ComputerInfo_Form.Close()
        }
    })

#endregion

#region Likely Values to be updated

# Used for Campus UI dropdown and useful Active Directory OU to computer name conversion
$CampusList = @(
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
    ('WC', 'West'), 
    ('PCC', 'West')
)

# Used for Suffix UI dropdowns and useful Active Directory OU to computer name conversion
$userSuffixList = @(
    ('A', 'Administrator'), 
    ('S', 'Staff'), 
    ('F', 'Faculty'), 
    ('I', 'Instructor'), 
    ('C', 'Class'), 
    ('L', 'Lab'), 
    ('P', 'Public'), 
    ('M', 'Meeting/conference'), 
    ('D', 'DPS')
)

$hardwareSuffixList = @(
    ('C', 'Computer'), 
    ('N', 'Notebook'),
    ('K', 'Kiosk'),
    ('T', 'Tablet'),
    ('V', 'Virtual Machine')
)

# Add items to dropdowns
$Campus_Dropdown.Items.AddRange(($CampusList | ForEach-Object { $_[0] }))
$userSuffix_Dropdown.Items.AddRange(($userSuffixList | ForEach-Object { $_[1] }))
$hardwareSuffix_Dropdown.Items.AddRange(($hardwareSuffixList | ForEach-Object { $_[1] }))

#endregion

# Launches main login window function which the gets AD creds needed for the rest of the script
Show-ADLoginWindow
# Enable to view computer info form for testing
#[void]$ComputerInfo_Form.ShowDialog()