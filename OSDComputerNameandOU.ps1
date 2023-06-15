# 0.7.0
# Will Crabtree
# 

Import-Module ActiveDirectory

#region Likely values to be updated

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
    ('AT', 'Downtown'),
    ('PCC', 'West')
)

# Used for Suffix UI dropdowns and useful Active Directory OU to computer name conversion
$userSuffixList = @(
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

$hardwareSuffixList = @(
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
$ComputerInfo_Form.Padding = New-Object System.Windows.Forms.Padding(10)

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
$Campus_Dropdown.Items.AddRange(($CampusList | ForEach-Object { $_[0] }))

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
$userSuffix_Label.Text = 'User Suffix'
$userSuffix_Label.AutoSize = $true
$userSuffix_Dropdown = New-Object System.Windows.Forms.ComboBox
$userSuffix_Dropdown.DropDownStyle = 'DropDown'
$userSuffix_Dropdown.AutoCompleteMode = 'SuggestAppend'
$userSuffix_Dropdown.AutoCompleteSource = 'ListItems'
$userSuffix_Dropdown.TabIndex = 4
$userSuffix_Dropdown.Dock = 'Top'
$userSuffix_Dropdown.MinimumSize = '50,50'
$userSuffix_Dropdown.Items.AddRange(($userSuffixList | ForEach-Object { $_[1] }))

$hardwareSuffix_Label = New-Object system.Windows.Forms.Label
$hardwareSuffix_Label.Text = 'Hardware Suffix'
$hardwareSuffix_Label.AutoSize = $true
$hardwareSuffix_Dropdown = New-Object System.Windows.Forms.ComboBox
$hardwareSuffix_Dropdown.DropDownStyle = 'DropDown'
$hardwareSuffix_Dropdown.AutoCompleteMode = 'SuggestAppend'
$hardwareSuffix_Dropdown.AutoCompleteSource = 'ListItems'
$hardwareSuffix_Dropdown.TabIndex = 5
$hardwareSuffix_Dropdown.Dock = 'Top'
$hardwareSuffix_Dropdown.MinimumSize = '50,50'
$hardwareSuffix_Dropdown.Items.AddRange(($hardwareSuffixList | ForEach-Object { $_[1] }))

$CheckPCC_Button = New-Object System.Windows.Forms.Button
$CheckPCC_Button.Text = '&Check Name'
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
$adTree.TabIndex = 7

$Submit_Button = New-Object System.Windows.Forms.Button
$Submit_Button.Text = '&Submit'
$Submit_Button.TabIndex = 8
$Submit_Button.Dock = 'Bottom'
$Submit_Button.AutoSize = $true
$ComputerInfo_Form.AcceptButton = $Submit_Button

#endregion
#region Login Window
$Login_Form = New-Object System.Windows.Forms.Form
$Login_Form.ControlBox = $false
$Login_Form.TopMost = $true
$Login_Form.Text = 'Sign In'
$Login_Form.Size = New-Object System.Drawing.Size(300,200)
$Login_Form.StartPosition = 'CenterScreen'
$Login_Form.AutoSizeMode = 'GrowAndShrink'
$Login_Form.MinimumSize = New-Object System.Drawing.Size(200, 150)  # Minimum form size
$Login_Form.Padding = New-Object System.Windows.Forms.Padding(10)


$Username_Label = New-Object System.Windows.Forms.Label
$Username_Label.Text = 'Username:'
$Username_Label.Anchor = 'None'

$Username_TextBox = New-Object System.Windows.Forms.TextBox
$Username_TextBox.Dock = 'Fill'
$Username_TextBox.Anchor = 'Left, Right'
$Username_TextBox.MinimumSize = New-Object System.Drawing.Size(100, 0)
$Username_TextBox.TabIndex = 1


$Password_Label = New-Object System.Windows.Forms.Label
$Password_Label.Text = 'Password:'
$Password_Label.Anchor = 'None'

$Password_TextBox = New-Object System.Windows.Forms.TextBox
$Password_TextBox.Dock = 'Fill'
$Password_TextBox.Anchor = 'Left, Right'
$Password_TextBox.UseSystemPasswordChar = $true
$Password_TextBox.MinimumSize = New-Object System.Drawing.Size(100, 0)
$Password_TextBox.TabIndex = 2


$EDU_RadioButton = New-Object System.Windows.Forms.RadioButton
$EDU_RadioButton.Text = 'EDU'
$EDU_RadioButton.Anchor = 'None'
$EDU_RadioButton.TabStop = $true
$EDU_RadioButton.Checked = $true

$PCC_RadioButton = New-Object System.Windows.Forms.RadioButton
$PCC_RadioButton.Text = 'PCC'
$PCC_RadioButton.Anchor = 'None'
$PCC_RadioButton.TabStop = $true


$Login_ButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$Login_ButtonPanel.Dock = 'Fill'
$Login_ButtonPanel.Anchor = 'None'
$Login_ButtonPanel.AutoSize = $true
$Login_ButtonPanel.AutoSizeMode = 'GrowAndShrink'
$Login_ButtonPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight

$SignIn_Button = New-Object System.Windows.Forms.Button
$SignIn_Button.Text = '&Sign In'
$SignIn_Button.Anchor = 'None'
$SignIn_Button.AutoSize = $true
$SignIn_Button.MinimumSize = New-Object System.Drawing.Size(100, 0)
$SignIn_Button.TabIndex = 3
$Login_Form.AcceptButton = $SignIn_Button
$Login_Form.AcceptButton.DialogResult = 'OK'

$Cancel_Button_Login = New-Object System.Windows.Forms.Button
$Cancel_Button_Login.Text = '&Cancel'
$Cancel_Button_Login.Anchor = 'None'
$Cancel_Button_Login.AutoSize = $true
$Cancel_Button_Login.MinimumSize = New-Object System.Drawing.Size(100, 0)
$Cancel_Button_Login.TabIndex = 4
$Login_Form.CancelButton = $Cancel_Button_Login
$Login_Form.CancelButton.DialogResult = 'Cancel'


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

$Login_Form.Controls.Add($Login_LayoutPanel)
$Login_LayoutPanel.Controls.Add($Username_Label, 0, 0)
$Login_LayoutPanel.Controls.Add($Username_TextBox, 1, 0)
$Login_LayoutPanel.Controls.Add($Password_Label, 0, 1)
$Login_LayoutPanel.Controls.Add($Password_TextBox, 1, 1)
$Login_LayoutPanel.Controls.Add($EDU_RadioButton, 0, 2)
$Login_LayoutPanel.Controls.Add($PCC_RadioButton, 1, 2)
$Login_LayoutPanel.Controls.Add($Login_ButtonPanel, 0, 3)
$Login_LayoutPanel.SetColumnSpan($Login_ButtonPanel, 2)

$Login_ButtonPanel.Controls.Add($SignIn_Button)
$Login_ButtonPanel.Controls.Add($Cancel_Button_Login)
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
        if (($ADDomain.Name -match $PCC_RadioButton.text) -or ($ADDomain.Name -match $EDU_RadioButton.text)) {
            [void]$ComputerInfo_Form.ShowDialog()
            break
        }
        # If login fails, ask user if they want to retry or cancel
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
    # If user clicks 'Cancel', prompt a warning message and reboot the computer after 5 seconds
    elseif ($Login_Form.DialogResult -eq 'Cancel') {
        [System.Windows.Forms.MessageBox]::Show("Login was cancelled, rebooting the computer.", 'Warning', 'Ok', 'Warning')
        Start-Sleep -Seconds 5
        Restart-Computer -Force -WhatIf
    }  
}

Function Confirm-ComputerName {
    # Clear errors first
    $ErrorProvider.SetError($ComputerForm_Label, '')

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
    $NodeSub = $Node.Nodes.Add($CurrentOU.DistinguishedName.toString(), $CurrentOU.Name)
    
    # Add a placeholder node that will be replaced when the user expands the node.
    $placeholder = $NodeSub.Nodes.Add("Loading...")
}

#endregion
#region GUI Actions
$adTree.Add_BeforeExpand({
        param($sender, $e)
    
        # The node that is being expanded.
        $node = $e.Node

        # Check if the node contains a placeholder.
        if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Text -eq "Loading...") {
            # Remove the placeholder.
            $node.Nodes.Clear()

            # Add the child nodes.
            Get-ADOrganizationalUnit -Filter * -SearchScope OneLevel -SearchBase $node.Name -Server $ADDomain.Forest | ForEach-Object { Get-ADTreeNode $node $_ }
        }
    })


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
        # populate the treeview with the OUs found
        Get-ADOrganizationalUnit -Filter * -SearchScope OneLevel -SearchBase $searchBase -Server $ADDomain.Forest | ForEach-Object { Get-ADTreeNode $adTree $_ }
        $adTree_Label.Text = 'Select an OU'
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


# Launches main login window function which the gets AD creds needed for the rest of the script
Show-ADLoginWindow
# Enable to view computer info form for testing
#[void]$ComputerInfo_Form.ShowDialog()