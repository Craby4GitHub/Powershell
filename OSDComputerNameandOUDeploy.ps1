# 0.6.0
# Will Crabtree
# 

import-module activedirectory

#region GUI

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

$screen = [System.Windows.Forms.Screen]::AllScreens
#region Computer Info Window
$ComputerInfo_Form = New-Object System.Windows.Forms.Form    
$ComputerInfo_Form.AutoScaleDimensions = '7,15'
$ComputerInfo_Form.AutoScaleMode = 'Font'
$ComputerInfo_Form.StartPosition = 'CenterScreen'
$ComputerInfo_Form.Width = $($screen[0].bounds.Width / 4)
$ComputerInfo_Form.Height = $($screen[0].bounds.Height / 3)
$ComputerInfo_Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSHome + '\powershell.exe')
$ComputerInfo_Form.Text = 'Active Directory Information'
$ComputerInfo_Form.ControlBox = $false
$ComputerInfo_Form.TopMost = $true

$ComputerForm_Label = New-Object system.Windows.Forms.Label
$ComputerForm_Label.Text = 'Computer Name'
$ComputerForm_Label.Font = 'Segoe UI, 10pt,style=bold'
$ComputerForm_Label.Dock = 'Fill'
$ComputerForm_Label.Anchor = 'Bottom'
$ComputerForm_Label.AutoSize = $true

# Used for UI dropdown and useful AD OU to computername conversion
$CampusList = @(('29', '29th St.'), ('ER', 'El Rio'), ('EP', 'El Pueblo'), ('DV', 'Desert Vista'), ('DO', 'District'), ('DC', 'Downtown'), ('DM', 'DM'), ('EC', 'East'), ('MS', 'Maintenance and Security'), ('NW', 'Northwest'), ('WC', 'West'), ('PCC', 'West'))

$Campus_Label = New-Object system.Windows.Forms.Label
$Campus_Label.Text = 'Campus'
$Campus_Label.Font = 'Segoe UI, 8pt'
$Campus_Label.AutoSize = $true
$Campus_Dropdown = New-Object System.Windows.Forms.ComboBox
$Campus_Dropdown.DropDownStyle = 'DropDown'
$Campus_Dropdown.Items.AddRange(($CampusList | ForEach-Object { $_[0] }))
$Campus_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Campus_Dropdown.AutoCompleteSource = 'ListItems'
$Campus_Dropdown.Font = 'Segoe UI, 8pt'
$Campus_Dropdown.TabIndex = 1
$Campus_Dropdown.Dock = 'Top'

$BuildingRoom_Label = New-Object system.Windows.Forms.Label
$BuildingRoom_Label.Text = 'Bldg/Room'
$BuildingRoom_Label.Font = 'Segoe UI, 8pt'
$BuildingRoom_Label.AutoSize = $true
$BuildingRoom_Textbox = New-Object System.Windows.Forms.TextBox
$BuildingRoom_Textbox.Font = 'Segoe UI, 8pt'
$BuildingRoom_Textbox.Dock = 'Top'
$BuildingRoom_Textbox.TabIndex = 2

$pccNumber_Label = New-Object system.Windows.Forms.Label
$pccNumber_Label.Text = 'PCC#'
$pccNumber_Label.Font = 'Segoe UI, 8pt'
$pccNumber_Label.AutoSize = $true
$pccNumber_Textbox = New-Object System.Windows.Forms.TextBox
$pccNumber_Textbox.Font = 'Segoe UI, 8pt'
$pccNumber_Textbox.TabIndex = 3
$pccNumber_Textbox.Dock = 'Top'

$Suffix_Label = New-Object system.Windows.Forms.Label
$Suffix_Label.Text = 'Suffix'
$Suffix_Label.Font = 'Segoe UI, 8pt'
$Suffix_Label.AutoSize = $true
$Suffix_Textbox = New-Object System.Windows.Forms.TextBox
$Suffix_Textbox.TabIndex = 4
$Suffix_Textbox.Font = 'Segoe UI, 8pt'
$Suffix_Textbox.Dock = 'Top'

$CheckPCC_Button = New-Object System.Windows.Forms.Button
$CheckPCC_Button.Text = 'Check PCC'
$CheckPCC_Button.Font = 'Segoe UI, 8pt'
$CheckPCC_Button.TabIndex = 6
$CheckPCC_Button.AutoSize = $true
$CheckPCC_Button.BackColor = 'LightGray'

$adTree_Label = New-Object system.Windows.Forms.Label
$adTree_Label.Text = 'Select an OU'
$adTree_Label.Font = 'Segoe UI, 10pt,style=bold'
$adTree_Label.Dock = 'Bottom'
$adTree_Label.Anchor = 'Bottom'
$adTree_Label.AutoSize = $true
$adTree = New-Object System.Windows.Forms.TreeView
$adTree.Dock = 'Fill'

$Submit_Button = New-Object System.Windows.Forms.Button
$Submit_Button.Text = 'Submit'
$Submit_Button.Font = 'Segoe UI, 8pt'
$Submit_Button.TabIndex = 6
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

$Username_TextBox = New-Object system.Windows.Forms.TextBox
$Username_TextBox.multiline = $false
$Username_TextBox.Text = 'Username'
$Username_TextBox.TabIndex = 1
$Username_TextBox.Dock = 'Bottom'
$Username_TextBox.Anchor = 'Bottom'
$Username_TextBox.AutoSize = $true

$Password_TextBox = New-Object system.Windows.Forms.TextBox
$Password_TextBox.multiline = $false
$Password_TextBox.Text = "PimaRocks"
$Password_TextBox.TabIndex = 2
$Password_TextBox.PasswordChar = '*'
$Password_TextBox.Dock = 'Bottom'
$Password_TextBox.Anchor = 'Bottom'
$Password_TextBox.AutoSize = $true

$DomainSelection_Label = New-Object system.Windows.Forms.Label
$DomainSelection_Label.Text = 'Domain'
$DomainSelection_Label.Font = 'Segoe UI, 10pt,style=bold'
$DomainSelection_Label.Dock = 'Bottom'
$DomainSelection_Label.Anchor = 'Bottom'
$DomainSelection_Label.AutoSize = $true

$EDU_RadioButton = New-Object System.Windows.Forms.RadioButton
$EDU_RadioButton.Text = 'EDU'
$EDU_RadioButton.Font = 'Segoe UI, 8pt'
$EDU_RadioButton.TabStop = $true
$EDU_RadioButton.Dock = 'Fill'
$EDU_RadioButton.AutoSize = $true
$EDU_RadioButton.CheckAlign = 'MiddleRight'
$EDU_RadioButton.TextAlign = 'MiddleRight'
$EDU_RadioButton.Checked = $true

$PCC_RadioButton = New-Object System.Windows.Forms.RadioButton
$PCC_RadioButton.Text = 'PCC'
$PCC_RadioButton.Font = 'Segoe UI, 8pt'
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
#$Domain_LayoutPanel.CellBorderStyle = 1
[void]$Domain_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$Domain_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$Domain_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$Domain_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$Domain_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))

$Domain_LayoutPanel.Controls.Add($DomainSelection_Label, 1, 0)
$Domain_LayoutPanel.Controls.Add($EDU_RadioButton, 0, 1)
$Domain_LayoutPanel.Controls.Add($PCC_RadioButton, 2, 1)

$Login_LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$Login_LayoutPanel.Dock = "Fill"
$Login_LayoutPanel.ColumnCount = 4
$Login_LayoutPanel.RowCount = 5
#$Login_LayoutPanel.CellBorderStyle = 1
[void]$Login_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$Login_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$Login_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$Login_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))

[void]$Login_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$Login_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$Login_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$Login_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$Login_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))

$Login_LayoutPanel.Controls.Add($Username_TextBox, 1, 0)
$Login_LayoutPanel.SetColumnSpan($Username_TextBox, 2)
$Login_LayoutPanel.Controls.Add($Password_TextBox, 1, 1)
$Login_LayoutPanel.SetColumnSpan($Password_TextBox, 2)
$Login_LayoutPanel.Controls.Add($Domain_LayoutPanel, 1, 2)
$Login_LayoutPanel.SetColumnSpan($Domain_LayoutPanel, 2)
$Login_LayoutPanel.Controls.Add($OK_Button_Login, 1, 3)
$Login_LayoutPanel.Controls.Add($Cancel_Button_Login, 2, 3)

$Login_Form.controls.Add($Login_LayoutPanel)

#endregion

#region ComputerInfo UI Layout
$ComputerInfo_LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$ComputerInfo_LayoutPanel.Dock = "Fill"
$ComputerInfo_LayoutPanel.ColumnCount = 4
$ComputerInfo_LayoutPanel.RowCount = 5
#$ComputerInfo_LayoutPanel.CellBorderStyle = 1
[void]$ComputerInfo_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 20)))
[void]$ComputerInfo_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 30)))
[void]$ComputerInfo_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$ComputerInfo_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))
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

$ComputerInfo_LayoutPanel.Controls.Add($Suffix_Label, 0, 4)
$ComputerInfo_LayoutPanel.Controls.Add($Suffix_Textbox, 1, 4)

$ComputerInfo_LayoutPanel.Controls.Add($CheckPCC_Button, 0, 5)

$ComputerInfo_LayoutPanel.Controls.Add($adTree_Label, 3, 0)

$ComputerInfo_LayoutPanel.Controls.Add($adTree, 3, 1)
$ComputerInfo_LayoutPanel.SetrowSpan($adTree, 4)

$ComputerInfo_LayoutPanel.Controls.Add($Submit_Button, 3, 5)
$ComputerInfo_Form.controls.Add($ComputerInfo_LayoutPanel)
#endregion
#endregion
#endregion
#region Functions

function Login-AD {
    # Show the login window and log the domain for later
    [void]$Login_Form.ShowDialog()
    if ($Login_Form.DialogResult -eq 'OK') {
        $Password = ConvertTo-SecureString $Password_TextBox.Text -AsPlainText -Force
        $Credentials = New-Object System.Management.Automation.PSCredential ($Username_TextBox.text, $Password)
    
        # Checks what domain is selected and used for various useful... uses
        try {
            if ($PCC_RadioButton.Checked) {
                $ADDomain = Get-ADDomain -Credential $Credentials -Server $($PCC_RadioButton.text + '-domain.pima.edu')
            }
            else {
                $ADDomain = Get-ADDomain -Credential $Credentials -Server $($EDU_RadioButton.text + '-domain.pima.edu')
            }
        }
        # Recurse when the login fails
        catch [System.Security.Authentication.AuthenticationException] {
            Login-AD
        }
        
        # Comparing pulled AD name against UI text to verify the login was successful. continue to main core script otherwise recurse
        if (($ADDomain.Name -match $PCC_RadioButton.text) -or ($ADDomain.Name -match $EDU_RadioButton.text)) {
            [void]$ComputerInfo_Form.ShowDialog()
            break
        }
        else {
            $RelogChoice = [System.Windows.Forms.MessageBox]::Show("Login Failed, please relaunch.", 'Warning', 'RetryCancel', 'Warning')
            switch ($RelogChoice) {
                'Retry' { 
                    Login-AD
                }
                'Cancel' {
                    exit
                }
            }        
        }
    }
    elseif ($Login_Form.DialogResult -eq 'Cancel') {
        # Wishlist: Reboot computer if cancelled
    }  
}

Function Confirm-ComputerName {

    # Verify each text box agaisnt regex to verify they are valid entries and if not verified, set error on that element
    $CheckPCC_Button.BackColor = 'LightGray'
    if ($Campus_Dropdown.Items -contains $Campus_Dropdown.Text) {
        $ErrorProvider.SetError($ComputerForm_Label, '')
        if ($BuildingRoom_Textbox.Text -match '^[a-z]{1}\d{3}$|^[a-z]{2}\d{2}$|^[a-z]{2}\d{3}$|^[a-z]{3}$') {
            $ErrorProvider.SetError($ComputerForm_Label, '')
            if ($pccNumber_Textbox.Text -match '^\d{6}$') {
                $ErrorProvider.SetError($ComputerForm_Label, '')
                # Suffix check
                if ($Suffix_Textbox.Text -match '^[a-z]{2}$|[v]\d') {
                    $ErrorProvider.SetError($ComputerForm_Label, '')

                    # Check to see if this is a DC computer, as thier naming scheme doesnt include a dash
                    switch ($Campus_Dropdown.Text) {
                        'DC' { 
                            $Global:ComputerName = $Campus_Dropdown.Text + $BuildingRoom_Textbox.Text + $pccNumber_Textbox.Text + $Suffix_Textbox.Text
                        }
                        Default {
                            $Global:ComputerName = $Campus_Dropdown.Text + '-' + $BuildingRoom_Textbox.Text + $pccNumber_Textbox.Text + $Suffix_Textbox.Text
                        }
                    }
                    if ($ComputerName.Length -gt 15) {
                        $ErrorProvider.SetError($ComputerForm_Label, 'Name too long')
                    }
                    else {
                        $CheckPCC_Button.BackColor = 'Green'
                    }
                }
                else {
                    $ErrorProvider.SetError($ComputerForm_Label, 'Enter a proper suffix')
                }
            }
            else {
                $ErrorProvider.SetError($ComputerForm_Label, 'Enter a proper PCC Number')
            }
        }
        else {
            $ErrorProvider.SetError($ComputerForm_Label, 'Enter a proper building/room')
        }
    }
    else {
        $ErrorProvider.SetError($ComputerForm_Label, 'Select a proper campus')
    }

    
    # Search AD to see if there is a computer with the supplied PCC number and notify the technician
    # Wishlist: Currently just notifies, deal with those assets?
    $PCCSearch = Get-ADComputer -Filter ('Name -Like "*' + $pccNumber_Textbox.Text + '*"') -Server $ADDomain.Forest
    if ($null -ne $PCCSearch) {
        # Wishlist: Loop through each entry to out put results on new line
        [System.Windows.Forms.MessageBox]::Show("The following system(s) matches the entered PCC Number:`n$($PCCSearch.Name)", 'Warning', 'Ok', 'Warning')
    }
}

Function AddNodes ( $Node, $CurrentOU) {
    # Used to populate AD tree
    $NodeSub = $Node.Nodes.Add($CurrentOU.DistinguishedName.toString(), $CurrentOU.Name)
    Get-ADOrganizationalUnit -Filter * -SearchScope OneLevel -SearchBase $CurrentOU -Server $ADDomain.Forest | ForEach-Object { AddNodes $NodeSub $_ $ADDomain.Forest }
}
#endregion
#region UI Actions

# Populates the AD tree based on the campus and domain selected
$Campus_Dropdown.Add_SelectedIndexChanged( {
        $adTree.Nodes.Clear()
        if ($EDU_RadioButton.Checked -eq $true) {
            Get-ADOrganizationalUnit -Filter * -SearchScope OneLevel -SearchBase "OU=EDU_Computers,DC=edu-domain,DC=pima,DC=edu" -Server $ADDomain.Forest | ForEach-Object { AddNodes $adTree $_ }
        }
        elseif ($PCC_RadioButton.Checked -eq $true) {
            # Verify submitted campus, otherwise search whole domain
            if ($Campus_Dropdown.Items -contains $Campus_Dropdown.Text) {
                Get-ADOrganizationalUnit -Filter * -SearchScope OneLevel -SearchBase "OU=$($CampusList[$Campus_Dropdown.SelectedIndex][1]),OU=PCC,DC=PCC-Domain,DC=pima,DC=edu" -Server $ADDomain.Forest | ForEach-Object { AddNodes $adTree $_ }
            }
            else {
                Get-ADOrganizationalUnit -Filter * -SearchScope OneLevel -SearchBase "OU=PCC,DC=PCC-Domain,DC=pima,DC=edu" -Server $ADDomain.Forest | ForEach-Object { AddNodes $adTree $_ }
            }
        }  
    })

# If the machine has a PCC number set in the BIOS, pull that and enter it into the PCC number field
$PCCNumber = (Get-WmiObject -Query "Select * from Win32_SystemEnclosure").SMBiosAssetTag
if ($PCCNumber -match '^\d{6}$') {
    $pccNumber_Textbox.Text = $PCCNumber
    $pccNumber_Textbox.ReadOnly = $true
}

$CheckPCC_Button.Add_Click( { 
        Confirm-ComputerName
    })

$Submit_Button.Add_Click( { 
        Confirm-ComputerName

        # Verify a target OU is selected in UI
        if ($null -eq $adTree.SelectedNode) {
            $ErrorProvider.SetError($adTree_Label, 'Select an OU')
        }
        else {
            $ErrorProvider.SetError($adTree_Label, '')
        }          
        
        # Checks to see if there are errors on UI elements
        if (-not($ErrorProvider.GetError($ComputerForm_Label) -or $ErrorProvider.GetError($adTree_Label))) {
            $TSEnvironment = New-Object -COMObject Microsoft.SMS.TSEnvironment 
            $TSEnvironment.Value("OSDComputerName") = "$($ComputerName.ToUpper())"
            $TSEnvironment.Value("OSDDomainOUName") = "$("LDAP://$($adTree.SelectedNode.Name)")"
            $TSEnvironment.Value("OSDDomainName") = "$($ADDomain.Forest)"               
            [void]$ComputerInfo_Form.Close()
        }
    })

#endregion

# Gotta login
Login-AD
#[void]$ComputerInfo_Form.ShowDialog()