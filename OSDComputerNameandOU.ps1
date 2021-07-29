function OSD-GUI {
    param (
        $Campus,
        $Bldg,
        $PCC,
        $Suffix,
        $Domain
    )

    #region GUI

    #[void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    #[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') 

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

    $screen = [System.Windows.Forms.Screen]::AllScreens
    #region Computer Info
    $ComputerInfo_Form = New-Object System.Windows.Forms.Form    
    $ComputerInfo_Form.AutoScaleDimensions = '7,15'
    $ComputerInfo_Form.AutoScaleMode = 'Font'
    $ComputerInfo_Form.StartPosition = 'CenterScreen'
    $ComputerInfo_Form.Width = $($screen[0].bounds.Width / 5)
    $ComputerInfo_Form.Height = $($screen[0].bounds.Height / 3)
    $ComputerInfo_Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSHome + '\powershell.exe')
    $ComputerInfo_Form.Text = 'Active Directory Information'
    $ComputerInfo_Form.ControlBox = $false
    $ComputerInfo_Form.TopMost = $true

    $ComputerName_Label = New-Object system.Windows.Forms.Label
    $ComputerName_Label.Text = 'Create Computer Name'
    $ComputerName_Label.Font = 'Segoe UI, 10pt,style=bold'
    $ComputerName_Label.Dock = 'Fill'
    $ComputerName_Label.Anchor = 'Bottom'
    $ComputerName_Label.AutoSize = $true

    $CampusList = @(('29', '29th St.'), ('ER', 'El Rio'), ('EP', 'El Pueblo'), ('DV', 'Desert Vista'), ('DO', 'District'), ('DC', 'Downtown'), ('DM', 'DM'), ('EC', 'East'), ('MS', 'Maintenance and Security'), ('NW', 'Northwest'), ('WC', 'West'), ('PCC', 'West'))

    $ComputerName_Campus_Label = New-Object system.Windows.Forms.Label
    $ComputerName_Campus_Label.Text = 'Campus'
    $ComputerName_Campus_Label.Font = 'Segoe UI, 8pt'
    $ComputerName_Campus_Label.Dock = 'Bottom'
    $ComputerName_Campus_Label.Anchor = 'Bottom'
    $ComputerName_Campus_Label.AutoSize = $true
    $ComputerName_Campus_Dropdown = New-Object System.Windows.Forms.ComboBox
    $ComputerName_Campus_Dropdown.DropDownStyle = 'DropDown'
    $ComputerName_Campus_Dropdown.Items.AddRange(($CampusList | ForEach-Object { $_[0] }))
    $ComputerName_Campus_Dropdown.AutoCompleteMode = 'SuggestAppend'
    $ComputerName_Campus_Dropdown.AutoCompleteSource = 'ListItems'
    $ComputerName_Campus_Dropdown.Font = 'Segoe UI, 8pt'
    $ComputerName_Campus_Dropdown.TabIndex = 1
    $ComputerName_Campus_Dropdown.Dock = 'Top'

    $ComputerName_BuildingRoom_Label = New-Object system.Windows.Forms.Label
    $ComputerName_BuildingRoom_Label.Text = 'Bldg/Room'
    $ComputerName_BuildingRoom_Label.Font = 'Segoe UI, 8pt'
    $ComputerName_BuildingRoom_Label.Dock = 'Bottom'
    $ComputerName_BuildingRoom_Label.Anchor = 'Bottom'
    $ComputerName_BuildingRoom_Label.AutoSize = $true
    $ComputerName_BuildingRoom_Textbox = New-Object System.Windows.Forms.TextBox
    $ComputerName_BuildingRoom_Textbox.Font = 'Segoe UI, 8pt'
    $ComputerName_BuildingRoom_Textbox.Dock = 'Top'
    $ComputerName_BuildingRoom_Textbox.TabIndex = 2

    $ComputerName_PCCNumber_Label = New-Object system.Windows.Forms.Label
    $ComputerName_PCCNumber_Label.Text = 'PCC#'
    $ComputerName_PCCNumber_Label.Font = 'Segoe UI, 8pt'
    $ComputerName_PCCNumber_Label.Dock = 'Bottom'
    $ComputerName_PCCNumber_Label.Anchor = 'Bottom'
    $ComputerName_PCCNumber_Label.AutoSize = $true
    $ComputerName_PCCNumber_Textbox = New-Object System.Windows.Forms.TextBox
    $ComputerName_PCCNumber_Textbox.Font = 'Segoe UI, 8pt'
    $ComputerName_PCCNumber_Textbox.TabIndex = 3
    $ComputerName_PCCNumber_Textbox.Dock = 'Top'

    $ComputerName_Suffix_Label = New-Object system.Windows.Forms.Label
    $ComputerName_Suffix_Label.Text = 'Suffix'
    $ComputerName_Suffix_Label.Font = 'Segoe UI, 8pt'
    $ComputerName_Suffix_Label.Dock = 'Bottom'
    $ComputerName_Suffix_Label.Anchor = 'Bottom'
    $ComputerName_Suffix_Label.AutoSize = $true
    $ComputerName_Suffix_Textbox = New-Object System.Windows.Forms.TextBox
    $ComputerName_Suffix_Textbox.TabIndex = 4
    $ComputerName_Suffix_Textbox.Font = 'Segoe UI, 8pt'
    $ComputerName_Suffix_Textbox.Dock = 'Top'

    $ComputerName_Check_Button = New-Object System.Windows.Forms.Button
    $ComputerName_Check_Button.Text = 'Verify'
    $ComputerName_Check_Button.Font = 'Segoe UI, 8pt'
    $ComputerName_Check_Button.TabIndex = 6
    $ComputerName_Check_Button.Dock = 'Bottom'
    $ComputerName_Check_Button.Anchor = 'Bottom'
    $ComputerName_Check_Button.AutoSize = $true
    $ComputerName_Check_Button.BackColor = 'LightGray'

    $ADOUTree_Label = New-Object system.Windows.Forms.Label
    $ADOUTree_Label.Text = 'Select an OU'
    $ADOUTree_Label.Font = 'Segoe UI, 10pt,style=bold'
    $ADOUTree_Label.Dock = 'Bottom'
    $ADOUTree_Label.Anchor = 'Bottom'
    $ADOUTree_Label.AutoSize = $true
    $ADOUTree = New-Object System.Windows.Forms.TreeView
    $ADOUTree.Dock = 'Fill'

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

    #EndRegion
    #region All UI Layouts
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

    $LayoutPanel_Login = New-Object System.Windows.Forms.TableLayoutPanel
    $LayoutPanel_Login.Dock = "Fill"
    $LayoutPanel_Login.ColumnCount = 4
    $LayoutPanel_Login.RowCount = 5
    #$LayoutPanel_Login.CellBorderStyle = 1
    [void]$LayoutPanel_Login.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
    [void]$LayoutPanel_Login.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
    [void]$LayoutPanel_Login.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
    [void]$LayoutPanel_Login.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))

    [void]$LayoutPanel_Login.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
    [void]$LayoutPanel_Login.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
    [void]$LayoutPanel_Login.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
    [void]$LayoutPanel_Login.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
    [void]$LayoutPanel_Login.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))

    $LayoutPanel_Login.Controls.Add($Username_TextBox, 1, 0)
    $LayoutPanel_Login.SetColumnSpan($Username_TextBox, 2)
    $LayoutPanel_Login.Controls.Add($Password_TextBox, 1, 1)
    $LayoutPanel_Login.SetColumnSpan($Password_TextBox, 2)
    $LayoutPanel_Login.Controls.Add($Domain_LayoutPanel, 1, 2)
    $LayoutPanel_Login.SetColumnSpan($Domain_LayoutPanel, 2)
    $LayoutPanel_Login.Controls.Add($OK_Button_Login, 1, 3)
    $LayoutPanel_Login.Controls.Add($Cancel_Button_Login, 2, 3)

    $Login_Form.controls.Add($LayoutPanel_Login)

    #endregion
    #region Main Layout
    #region Computer Name
    $ComputerName_LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $ComputerName_LayoutPanel.Dock = "Fill"
    $ComputerName_LayoutPanel.ColumnCount = 5
    $ComputerName_LayoutPanel.RowCount = 3
    #$ComputerName_LayoutPanel.CellBorderStyle = 1
    [void]$ComputerName_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 20)))
    [void]$ComputerName_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
    [void]$ComputerName_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 20)))
    [void]$ComputerName_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 15)))
    [void]$ComputerName_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 20)))
    [void]$ComputerName_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20)))
    [void]$ComputerName_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
    [void]$ComputerName_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))

    $ComputerName_LayoutPanel.Controls.Add($ComputerName_Label, 0, 0)
    $ComputerName_LayoutPanel.SetColumnSpan($ComputerName_Label, 5)

    $ComputerName_LayoutPanel.Controls.Add($ComputerName_Campus_Label, 0, 1)
    $ComputerName_LayoutPanel.Controls.Add($ComputerName_Campus_Dropdown, 0, 2)

    $ComputerName_LayoutPanel.Controls.Add($ComputerName_BuildingRoom_Label, 1, 1)
    $ComputerName_LayoutPanel.Controls.Add($ComputerName_BuildingRoom_Textbox, 1, 2)

    $ComputerName_LayoutPanel.Controls.Add($ComputerName_PCCNumber_Label, 2, 1)
    $ComputerName_LayoutPanel.Controls.Add($ComputerName_PCCNumber_Textbox, 2, 2)

    $ComputerName_LayoutPanel.Controls.Add($ComputerName_Suffix_Label, 3, 1)
    $ComputerName_LayoutPanel.Controls.Add($ComputerName_Suffix_Textbox, 3, 2)

    $ComputerName_LayoutPanel.Controls.Add($ComputerName_Check_Button, 4, 2)

    #endregion
    $ComputerInfo_LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $ComputerInfo_LayoutPanel.Dock = "Fill"
    $ComputerInfo_LayoutPanel.ColumnCount = 3
    $ComputerInfo_LayoutPanel.RowCount = 4
    #$ComputerInfo_LayoutPanel.CellBorderStyle = 1
    [void]$ComputerInfo_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
    [void]$ComputerInfo_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
    [void]$ComputerInfo_LayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
    [void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
    [void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15)))
    [void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    [void]$ComputerInfo_LayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15)))

    $ComputerInfo_LayoutPanel.Controls.Add($ComputerName_LayoutPanel, 0, 0)
    $ComputerInfo_LayoutPanel.SetColumnSpan($ComputerName_LayoutPanel, 3)

    $ComputerInfo_LayoutPanel.Controls.Add($ADOUTree_Label, 0, 1)
    $ComputerInfo_LayoutPanel.SetColumnSpan($ADOUTree_Label, 3)
    $ComputerInfo_LayoutPanel.Controls.Add($ADOUTree, 0, 2)
    $ComputerInfo_LayoutPanel.SetColumnSpan($ADOUTree, 3)

    $ComputerInfo_LayoutPanel.Controls.Add($Submit_Button, 1, 3)
    $ComputerInfo_Form.controls.Add($ComputerInfo_LayoutPanel)
    #endregion
    #endregion
    #endregion
    #region Functions
    function Confirm-NoError {
        if ($ErrorProvider.GetError($ComputerName_Label) -or $ErrorProvider.GetError($ADOUTree_Label)) {
            return $false
        }
        else {
            return $true
        }
    }

    Function Confirm-ComputerName {
        $ComputerName_Check_Button.BackColor = 'LightGray'
        if ($ComputerName_Campus_Dropdown.Items -contains $ComputerName_Campus_Dropdown.Text) {
            $ErrorProvider.SetError($ComputerName_Label, '')
            if ($ComputerName_BuildingRoom_Textbox.Text -match '^[a-z]{1}\d{3}$|^[a-z]{2}\d{2}$|^[a-z]{2}\d{3}$|^[a-z]{3}$') {
                $ErrorProvider.SetError($ComputerName_Label, '')
                if ($ComputerName_PCCNumber_Textbox.Text -match '^\d{6}$') {
                    $ErrorProvider.SetError($ComputerName_Label, '')
                    $PCCSearch = Get-ADComputer -Filter ('Name -Like "*' + $ComputerName_PCCNumber_Textbox.Text + '*"') -Server $ADDomain.Forest
                    if ($null -ne $PCCSearch) {
                        [System.Windows.Forms.MessageBox]::Show("The following system(s) matches the entered PCC Number:`n$($PCCSearch.Name)`nImaging will continue", 'Warning', 'Ok', 'Warning')
                    }
                    if ($ComputerName_Suffix_Textbox.Text -match '^[a-z]{2}$|') {
                        $ErrorProvider.SetError($ComputerName_Label, '')
                        if ($ComputerName_Campus_Dropdown.Text -ne 'DC') {
                            $Global:ComputerName = $ComputerName_Campus_Dropdown.Text + '-' + $ComputerName_BuildingRoom_Textbox.Text + $ComputerName_PCCNumber_Textbox.Text + $ComputerName_Suffix_Textbox.Text
                            if ($ComputerName.Length -gt 15) {
                                $ErrorProvider.SetError($ComputerName_Label, 'Name too long')
                            }else {
                                $ComputerName_Check_Button.BackColor = 'Green'
                            }
                        }
                        else {
                            $Global:ComputerName = $ComputerName_Campus_Dropdown.Text + $ComputerName_BuildingRoom_Textbox.Text + $ComputerName_PCCNumber_Textbox.Text + $ComputerName_Suffix_Textbox.Text
                            if ($ComputerName.Length -gt 15) {
                                $ErrorProvider.SetError($ComputerName_Label, 'Name too long')
                            }else {
                                $ComputerName_Check_Button.BackColor = 'Green'
                            }
                        }
                    }
                    else {
                        $ErrorProvider.SetError($ComputerName_Label, 'Enter a proper suffix')
                    }
                }
                else {
                    $ErrorProvider.SetError($ComputerName_Label, 'Enter a proper PCC Number')
                }
            }
            else {
                $ErrorProvider.SetError($ComputerName_Label, 'Enter a proper building/room')
            }
        }
        else {
            $ErrorProvider.SetError($ComputerName_Label, 'Select a proper campus')
        }
    }

    Function AddNodes ( $Node, $CurrentOU) {
        $NodeSub = $Node.Nodes.Add($CurrentOU.DistinguishedName.toString(), $CurrentOU.Name)
        Get-ADOrganizationalUnit -Filter * -SearchScope OneLevel -SearchBase $CurrentOU -Server $ADDomain.Forest | ForEach-Object { AddNodes $NodeSub $_ $ADDomain.Forest }
    }
    #endregion
    #region Actions
    $ComputerName_Campus_Dropdown.Add_SelectedIndexChanged( {
            $ADOUTree.Nodes.Clear()
            if ($EDU_RadioButton.Checked -eq $true) {
                Get-ADOrganizationalUnit -Filter * -SearchScope OneLevel -SearchBase "OU=EDU_Computers,DC=edu-domain,DC=pima,DC=edu" -Server $ADDomain.Forest | ForEach-Object { AddNodes $ADOUTree $_ }
            }
            elseif ($PCC_RadioButton.Checked -eq $true) {
                if ($ComputerName_Campus_Dropdown.Items -contains $ComputerName_Campus_Dropdown.Text) {
                    Get-ADOrganizationalUnit -Filter * -SearchScope OneLevel -SearchBase "OU=$($CampusList[$ComputerName_Campus_Dropdown.SelectedIndex][1]),OU=PCC,DC=PCC-Domain,DC=pima,DC=edu" -Server $ADDomain.Forest | ForEach-Object { AddNodes $ADOUTree $_ }
                }
                else {
                    Get-ADOrganizationalUnit -Filter * -SearchScope OneLevel -SearchBase "OU=PCC,DC=PCC-Domain,DC=pima,DC=edu" -Server $ADDomain.Forest | ForEach-Object { AddNodes $ADOUTree $_ }
                }
            }  
        })

    $PCCNumber = (Get-WmiObject -Query "Select * from Win32_SystemEnclosure").SMBiosAssetTag
    if ($PCCNumber -match '^\d{6}$') {
        $ComputerName_PCCNumber_Textbox.Text = $PCCNumber
        $ComputerName_PCCNumber_Textbox.ReadOnly = $true
    }
    $ComputerName_Check_Button.Add_Click( { 
        Confirm-ComputerName
        })
    $Submit_Button.Add_Click( { 
        Confirm-ComputerName
            if ($null -eq $ADOUTree.SelectedNode) {
                $ErrorProvider.SetError($ADOUTree_Label, 'Select an OU')
            }
            else {
                $ErrorProvider.SetError($ADOUTree_Label, '')
            }          
        
            if (Confirm-NoError) {
                # Output for testing
                $OULocation = "LDAP://$($ADOUTree.SelectedNode.Name)"
                $ComputerName.ToUpper(), $OULocation, 'Passed' | Out-File '.\OSDTest.csv' -Append

                #Enable when deployed
                #$TSEnvironment = New-Object -COMObject Microsoft.SMS.TSEnvironment 
                #$TSEnvironment.Value("OSDComputerName") = "$($ComputerName.ToUpper())"
                #$TSEnvironment.Value("OSDDomainOUName") = "$($OULocation)"
                #$TSEnvironment.Value("OSDDomainName") = "$($Domain)"               
                [void]$ComputerInfo_Form.Close()
            }
            else {
                $OULocation = "LDAP://$($ADOUTree.SelectedNode.Name)"
                $ComputerName.ToUpper(), $OULocation, 'Failed' | Out-File '.\OSDTest.csv' -Append
            }
        })

    #endregion
    #region Main
    [void]$Login_Form.ShowDialog()
    if ($Login_Form.DialogResult -eq 'OK') {
        $Password = ConvertTo-SecureString $Password_TextBox.Text -AsPlainText -Force
        $global:Credentials = New-Object System.Management.Automation.PSCredential ($Username_TextBox.text, $Password)
        try {
            if ($PCC_RadioButton.Checked) {
                $ADDomain = Get-ADDomain -Credential $Credentials -Server $($PCC_RadioButton.text + '-domain.pima.edu')
            }
            else {
                $ADDomain = Get-ADDomain -Credential $Credentials -Server $($EDU_RadioButton.text + '-domain.pima.edu')
            }
        }
        catch [System.Security.Authentication.AuthenticationException] {
            Write-Host "incorrect login"
        }
        catch {
            Write-Host 'hm...'
        }

        if (($ADDomain.Name -match $PCC_RadioButton.text) -or ($ADDomain.Name -match $EDU_RadioButton.text)) {
            # Test Values
            $ComputerName_BuildingRoom_Textbox.Text = $Bldg
            $ComputerName_PCCNumber_Textbox.Text = $PCC
            $ComputerName_Suffix_Textbox.Text = $Suffix

            [void]$ComputerInfo_Form.ShowDialog()
            break
        }
        else {
            $RelogChoice = [System.Windows.Forms.MessageBox]::Show("Login Failed, please relaunch.", 'Warning', 'RetryCancel', 'Warning')
            switch ($RelogChoice) {
                # Need to figure out logic to re-enter creds and what not
                'Retry' { 
                    # Retry is not working
                    [void]$Login_Form.ShowDialog() 
                    break
                }
                'Cancel' {
                    #Write-Log -Message "$($Credentials.UserName) failed to login" -Level WARN
                    #exit
                }
            }        
        }
    }
    elseif ($Login_Form.DialogResult -eq 'Cancel') {
    }
    #endregion
}

$CampusShortList = @('29', 'ER', 'EP', 'DV', 'DO', 'DC', 'EC', 'MS', 'NW', 'WC', 'PCC')
$RandomRooms = @('CG11', 'E513', 'AH321', 'emp', 'stu')

for ($i = 0; $i -lt 1; $i++) {
    OSD-GUI -Campus $(Get-Random -InputObject $CampusShortList) -Bldg $(Get-Random -InputObject $RandomRooms) -PCC $(Get-Random -Maximum 999999 -Minimum 111111) -Suffix $( -join ((65..90) + (97..121) | Get-Random -Count 2 | ForEach-Object { [char]$_ })) -Domain $(Get-Random -InputObject 'EDU', 'PCC')
}