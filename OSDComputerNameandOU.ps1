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

    $Form = New-Object System.Windows.Forms.Form    
    $Form.AutoScaleDimensions = '7,15'
    $Form.AutoScaleMode = 'Font'
    $Form.StartPosition = 'CenterScreen'
    $Form.Width = $($screen[0].bounds.Width / 5)
    $Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSHome + '\powershell.exe')
    $Form.Text = 'Active Directory Information'
    $Form.ControlBox = $false
    $Form.TopMost = $true

    $Main_LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $Main_LayoutPanel.Dock = "Fill"
    $Main_LayoutPanel.ColumnCount = 3
    $Main_LayoutPanel.RowCount = 4
    #$Main_LayoutPanel.CellBorderStyle = 1
    [void]$Main_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
    [void]$Main_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
    [void]$Main_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
    [void]$Main_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
    [void]$Main_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
    [void]$Main_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15)))


    $ComputerName_LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $ComputerName_LayoutPanel.Dock = "Fill"
    $ComputerName_LayoutPanel.ColumnCount = 4
    $ComputerName_LayoutPanel.RowCount = 3
    #$ComputerName_LayoutPanel.CellBorderStyle = 1
    [void]$ComputerName_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
    [void]$ComputerName_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
    [void]$ComputerName_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
    [void]$ComputerName_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
    [void]$ComputerName_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20)))
    [void]$ComputerName_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
    [void]$ComputerName_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))

    $ComputerName_Label = New-Object system.Windows.Forms.Label
    $ComputerName_Label.Text = 'Create Computer Name'
    $ComputerName_Label.Font = 'Segoe UI, 10pt,style=bold'
    $ComputerName_Label.Dock = 'Bottom'
    $ComputerName_Label.Anchor = 'Bottom'
    $ComputerName_Label.AutoSize = $true

    $CampusList = @(('29', 'Adult Education'), ('ER', 'Adult Education'), ('EP', 'Adult Education'), ('DV', 'Desert Vista'), ('DO', 'District'), ('DC', 'Downtown'), ('EC', 'East'), ('MS', 'Maintenance and Security'), ('NW', 'Northwest'), ('WC', 'West'), ('PCC', 'West'))

    $ComputerName_Campus_Label = New-Object system.Windows.Forms.Label
    $ComputerName_Campus_Label.Text = 'Campus'
    $ComputerName_Campus_Label.Font = 'Segoe UI, 8pt'
    $ComputerName_Campus_Label.Dock = 'Bottom'
    $ComputerName_Campus_Label.Anchor = 'Bottom'
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


    $Domain_LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $Domain_LayoutPanel.Dock = "Fill"
    $Domain_LayoutPanel.ColumnCount = 3
    $Domain_LayoutPanel.RowCount = 2
    #$Domain_LayoutPanel.CellBorderStyle = 1
    [void]$Domain_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
    [void]$Domain_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
    [void]$Domain_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
    [void]$Domain_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    [void]$Domain_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))

    $DomainSelection_Label = New-Object system.Windows.Forms.Label
    $DomainSelection_Label.Text = 'Select a Domain'
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

    $PCC_RadioButton = New-Object System.Windows.Forms.RadioButton
    $PCC_RadioButton.Text = 'PCC'
    $PCC_RadioButton.Font = 'Segoe UI, 8pt'
    $PCC_RadioButton.TabStop = $true
    $PCC_RadioButton.Dock = 'Fill'
    $PCC_RadioButton.AutoSize = $true

    $Submit_Button = New-Object System.Windows.Forms.Button
    $Submit_Button.Text = 'Submit'
    $Submit_Button.Font = 'Segoe UI, 8pt'
    $Submit_Button.TabIndex = 6
    $Submit_Button.Dock = 'Bottom'
    $Submit_Button.AutoSize = $true
    $Form.AcceptButton = $Submit_Button

    $Form.controls.Add($Main_LayoutPanel)
    #region UI Layout

    #region Main Layout

    $Main_LayoutPanel.Controls.Add($ComputerName_LayoutPanel, 0, 0)
    $Main_LayoutPanel.SetColumnSpan($ComputerName_LayoutPanel, 3)

    $Main_LayoutPanel.Controls.Add($Domain_LayoutPanel, 0, 1)
    $Main_LayoutPanel.SetColumnSpan($Domain_LayoutPanel, 3)

    $Main_LayoutPanel.Controls.Add($Submit_Button, 1, 2)
    #endregion

    #region Computer Name
    $ComputerName_LayoutPanel.Controls.Add($ComputerName_Label, 1, 0)
    $ComputerName_LayoutPanel.SetColumnSpan($ComputerName_Label, 2)
    $ComputerName_LayoutPanel.Controls.Add($ComputerName_Campus_Label, 0, 1)
    $ComputerName_LayoutPanel.Controls.Add($ComputerName_Campus_Dropdown, 0, 2)

    $ComputerName_LayoutPanel.Controls.Add($ComputerName_BuildingRoom_Label, 1, 1)
    $ComputerName_LayoutPanel.Controls.Add($ComputerName_BuildingRoom_Textbox, 1, 2)

    $ComputerName_LayoutPanel.Controls.Add($ComputerName_PCCNumber_Label, 2, 1)
    $ComputerName_LayoutPanel.Controls.Add($ComputerName_PCCNumber_Textbox, 2, 2)

    $ComputerName_LayoutPanel.Controls.Add($ComputerName_Suffix_Label, 3, 1)
    $ComputerName_LayoutPanel.Controls.Add($ComputerName_Suffix_Textbox, 3, 2)
    #endregion

    #region Domain Selection
    $Domain_LayoutPanel.Controls.Add($DomainSelection_Label, 1, 0)
    $Domain_LayoutPanel.Controls.Add($EDU_RadioButton, 0, 1)
    $Domain_LayoutPanel.Controls.Add($PCC_RadioButton, 2, 1)
    #endregion
    #endregion
    #endregion

    #region Functions

    function Confirm-NoError {
        if ($ErrorProvider.GetError($ComputerName_Label) -or $ErrorProvider.GetError($DomainSelection_Label)) {
            return $false
        }
        else {
            return $true
        }
    }
    #endregion

    #region Actions

    $PCCNumber = (Get-WmiObject -Query "Select * from Win32_SystemEnclosure").SMBiosAssetTag
    if ($PCCNumber -match '^\d{6}$') {
        $ComputerName_PCCNumber_Textbox.Text = $PCCNumber
    }

    $Submit_Button.Add_Click( { 
            if ($ComputerName_Campus_Dropdown.Items -contains $ComputerName_Campus_Dropdown.Text) {
                $ErrorProvider.SetError($ComputerName_Label, '')
                if ($ComputerName_BuildingRoom_Textbox.Text -match '^[a-z]{1}\d{3}$|^[a-z]{2}\d{2}$|^[a-z]{2}\d{3}$|^[a-z]{3}$') {
                    $ErrorProvider.SetError($ComputerName_Label, '')
                    if ($ComputerName_PCCNumber_Textbox.Text -match '^\d{6}$') {
                        $ErrorProvider.SetError($ComputerName_Label, '')
                        if ($ComputerName_Suffix_Textbox.Text -match '^[a-z]{2}$') {
                            $ErrorProvider.SetError($ComputerName_Label, '')
                            if ($ComputerName_Campus_Dropdown.Text -ne 'DC') {
                                $ComputerName = $ComputerName_Campus_Dropdown.Text + '-' + $ComputerName_BuildingRoom_Textbox.Text + $ComputerName_PCCNumber_Textbox.Text + $ComputerName_Suffix_Textbox.Text
                                if ($ComputerName.Length -gt 15) {
                                    $ErrorProvider.SetError($ComputerName_Label, 'Name too long')
                                }
                            }
                            else {
                                $ComputerName = $ComputerName_Campus_Dropdown.Text + $ComputerName_BuildingRoom_Textbox.Text + $ComputerName_PCCNumber_Textbox.Text + $ComputerName_Suffix_Textbox.Text
                                if ($ComputerName.Length -gt 15) {
                                    $ErrorProvider.SetError($ComputerName_Label, 'Name too long')
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
        
            # Verify Domain
            $ErrorProvider.SetError($DomainSelection_Label, '')
            if ($EDU_RadioButton.Checked -eq $true) {
                $Domain = 'EDU-Domain.pima.edu'
                $OULocation = "LDAP://OU=$($CampusList[$ComputerName_Campus_Dropdown.SelectedIndex][1]),OU=Staging,DC=$Domain,DC=pima,DC=edu"
            }
            if ($PCC_RadioButton.Checked -eq $true) {
                $Domain = 'PCC-Domain.pima.edu'
                $OULocation = "LDAP://OU=$($CampusList[$ComputerName_Campus_Dropdown.SelectedIndex][1]),OU=Staging,DC=$Domain,DC=pima,DC=edu"
            }
            elseif (($EDU_RadioButton.Checked -or $PCC_RadioButton.Checked) -eq $false) {
                $ErrorProvider.SetError($DomainSelection_Label, 'Select a domain')
            } 
        
            if (Confirm-NoError) {
                # Output for testing
                $ComputerName.ToUpper(), $OULocation, 'Passed' | Out-File '.\OSDTest.csv' -Append

                #Enable when deployed
                #$TSEnvironment = New-Object -COMObject Microsoft.SMS.TSEnvironment 
                #$TSEnvironment.Value("OSDComputerName") = "$($ComputerName.ToUpper())"
                #$TSEnvironment.Value("OSDDomainOUName") = "$($OULocation)"
                #$TSEnvironment.Value("OSDDomainName") = "$($Domain)"               
                [void]$Form.Close()
            }
        })

    #endregion

    # Test Values
    $ComputerName_Campus_Dropdown.SelectedItem = $Campus
    $ComputerName_BuildingRoom_Textbox.Text = $Bldg
    $ComputerName_PCCNumber_Textbox.Text = $PCC
    $ComputerName_Suffix_Textbox.Text = $Suffix
    switch -regex ($Domain) {
        'PCC' { $PCC_RadioButton.Checked = $true }
        'EDU' { $EDU_RadioButton.Checked = $true }
        Default {}
    }
    [void]$Form.ShowDialog()
    $Submit_Button.PerformClick()
}

$CampusShortList = @('29', 'ER', 'EP', 'DV', 'DO', 'DC', 'EC', 'MS', 'NW', 'WC', 'PCC')
$RandomRooms = @('CG11', 'E513', 'AH321', 'emp', 'stu')

for ($i = 0; $i -lt 1; $i++) {
    OSD-GUI -Campus $(get-random -InputObject $CampusShortList) -Bldg $(get-random -InputObject $RandomRooms) -PCC $(get-random -Maximum 999999) -Suffix $( -join ((65..90) + (97..121) | Get-Random -Count 2 | % { [char]$_ })) -Domain $(get-random -InputObject 'EDU', 'PCC')
}