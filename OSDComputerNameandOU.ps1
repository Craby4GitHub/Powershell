#region GUI

#[void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
#[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') 

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

$screen = [System.Windows.Forms.Screen]::AllScreens

$Form = New-Object System.Windows.Forms.Form    
$Form.FormBorderStyle = 'FixedDialog'
$Form.StartPosition = 'CenterScreen'
$Form.Width = $($screen[0].bounds.Width / 5)
$Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSHome + '\powershell.exe')
$Form.Text = 'Active Directory Information'
$Form.ControlBox = $false
$Form.TopMost = $true
$Form.Font = 'Microsoft Sans Serif,8'

$CampusList = @(('29','Adult Education'), ('ER','Adult Education'), ('EP','Adult Education'), ('DV','Desert Vista'), ('DO', 'District'), ('DC', 'Downtown'), ('EC','East'), ('MS','Maintenance and Security'), ('NW','Northwest'), ('WC','West'), ('PCC','West'))

$ComputerName_Campus_Label = New-Object system.Windows.Forms.Label
$ComputerName_Campus_Label.Text = 'Campus'
$ComputerName_Campus_Label.Size = New-Object System.Drawing.Size($(([System.Windows.Forms.TextRenderer]::MeasureText($ComputerName_Campus_Label.Text, $ComputerName_Campus_Label.Font)).Width), 20)
$ComputerName_Campus_Dropdown = New-Object System.Windows.Forms.ComboBox
$ComputerName_Campus_Dropdown.DropDownStyle = 'DropDown'
$ComputerName_Campus_Dropdown.Items.AddRange(($CampusList | ForEach-Object {$($_[0])}))
$ComputerName_Campus_Dropdown.AutoCompleteMode = 'SuggestAppend'
$ComputerName_Campus_Dropdown.AutoCompleteSource = 'ListItems'
$ComputerName_Campus_Dropdown.TabIndex = 1

$ComputerName_BuildingRoom_Label = New-Object system.Windows.Forms.Label
$ComputerName_BuildingRoom_Label.Text = 'Bldg/Room'
$ComputerName_BuildingRoom_Label.Size = New-Object System.Drawing.Size($(([System.Windows.Forms.TextRenderer]::MeasureText($ComputerName_BuildingRoom_Label.Text, $ComputerName_BuildingRoom_Label.Font)).Width), 20)
$ComputerName_BuildingRoom_Textbox = New-Object System.Windows.Forms.TextBox
$ComputerName_BuildingRoom_Textbox.TabIndex = 2

$ComputerName_PCCNumber_Label = New-Object system.Windows.Forms.Label
$ComputerName_PCCNumber_Label.Text = 'PCC#'
$ComputerName_PCCNumber_Label.Size = New-Object System.Drawing.Size($(([System.Windows.Forms.TextRenderer]::MeasureText($ComputerName_PCCNumber_Label.Text, $ComputerName_PCCNumber_Label.Font)).Width), 20)
$ComputerName_PCCNumber_Textbox = New-Object System.Windows.Forms.TextBox
$ComputerName_PCCNumber_Textbox.TabIndex = 3

$ComputerName_Suffix_Label = New-Object system.Windows.Forms.Label
$ComputerName_Suffix_Label.Text = 'Suffix'
$ComputerName_Suffix_Label.Size = New-Object System.Drawing.Size($(([System.Windows.Forms.TextRenderer]::MeasureText($ComputerName_Suffix_Label.Text, $ComputerName_Suffix_Label.Font)).Width), 20)
$ComputerName_Suffix_Textbox = New-Object System.Windows.Forms.TextBox
$ComputerName_Suffix_Textbox.TabIndex = 4

$ComputerName_Group = New-Object System.Windows.Forms.GroupBox
$ComputerName_Group.Text = 'Computer name:'
$ComputerName_Group.Size = New-Object System.Drawing.Size($($Form.Width * .8), 50)
$ComputerName_Group.Location = New-Object System.Drawing.Point($(($Form.Width - $ComputerName_Group.Width) * .25), 10)

$ComputerName_Campus_Dropdown.Size = New-Object System.Drawing.Size($($ComputerName_Group.Width * .2), 50)
$ComputerName_BuildingRoom_Textbox.Size = New-Object System.Drawing.Size($($ComputerName_Group.Width * .2), 50)
$ComputerName_PCCNumber_Textbox.Size = New-Object System.Drawing.Size($($ComputerName_Group.Width * .3), 50)
$ComputerName_Suffix_Textbox.Size = New-Object System.Drawing.Size($($ComputerName_Group.Width * .15), 50)

$ComputerName_Group.Height = ($ComputerName_Campus_Dropdown.Height + $ComputerName_Campus_Label.Height) * 1.8

$ComputerName_Campus_Dropdown.Location = New-Object System.Drawing.Point($(($ComputerName_Group.Width - $ComputerName_Campus_Dropdown.Width) * .05), $(($ComputerName_Group.Height - $ComputerName_Campus_Dropdown.Height) * .75))
$ComputerName_BuildingRoom_Textbox.Location = New-Object System.Drawing.Point($(($ComputerName_Campus_Dropdown.Location.X + $ComputerName_Campus_Dropdown.Width) + 10), $(($ComputerName_Group.Height - $ComputerName_BuildingRoom_Textbox.Height) * .75))
$ComputerName_PCCNumber_Textbox.Location = New-Object System.Drawing.Point($(($ComputerName_BuildingRoom_Textbox.Location.X + $ComputerName_BuildingRoom_Textbox.Width) + 10), $(($ComputerName_Group.Height - $ComputerName_PCCNumber_Textbox.Height) * .75))
$ComputerName_Suffix_Textbox.Location = New-Object System.Drawing.Point($(($ComputerName_PCCNumber_Textbox.Location.X + $ComputerName_PCCNumber_Textbox.Width) + 10), $(($ComputerName_Group.Height - $ComputerName_Suffix_Textbox.Height) * .75))

$ComputerName_Campus_Label.Location = New-Object System.Drawing.Point($($ComputerName_Campus_Dropdown.Location.X + ($ComputerName_Campus_Dropdown.Width - $ComputerName_Campus_Label.Width) * .25), $($ComputerName_Campus_Dropdown.Location.Y - 20))
$ComputerName_BuildingRoom_Label.Location = New-Object System.Drawing.Point($($ComputerName_BuildingRoom_Textbox.Location.X + ($ComputerName_BuildingRoom_Textbox.Width - $ComputerName_BuildingRoom_Label.Width) * .25), $($ComputerName_BuildingRoom_Textbox.Location.Y - 20))
$ComputerName_PCCNumber_Label.Location = New-Object System.Drawing.Point($($ComputerName_PCCNumber_Textbox.Location.X + ($ComputerName_PCCNumber_Textbox.Width - $ComputerName_PCCNumber_Label.Width) * .25), $($ComputerName_PCCNumber_Textbox.Location.Y - 20))
$ComputerName_Suffix_Label.Location = New-Object System.Drawing.Point($($ComputerName_Suffix_Textbox.Location.X + ($ComputerName_Suffix_Textbox.Width - $ComputerName_Suffix_Label.Width) * .25), $($ComputerName_Suffix_Textbox.Location.Y - 20))

$EDU_RadioButton = New-Object System.Windows.Forms.RadioButton
$EDU_RadioButton.Text = 'EDU'
$EDU_RadioButton.Size = New-Object System.Drawing.Size($(([System.Windows.Forms.TextRenderer]::MeasureText($EDU_RadioButton.Text, $EDU_RadioButton.Font)).Width + 30), 30)
$EDU_RadioButton.TabStop = $true

$PCC_RadioButton = New-Object System.Windows.Forms.RadioButton
$PCC_RadioButton.Text = 'PCC'
$PCC_RadioButton.Size = New-Object System.Drawing.Size($(([System.Windows.Forms.TextRenderer]::MeasureText($PCC_RadioButton.Text, $PCC_RadioButton.Font)).Width + 30), 30)
$PCC_RadioButton.TabStop = $true

$Domain_Group = New-Object System.Windows.Forms.GroupBox
$Domain_Group.Text = 'Select Domain'
$Domain_Group.Size = New-Object System.Drawing.Size($($Form.Width * .8), $($EDU_RadioButton.Height + $PCC_RadioButton.Height))
$Domain_Group.Location = New-Object System.Drawing.Point($(($Form.Width - $Domain_Group.Width) * .25), $($($ComputerName_Group.Location.Y + $ComputerName_Group.Size.Height + 5)))
$Domain_Group.TabIndex = 5

$EDU_RadioButton.Location = New-Object System.Drawing.Point($(($Domain_Group.Width - ($EDU_RadioButton.Width + $PCC_RadioButton.Width)) * .5), $(($Domain_Group.Height - $EDU_RadioButton.Height) * .75))
$PCC_RadioButton.Location = New-Object System.Drawing.Point($($EDU_RadioButton.Location.X + $EDU_RadioButton.Width), $(($Domain_Group.Height - $PCC_RadioButton.Height) * .75))

$Submit_Button = New-Object System.Windows.Forms.Button
$Submit_Button.Text = 'Submit'
$Submit_Button.Size = New-Object System.Drawing.Size(80, 25)
$Submit_Button.Location = New-Object System.Drawing.Point($((($Domain_Group.Width - $Submit_Button.Width) * .5) + $Domain_Group.Location.X), $($Domain_Group.Location.Y + $Domain_Group.Size.Height + 5))
$Submit_Button.TabIndex = 6
$Form.AcceptButton = $Submit_Button

$Form.Height = ($Submit_Button.Location.Y + $Submit_Button.Height) * 1.1
$Form.Controls.AddRange(@($ComputerName_Group, $Domain_Group, $Submit_Button))
$ComputerName_Group.Controls.AddRange(@($ComputerName_Campus_Dropdown, $ComputerName_Campus_Label, $ComputerName_BuildingRoom_Textbox, $ComputerName_BuildingRoom_Label, $ComputerName_PCCNumber_Textbox, $ComputerName_PCCNumber_Label, $ComputerName_Suffix_Textbox, $ComputerName_Suffix_Label))
$Domain_Group.Controls.AddRange(@($EDU_RadioButton, $PCC_RadioButton))

#endregion

#region Functions

function Find-Group() {
    $FormGroups = @()
    ForEach ($control in $Form.Controls) {
        if ($control.ToString().StartsWith('System.Windows.Forms.GroupBox')) {
            $FormGroups += $control
        }
    }
    return $FormGroups
}

function Confirm-NoError {
    $i = 0
    foreach ($group in Find-Group) {
        if ($ErrorProvider.GetError($group)) {
            $i++
        }
    }
    if ($i -gt 0) {
        return $false
    }
    else {
        return $true
    }
}
#endregion

#region Actions

# Test Values
$ComputerName_BuildingRoom_Textbox.Text = 'R016'
$ComputerName_PCCNumber_Textbox.Text = '123456'
$ComputerName_Suffix_Textbox.Text = 'LL'

$PCCNumber = (Get-WmiObject -Query "Select * from Win32_SystemEnclosure").SMBiosAssetTag
if ($PCCNumber -match '\d{6}') {
    $ComputerName_PCCNumber_Textbox.Text = $PCCNumber
}

$Submit_Button.Add_Click( { 

        if ($ComputerName_Campus_Dropdown.Items -contains $ComputerName_Campus_Dropdown.Text) {
            $ErrorProvider.SetError($ComputerName_Group, '')
            if ($ComputerName_BuildingRoom_Textbox.Text -match '^[a-z]{1}\d{3}$|^[a-z]{2}\d{2}$|^[a-z]{2}\d{3}$|^[a-z]{3}$') {
                if ($ComputerName_PCCNumber_Textbox.Text -match '^\d{6}$') {
                    if ($ComputerName_Suffix_Textbox.Text -match '^[a-z]{2}$') {
                        if ($ComputerName_Campus_Dropdown.Text -ne 'DC') {
                            $ComputerName = $ComputerName_Campus_Dropdown.Text + '-' + $ComputerName_BuildingRoom_Textbox.Text + $ComputerName_PCCNumber_Textbox.Text + $ComputerName_Suffix_Textbox.Text
                            if ($ComputerName.Length -ne 15) {
                                $ErrorProvider.SetError($ComputerName_Group, 'Name too long')
                            }
                        }
                        else {
                            $ComputerName = $ComputerName_Campus_Dropdown.Text + $ComputerName_BuildingRoom_Textbox.Text + $ComputerName_PCCNumber_Textbox.Text + $ComputerName_Suffix_Textbox.Text
                            if ($ComputerName.Length -ne 15) {
                                $ErrorProvider.SetError($ComputerName_Group, 'Name too long')
                            }
                        }
                    }
                    else {
                        $ErrorProvider.SetError($ComputerName_Group, 'Enter a proper suffix')
                    }
                }
                else {
                    $ErrorProvider.SetError($ComputerName_Group, 'Enter a proper PCC Number')
                }
            }
            else {
                $ErrorProvider.SetError($ComputerName_Group, 'Enter a proper building/room')
            }
        }
        else {
            $ErrorProvider.SetError($ComputerName_Group, 'Select a proper campus')
        }            
        
        # Verify Location
        if ($Location_Dropdown.Items -contains $Location_Dropdown.Text) {
            $ErrorProvider.SetError($Location_Group, '')
        }
        else {
            $ErrorProvider.SetError($Location_Group, 'Select a location')
        }

        # Verify Domain
        $ErrorProvider.SetError($Domain_Group, '')
        if ($EDU_RadioButton.Checked -eq $true) {
            $Domain = 'EDU-Domain.pima.edu'
            $OULocation = "LDAP://OU=$($CampusList[$ComputerName_Campus_Dropdown.SelectedIndex][1]),OU=Staging,DC=$Domain,DC=pima,DC=edu"
        }
        if ($PCC_RadioButton.Checked -eq $true) {
            $Domain = 'PCC-Domain.pima.edu'
            $OULocation = "LDAP://OU=$($CampusList[$ComputerName_Campus_Dropdown.SelectedIndex][1]),OU=Staging,DC=$Domain,DC=pima,DC=edu"
        }
        elseif (($EDU_RadioButton.Checked -or $PCC_RadioButton.Checked) -eq $false) {
            $ErrorProvider.SetError($Domain_Group, 'Select a domain')
        } 
        
        if (Confirm-NoError) {
            # Temp Messagebox, for testing
            [void][System.Windows.Forms.MessageBox]::Show("Computer Name: $($ComputerName) `nOU: $($OULocation)", "Test Submission")
            #$TSEnvironment = New-Object -COMObject Microsoft.SMS.TSEnvironment 
            #$TSEnvironment.Value("OSDComputerName") = "$($ComputerName)"
            #$TSEnvironment.Value("OSDDomainOUName") = "$($OULocation)"
            #$TSEnvironment.Value("OSDDomainName") = "$($Domain)"
            $Form.Close()
        }
    })

#endregion

[void]$Form.ShowDialog()