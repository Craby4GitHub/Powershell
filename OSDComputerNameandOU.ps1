#region GUI

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

$screen = [System.Windows.Forms.Screen]::AllScreens


$Form = New-Object System.Windows.Forms.Form    
$Form.FormBorderStyle = "FixedDialog"
$Form.StartPosition = "CenterScreen"
$Form.Width = $($screen[0].bounds.Width / 7)
#$Form.Height = $($screen[0].bounds.Height / 4.5)
$Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSHome + "\powershell.exe")
$Form.Text = "Active Directory Information"
$Form.ControlBox = $false
$Form.TopMost = $true
$Form.Font = 'Microsoft Sans Serif,8'

$ComputerName_Textbox = New-Object System.Windows.Forms.TextBox
$ComputerName_Textbox.TabIndex = 1
#Temp name for testing
$ComputerName_Textbox.text = 'wc-r015123123sc'

$ComputerName_Group = New-Object System.Windows.Forms.GroupBox
$ComputerName_Group.Size = New-Object System.Drawing.Size($($Form.Width * .9), 50)
$ComputerName_Group.Location = New-Object System.Drawing.Size(5, 10)
$ComputerName_Group.Text = 'Computer name:'

$ComputerName_Textbox.Size = New-Object System.Drawing.Size($($ComputerName_Group.Width * .9), 50)
$ComputerName_Group.Height = $ComputerName_Textbox.Height * 2.2
$ComputerName_Textbox.Location = New-Object System.Drawing.Size($(($ComputerName_Group.Width - $ComputerName_Textbox.Width) * .5), $(($ComputerName_Group.Height - $ComputerName_Textbox.Height) * .75))

$EDU_RadioButton = New-Object System.Windows.Forms.RadioButton
$EDU_RadioButton.Size = New-Object System.Drawing.Size(100, 30)
$EDU_RadioButton.TabStop = $true
$EDU_RadioButton.Text = 'EDU'

$PCC_RadioButton = New-Object System.Windows.Forms.RadioButton
$PCC_RadioButton.Size = New-Object System.Drawing.Size(100, 30)
$PCC_RadioButton.TabStop = $true
$PCC_RadioButton.Text = 'PCC'

$Domain_Group = New-Object System.Windows.Forms.GroupBox
$Domain_Group.Size = New-Object System.Drawing.Size($($Form.Width * .9), $($EDU_RadioButton.Height + $PCC_RadioButton.Height))
$Domain_Group.Location = New-Object System.Drawing.Size(5, $($($ComputerName_Group.Location.Y + $ComputerName_Group.Size.Height + 5)))
$Domain_Group.TabIndex = 2
$Domain_Group.Text = 'Select Domain'

$EDU_RadioButton.Location = New-Object System.Drawing.Size(5, $(($Domain_Group.Height - $EDU_RadioButton.Height) * .75))
$PCC_RadioButton.Location = New-Object System.Drawing.Size($(($Domain_Group.Width - $PCC_RadioButton.Width) - 5), $(($Domain_Group.Height - $PCC_RadioButton.Height) * .75))

$Location_Dropdown = New-Object System.Windows.Forms.ComboBox
$Location_Dropdown.DropDownStyle = "DropDown"
$Location_Dropdown.Items.AddRange(@("Adult Education", "Desert Vista", "District", "Downtown", "East", "Maintenance and Security", "Northwest", "West"))
$Location_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Location_Dropdown.AutoCompleteSource = 'ListItems'
$Location_Dropdown.TabIndex = 3

$Location_Group = New-Object System.Windows.Forms.GroupBox
$Location_Group.Size = New-Object System.Drawing.Size($($Form.Width * .9), 50)
$Location_Group.Location = New-Object System.Drawing.Size(5, $($($Domain_Group.Location.Y + $Domain_Group.Size.Height + 5)))
$Location_Group.Text = "Select Campus"

$Location_Dropdown.Size = New-Object System.Drawing.Size($($Location_Group.Width * .9), 50)
$Location_Group.Height = $Location_Dropdown.Height * 2.2
$Location_Dropdown.Location = New-Object System.Drawing.Size($(($Location_Group.Width - $Location_Dropdown.Width) * .5), $(($Location_Group.Height - $Location_Dropdown.Height) * .75))

$Submit_Button = New-Object System.Windows.Forms.Button
$Submit_Button.Size = New-Object System.Drawing.Size(80, 25)
$Submit_Button.Location = New-Object System.Drawing.Size($(($Location_Group.Width - $Submit_Button.Width) * .5), $($($Location_Group.Location.Y + $Location_Group.Size.Height + 5)))
$Submit_Button.Text = "OK"
$Submit_Button.TabIndex = 4
$Form.AcceptButton = $Submit_Button

$Form.Height = $($ComputerName_Group.Size.Height + $Domain_Group.Size.Height + $Location_Group.Size.Height + $Submit_Button.Size.Height) * 1.5

$Form.Controls.AddRange(@($ComputerName_Group, $Domain_Group, $Location_Group, $Submit_Button))
$ComputerName_Group.Controls.Add($ComputerName_Textbox)
$Domain_Group.Controls.AddRange(@($EDU_RadioButton, $PCC_RadioButton))
$Location_Group.Controls.Add($Location_Dropdown)
#endregion

#region Functions
function Set-OULocation($Location) {
    $ErrorProvider.SetError($Domain_Group,'')
    if ($EDU_RadioButton.Checked -eq $true) {
        $Domain = "EDU-Domain.pima.edu"
        # Temp Global, for testing
        $Global:OULocation = "LDAP://OU=$($Location),OU=Staging,DC=$($Domain),DC=pima,DC=edu"
    }
    if ($PCC_RadioButton.Checked -eq $true) {
        $Domain = "PCC-Domain.pima.edu"
        # Temp Global, for testing
        $Global:OULocation = "LDAP://OU=$($Location),OU=Staging,DC=$($Domain),DC=pima,DC=edu"
    }
    elseif (($EDU_RadioButton.Checked -or $PCC_RadioButton.Checked) -eq $false) {
        $ErrorProvider.SetError($Domain_Group, 'Select a domain')
    }
    #$TSEnvironment = New-Object -COMObject Microsoft.SMS.TSEnvironment 
    #$TSEnvironment.Value("OSDDomainOUName") = "$($OULocation)"
    #$TSEnvironment.Value("OSDDomainName") = "$($Domain)"
}

Function Set-OSDComputerName {
    $ErrorProvider.SetError($ComputerName_Group,'')

    Switch -regex ($ComputerName_Textbox.Text) {
        #Normal Campus
        '([a-z]{4}|(([a-z]{2}|\d{2})-[a-z]{1,2}))\d{8,9}([a-z]{2}|v\d{1})' {
            $Global:OSDComputerName = $ComputerName_Textbox.Text.ToUpper()
            #$TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment
            #$TSEnv.Value("OSDComputerName") = "$($OSDComputerName)"
            break
        }
        #CARES Act
        '[a-z]{3}-[a-z]{3}\d{6}[a-z]{2}' {
            $Global:OSDComputerName = $ComputerName_Textbox.Text.ToUpper()
            #$TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment
            #$TSEnv.Value("OSDComputerName") = "$($OSDComputerName)"
            break
        }
        default {
            $ErrorProvider.SetError($ComputerName_Group, "Computer name invalid, please correct the computer name.")
        }
    #Validation Rule for computer names.
    }
}

function Find-Group() {
    $FormGroups = @()
    ForEach ($control in $Form.Controls) {
        if ($control.ToString().StartsWith("System.Windows.Forms.GroupBox")) {
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
$Submit_Button.Add_Click( { 
        if ($Location_Dropdown.Items -contains $Location_Dropdown.Text) {
            $ErrorProvider.SetError($Location_Group,'')
            Set-OULocation -Location $Location_Dropdown.Text
        }
        else {
            $ErrorProvider.SetError($Location_Group, 'Select a location')
        } 
    })
$Submit_Button.Add_Click( { Set-OSDComputerName }) 
$Submit_Button.Add_Click( {
        if (Confirm-NoError) {
            # Temp Messagebox, for testing
            [void][System.Windows.Forms.MessageBox]::Show("Computer Name: $($OSDComputerName) `nOU: $($OULocation)", "Test Submission")
            $Form.Close()
        }
    })
    
#endregion

[void]$Form.ShowDialog()