$LocationList = @("Adult Education","Community","Desert Vista","District","Downtown","East","Maintenance and Security","Northwest","West")
#region GUI

# Assemblies
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

# Form
$Form                   = New-Object System.Windows.Forms.Form    
$Form.ClientSize        = '300,300'
$Form.FormBorderStyle   = "FixedDialog"
$Form.StartPosition     = "CenterScreen"
$Form.Icon              = [System.Drawing.Icon]::ExtractAssociatedIcon($PSHome + "\powershell.exe")
$Form.Text              = "Active Directory Location"
$Form.ControlBox        = $false
$Form.TopMost           = $true

$ComputerName_Textbox           = New-Object System.Windows.Forms.TextBox
$ComputerName_Textbox.Size      = New-Object System.Drawing.Size(240,50)
$ComputerName_Textbox.TabIndex = '1'

$ComputerName_Group                 = New-Object System.Windows.Forms.GroupBox
$ComputerName_Group.Size            = New-Object System.Drawing.Size($($ComputerName_Textbox.Width+10),$($ComputerName_Textbox.Height*2.2))
$ComputerName_Group.Location        = New-Object System.Drawing.Size($(($Form.Width-$ComputerName_Group.Width)/2),$($Form.Height / 30))
$ComputerName_Group.Text            = 'Computer name:'

$ComputerName_Textbox.Location  = New-Object System.Drawing.Size(5,$(($ComputerName_Group.Height-$ComputerName_Textbox.Height)/2))

$Domain_Group               = New-Object System.Windows.Forms.GroupBox
$Domain_Group.Size          = New-Object System.Drawing.Size(260,40)
$Domain_Group.Location      = New-Object System.Drawing.Size($(($Form.Width-$Domain_Group.Width)/2),$($Form.Height/4.5))
$Domain_Group.Text          = 'Select Domain'

$EDU_RadioButton              = New-Object System.Windows.Forms.RadioButton
$EDU_RadioButton.Location     = New-Object System.Drawing.Size(10,15)
$EDU_RadioButton.Size         = New-Object System.Drawing.Size(100,20)
$EDU_RadioButton.Text         = 'EDU-Domain'
$EDU_RadioButton.TabIndex = '2'

$PCC_RadioButton              = New-Object System.Windows.Forms.RadioButton
$PCC_RadioButton.Location     = New-Object System.Drawing.Size(110,15)
$PCC_RadioButton.Size         = New-Object System.Drawing.Size(100,20)
$PCC_RadioButton.Text         = 'PCC-Domain'
$PCC_RadioButton.TabIndex = '3'

$Location_Group             = New-Object System.Windows.Forms.GroupBox
$Location_Group.Size        = New-Object System.Drawing.Size(260,50)
$Location_Group.Location    = New-Object System.Drawing.Size($(($Form.Width-$Location_Group.Width)/2),$(($Form.Height-$Location_Group.Height)/2))
$Location_Group.Text        = "Select Campus"

$Location_Dropdown               = New-Object System.Windows.Forms.ComboBox
$Location_Dropdown.Location      = New-Object System.Drawing.Size(10,15)
$Location_Dropdown.Size          = New-Object System.Drawing.Size(240,30)
$Location_Dropdown.DropDownStyle = "DropDown"
$Location_Dropdown.Items.AddRange($LocationList)
$Location_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Location_Dropdown.AutoCompleteSource = 'ListItems'
$Location_Dropdown.TabIndex = '4'

$Submit_Button                 = New-Object System.Windows.Forms.Button
$Submit_Button.Size            = New-Object System.Drawing.Size(80,25)
$Submit_Button.Location        = New-Object System.Drawing.Size($(($Form.Width-$Submit_Button.Width)/2),215)
$Submit_Button.Text            = "OK"
$Submit_Button.Enabled = $false
$Submit_Button.TabIndex = '5'
$Submit_Button.TabStop = $true

$Form.Controls.AddRange(@($ComputerName_Group,$Domain_Group,$Location_Group,$Submit_Button))
$ComputerName_Group.controls.AddRange(@($ComputerName_Textbox))
$Domain_Group.controls.AddRange(@($EDU_RadioButton, $PCC_RadioButton))
$Location_Group.controls.AddRange(@($Location_Dropdown))
#endregion

function Set-OULocation {
    param(
    [parameter(Mandatory=$true)]
    $Location
    )
    if ($EDU_RadioButton.Checked -eq $true) {
        $Global:OULocation = "LDAP://OU=$($Location),OU=Staging,DC=$($EDU_RadioButton.Text),DC=pima,DC=edu"
        $Domain = "EDU-Domain.pima.edu"
    }
    if ($PCC_RadioButton.Checked -eq $true) {
        $Global:OULocation = "LDAP://OU=$($Location),OU=Staging,DC=$($PCC_RadioButton.Text),DC=pima,DC=edu"
        $Domain = "PCC-Domain.pima.edu"
    }
    elseif (($EDU_RadioButton.Checked -or $PCC_RadioButton.Checked) -eq $false) {
        $ErrorProvider.SetError($Domain_Group, 'Select a domain')
    }
    #$TSEnvironment = New-Object -COMObject Microsoft.SMS.TSEnvironment 
    #$TSEnvironment.Value("OSDDomainOUName") = "$($OULocation)"
    #$TSEnvironment.Value("OSDDomainName") = "$($Domain)"
}

Function Set-OSDComputerName 
{
    $ErrorProvider.Clear()
    #Validation Rule for computer names.
    if ($ComputerName_Textbox.Text -notmatch "([a-z]{4}|(([a-z]{2}|\d{2})-[a-z]{1,2}))\d{8,9}([a-z]{2}|v\d{1})"){
        $ErrorProvider.SetError($ComputerName_Group, "Computer name invalid, please correct the computer name.")
    }

    else {
        $Global:OSDComputerName = $ComputerName_Textbox.Text.ToUpper()
        #$TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment
        #$TSEnv.Value("OSDComputerName") = "$($OSDComputerName)"
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
        if ($ErrorProvider.GetError($group).length -gt 0) {
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

#region Actions
$Location_Dropdown.Add_SelectedValueChanged({$Submit_Button.Enabled = $true})
$Submit_Button.Add_Click({Set-OSDComputerName})
$Submit_Button.Add_Click({Set-OULocation -Location $Location_Dropdown.SelectedItem.ToString()}) 
$Submit_Button.Add_Click( {
    if (Confirm-NoError) {
        [void][System.Windows.Forms.MessageBox]::Show("Computer Name: $($OSDComputerName) - OU: $($OULocation)","Test Submission")
        $Form.Close()
    }
})
#endregion

[void]$Form.ShowDialog()