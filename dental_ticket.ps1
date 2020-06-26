<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Dental Ticket
#>

# https://stackoverflow.com/a/40621143/20267
# hide the powerhsell console

Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    '
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0) # hide

#region GUI
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

$Form = New-Object system.Windows.Forms.Form
$Form.FormBorderStyle = "FixedDialog"
$Form.ClientSize = "400,400"
$Form.text = "Equipment Repair Form"
$Form.TopMost = $true
$Form.StartPosition = 'CenterScreen'
$Form.Font = 'Microsoft Sans Serif,10'

$ID_Num_Text = New-Object system.Windows.Forms.TextBox
$ID_Num_Text.multiline = $false
$ID_Num_Text.width = 100
$ID_Num_Text.height = 20

$ID_Num_Group = New-Object system.Windows.Forms.Groupbox
$ID_Num_Group.height = 50
$ID_Num_Group.width = 120
$ID_Num_Group.text = "ID Number"
$ID_Num_Group.location = New-Object System.Drawing.Point(10, 10)

$ID_Num_Text.location = New-Object System.Drawing.Point(($($ID_Num_Group.Width - $ID_Num_Text.Width) / 2), ($($ID_Num_Group.Height - $ID_Num_Text.Height) / 1.5))

$Location_Dropdown              = New-Object system.Windows.Forms.ComboBox
$Location_Dropdown.width        = 100
$Location_Dropdown.height       = 20
$Location_Dropdown.Text = (Get-WmiObject -Class Win32_OperatingSystem).Description
$Location_Dropdown.DropDownStyle = "DropDown"
$Location_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Location_Dropdown.AutoCompleteSource = 'ListItems'

$Location_Group = New-Object system.Windows.Forms.Groupbox
$Location_Group.height = 50
$Location_Group.width = 120
$Location_Group.text = "Location"
$Location_Group.location = New-Object System.Drawing.Point(140, 10)

$Location_Dropdown.location = New-Object System.Drawing.Point(($($Location_Group.Width - $Location_Dropdown.Width) / 2), ($($Location_Group.Height - $Location_Dropdown.Height) / 1.5))

$Equipment_Dropdown              = New-Object system.Windows.Forms.ComboBox
$Equipment_Dropdown.width        = 100
$Equipment_Dropdown.height       = 20
$Equipment_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Equipment_Dropdown.AutoCompleteSource = 'ListItems'

$Equipment_Group = New-Object system.Windows.Forms.Groupbox
$Equipment_Group.height = 50
$Equipment_Group.width = 120
$Equipment_Group.text = "Equipment"
$Equipment_Group.location = New-Object System.Drawing.Point(270, 10)

$Equipment_Dropdown.location = New-Object System.Drawing.Point(($($Equipment_Group.Width - $Equipment_Dropdown.Width) / 2), ($($Equipment_Group.Height - $Equipment_Dropdown.Height) / 1.5))

$Desc_Text = New-Object system.Windows.Forms.TextBox
$Desc_Text.multiline = $true
$Desc_Text.height = 70

$Desc_Group = New-Object system.Windows.Forms.Groupbox
$Desc_Group.height = 100
$Desc_Group.width = $($Form.Width - 40)
$Desc_Group.text = "Description of Issue"
$Desc_Group.location = New-Object System.Drawing.Point(10, 76)

$Desc_Text.width = $($Desc_Group.Width - 20)
$Desc_Text.location = New-Object System.Drawing.Point(($($Desc_Group.Width - $Desc_Text.Width) / 2), ($($Desc_Group.Height - $Desc_Text.Height) / 1.5))

$Issue_History = New-Object system.Windows.Forms.DataGridView
$Issue_History.ScrollBars = "Vertical"
$Issue_History.AutoGenerateColumns = $true
$Issue_History.AutoSizeColumnsMode = 'AllCells'
$Issue_History.RowHeadersVisible = $false
$Issue_History.ColumnCount = 3
$Issue_History.Columns[0].Name = 'Equipment'
$Issue_History.Columns[1].Name = 'Issue'
# Isnt formating...
$Issue_History.Columns[2].DefaultCellStyle.Format = "m"
$Issue_History.Columns[2].Name = 'Submitted'
$Issue_History.TabStop = $false

$Issue_History_Group = New-Object system.Windows.Forms.Groupbox
$Issue_History_Group.height = 150
$Issue_History_Group.width = $($Form.Width - 40)
$Issue_History_Group.text = "Current Issues"
$Issue_History_Group.location = New-Object System.Drawing.Point(10, 180)

$Issue_History.height = $($Issue_History_Group.Height - 20)
$Issue_History.width = $($Issue_History_Group.Width - 20)
$Issue_History.location = New-Object System.Drawing.Point(($($Issue_History_Group.Width - $Issue_History.Width) / 2), ($($Issue_History_Group.Height - $Issue_History.Height) / 1.1))

$Submit_Button = New-Object system.Windows.Forms.Button
$Submit_Button.text = "Submit"
$Submit_Button.width = 60
$Submit_Button.height = 30
$Submit_Button.location = New-Object System.Drawing.Point(30, 350)

$Sumbit_Status = New-Object system.Windows.Forms.Label
$Sumbit_Status.AutoSize = $true
$Sumbit_Status.width = 250
$Sumbit_Status.height = 10
$Sumbit_Status.location = New-Object System.Drawing.Point(100, 350)
$Form.AcceptButton = $Submit_Button

$Clear_Button = New-Object system.Windows.Forms.Button
$Clear_Button.text = "Clear"
$Clear_Button.width = 60
$Clear_Button.height = 30
$Clear_Button.location = New-Object System.Drawing.Point(300, 350)

$Form.controls.AddRange(@($ID_Num_Group, $Location_Group, $Equipment_Group, $Desc_Group, $Issue_History_Group, $Submit_Button, $Sumbit_Status, $Clear_Button))
$ID_Num_Group.controls.Add($ID_Num_Text)
$Location_Group.controls.Add($Location_Dropdown)
$Equipment_Group.controls.Add($Equipment_Dropdown)
$Desc_Group.Controls.Add($Desc_Text)
$Issue_History_Group.Controls.Add($Issue_History)
#endregion

#region Functions
function Get-File($filePath) {   
    try {
        $file = Import-Csv -Path $filePath
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message, 'Critical Issue', 'OK', 'Error')
        exit
    }
    return $file
}

function Update-CurrentIssues {
    $Issue_History.Rows.Clear()
    foreach ($issue in Get-File -filePath $IssuePath) {
        if (($Location_Dropdown.Text -eq $issue.Location) -and ($issue.status -eq '')) {
            [void]$Issue_History.Rows.Add($($issue.'Equipment'), 
                                          $($issue.'Issue Description'),
                                          $($issue.TimeStamp))
        }  
    }
    # Sort the issues by most recent
    $Issue_History.Sort($Issue_History.Columns[2],[System.ComponentModel.ListSortDirection]::Descending)
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

function Clear-Fields {
    $ID_Num_Text.Text = ''
    $Location_Dropdown.Text = (Get-WmiObject -Class Win32_OperatingSystem).description
    $Equipment_Dropdown.Text = ''
    $Desc_Text.Text = ''
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

function Confirm-ID($CurrentField,$Group, $ErrorMSG) {
    Switch -regex ($CurrentField.Text) {
        #FACULTY
        '^AJ((0[1-9])|(1[0-9])|20)$' {
            $ErrorProvider.SetError($Group, '')
            break
        }
        '^DAE[1-5]$' {
            $ErrorProvider.SetError($Group, '')
            break
        }
        '^DHE[1-5]$' {
            $ErrorProvider.SetError($Group, '')
            break
        }
        '^DR((0[1-9])|1[0-5])$' {
            $ErrorProvider.SetError($Group, '')
            break
        }
        #Student
        '^DA(0[1-9]|((1|2)[0-9])|(30))$' {
            $ErrorProvider.SetError($Group, '')
            break
        }
        '^Y1((0[1-9])|((1|2)[0-9])|(30))$' {
            $ErrorProvider.SetError($Group, '')
            break
        }
        '^Y2((0[1-9])|((1|2)[0-9])|(30))$' {
            $ErrorProvider.SetError($Group, '')
            break
        }
        #ADMINISTRATIVE
        '^ADM1$' {
            $ErrorProvider.SetError($Group, '')
            break
        }
        '^TECH$' {
            $ErrorProvider.SetError($Group, '')
            break
        }
        '^BA0[1-2]$' {
            $ErrorProvider.SetError($Group, '')
            break
        }
        '^FO0[1-2]$' {
            $ErrorProvider.SetError($Group, '')
            break
        }
        default {
            $ErrorProvider.SetError($Group, $ErrorMSG)
        }
    }

}

function Confirm-UserInput($Regex, $CurrentField, $ErrorMSG) {
    if ($CurrentField.Text -Notmatch $Regex) {
        $ErrorProvider.SetError($CurrentField, $ErrorMSG)
    }
    else {
        $ErrorProvider.SetError($CurrentField, '')
    }
}

$Submission = [pscustomobject]@{
    'ID'    = ''
    'Location'         = ''
    'Equipment'         = ''
    'Issue Description' = ''
    TimeStamp           = ''
    'Status'            = ''
    'Note'              = ''
}

function Update-Submission {
    $Submission.'Issue Description' = $Desc_Text.Text
    $Submission.'Location' = $Location_Dropdown.Text
    $Submission.'Equipment' = $Equipment_Dropdown.Text
    $Submission.'ID' = $ID_Num_Text.Text
    $Submission.TimeStamp = Get-Date
    return $Submission
}
function Confirm-Dropdown($Dropdown, $Group, $ErrorMSG) {
    if ($Dropdown.Items -contains $Dropdown.Text) {
        $ErrorProvider.SetError($Group,'')
        return $true
    }
    else {
        $ErrorProvider.SetError($Group, $ErrorMSG)
        return $false
    }     
}
#endregion

# Generate ticket file
# Export-Csv -InputObject $Submission -Path $IssuePath -NoTypeInformation

#$IssuePath = "\\dentrix-prod-1\staff\front desk\tickets.csv"
$IssuePath = "$PSScriptRoot\tickets.csv"
Update-CurrentIssues

$dentalArea = Get-File -filePath "$PSScriptRoot\Equipment and Locations.csv"

#region Actions
$Clear_Button.Add_MouseUp( { Clear-Fields })
$dentalArea[0].PSObject.Properties.Name[1..$dentalArea[0].PSObject.Properties.Name.count] | ForEach-Object {[void] $Location_Dropdown.Items.Add($_)}
$Location_Dropdown.Add_SelectedValueChanged({
    Update-CurrentIssues
    $Equipment_Dropdown.Items.Clear()
    $i=0
    foreach ($equipment in $dentalArea.$($Location_Dropdown.SelectedItem)) {
        if ($equipment -eq 'x') {
            $Equipment_Dropdown.Items.Add($dentalArea[$i].'Equipment')
        }
        $i++
    }
})
$Submit_Button.Add_MouseUp( { Confirm-ID -CurrentField $ID_Num_Text -Group $ID_Num_Group -ErrorMSG 'INVALID STUDENT NUMBER' })
$Submit_Button.Add_MouseUp( { Confirm-Dropdown -Dropdown $Location_Dropdown -Group $Location_Group -ErrorMSG 'INVALID LOCATION' })
$Submit_Button.Add_MouseUp( { Confirm-Dropdown -Dropdown $Equipment_Dropdown -Group $Equipment_Group -ErrorMSG 'INVALID EQUIPMENT' })
$Submit_Button.Add_MouseUp( { Confirm-UserInput -regex "." -CurrentField $Desc_Text -ErrorMSG 'INVALID DESCRIPTION' })
$Submit_Button.Add_MouseUp( {
        if (Confirm-NoError) {
            $ErrorProvider.Clear()
            try {
                Update-Submission | Export-Csv -path $IssuePath -append -NoTypeInformation -force
                Update-CurrentIssues
                $Sumbit_Status.Text = ''
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Submission Error', 'OK', 'Error')
            }
        }
        else {
            $Sumbit_Status.Text = 'Please fix errors'
        }
        Start-Sleep -Seconds 2
    })
#endregion

[void]$Form.ShowDialog()