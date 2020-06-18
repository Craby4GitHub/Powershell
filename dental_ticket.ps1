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

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object system.Windows.Forms.Form
$Form.FormBorderStyle = "FixedDialog"
$Form.ClientSize = "400,350"
$Form.text = "Equipment Repair Form"
$Form.TopMost = $true
$Form.StartPosition = 'CenterScreen'

$ID_Num_Group = New-Object system.Windows.Forms.Groupbox
$ID_Num_Group.height = 50
$ID_Num_Group.width = 120
$ID_Num_Group.text = "ID Number"
$ID_Num_Group.location = New-Object System.Drawing.Point(10, 10)

$ID_Num_Text = New-Object system.Windows.Forms.TextBox
$ID_Num_Text.multiline = $false
$ID_Num_Text.width = 100
$ID_Num_Text.height = 20
$ID_Num_Text.location = New-Object System.Drawing.Point(5, 17)
$ID_Num_Text.Font = 'Microsoft Sans Serif,10'

$Location_Group = New-Object system.Windows.Forms.Groupbox
$Location_Group.height = 50
$Location_Group.width = 120
$Location_Group.text = "Location"
$Location_Group.location = New-Object System.Drawing.Point(140, 10)

$Location_Dropdown              = New-Object system.Windows.Forms.ComboBox
$Location_Dropdown.width        = 100
$Location_Dropdown.height       = 20
$Location_Dropdown.location     = New-Object System.Drawing.Point(5,17)
$Location_Dropdown.Font         = 'Microsoft Sans Serif,10'
$Location_Dropdown.Text = (Get-WmiObject -Class Win32_OperatingSystem).Description
$Location_Dropdown.AutoCompleteMode = 'Suggest'
$Location_Dropdown.AutoCompleteSource = 'ListItems'

$Equipment_Group = New-Object system.Windows.Forms.Groupbox
$Equipment_Group.height = 50
$Equipment_Group.width = 120
$Equipment_Group.text = "Equipment"
$Equipment_Group.location = New-Object System.Drawing.Point(270, 10)

$Equipment_Dropdown              = New-Object system.Windows.Forms.ComboBox
$Equipment_Dropdown.width        = 100
$Equipment_Dropdown.height       = 20
$Equipment_Dropdown.location     = New-Object System.Drawing.Point(5,17)
$Equipment_Dropdown.Font         = 'Microsoft Sans Serif,10'
$Equipment_Dropdown.AutoCompleteMode = 'Suggest'
$Equipment_Dropdown.AutoCompleteSource = 'ListItems'

$Desc_Label = New-Object system.Windows.Forms.Label
$Desc_Label.text = "Description of Problem"
$Desc_Label.AutoSize = $true
$Desc_Label.width = 25
$Desc_Label.height = 10
$Desc_Label.location = New-Object System.Drawing.Point(140, 76)
$Desc_Label.Font = 'Microsoft Sans Serif,10'

$Desc_Text = New-Object system.Windows.Forms.TextBox
$Desc_Text.multiline = $true
$Desc_Text.width = 300
$Desc_Text.height = 70
$Desc_Text.location = New-Object System.Drawing.Point(50, 100)
$Desc_Text.Font = 'Microsoft Sans Serif,10'

$Issue_History_Label = New-Object system.Windows.Forms.Label
$Issue_History_Label.text = "Current Issues"
$Issue_History_Label.AutoSize = $true
$Issue_History_Label.width = 25
$Issue_History_Label.height = 10
$Issue_History_Label.location = New-Object System.Drawing.Point(140, 180)
$Issue_History_Label.Font = 'Microsoft Sans Serif,10'

$Issue_History = New-Object system.Windows.Forms.DataGridView
$Issue_History.width = 360
$Issue_History.height = 70
$Issue_History.location = New-Object System.Drawing.Point(20, 200)
$Issue_History.ScrollBars = "Vertical"
$Issue_History.AutoGenerateColumns = $true
$Issue_History.ColumnCount = 3
$Issue_History.Columns[0].Name = 'Equipment'
$Issue_History.Columns[1].Name = 'Issue'
$Issue_History.Columns[2].Name = 'Submitted'


$Submit_Button = New-Object system.Windows.Forms.Button
$Submit_Button.text = "Submit"
$Submit_Button.width = 60
$Submit_Button.height = 30
$Submit_Button.location = New-Object System.Drawing.Point(30, 310)
$Submit_Button.Font = 'Microsoft Sans Serif,10'

#not submitting, work on this
$Form.AcceptButton = $Submit_Button  

$Sumbit_Status = New-Object system.Windows.Forms.Label
$Sumbit_Status.AutoSize = $true
$Sumbit_Status.width = 250
$Sumbit_Status.height = 10
$Sumbit_Status.location = New-Object System.Drawing.Point(100, 310)
$Sumbit_Status.Font = 'Microsoft Sans Serif,10'

$Clear_Button = New-Object system.Windows.Forms.Button
$Clear_Button.text = "Clear"
$Clear_Button.width = 60
$Clear_Button.height = 30
$Clear_Button.location = New-Object System.Drawing.Point(300, 310)
$Clear_Button.Font = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($ID_Num_Group, $Location_Group, $Equipment_Group, $Desc_Label, $Desc_Text, $Submit_Button, $Sumbit_Status, $Clear_Button, $Issue_History, $Issue_History_Label))
$Equipment_Group.controls.AddRange(@($Equipment_Dropdown))
$ID_Num_Group.controls.AddRange(@($ID_Num_Text))
$Location_Group.controls.AddRange(@($Location_Dropdown))

$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

function Get-Issues {   
    try {
        $issues = Import-Csv -Path $IssuePath
    }
    catch {
        Export-Csv -InputObject $Submission -Path $IssuePath -NoTypeInformation
        [System.Windows.Forms.MessageBox]::Show("Please relaunch. Error: " + $_.Exception.Message, 'Critical Issue', 'OK', 'Error')
        exit
    }
    return $issues
}
function Update-CurrentIssues {
    $Issue_History.Rows.Clear()
    foreach ($issue in Get-Issues) {
        if (($OP_Text.Text -eq $issue.Operatory) -and ($issue.status -eq '')) {
            [void]$Issue_History.Rows.Add($($issue.'Equipment'), 
                                          $($issue.'Issue Description'),
                                          $($issue.TimeStamp))
        }  
    }
    # Sort the issues by most recent
    $Issue_History.Sort($Issue_History.Columns[2],[System.ComponentModel.ListSortDirection]::Descending)
}

function Find-Control($ControlType) {
    $FormControls = @()
    ForEach ($control in $Form.Controls) {
        if ($control.ToString().StartsWith("System.Windows.Forms.GroupBox")) {
            foreach ($gcontrol in $control.controls) {
                if ($gcontrol.GetType().Name -eq $ControlType) {
                    $FormControls += $gcontrol
                }
            }
        }
        elseif ($control.GetType().Name -eq $ControlType) {
            $FormControls += $control
        }
    }
    return $FormControls
}

function Clear-TextFields($Fields) {
    foreach ($control in Find-Control -ControlType 'TextBox') {
        $control.Clear()
        $OP_Text.Text = (Get-WmiObject -Class Win32_OperatingSystem).description
    }
}

function Confirm-NoTextError {
    $i = 0
    foreach ($control in Find-Control -ControlType 'TextBox') {
        if ($ErrorProvider.GetError($control).length -gt 0) {
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
    $Submission.'Location' = $Location_Text.Text
    $Submission.'Equipment' = $Equipment_Dropdown.Text
    $Submission.'ID' = $ID_Num_Text.Text
    $Submission.TimeStamp = Get-Date
    return $Submission
}
#$IssuePath = "\\dentrix-prod-1\staff\front desk\tickets.csv"
$IssuePath = "$PSScriptRoot\test.csv"
Update-CurrentIssues

$Clear_Button.Add_MouseUp( { Clear-TextFields })
@('1','2') | ForEach-Object {[void] $Location_Dropdown.Items.Add($_)}
@('xray','thingy') | ForEach-Object {[void] $Equipment_Dropdown.Items.Add($_)}

function Confirm-ID($CurrentField,$ErrorMSG) {
    Switch -regex ($CurrentField.Text) {
        #FACULTY
        '^AJ((0[1-9])|(1[0-9])|20)$' {
            $ErrorProvider.SetError($CurrentField, '')
            break
        }
        '^DAE[1-5]$' {
            $ErrorProvider.SetError($CurrentField, '')
            break
        }
        '^DHE[1-5]$' {
            $ErrorProvider.SetError($CurrentField, '')
            break
        }
        '^DR((0[1-9])|1[0-5])$' {
            $ErrorProvider.SetError($CurrentField, '')
            break
        }
        #Student
        '^DA(0[1-9]|((1|2)[0-9])|(30))$' {
            $ErrorProvider.SetError($CurrentField, '')
            break
        }
        '^Y1((0[1-9])|((1|2)[0-9])|(30))$' {
            $ErrorProvider.SetError($CurrentField, '')
            break
        }
        '^Y2((0[1-9])|((1|2)[0-9])|(30))$' {
            $ErrorProvider.SetError($CurrentField, '')
            break
        }
        #ADMINISTRATIVE
        '^ADM1$' {
            $ErrorProvider.SetError($CurrentField, '')
            break
        }
        '^TECH$' {
            $ErrorProvider.SetError($CurrentField, '')
            break
        }
        '^BA0[1-2]$' {
            $ErrorProvider.SetError($CurrentField, '')
            break
        }
        '^FO0[1-2]$' {
            $ErrorProvider.SetError($CurrentField, '')
            break
        }
        default {
            $ErrorProvider.SetError($CurrentField, $ErrorMSG)
        }
    }

}


$Submit_Button.Add_MouseUp( { Confirm-ID -CurrentField $ID_Num_Text -ErrorMSG 'INVALID STUDENT NUMBER' })
$Submit_Button.Add_MouseUp( { Confirm-UserInput -regex "\w*" -CurrentField $Location_Dropdown -ErrorMSG 'INVALID LOCATION' })
$Submit_Button.Add_MouseUp( { Confirm-UserInput -regex "." -CurrentField $Equipment_Dropdown -ErrorMSG 'INVALID EQUIPMENT' })
$Submit_Button.Add_MouseUp( { Confirm-UserInput -regex "." -CurrentField $Desc_Text -ErrorMSG 'INVALID DESCRIPTION' })
$Submit_Button.Add_MouseUp( {
        if (Confirm-NoTextError) {
            $error.Clear()
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

[void]$Form.ShowDialog()