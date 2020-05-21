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

#todo
#Equipment field - link to csv, clear
#add pk(i forgot what pk stood for, oops)
#ticket history | Done


$Form = New-Object system.Windows.Forms.Form
$Form.FormBorderStyle = "FixedDialog"
$Form.ClientSize = "400,350"
$Form.text = "Equipment Repair Form"
$Form.TopMost = $true
$Form.StartPosition = 'CenterScreen'

<#
$Stu_Num_Label                   = New-Object system.Windows.Forms.Label
$Stu_Num_Label.text              = "Student Number"
$Stu_Num_Label.AutoSize          = $true
$Stu_Num_Label.width             = 25
$Stu_Num_Label.height            = 10
$Stu_Num_Label.location          = New-Object System.Drawing.Point(25,10)
$Stu_Num_Label.Font              = 'Microsoft Sans Serif,10'
#>

$Stu_Num_Group = New-Object system.Windows.Forms.Groupbox
$Stu_Num_Group.height = 50
$Stu_Num_Group.width = 110
$Stu_Num_Group.text = "Student Number"
$Stu_Num_Group.location = New-Object System.Drawing.Point(25, 10)


$Stu_Num_Text = New-Object system.Windows.Forms.TextBox
$Stu_Num_Text.multiline = $false
$Stu_Num_Text.width = 100
$Stu_Num_Text.height = 20
$Stu_Num_Text.location = New-Object System.Drawing.Point(5, 17)
$Stu_Num_Text.Font = 'Microsoft Sans Serif,10'
$Stu_Num_Text.text = 'A01070484'

<#
$OP_Label                       = New-Object system.Windows.Forms.Label
$OP_Label.text                  = "Operatory"
$OP_Label.AutoSize              = $true
$OP_Label.width                 = 25
$OP_Label.height                = 10
$OP_Label.location              = New-Object System.Drawing.Point(150,10)
$OP_Label.Font                  = 'Microsoft Sans Serif,10'
#>

$OP_Group = New-Object system.Windows.Forms.Groupbox
$OP_Group.height = 50
$OP_Group.width = 110
$OP_Group.text = "Operatory"
$OP_Group.location = New-Object System.Drawing.Point(140, 10)

$OP_Text = New-Object system.Windows.Forms.TextBox
$OP_Text.multiline = $false
$OP_Text.width = 100
$OP_Text.height = 20
$OP_Text.location = New-Object System.Drawing.Point(5, 17)
$OP_Text.Font = 'Microsoft Sans Serif,10'
$OP_Text.Text = (Get-WmiObject -Class Win32_OperatingSystem).description

$Equipment_Group = New-Object system.Windows.Forms.Groupbox
$Equipment_Group.height = 50
$Equipment_Group.width = 110
$Equipment_Group.text = "Equipment"
$Equipment_Group.location = New-Object System.Drawing.Point(260, 10)
<#
$Equipment_Label                       = New-Object system.Windows.Forms.Label
$Equipment_Label.text                  = "Equipment"
$Equipment_Label.AutoSize              = $true
$Equipment_Label.width                 = 25
$Equipment_Label.height                = 10
$Equipment_Label.location              = New-Object System.Drawing.Point(275,10)
$Equipment_Label.Font                  = 'Microsoft Sans Serif,10'
#>
$Equipment_Text = New-Object system.Windows.Forms.TextBox
$Equipment_Text.multiline = $false
$Equipment_Text.width = 100
$Equipment_Text.height = 20
$Equipment_Text.location = New-Object System.Drawing.Point(5, 17)
$Equipment_Text.Font = 'Microsoft Sans Serif,10'

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

$Form.controls.AddRange(@($Stu_Num_Group, $OP_Group, $Equipment_Group, $Desc_Label, $Desc_Text, $Submit_Button, $Sumbit_Status, $Clear_Button, $Issue_History, $Issue_History_Label))
$Equipment_Group.controls.AddRange(@($Equipment_Text))
$Stu_Num_Group.controls.AddRange(@($Stu_Num_Text))
$OP_Group.controls.AddRange(@($OP_Text))

$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

function Update-CurrentIssues {
    $Issue_History.Rows.Clear()
    #$filePath = "\\dentrix-prod-1\staff\front desk\tickets.csv"
    $filePath = '\\wc-vm-prtsvr\c$\usmt\tester.csv'
    $issues = Import-Csv -Path $filePath
    foreach ($issue in $issues) {
        if (($OP_Text.Text -eq $issue.Operatory) -and ($issue.status -eq '')) {
            [void]$Issue_History.Rows.Add($($issue.'Equipment'), $($issue.'Issue Description'),$($issue.TimeStamp))
        }  
    }
    $Issue_History.Sort($Issue_History.Columns[2],[System.ComponentModel.ListSortDirection]::Descending)
}
Update-CurrentIssues

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

$Clear_Button.Add_MouseUp( { Clear-TextFields })
$Clear_Button.Add_MouseUp( { $Sumbit_Status.Text = '' }) # Make cleaner... later

function Confirm-NoTextError {
    foreach ($control in Find-Control -ControlType 'TextBox') {
        if ($ErrorProvider.GetError($control).length -gt 0) {
            return $false
        }
        else {
            return $true
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
    'Student Number'    = ''
    'Operatory'         = ''
    'Equipment'         = ''
    'Issue Description' = ''
    TimeStamp           = ''
    'Status'            = ''
    'Note'              = ''
}

function Update-Submission {
    $Submission.'Issue Description' = $Desc_Text.Text
    $Submission.'Operatory' = $OP_Text.Text
    $Submission.'Equipment' = $Equipment_Text.Text
    $Submission.'Student Number' = $Stu_Num_Text.Text
    $Submission.TimeStamp = Get-Date
    return $Submission
}

$Submit_Button.Add_MouseUp( { Confirm-UserInput -regex "^[Aa]{0,1}\d{8}$" -CurrentField $Stu_Num_Text -ErrorMSG 'INVALID STUDENT NUMBER: A12345678' })
$Submit_Button.Add_MouseUp( { Confirm-UserInput -regex "\w*" -CurrentField $OP_Text -ErrorMSG 'INVALID LOCATION' })
$Submit_Button.Add_MouseUp( { Confirm-UserInput -regex "." -CurrentField $Desc_Text -ErrorMSG 'INVALID DESCRIPTION' })
$Submit_Button.Add_MouseUp( {
        if (Confirm-NoTextError) {
            $error.Clear()
            try {
                Update-Submission | Export-Csv -path $filePath -append -NoTypeInformation -force
                Update-CurrentIssues
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Submission Error', 'OK', 'Error')
            }
        }
        else {
            $Sumbit_Status.Text = 'Please fix errors'
        }
        if (!$error) {
            #$Sumbit_Status.Text = 'Submitted'
        }
        Start-Sleep -Seconds 2
    })

[void]$Form.ShowDialog()