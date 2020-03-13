<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled
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


$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#todo
#remove varible location
#equpoiment field
#auto load op



$Form                            = New-Object system.Windows.Forms.Form
$Form.FormBorderStyle            = "FixedDialog"
$Form.ClientSize                 = "350,350"
$Form.text                       = "Equipment Repair Form"
$Form.TopMost                    = $true
$Form.StartPosition              = 'CenterScreen'

$Stu_Num_Label                   = New-Object system.Windows.Forms.Label
$Stu_Num_Label.text              = "Student Number"
$Stu_Num_Label.AutoSize          = $true
$Stu_Num_Label.width             = 25
$Stu_Num_Label.height            = 10
$Stu_Num_Label.location          = New-Object System.Drawing.Point(15,10)
$Stu_Num_Label.Font              = 'Microsoft Sans Serif,10'

$Stu_Num_Text                    = New-Object system.Windows.Forms.TextBox
$Stu_Num_Text.multiline          = $false
$Stu_Num_Text.width              = 100
$Stu_Num_Text.height             = 20
$Stu_Num_Text.location           = New-Object System.Drawing.Point($(($Form.ClientSize.Width-$Stu_Num_Text.width)/5),30)
$Stu_Num_Text.Font               = 'Microsoft Sans Serif,10'

$Loc_Label                       = New-Object system.Windows.Forms.Label
$Loc_Label.text                  = "Location/Op/Area"
$Loc_Label.AutoSize              = $true
$Loc_Label.width                 = 25
$Loc_Label.height                = 10
$Loc_Label.location              = New-Object System.Drawing.Point(150,10)
$Loc_Label.Font                  = 'Microsoft Sans Serif,10'

$Loc_text                        = New-Object system.Windows.Forms.TextBox
$Loc_text.multiline              = $false
$Loc_text.width                  = 100
$Loc_text.height                 = 20
$Loc_text.location               = New-Object System.Drawing.Point($(($Form.ClientSize.Width-$Loc_text.width)/1.25),30)
$Loc_text.Font                   = 'Microsoft Sans Serif,10'

$Desc_Label                      = New-Object system.Windows.Forms.Label
$Desc_Label.text                 = "Description of Problem"
$Desc_Label.AutoSize             = $true
$Desc_Label.width                = 25
$Desc_Label.height               = 10
$Desc_Label.location             = New-Object System.Drawing.Point(140,76)
$Desc_Label.Font                 = 'Microsoft Sans Serif,10'

$Desc_Text                       = New-Object system.Windows.Forms.TextBox
$Desc_Text.multiline             = $true
$Desc_Text.width                 = 250
$Desc_Text.height                = 200
$Desc_Text.location              = New-Object System.Drawing.Point($(($Form.ClientSize.Width-$Desc_Text.width)/2),100)
$Desc_Text.Font                  = 'Microsoft Sans Serif,10'

$Submit_Button                   = New-Object system.Windows.Forms.Button
$Submit_Button.text              = "Submit"
$Submit_Button.width             = 60
$Submit_Button.height            = 30
$Submit_Button.location          = New-Object System.Drawing.Point(30,310)
$Submit_Button.Font              = 'Microsoft Sans Serif,10'

$Sumbit_Status                   = New-Object system.Windows.Forms.Label
$Sumbit_Status.AutoSize          = $true
$Sumbit_Status.width             = 250
$Sumbit_Status.height            = 10
$Sumbit_Status.location          = New-Object System.Drawing.Point(100,310)
$Sumbit_Status.Font              = 'Microsoft Sans Serif,10'

$Clear_Button                    = New-Object system.Windows.Forms.Button
$Clear_Button.text               = "Clear"
$Clear_Button.width              = 60
$Clear_Button.height             = 30
$Clear_Button.location           = New-Object System.Drawing.Point($(($Form.ClientSize.Width-$Clear_Button.width)/1.35),310)
$Clear_Button.Font               = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($Stu_Num_Text,$Loc_text,$Desc_Text,$Stu_Num_Label,$Desc_Label,$Loc_Label,$Submit_Button,$Sumbit_Status,$Clear_Button))

function Find-Control($ControlType) {
    $FormControls = @()
    ForEach ($control in $Form.Controls) {
        if ($control.GetType().Name -eq $ControlType) {
            $FormControls += $control
        }
    }
    return $FormControls
}

function Clear-TextFields{
    foreach ($control in Find-Control -ControlType 'TextBox') {
        $control.Clear()
    }
}

$Clear_Button.Add_MouseUp({Clear-TextFields})
$Clear_Button.Add_MouseUp({$Sumbit_Status.Text = ''}) # Make cleaner... later

function Confirm-NoTextError {
    foreach ($control in Find-Control -ControlType 'TextBox') {
        if ($ErrorProvider.GetError($control).length -gt 0) {
            return $false
        }else{
            return $true
        }
    }
}

function Confirm-UserInput($Regex,$CurrentField,$ErrorMSG){
    if ($CurrentField.Text -Notmatch $Regex){
        $ErrorProvider.SetError($CurrentField, $ErrorMSG)
    }else {
        $ErrorProvider.SetError($CurrentField, '')
    }
}

$Submission = [pscustomobject]@{
    'Student Number'    = ''
    'Location'          = ''
    'Issue Description' = ''
    TimeStamp           = ''
}

function Update-Submission{
    $Submission.'Issue Description' = $Desc_Text.Text
    $Submission.'Location'          = $Loc_text.Text
    $Submission.'Student Number'    = $Stu_Num_Text.Text
    $Submission.TimeStamp           = Get-Date
    return $Submission
}

$Submit_Button.Add_MouseUp({Confirm-UserInput -regex "^[Aa]{0,1}\d{8}$" -CurrentField $Stu_Num_Text -ErrorMSG 'INVALID STUDENT NUMBER: A12345678'})
$Submit_Button.Add_MouseUp({Confirm-UserInput -regex "." -CurrentField $Loc_text -ErrorMSG 'INVALID LOCATION'})
$Submit_Button.Add_MouseUp({Confirm-UserInput -regex "." -CurrentField $Desc_Text -ErrorMSG 'INVALID DESCRIPTION'})

$filePath = "\\dentrix-prod-1\staff\front desk\tickets.csv"
#$filePath = '\\wc-vm-prtsvr\c$\usmt\tester.csv'

$Submit_Button.Add_MouseUp({
    if (Confirm-NoTextError) {
        $error.Clear()
        try {
            Update-Submission | Export-Csv -path $filePath -append -NoTypeInformation -force
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Submission Error','OK','Error')
        }
    }
    if (!$error) {
        $Sumbit_Status.Text = 'Submitted'
    }
    Start-Sleep -Seconds 5
})

[void]$Form.ShowDialog()