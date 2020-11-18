<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Dental Ticket
#>

#region hide the powerhsell console
# https://stackoverflow.com/a/40621143/20267
Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    '
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0) # hide
#endregion

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
$Form.Font = 'Segoe UI, 8pt'

$LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$LayoutPanel.Dock = "Fill"
$LayoutPanel.ColumnCount = 7
$LayoutPanel.RowCount = 4
#$LayoutPanel.CellBorderStyle = 1
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))

$ID_Num_Text = New-Object system.Windows.Forms.TextBox
$ID_Num_Text.multiline = $false
$ID_Num_Text.Dock = 'Fill'
$ID_Num_Group = New-Object system.Windows.Forms.Groupbox
$ID_Num_Group.text = "ID Number"
$ID_Num_Group.Dock = 'Fill'
$ID_Num_Group.controls.Add($ID_Num_Text)

$Location_Dropdown = New-Object system.Windows.Forms.ComboBox
$Location_Dropdown.DropDownStyle = "DropDown"
$Location_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Location_Dropdown.AutoCompleteSource = 'ListItems'
$Location_Dropdown.DropDownWidth = '100'
$Location_Dropdown.Dock = 'Fill'
$Location_Group = New-Object system.Windows.Forms.Groupbox
$Location_Group.text = "Location"
$Location_Group.Dock = 'Fill'
$Location_Group.controls.Add($Location_Dropdown)

$Equipment_Dropdown = New-Object system.Windows.Forms.ComboBox
$Equipment_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Equipment_Dropdown.AutoCompleteSource = 'ListItems'
$Equipment_Dropdown.Dock = 'Fill'
$Equipment_Group = New-Object system.Windows.Forms.Groupbox
$Equipment_Group.text = "Equipment"
$Equipment_Group.Dock = 'Fill'
$Equipment_Group.controls.Add($Equipment_Dropdown)

$Desc_Text = New-Object system.Windows.Forms.TextBox
$Desc_Text.multiline = $true
$Desc_Text.Dock = 'Fill'
$Desc_Group = New-Object system.Windows.Forms.Groupbox
$Desc_Group.text = "Description of Issue"
$Desc_Group.Dock = 'Fill'
$Desc_Group.Controls.Add($Desc_Text)

$Issue_History = New-Object system.Windows.Forms.DataGridView
$Issue_History.ScrollBars = "Vertical"
$Issue_History.AutoGenerateColumns = $true
$Issue_History.AutoSizeColumnsMode = 'Fill'
$Issue_History.RowHeadersVisible = $false
$Issue_History.ColumnCount = 3
$Issue_History.Columns[0].Name = 'Equipment'
$Issue_History.Columns[1].Name = 'Issue'
$Issue_History.Columns[2].DefaultCellStyle.Format = "ddMMMyy hh:mm tt"
$Issue_History.Columns[2].Name = 'Submitted'
$Issue_History.TabStop = $false
$Issue_History.Dock = 'Fill'
$Issue_History.Anchor = 'Left,Right'
$Issue_History_Group = New-Object system.Windows.Forms.Groupbox
$Issue_History_Group.text = "Current Issues"
$Issue_History_Group.Dock = 'Fill'
$Issue_History_Group.Controls.Add($Issue_History)

$Submit_Button = New-Object system.Windows.Forms.Button
$Submit_Button.text = "Submit"
$Submit_Button.Dock = 'Fill'
$Submit_Button.Anchor = 'Left,Right'

$Sumbit_Status = New-Object system.Windows.Forms.Label
$Sumbit_Status.AutoSize = $true
$Form.AcceptButton = $Submit_Button

$Form.controls.Add($LayoutPanel)
#endregion

#region Main Layout
$LayoutPanel.Controls.Add($ID_Num_Group, 1, 0)
$LayoutPanel.Controls.Add($Location_Group, 3, 0)
$LayoutPanel.Controls.Add($Equipment_Group, 5, 0)

$LayoutPanel.Controls.Add($Desc_Group, 1, 1)
$LayoutPanel.SetColumnSpan($Desc_Group, 5)

$LayoutPanel.Controls.Add($Issue_History_Group, 1, 2)
$LayoutPanel.SetColumnSpan($Issue_History_Group, 5)

$LayoutPanel.Controls.Add($Submit_Button, 3, 3)
#endregion

#region Functions
function Get-File($filePath) {   
    try {
        $file = Import-Csv -Path $filePath
    }
    catch {
        Write-Log -Level 'FATAL' -Message $_.Exception.Message
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
                $([datetime]$issue.'TimeStamp'))
        }  
    }
    # Sort the issues by most recent
    $Issue_History.Sort($Issue_History.Columns[2], [System.ComponentModel.ListSortDirection]::Descending)
}

function Update-CurrentEquipment {
    $Equipment_Dropdown.Items.Clear()
    $i = 0
    foreach ($equipment in $dentalArea.$($Location_Dropdown.SelectedItem)) {
        if ($equipment -eq 'x') {
            [void]$Equipment_Dropdown.Items.Add($dentalArea[$i].'Equipment')
        }
        $i++
    }
}

function Confirm-NoError {
    $i = 0
    foreach ($control in $LayoutPanel.Controls) {
        if ($ErrorProvider.GetError($control)) {
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

function Confirm-ID($CurrentField, $Group, $ErrorMSG) {
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
            Write-Log -Level INFO -Message 'Invalid ID' -Element $CurrentField.Text
            $ErrorProvider.SetError($Group, $ErrorMSG)
        }
    }
}

function Confirm-UserInput($Regex, $CurrentField, $ErrorMSG) {
    if ($CurrentField.Text -Notmatch $Regex) {
        Write-Log -Level INFO -Message 'Invalid Input' -Element $CurrentField.Text
        $ErrorProvider.SetError($CurrentField, $ErrorMSG)
    }
    else {
        $ErrorProvider.SetError($CurrentField, '')
    }
}

$Submission = [pscustomobject]@{
    'ID'                = ''
    'Location'          = ''
    'Equipment'         = ''
    'Issue Description' = ''
    'TimeStamp'         = ''
    'Status'            = ''
    'Res Date'          = ''
    'Resolution'        = ''
    'Who'               = ''
    'Note'              = ''
}

function Update-Submission {
    $Submission.'Issue Description' = $Desc_Text.Text
    $Submission.'Location' = $Location_Dropdown.Text
    $Submission.'Equipment' = $Equipment_Dropdown.Text
    $Submission.'ID' = $ID_Num_Text.Text
    $Submission.'TimeStamp' = Get-Date
    return $Submission
}

function Confirm-Dropdown($Dropdown, $Group, $ErrorMSG) {
    if ($Dropdown.Items -contains $Dropdown.Text) {
        $ErrorProvider.SetError($Group, '')
        return $true
    }
    else {
        Write-Log -Level INFO -Message 'Invalid Selection' -Element $Dropdown.Text
        $ErrorProvider.SetError($Group, $ErrorMSG)
        return $false
    }     
}

function Check-DuplicateIssue {
    foreach ($row in $Issue_History.Rows) {
        if ($Equipment_Dropdown.Text -eq 'Other') {
            break
        }
        if ($Equipment_Dropdown.Text -eq $row.Cells.Value[0]) {           
            $DuplicateTicket = [System.Windows.Forms.MessageBox]::Show("A ticket has already been submitted for $($Equipment_Dropdown.Text):`n`n$($row.Cells.Value[1])`n`nAre you having this issue?", 'Warning', 'YesNo', 'Warning')
            if ($DuplicateTicket -eq 'Yes') {
                $Equipment_Dropdown.SelectedIndex = -1
            }
        }
    }
}

Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $False)]
        [ValidateSet("INFO", "WARN", "ERROR", "FATAL", "DEBUG")]
        [String]
        $Level = "INFO",

        [Parameter(Mandatory = $True)]
        [string]
        $Message,

        [Parameter(Mandatory = $false)]
        [string]
        $Element
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp,$Level,$env:COMPUTERNAME,$Message,$Element"

    Add-Content $ErrorPath -Value $Line
}

#endregion

# Generate ticket file
# Export-Csv -InputObject $Submission -Path $IssuePath -NoTypeInformation
#\\dentrix-prod-1\staff\front desk\tickets.csv
$IssuePath = "$PSScriptRoot\tickets.csv"
$ErrorPath = "$PSScriptRoot\errors.csv"

$dentalArea = Get-File -filePath "$PSScriptRoot\Equipment and Locations.csv"
$dentalArea[0].PSObject.Properties.Name[1..$dentalArea[0].PSObject.Properties.Name.count] | ForEach-Object { [void] $Location_Dropdown.Items.Add($_) }

$Location_Dropdown.Text = (Get-WmiObject -Class Win32_OperatingSystem).Description

Update-CurrentIssues
Update-CurrentEquipment

#region Actions

$Location_Dropdown.Add_SelectedValueChanged( {
        Update-CurrentIssues
        Update-CurrentEquipment
    })

$Equipment_Dropdown.Add_SelectedValueChanged( { Check-DuplicateIssue })

$Submit_Button.Add_MouseUp( { Confirm-ID -CurrentField $ID_Num_Text -Group $ID_Num_Group -ErrorMSG 'INVALID STUDENT NUMBER' })
$Submit_Button.Add_MouseUp( { Confirm-Dropdown -Dropdown $Location_Dropdown -Group $Location_Group -ErrorMSG 'INVALID LOCATION' })
$Submit_Button.Add_MouseUp( { Confirm-Dropdown -Dropdown $Equipment_Dropdown -Group $Equipment_Group -ErrorMSG 'INVALID EQUIPMENT' })
$Submit_Button.Add_MouseUp( { Confirm-UserInput -regex '' -CurrentField $Desc_Text -ErrorMSG 'INVALID DESCRIPTION' })
$Submit_Button.Add_MouseUp( {
        if (Confirm-NoError) {
            $ErrorProvider.Clear()
            try {
                Update-Submission | Export-Csv -Path $IssuePath -Append -NoTypeInformation -Force
                Update-CurrentIssues
                $Sumbit_Status.Text = ''
            }
            catch {
                Write-Log -Level 'FATAL' -Message $_.Exception.Message
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