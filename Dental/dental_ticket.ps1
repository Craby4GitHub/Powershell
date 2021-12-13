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
[Console.Window]::ShowWindow($consolePtr, 0) | Out-Null
#endregion

#region GUI
. (Join-Path $PSSCRIPTROOT "GUI.ps1")
#endregion

#region Functions
function Get-File($filePath, $fileName) {   
    try {
        $file = Import-Csv -Path $filePath
    }
    catch {
        #Write-Log -Level 'FATAL' -Message $_.Exception.InnerException.Message.toString
        [System.Windows.Forms.MessageBox]::Show("Error: Could not open $($fileName), please contact the front desk for help. ", 'Critical Issue', 'OK', 'Error')
        exit
    }
    return $file
}

function Update-CurrentIssues {
    $Issue_History.Rows.Clear()

    foreach ($issue in Get-File -filePath $TicketPath -fileName "Tickets") {
        if (($Location_Dropdown.Text -eq $issue.Location) -and ($issue.status -eq '')) {
            try {
                [void]$Issue_History.Rows.Add(
                    $($issue.'Equipment'), 
                    $($issue.'Issue Description'),
                    $([datetime]$issue.'TimeStamp'))
            }
            catch {
                #$Error[0].Exception.GetType()
                Write-Log -Message $_.Exception.InnerException.Message.toString
                #$_.ScriptStackTrace.toString    
            }
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


function Update-Submission {

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
    $Line = "$Stamp,$Level,$env:COMPUTERNAME,$($Message -replace','),$Element"

    try {
        Add-Content $ErrorPath -Value $Line
    }
    catch {
        Write-Host 'Unable to open Error Log, close any open instances.'
    }  
}

#endregion

#\\dentrix-prod-1\staff\front desk\tickets.csv
$TicketPath = "$PSScriptRoot\tickets.csv"
$ErrorPath = "$PSScriptRoot\errors.csv"

$dentalArea = Get-File -filePath "$PSScriptRoot\Equipment and Locations.csv" -fileName "Equipment List"
$dentalArea[0].PSObject.Properties.Name[1..$dentalArea[0].PSObject.Properties.Name.count] | ForEach-Object { [void] $Location_Dropdown.Items.Add($_) }

$Location_Dropdown.Text = (Get-WmiObject -Class Win32_OperatingSystem).Description

Update-CurrentIssues
Update-CurrentEquipment

#region Actions

$Location_Dropdown.Add_SelectedValueChanged( {
        Update-CurrentIssues
        Update-CurrentEquipment
        $Issue_History_Group.text = "Current issues in $($Location_Dropdown.Text)"
    })

$Equipment_Dropdown.Add_SelectedValueChanged( { 
        Check-DuplicateIssue 
    })

$Submit_Button.Add_MouseUp( { 
        Confirm-ID -CurrentField $ID_Num_Text -Group $ID_Num_Group -ErrorMSG 'INVALID STUDENT NUMBER' 
        Confirm-Dropdown -Dropdown $Location_Dropdown -Group $Location_Group -ErrorMSG 'INVALID LOCATION'
        Confirm-Dropdown -Dropdown $Equipment_Dropdown -Group $Equipment_Group -ErrorMSG 'INVALID EQUIPMENT'
        Confirm-UserInput -regex '' -CurrentField $Desc_Text -ErrorMSG 'INVALID DESCRIPTION'
    })
$Submit_Button.Add_MouseUp( {
        if (Confirm-NoError) {
            $ErrorProvider.Clear()
            try {
                Update-Submission | Export-Csv -Path $TicketPath -Append -NoTypeInformation -Force
                Update-CurrentIssues
            }
            catch {
                Write-Log -Level 'FATAL' -Message $_.Exception.Message
                [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Submission Error', 'OK', 'Error')
            }
        }
        Start-Sleep -Seconds 2
    })
#endregion

[void]$Form.ShowDialog()