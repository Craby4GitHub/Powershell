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
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

$AdminForm = New-Object system.Windows.Forms.Form
$AdminForm.FormBorderStyle = "Sizable"
$AdminForm.ClientSize = "400,400"
$AdminForm.text = "Equipment Repair Form"
$AdminForm.TopMost = $true
$AdminForm.StartPosition = 'CenterScreen'

$LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$LayoutPanel.Dock = "Fill"
$LayoutPanel.ColumnCount = 7
$LayoutPanel.RowCount = 4
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
$AdminForm.controls.Add($LayoutPanel)

$WorkReport_Button = New-Object system.Windows.Forms.Button
$WorkReport_Button.text = "Work Report"
$WorkReport_Button.Dock = 'Fill'
#$WorkReport_Button.Anchor = 'Left,Right'

$TicketFile_Button = New-Object system.Windows.Forms.Button
$TicketFile_Button.text = "Default Ticket File"
$TicketFile_Button.Dock = 'Left'

$Generate_Group = New-Object system.Windows.Forms.Groupbox
$Generate_Group.text = "Generate Default Files"
$Generate_Group.Dock = 'Fill'
$Generate_Group.controls.AddRange(@($TicketFile_Button))

$Modify_Theme_Button = New-Object system.Windows.Forms.Button
$Modify_Theme_Button.text = "Modify Theme"
$Modify_Theme_Button.Dock = 'Fill'
$Modify_Theme_Button.Anchor = 'Left,Right'

$ColorSelector_Popup = New-Object system.Windows.Forms.Form
$ColorSelector_Popup.Text = 'Color Selector'
$ColorSelector_Popup.FormBorderStyle = "FixedDialog"
$ColorSelector_Popup.ClientSize = "400,400"
$ColorSelector_Popup.TopMost = $true
$ColorSelector_Popup.ControlBox = $false
$ColorSelector_Popup.AutoSize = $true

$ColorSelector_Popup_Layout = New-Object System.Windows.Forms.TableLayoutPanel
$ColorSelector_Popup_Layout.Dock = "Fill"
$ColorSelector_Popup_Layout.ColumnCount = 3
$ColorSelector_Popup_Layout.RowCount = 4
[void]$ColorSelector_Popup_Layout.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$ColorSelector_Popup_Layout.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$ColorSelector_Popup_Layout.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$ColorSelector_Popup_Layout.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15)))
[void]$ColorSelector_Popup_Layout.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$ColorSelector_Popup_Layout.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$ColorSelector_Popup_Layout.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
$ColorSelector_Popup.controls.Add($ColorSelector_Popup_Layout)

$ColorSelector_IDNumber_Color = New-Object System.Windows.Forms.RadioButton
$ColorSelector_IDNumber_Color.Text = 'Color'
$ColorSelector_IDNumber_Color.Font = 'Segoe UI, 8pt'
$ColorSelector_IDNumber_Color.Dock = 'Left'
$ColorSelector_IDNumber_Color.AutoSize = $true
$ColorSelector_IDNumber_Color.CheckAlign = 'BottomCenter'

$ColorSelector_IDNumber_Image = New-Object System.Windows.Forms.RadioButton
$ColorSelector_IDNumber_Image.Text = 'Image'
$ColorSelector_IDNumber_Image.Font = 'Segoe UI, 8pt'
$ColorSelector_IDNumber_Image.Dock = 'Right'
$ColorSelector_IDNumber_Image.AutoSize = $true
$ColorSelector_IDNumber_Image.CheckAlign = 'BottomCenter'

$ColorSelector_IDNumber_Group = New-Object system.Windows.Forms.Groupbox
$ColorSelector_IDNumber_Group.text = "ID Number"
$ColorSelector_IDNumber_Group.Dock = 'Fill'
$ColorSelector_IDNumber_Group.controls.AddRange(@($ColorSelector_IDNumber_Color, $ColorSelector_IDNumber_Image))

#region Layout
$LayoutPanel.Controls.Add($Generate_Group, 1, 0)
$LayoutPanel.SetColumnSpan($Generate_Group, 4)
$LayoutPanel.Controls.Add($WorkReport_Button, 1, 1)
#$LayoutPanel.Controls.Add($Generate_EquipmentFile_Button, 5, 0)

$LayoutPanel.Controls.Add($Modify_Theme_Button, 1, 2)

#endregion


#endregion

#\\dentrix-prod-1\staff\front desk\tickets.csv
$TicketPath = "$PSScriptRoot\tickets.csv"

$TicketFile_Button.Add_MouseUp( {
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

        Export-Csv -InputObject $Submission -Path $TicketPath -NoTypeInformation
    })
$WorkReport_Button.Add_MouseUp( {
        # . (Join-Path $PSSCRIPTROOT "Dental.ps1")
        #$Window.Show()
        foreach ($issue in Get-File -filePath $TicketPath -fileName "Tickets") {
            if (!$issue.'Res Date') {
                # save file with OP, issue and a technotes field
            }  
        }
    })
<#
$Modify_Theme_Button.Add_MouseUp( { 
        . (Join-Path $PSSCRIPTROOT "GUI.ps1")
        $Form.Show($AdminForm)
        $Form.Location = New-Object System.Drawing.Point($($AdminForm.Location.X + $AdminForm.Width), $($AdminForm.Location.Y))

        $ID_Num_Text.Text = 'Y101'
        $Location_Dropdown.Text = 'OP1'
        $Equipment_Dropdown.Text = 'Computer'
        $Desc_Text.Text = 'The computer still isnt working.'
        [void]$Issue_History.Rows.Add('Computer', 'Computer isnt working.', '03May44')
        $Submit_Button.Add_MouseUp( { $Form.Close() })

        $ColorSelector_Popup.Show($Form)
        $ColorSelector_Popup.Location = New-Object System.Drawing.Point($($AdminForm.Location.X + $AdminForm.Width), $($AdminForm.Location.Y + $AdminForm.Height))
        $ColorSelector_Popup_Layout.Controls.Add($ColorSelector_IDNumber_Group, 0, 0)
        $ColorSelector_Popup_Layout.Controls.Add($Location_Group, 1, 0)
        $ColorSelector_Popup_Layout.Controls.Add($Equipment_Group, 2, 0)

        $ColorSelector_Popup_Layout.Controls.Add($Desc_Group, 0, 1)
        $ColorSelector_Popup_Layout.SetColumnSpan($Desc_Group, 3)

        $ColorSelector_Popup_Layout.Controls.Add($Issue_History_Group, 0, 2)
        $ColorSelector_Popup_Layout.SetColumnSpan($Issue_History_Group, 3)

        $ColorSelector_Popup_Layout.Controls.Add($Submit_Button, 1, 3)
    })
#>

[void]$AdminForm.ShowDialog()