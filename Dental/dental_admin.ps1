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

. (Join-Path $PSSCRIPTROOT "GUI.ps1")



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

$EquipmentFile_Button = New-Object system.Windows.Forms.Button
$EquipmentFile_Button.text = "Default Equipment File"
$EquipmentFile_Button.Dock = 'Right'

$Generate_Group = New-Object system.Windows.Forms.Groupbox
$Generate_Group.text = "Generate Default Files"
$Generate_Group.Dock = 'Fill'
$Generate_Group.controls.AddRange(@($EquipmentFile_Button,$TicketFile_Button))

$Modify_Theme_Button = New-Object system.Windows.Forms.Button
$Modify_Theme_Button.text = "Modify Theme"
$Modify_Theme_Button.Dock = 'Fill'
$Modify_Theme_Button.Anchor = 'Left,Right'

#region Main Layout
$LayoutPanel.Controls.Add($Generate_Group, 1, 0)
$LayoutPanel.SetColumnSpan($Generate_Group,2)
$LayoutPanel.Controls.Add($WorkReport_Button, 1, 1)
#$LayoutPanel.Controls.Add($Generate_EquipmentFile_Button, 5, 0)

$LayoutPanel.Controls.Add($Modify_Theme_Button, 1, 2)

#endregion


#endregion



$Modify_Theme_Button.Add_MouseUp( { 
    $Form.show()
    $Form.Location = New-Object System.Drawing.Point($($AdminForm.Location.X + $AdminForm.Width),$($AdminForm.Location.Y))
})


[void]$AdminForm.ShowDialog()