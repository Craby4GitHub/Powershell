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
$Form.FormBorderStyle = "Sizable"
$Form.ClientSize = "400,400"
$Form.text = "Equipment Repair Form"
$Form.TopMost = $true
$Form.StartPosition = 'CenterScreen'

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


$Generate_WorkReport_Button = New-Object system.Windows.Forms.Button
$Generate_WorkReport_Button.text = "Generate Work Report"
$Generate_WorkReport_Button.Dock = 'Fill'
$Generate_WorkReport_Button.Anchor = 'Left,Right'


$Generate_TicketFile_Button = New-Object system.Windows.Forms.Button
$Generate_TicketFile_Button.text = "Generate Default Ticket File"
$Generate_TicketFile_Button.Dock = 'Fill'
$Generate_TicketFile_Button.Anchor = 'Left,Right'

$Generate_EquipmentFile_Button = New-Object system.Windows.Forms.Button
$Generate_EquipmentFile_Button.text = "Generate Default Equipment File"
$Generate_EquipmentFile_Button.Dock = 'Fill'
$Generate_EquipmentFile_Button.Anchor = 'Left,Right'

$Modify_Theme_Button = New-Object system.Windows.Forms.Button
$Modify_Theme_Button.text = "Modify Theme"
$Modify_Theme_Button.Dock = 'Fill'
$Modify_Theme_Button.Anchor = 'Left,Right'

$Form.controls.Add($LayoutPanel)



#region Main Layout
$LayoutPanel.Controls.Add($Generate_WorkReport_Button, 1, 0)
$LayoutPanel.Controls.Add($Generate_TicketFile_Button, 3, 0)
$LayoutPanel.Controls.Add($Generate_EquipmentFile_Button, 5, 0)

$LayoutPanel.Controls.Add($Modify_Theme_Button, 1, 1)

#endregion


#endregion
[void]$Form.ShowDialog()