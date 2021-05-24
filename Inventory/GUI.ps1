#region GUI
#region Asset Search Window
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

$Form = New-Object system.Windows.Forms.Form
$Form.AutoScaleMode = 'Font'
$Form.StartPosition = 'Manual'
$Form.Text = 'Inventory Helper Beta 0.4.2'
$Form.ClientSize = "180,300"
$Form.TopMost = $true
$Form.FormBorderStyle = 'Sizable'

$Campus_Dropdown = New-Object System.Windows.Forms.ComboBox
$Campus_Dropdown.DropDownStyle = 'DropDown'
$Campus_Dropdown.Text = 'Select Campus'
$Campus_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Campus_Dropdown.AutoCompleteSource = 'ListItems'
$Campus_Dropdown.TabIndex = 1

$Room_Dropdown = New-Object System.Windows.Forms.ComboBox
$Room_Dropdown.DropDownStyle = 'DropDown'
#$Room_Dropdown.DropDownHeight = $Room_Dropdown.ItemHeight * 5
$Room_Dropdown.ItemHeight = 3000
$Room_Dropdown.Text = 'Select Room'
$Room_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Room_Dropdown.AutoCompleteSource = 'ListItems'
$Room_Dropdown.TabIndex = 2
$Room_Dropdown.Enabled = $false

$PCC_TextBox = New-Object system.Windows.Forms.TextBox
$PCC_TextBox.multiline = $false
$PCC_TextBox.Text = 'PCC or Serial Number'
$PCC_TextBox.TabIndex = 3

$Search_Button = New-Object system.Windows.Forms.Button
$Search_Button.Text = "Search"
$Search_Button.TabIndex = 4
$Form.AcceptButton = $Search_Button

$Option_Button = New-Object system.Windows.Forms.Button
$Option_Button.Text = "Options"
$Option_Button.TabIndex = 5

$StatusBar = New-Object System.Windows.Forms.Label
$StatusBar.Text = "Ready"
#$StatusBar.SizingGrip = $false


#EndRegion

#region Option Popup
$Option_Popup = New-Object system.Windows.Forms.Form
$Option_Popup.Text = 'Options'
$Option_Popup.FormBorderStyle = "FixedDialog"
$Option_Popup.ClientSize = "$($Form.Size.Width),220"
$Option_Popup.TopMost = $true
$Option_Popup.StartPosition = 'Manual'
$Option_Popup.AutoSize = $true

$ScanLog_Button = New-Object system.Windows.Forms.Button
$ScanLog_Button.Text = "Open Scan Log"
$ScanLog_Button.TabIndex = 3

$ErrorLog_Button = New-Object system.Windows.Forms.Button
$ErrorLog_Button.Text = "Open Error Log"
$ErrorLog_Button.FlatAppearance.BorderSize = 0
#EndRegion

#region Asset Update Popup
$AssetUpdate_Popup = New-Object system.Windows.Forms.Form
$AssetUpdate_Popup.Text = 'Asset Update'
$AssetUpdate_Popup.FormBorderStyle = "FixedDialog"
$AssetUpdate_Popup.ClientSize = "$($Form.Size.Width),220"
$AssetUpdate_Popup.TopMost = $true
$AssetUpdate_Popup.StartPosition = 'Manual'
$AssetUpdate_Popup.ControlBox = $false
$AssetUpdate_Popup.AutoSize = $true

$Assigneduser_TextBox_Popup = New-Object system.Windows.Forms.TextBox
$Assigneduser_TextBox_Popup.multiline = $false
$Assigneduser_TextBox_Popup.Text = "Assigned User"
$Assigneduser_TextBox_Popup.TabIndex = 1

$Status_Dropdown_Popup = New-Object System.Windows.Forms.ComboBox
$Status_Dropdown_Popup.DropDownStyle = 'DropDown'
$Status_Dropdown_Popup.Text = "Status"
$Status_Dropdown_Popup.AutoCompleteMode = 'SuggestAppend'
$Status_Dropdown_Popup.AutoCompleteSource = 'ListItems'
$Status_Dropdown_Popup.TabIndex = 2

$OK_Button_Popup = New-Object system.Windows.Forms.Button
$OK_Button_Popup.Text = "OK"
$OK_Button_Popup.TabIndex = 3
$AssetUpdate_Popup.AcceptButton = $OK_Button_Popup
$AssetUpdate_Popup.AcceptButton.DialogResult = 'OK'

$Cancel_Button_Popup = New-Object system.Windows.Forms.Button
$Cancel_Button_Popup.Text = "Cancel"
$Cancel_Button_Popup.TabIndex = 4
$AssetUpdate_Popup.CancelButton = $Cancel_Button_Popup
$AssetUpdate_Popup.CancelButton.DialogResult = 'Cancel'
#EndRegion

#region Login Window
$Login_Form = New-Object system.Windows.Forms.Form
$Login_Form.FormBorderStyle = "FixedDialog"
$Login_Form.ClientSize = "400,220"
$Login_Form.TopMost = $true
$Login_Form.StartPosition = 'CenterScreen'
$Login_Form.ControlBox = $false
$Login_Form.AutoSize = $true

$Username_TextBox = New-Object system.Windows.Forms.TextBox
$Username_TextBox.multiline = $false
$Username_TextBox.Text = $env:USERNAME
$Username_TextBox.TabIndex = 1

$Password_TextBox = New-Object system.Windows.Forms.TextBox
$Password_TextBox.multiline = $false
$Password_TextBox.Text = "PimaRocks"
$Password_TextBox.TabIndex = 2
$Password_TextBox.PasswordChar = '*'
$Password_TextBox.Select()

$OK_Button_Login = New-Object system.Windows.Forms.Button
$OK_Button_Login.Text = "Login"

$OK_Button_Login.Dock = 'Fill'
$OK_Button_Login.TabIndex = 3

$OK_Button_Login.FlatStyle = 1
$OK_Button_Login.FlatAppearance.BorderSize = 0
$Login_Form.AcceptButton = $OK_Button_Login
$Login_Form.AcceptButton.DialogResult = 'OK'

$Cancel_Button_Login = New-Object system.Windows.Forms.Button
$Cancel_Button_Login.Text = "Cancel"
$Cancel_Button_Login.TabIndex = 4

$Login_Form.CancelButton = $Cancel_Button_Login
$Login_Form.CancelButton.DialogResult = 'Cancel'

$LayoutPanel_Login = New-Object System.Windows.Forms.TableLayoutPanel
$LayoutPanel_Login.Dock = "Fill"
$LayoutPanel_Login.ColumnCount = 4
$LayoutPanel_Login.RowCount = 4
#$LayoutPanel_Login.CellBorderStyle = 1
[void]$LayoutPanel_Login.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel_Login.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel_Login.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel_Login.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))

[void]$LayoutPanel_Login.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$LayoutPanel_Login.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$LayoutPanel_Login.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$LayoutPanel_Login.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))

$LayoutPanel_Login.Controls.Add($Username_TextBox, 1, 0)
$LayoutPanel_Login.SetColumnSpan($Username_TextBox, 2)
$LayoutPanel_Login.Controls.Add($Password_TextBox, 1, 1)
$LayoutPanel_Login.SetColumnSpan($Password_TextBox, 2)
$LayoutPanel_Login.Controls.Add($OK_Button_Login, 1, 2)
$LayoutPanel_Login.Controls.Add($Cancel_Button_Login, 2, 2)
$Login_Form.controls.Add($LayoutPanel_Login)
#EndRegion
#EndRegion

#region UI Theme
$Theme = Get-Content -Path .\theme.json | ConvertFrom-Json

$Login_Form.Backcolor = $Theme.Login_Form.BackColor
$Login_Form.ForeColor = $Theme.Login_Form.ForeColor 

$Username_TextBox.Font = $Theme.Username_TextBox.Font
$Username_TextBox.Backcolor = $Theme.Username_TextBox.Backcolor
$Username_TextBox.ForeColor = $Theme.Username_TextBox.ForeColor
$Username_TextBox.BorderStyle = 1
$Username_TextBox.Dock = 'Top'
$Username_TextBox.Anchor = 'Left,Right'

$Password_TextBox.Font = $Theme.Password_TextBox.Font
$Password_TextBox.Backcolor = $Theme.Password_TextBox.Backcolor
$Password_TextBox.ForeColor = $Theme.Password_TextBox.ForeColor
$Password_TextBox.BorderStyle = 1
$Password_TextBox.Dock = 'Top'
$Password_TextBox.Anchor = 'Left,Right'

$OK_Button_Login.Font = $Theme.OK_Button_Login.Font
$OK_Button_Login.Backcolor = $Theme.OK_Button_Login.Backcolor
$OK_Button_Login.ForeColor = $Theme.OK_Button_Login.ForeColor

$Cancel_Button_Login.Font = $Theme.Cancel_Button_Login.Font
$Cancel_Button_Login.Backcolor = $Theme.Cancel_Button_Login.Backcolor
$Cancel_Button_Login.ForeColor = $Theme.Cancel_Button_Login.ForeColor
$Cancel_Button_Login.FlatStyle = 1
$Cancel_Button_Login.FlatAppearance.BorderSize = 0
$Cancel_Button_Login.Dock = 'Fill'

$Form.Font = $Theme.Form.Font
$Form.BackColor = $Theme.Form.BackColor
$Form.ForeColor = $Theme.Form.ForeColor

$Campus_Dropdown.Font = $Theme.Campus_Dropdown.Font
$Campus_Dropdown.Backcolor = $Theme.Campus_Dropdown.Backcolor
$Campus_Dropdown.ForeColor = $Theme.Campus_Dropdown.ForeColor
$Campus_Dropdown.FlatStyle = 0
#$Campus_Dropdown.Dock = "Fill"
$Campus_Dropdown.Anchor = 'left, Right'

$Room_Dropdown.Font = $Theme.Room_Dropdown.Font
$Room_Dropdown.Backcolor = $Theme.Room_Dropdown.Backcolor
$Room_Dropdown.ForeColor = $Theme.Room_Dropdown.ForeColor
$Room_Dropdown.FlatStyle = 0
$Room_Dropdown.Dock = "Fill"
$Room_Dropdown.Anchor = 'Left,Right'

$PCC_TextBox.Font = $Theme.PCC_TextBox.Font
$PCC_TextBox.Backcolor = $Theme.PCC_TextBox.Backcolor
$PCC_TextBox.ForeColor = $Theme.PCC_TextBox.ForeColor
$PCC_TextBox.BorderStyle = 1
$PCC_TextBox.Dock = 'Fill'
$PCC_TextBox.Anchor = 'Left,Right'

$Search_Button.Font = $Theme.Search_Button.Font
$Search_Button.Backcolor = $Theme.Search_Button.Backcolor
$Search_Button.ForeColor = $Theme.Search_Button.ForeColor
$Search_Button.FlatStyle = 1
$Search_Button.FlatAppearance.BorderSize = 0
$Search_Button.Dock = 'Fill'

$Option_Button.Backcolor = $Theme.Option_Button.BackColor
$Option_Button.ForeColor = $Theme.Option_Button.ForeColor
$Option_Button.FlatStyle = 1
$Option_Button.FlatAppearance.BorderSize = 0
$Option_Button.Dock = 'Fill'

$StatusBar.Font = $Theme.StatusBar.Font
$StatusBar.Backcolor = $Theme.StatusBar.Backcolor
$StatusBar.ForeColor = $Theme.StatusBar.ForeColor
$StatusBar.Dock = 'Bottom'

$AssetUpdate_Popup.Backcolor = $Theme.AssetUpdate_Popup.BackColor
$AssetUpdate_Popup.ForeColor = $Theme.AssetUpdate_Popup.ForeColor

$Assigneduser_TextBox_Popup.Font = $Theme.Assigneduser_TextBox_Popup.Font
$Assigneduser_TextBox_Popup.Backcolor = $Theme.Assigneduser_TextBox_Popup.Backcolor
$Assigneduser_TextBox_Popup.ForeColor = $Theme.Assigneduser_TextBox_Popup.ForeColor
$Assigneduser_TextBox_Popup.BorderStyle = 1
$Assigneduser_TextBox_Popup.Dock = 'Top'
$Assigneduser_TextBox_Popup.Anchor = 'Left,Right'

$Status_Dropdown_Popup.Font = $Theme.Status_Dropdown_Popup.Font
$Status_Dropdown_Popup.Backcolor = $Theme.Status_Dropdown_Popup.Backcolor
$Status_Dropdown_Popup.ForeColor = $Theme.Status_Dropdown_Popup.ForeColor
$Status_Dropdown_Popup.FlatStyle = 0
$Status_Dropdown_Popup.Dock = "Fill"
$Status_Dropdown_Popup.Anchor = 'Top, Left, Right'

$OK_Button_Popup.Font = $Theme.OK_Button_Popup.Font
$OK_Button_Popup.Backcolor = $Theme.OK_Button_Popup.Backcolor
$OK_Button_Popup.ForeColor = $Theme.OK_Button_Popup.ForeColor
$OK_Button_Popup.FlatStyle = 1
$OK_Button_Popup.FlatAppearance.BorderSize = 0
$OK_Button_Popup.Dock = 'Fill'

$Cancel_Button_Popup.Font = $Theme.Cancel_Button_Popup.Font
$Cancel_Button_Popup.Backcolor = $Theme.Cancel_Button_Popup.Backcolor
$Cancel_Button_Popup.ForeColor = $Theme.Cancel_Button_Popup.ForeColor
$Cancel_Button_Popup.FlatStyle = 1
$Cancel_Button_Popup.FlatAppearance.BorderSize = 0
$Cancel_Button_Popup.Dock = 'Fill'

$Option_Popup.Backcolor = $Theme.Option_Popup.BackColor
$Option_Popup.ForeColor = $Theme.Option_Popup.ForeColor

$ScanLog_Button.Font = $Theme.ScanLog_Button.Font
$ScanLog_Button.Backcolor = $Theme.ScanLog_Button.Backcolor
$ScanLog_Button.ForeColor = $Theme.ScanLog_Button.ForeColor
$ScanLog_Button.FlatStyle = 1
$ScanLog_Button.FlatAppearance.BorderSize = 0
$ScanLog_Button.Dock = 'Fill'

$ErrorLog_Button.Font = $Theme.ErrorLog_Button.Font
$ErrorLog_Button.Backcolor = $Theme.ErrorLog_Button.BackColor
$ErrorLog_Button.ForeColor = $Theme.ErrorLog_Button.ForeColor
$ErrorLog_Button.TabIndex = 3
$ErrorLog_Button.FlatStyle = 1
$ErrorLog_Button.Dock = 'Fill'

#Endregion

#region Layouts
$Main_LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$Main_LayoutPanel.ColumnCount = 3
$Main_LayoutPanel.RowCount = 5
#$Main_LayoutPanel.CellBorderStyle = 1
$Main_LayoutPanel.Dock = "Fill"
[void]$Main_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, .5)))
[void]$Main_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$Main_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, .5)))
[void]$Main_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$Main_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$Main_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$Main_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 2)))
[void]$Main_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, .25)))

$Main_LayoutPanel.Controls.Add($Campus_Dropdown, 1, 0)
$Main_LayoutPanel.Controls.Add($Room_Dropdown, 1, 1)
$Main_LayoutPanel.Controls.Add($PCC_TextBox, 1, 2)
$Main_LayoutPanel.Controls.Add($Search_Button, 1, 3)
$Main_LayoutPanel.Controls.Add($Option_Button, 2, 4)
$Form.controls.AddRange(@($Main_LayoutPanel, $StatusBar))


$AssetUpdate_LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$AssetUpdate_LayoutPanel.Dock = "Fill"
$AssetUpdate_LayoutPanel.ColumnCount = 4
$AssetUpdate_LayoutPanel.RowCount = 4
#$AssetUpdate_LayoutPanel.CellBorderStyle = 1
[void]$AssetUpdate_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$AssetUpdate_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$AssetUpdate_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$AssetUpdate_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))

[void]$AssetUpdate_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 6)))
[void]$AssetUpdate_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 6)))
[void]$AssetUpdate_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 6)))
[void]$AssetUpdate_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 1)))

$AssetUpdate_LayoutPanel.Controls.Add($Assigneduser_TextBox_Popup, 1, 0)
$AssetUpdate_LayoutPanel.SetColumnSpan($Assigneduser_TextBox_Popup, 2)
$AssetUpdate_LayoutPanel.Controls.Add($Status_Dropdown_Popup, 1, 1)
$AssetUpdate_LayoutPanel.SetColumnSpan($Status_Dropdown_Popup, 2)
$AssetUpdate_LayoutPanel.Controls.Add($OK_Button_Popup, 1, 2)
$AssetUpdate_LayoutPanel.Controls.Add($Cancel_Button_Popup, 2, 2)
$AssetUpdate_Popup.controls.Add($AssetUpdate_LayoutPanel)


$Options_LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$Options_LayoutPanel.Dock = "Fill"
$Options_LayoutPanel.ColumnCount = 4
$Options_LayoutPanel.RowCount = 4
#$Options_LayoutPanel.CellBorderStyle = 1
[void]$Options_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$Options_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$Options_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$Options_LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))

[void]$Options_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$Options_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$Options_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$Options_LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))

$Options_LayoutPanel.Controls.Add($ScanLog_Button, 1, 0)
$Options_LayoutPanel.SetColumnSpan($ScanLog_Button, 2)
$Options_LayoutPanel.Controls.Add($ErrorLog_Button, 1, 1)
$Options_LayoutPanel.SetColumnSpan($ErrorLog_Button, 2)
$Option_Popup.controls.Add($Options_LayoutPanel)


#EndRegion

# Comment line out below if code in production
#[void]$Form.ShowDialog()