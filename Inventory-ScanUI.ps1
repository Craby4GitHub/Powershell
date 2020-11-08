$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

#region UI
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#https://colorhunt.co/palette/10792

$Form = New-Object system.Windows.Forms.Form
$Form.AutoScaleMode = 'Font'
$Form.StartPosition = 'CenterScreen'
$Form.ClientSize = "350,150"
$Form.text = "ITAM - Inventory Automation"
$Form.Font = 'Segoe UI, 8pt'
$Form.TopMost = $true
$Form.BackColor = '#303841'
$Form.ForeColor = '#eeeeee' 
#$Form.Icon = New-Object System.Drawing.Icon ("$PSScriptRoot\favicon.ico")

$Campus_Dropdown = New-Object System.Windows.Forms.ComboBox
$Campus_Dropdown.DropDownStyle = 'DropDown'
$Campus_Dropdown.DropDownHeight = $Campus_Dropdown.ItemHeight * 5
$Campus_Dropdown.Text = 'Select Campus'
$Campus_Dropdown.Font = 'Segoe UI, 8pt'
$Campus_Dropdown.BackColor = '#3a4750'
$Campus_Dropdown.ForeColor = '#eeeeee' 
$Campus_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Campus_Dropdown.AutoCompleteSource = 'ListItems'
$Campus_Dropdown.TabIndex = 1
$Campus_Dropdown.Dock = "Bottom"

$Room_Dropdown = New-Object System.Windows.Forms.ComboBox
$Room_Dropdown.DropDownStyle = 'DropDown'
$Room_Dropdown.DropDownHeight = $Room_Dropdown.ItemHeight * 5
$Room_Dropdown.Text = 'Select Room'
$Room_Dropdown.Font = 'Segoe UI, 8pt'
$Room_Dropdown.BackColor = '#3a4750'
$Room_Dropdown.ForeColor = '#eeeeee' 
$Room_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Room_Dropdown.AutoCompleteSource = 'ListItems'
$Room_Dropdown.TabIndex = 2
$Room_Dropdown.Dock = "Bottom"
$Room_Dropdown.Enabled = $false

$PCC_Label = New-Object system.Windows.Forms.Label
$PCC_Label.text = "PCC Number:"
$PCC_Label.Font = 'Segoe UI, 10pt, style=Bold'
$PCC_Label.AutoSize = $true
$PCC_Label.Dock = 'Bottom'

$PCC_TextBox = New-Object system.Windows.Forms.TextBox
$PCC_TextBox.multiline = $false
$PCC_TextBox.Font = 'Segoe UI, 8pt'
$PCC_TextBox.BackColor = '#3a4750'
$PCC_TextBox.ForeColor = '#eeeeee' 
$PCC_TextBox.Dock = 'Fill'
$PCC_TextBox.TabIndex = 3

$Search_Button = New-Object system.Windows.Forms.Button
$Search_Button.text = "Search"
$Search_Button.Dock = 'Fill'
$Search_Button.TabIndex = 4
$Search_Button.BackColor = '#00adb5'
$Search_Button.Font = 'Segoe UI, 10pt, style=Bold'
$Form.AcceptButton = $Search_Button

$StatusBar = New-Object System.Windows.Forms.StatusBar
$StatusBar.Text = "Ready"
$StatusBar.SizingGrip = $false
$StatusBar.Dock = 'Bottom'

#Region Panel
$LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$LayoutPanel.Dock = "Fill"
$LayoutPanel.ColumnCount = 5
$LayoutPanel.RowCount = 6
#$LayoutPanel.CellBorderStyle = 1
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 2)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 2)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 2)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))

$LayoutPanel.Controls.Add($Campus_Dropdown, 1, 1)
$LayoutPanel.Controls.Add($Room_Dropdown, 3, 1)
$LayoutPanel.Controls.Add($PCC_Label, 1, 3)
$LayoutPanel.Controls.Add($PCC_TextBox, 1, 4)
$LayoutPanel.Controls.Add($Search_Button, 3, 3)
$LayoutPanel.SetRowSpan($Search_Button, 3)

$Form.controls.AddRange(@($LayoutPanel, $StatusBar))
#EndRegion 

#region Assest Update Popup
$AssetUpdate_Popup = New-Object system.Windows.Forms.Form
$AssetUpdate_Popup.Text = 'Asset Update'
$AssetUpdate_Popup.BackColor = '#303841'
$AssetUpdate_Popup.ForeColor = '#eeeeee' 
$AssetUpdate_Popup.FormBorderStyle = "FixedDialog"
$AssetUpdate_Popup.ClientSize = "250,100"
$AssetUpdate_Popup.TopMost = $true
$AssetUpdate_Popup.StartPosition = 'CenterScreen'
$AssetUpdate_Popup.ControlBox = $false

$Assigneduser_TextBoxLabel_Popup = New-Object system.Windows.Forms.Label
$Assigneduser_TextBoxLabel_Popup.text = "Assigned User"
$Assigneduser_TextBoxLabel_Popup.ForeColor = '#eeeeee' 
$Assigneduser_TextBoxLabel_Popup.AutoSize = $true
$Assigneduser_TextBoxLabel_Popup.Dock = 'Bottom'

$Assigneduser_TextBox_Popup = New-Object system.Windows.Forms.TextBox
$Assigneduser_TextBox_Popup.multiline = $false
$Assigneduser_TextBox_Popup.BackColor = '#3a4750'
$Assigneduser_TextBox_Popup.ForeColor = '#eeeeee' 
$Assigneduser_TextBox_Popup.Dock = 'Top'
$Assigneduser_TextBox_Popup.TabIndex = 1

$Status_DropdownLabel_Popup = New-Object system.Windows.Forms.Label
$Status_DropdownLabel_Popup.text = "Status"
$Status_DropdownLabel_Popup.ForeColor = '#eeeeee' 
$Status_DropdownLabel_Popup.AutoSize = $true
$Status_DropdownLabel_Popup.Dock = 'Bottom'

$Status_Dropdown_Popup = New-Object System.Windows.Forms.ComboBox
$Status_Dropdown_Popup.DropDownStyle = 'DropDown'
$Status_Dropdown_Popup.BackColor = '#3a4750'
$Status_Dropdown_Popup.ForeColor = '#eeeeee' 
$Status_Dropdown_Popup.AutoCompleteMode = 'SuggestAppend'
$Status_Dropdown_Popup.AutoCompleteSource = 'ListItems'
$Status_Dropdown_Popup.TabIndex = 2

$OK_Button_Popup = New-Object system.Windows.Forms.Button
$OK_Button_Popup.text = "OK"
$OK_Button_Popup.BackColor = '#00adb5'
$OK_Button_Popup.ForeColor = '#eeeeee' 
$OK_Button_Popup.Dock = 'Bottom'
$OK_Button_Popup.TabIndex = 4

$Cancel_Button_Popup = New-Object system.Windows.Forms.Button
$Cancel_Button_Popup.text = "Cancel"
$Cancel_Button_Popup.BackColor = '#3a4750'
$Cancel_Button_Popup.ForeColor = '#eeeeee' 
$Cancel_Button_Popup.Dock = 'Bottom'
$Cancel_Button_Popup.TabIndex = 4

$LayoutPanel_Popup = New-Object System.Windows.Forms.TableLayoutPanel
$LayoutPanel_Popup.Dock = "Fill"
$LayoutPanel_Popup.ColumnCount = 2
$LayoutPanel_Popup.RowCount = 3
#$LayoutPanel_Popup.CellBorderStyle = 1
[void]$LayoutPanel_Popup.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$LayoutPanel_Popup.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$LayoutPanel_Popup.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 30)))
[void]$LayoutPanel_Popup.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$LayoutPanel_Popup.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 30)))

$LayoutPanel_Popup.Controls.Add($Assigneduser_TextBox_Popup, 1, 1)
$LayoutPanel_Popup.Controls.Add($Status_Dropdown_Popup, 0, 1)
$LayoutPanel_Popup.Controls.Add($Assigneduser_TextBoxLabel_Popup, 1, 0)
$LayoutPanel_Popup.Controls.Add($Status_DropdownLabel_Popup, 0, 0)
$LayoutPanel_Popup.Controls.Add($OK_Button_Popup, 0, 2)
$LayoutPanel_Popup.Controls.Add($Cancel_Button_Popup, 1, 2)

$AssetUpdate_Popup.controls.Add($LayoutPanel_Popup)

$Search_Button.Add_MouseUp( {
    $AssetUpdate_Popup.ShowDialog()
})

$OK_Button_Popup.Add_MouseUp( {
    $AssetUpdate_Popup.Close()
})

[void]$Form.ShowDialog()