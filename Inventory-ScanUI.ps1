#Requires -Modules Selenium
If (-not(Get-InstalledModule Selenium -ErrorAction silentlycontinue)) {
    Install-Module Selenium -Confirm:$False -Force -Scope CurrentUser
}

#region UI
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

#https://colorhunt.co/palette/10792

$Form = New-Object system.Windows.Forms.Form
$Form.AutoScaleMode = 'Font'
$Form.StartPosition = 'CenterScreen'
$Form.Text = 'Inventory Scanning'
$Form.ClientSize = "200,330"
$Form.Font = 'Segoe UI, 18pt'
$Form.TopMost = $true
$Form.BackColor = '#324e7a'
$Form.ForeColor = '#eeeeee' 
$Form.FormBorderStyle = 'Sizable'

$Campus_Dropdown = New-Object System.Windows.Forms.ComboBox
$Campus_Dropdown.DropDownStyle = 'DropDown'
$Campus_Dropdown.Text = 'Select Campus'
#$Campus_Dropdown.Font = 'Segoe UI, 18pt'
$Campus_Dropdown.Backcolor = '#1b3666'
$Campus_Dropdown.ForeColor = '#eeeeee' 
$Campus_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Campus_Dropdown.AutoCompleteSource = 'ListItems'
$Campus_Dropdown.TabIndex = 1
#$Campus_Dropdown.Dock = "Fill"
$Campus_Dropdown.FlatStyle = 0
$Campus_Dropdown.Anchor = 'left, Right'

$Room_Label = New-Object system.Windows.Forms.Label
$Room_Label.Text = "Room:" 
#$Room_Label.Font = 'Segoe UI, 10pt, style=Bold'
$Room_Label.AutoSize = $true
$Room_Label.Dock = 'Bottom'
$Room_Label.Anchor = 'Left,Right,Bottom'

$Room_Dropdown = New-Object System.Windows.Forms.ComboBox
$Room_Dropdown.DropDownStyle = 'DropDown'
#$Room_Dropdown.DropDownHeight = $Room_Dropdown.ItemHeight * 5
$Room_Dropdown.ItemHeight = 3000
$Room_Dropdown.Text = 'Select Room'
#$Room_Dropdown.Font = 'Segoe UI, 18pt'
$Room_Dropdown.Backcolor = '#1b3666'
$Room_Dropdown.ForeColor = '#eeeeee' 
$Room_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Room_Dropdown.AutoCompleteSource = 'ListItems'
$Room_Dropdown.TabIndex = 2
$Room_Dropdown.Dock = "Fill"
$Room_Dropdown.Enabled = $false
$Room_Dropdown.FlatStyle = 0
$Room_Dropdown.Anchor = 'Left,Right'

$PCC_TextBox = New-Object system.Windows.Forms.TextBox
$PCC_TextBox.multiline = $false
#$PCC_TextBox.Font = 'Segoe UI, 15pt'
$PCC_TextBox.Backcolor = '#1b3666'
$PCC_TextBox.Text = 'PCC Number'
$PCC_TextBox.ForeColor = '#a3a3a3' 
$PCC_TextBox.Dock = 'Fill'
$PCC_TextBox.TabIndex = 3
$PCC_TextBox.BorderStyle = 1
$PCC_TextBox.Anchor = 'Left,Right'

$Search_Button = New-Object system.Windows.Forms.Button
$Search_Button.Text = "Search"
$Search_Button.Dock = 'Fill'
$Search_Button.TabIndex = 4
$Search_Button.Backcolor = '#616161'
$Search_Button.ForeColor = '#eeeeee' 
#$Search_Button.Font = 'Segoe UI, 18pt, style=Bold'
$Search_Button.FlatStyle = 1
$Search_Button.FlatAppearance.BorderSize = 0
$Form.AcceptButton = $Search_Button

$StatusBar = New-Object System.Windows.Forms.StatusBar
$StatusBar.Text = "Ready"
$StatusBar.SizingGrip = $false
$StatusBar.Font = 'Segoe UI, 12pt'
$StatusBar.Dock = 'Bottom'
$StatusBar.Backcolor = '#1b3666'
#$StatusBar.ForeColor = '#eeeeee'

$LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$LayoutPanel.Dock = "Fill"
$LayoutPanel.ColumnCount = 3
$LayoutPanel.RowCount = 5
#$LayoutPanel.CellBorderStyle = 1

[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, .5)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, .5)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 2)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, .25)))


#region Asset Update Popup
$AssetUpdate_Popup = New-Object system.Windows.Forms.Form
$AssetUpdate_Popup.Text = 'Asset Update'
$AssetUpdate_Popup.Backcolor = '#324e7a'
$AssetUpdate_Popup.ForeColor = '#eeeeee' 
$AssetUpdate_Popup.FormBorderStyle = "FixedDialog"
$AssetUpdate_Popup.ClientSize = "400,220"
$AssetUpdate_Popup.TopMost = $true
$AssetUpdate_Popup.StartPosition = 'CenterScreen'
$AssetUpdate_Popup.ControlBox = $false
$AssetUpdate_Popup.AutoSize = $true

$Status_DropdownLabel_Popup = New-Object system.Windows.Forms.Label
$Status_DropdownLabel_Popup.Font = 'Segoe UI, 8pt'
$Status_DropdownLabel_Popup.ForeColor = '#eeeeee' 
$Status_DropdownLabel_Popup.AutoSize = $true
$Status_DropdownLabel_Popup.Dock = 'Bottom'

$Status_Dropdown_Popup = New-Object System.Windows.Forms.ComboBox
$Status_Dropdown_Popup.DropDownStyle = 'DropDown'
$Status_Dropdown_Popup.Text = "Status"
$Status_Dropdown_Popup.Backcolor = '#1b3666'
$Status_Dropdown_Popup.ForeColor = '#eeeeee' 
$Status_Dropdown_Popup.AutoCompleteMode = 'SuggestAppend'
$Status_Dropdown_Popup.AutoCompleteSource = 'ListItems'
$Status_Dropdown_Popup.TabIndex = 1
$Status_Dropdown_Popup.Dock = "Fill"
$Status_Dropdown_Popup.FlatStyle = 0
$Status_Dropdown_Popup.Anchor = 'Top, Left, Right'
$Status_Dropdown_Popup.Font = 'Segoe UI, 18pt'

$Assigneduser_TextBoxLabel_Popup = New-Object system.Windows.Forms.Label
$Assigneduser_TextBoxLabel_Popup.Text = "Assigned User"
$Assigneduser_TextBoxLabel_Popup.Font = 'Segoe UI, 8pt'
$Assigneduser_TextBoxLabel_Popup.ForeColor = '#eeeeee' 
$Assigneduser_TextBoxLabel_Popup.AutoSize = $true
$Assigneduser_TextBoxLabel_Popup.Dock = 'Bottom'

$Assigneduser_TextBox_Popup = New-Object system.Windows.Forms.TextBox
$Assigneduser_TextBox_Popup.multiline = $false
$Assigneduser_TextBox_Popup.Text = "Assigned User"
$Assigneduser_TextBox_Popup.Font = 'Segoe UI, 18pt'
$Assigneduser_TextBox_Popup.Backcolor = '#1b3666'
$Assigneduser_TextBox_Popup.ForeColor = '#a3a3a3' 
$Assigneduser_TextBox_Popup.Dock = 'Top'
$Assigneduser_TextBox_Popup.TabIndex = 2
$Assigneduser_TextBox_Popup.BorderStyle = 1
$Assigneduser_TextBox_Popup.Anchor = 'Left,Right'

$OK_Button_Popup = New-Object system.Windows.Forms.Button
$OK_Button_Popup.Text = "OK"
$OK_Button_Popup.Backcolor = '#616161'
$OK_Button_Popup.ForeColor = '#eeeeee' 
$OK_Button_Popup.Dock = 'Fill'
$OK_Button_Popup.TabIndex = 3
$OK_Button_Popup.Font = 'Segoe UI, 18pt'
$OK_Button_Popup.FlatStyle = 1
$OK_Button_Popup.FlatAppearance.BorderSize = 0
$AssetUpdate_Popup.AcceptButton = $OK_Button_Popup

$Cancel_Button_Popup = New-Object system.Windows.Forms.Button
$Cancel_Button_Popup.Text = "Cancel"
$Cancel_Button_Popup.Font = 'Segoe UI, 18pt'
$Cancel_Button_Popup.Backcolor = '#1b3666'
$Cancel_Button_Popup.ForeColor = '#eeeeee' 
$Cancel_Button_Popup.Dock = 'Fill'
$Cancel_Button_Popup.TabIndex = 4
$Cancel_Button_Popup.FlatStyle = 1
$Cancel_Button_Popup.FlatAppearance.BorderSize = 0

$LayoutPanel_Popup = New-Object System.Windows.Forms.TableLayoutPanel
$LayoutPanel_Popup.Dock = "Fill"
$LayoutPanel_Popup.ColumnCount = 4
$LayoutPanel_Popup.RowCount = 4
#$LayoutPanel_Popup.CellBorderStyle = 1
[void]$LayoutPanel_Popup.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel_Popup.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel_Popup.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel_Popup.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))

[void]$LayoutPanel_Popup.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$LayoutPanel_Popup.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$LayoutPanel_Popup.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$LayoutPanel_Popup.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))

$LayoutPanel_Popup.Controls.Add($Assigneduser_TextBox_Popup, 1, 0)
$LayoutPanel_Popup.SetColumnSpan($Assigneduser_TextBox_Popup, 2)
$LayoutPanel_Popup.Controls.Add($Status_Dropdown_Popup, 1, 1)
$LayoutPanel_Popup.SetColumnSpan($Status_Dropdown_Popup, 2)
$LayoutPanel_Popup.Controls.Add($OK_Button_Popup, 1, 2)
$LayoutPanel_Popup.Controls.Add($Cancel_Button_Popup, 2, 2)
$AssetUpdate_Popup.controls.Add($LayoutPanel_Popup)

$LayoutPanel.Controls.Add($Campus_Dropdown, 1, 0)
$LayoutPanel.Controls.Add($Room_Dropdown, 1, 1)
$LayoutPanel.Controls.Add($PCC_TextBox, 1, 2)
$LayoutPanel.Controls.Add($Search_Button, 1, 3)
$Form.controls.AddRange(@($LayoutPanel, $StatusBar))
#EndRegion
#EndRegion


@('1', '2', '3', '4', '5', '6', '7', '8') | ForEach-Object { [void] $Campus_Dropdown.Items.Add($_) }
@('1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15') | ForEach-Object { [void] $Room_Dropdown.Items.Add($_) }
@('1', '2', '3', '4', '5') | ForEach-Object { [void] $Status_Dropdown_Popup.Items.Add($_) }

$Search_Button.Add_MouseUp( {
        $AssetUpdate_Popup.ShowDialog()
    })

$OK_Button_Popup.Add_MouseUp( {
        $AssetUpdate_Popup.Close()
    })

$PCC_TextBox.Add_MouseDown( {
        $PCC_TextBox.clear()
        $PCC_TextBox.Forecolor = '#eeeeee'
    })

$Assigneduser_TextBox_Popup.Add_MouseDown( {
        $Assigneduser_TextBox_Popup.clear()
        $Assigneduser_TextBox_Popup.Forecolor = '#eeeeee'
    })

[void]$Form.ShowDialog()