If (-not(Get-InstalledModule Selenium -ErrorAction silentlycontinue)) {
    Install-Module Selenium -Confirm:$False -Force -Scope CurrentUser
}

#region UI
#region Asset Search Window
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

$Form = New-Object system.Windows.Forms.Form
$Form.AutoScaleMode = 'Font'
$Form.StartPosition = 'Manual'
$Form.Text = 'Inventory Helper Beta 0.3.2'
$Form.ClientSize = "180,250"
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
$PCC_TextBox.Text = 'PCC/Serial Number'
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
#$Form.AcceptButton.DialogResult = 'OK'

$Option_Button = New-Object system.Windows.Forms.Button
$Option_Button.Text = "Options"
$Option_Button.Dock = 'Fill'
$Option_Button.TabIndex = 5
$Option_Button.Backcolor = '#616161'
$Option_Button.ForeColor = '#eeeeee' 
$Option_Button.FlatStyle = 1
$Option_Button.FlatAppearance.BorderSize = 0
#$Form.AcceptButton.DialogResult = 'OK'

$StatusBar = New-Object System.Windows.Forms.Label
$StatusBar.Text = "Ready"
#$StatusBar.SizingGrip = $false
$StatusBar.Font = 'Segoe UI, 12pt'
$StatusBar.Dock = 'Bottom'
$StatusBar.Backcolor = '#1b3666'
#$StatusBar.ForeColor = '#eeeeee'

$LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$LayoutPanel.Dock = "Fill"
$LayoutPanel.ColumnCount = 3
$LayoutPanel.RowCount = 5
$LayoutPanel.CellBorderStyle = 1

[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, .5)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, .5)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 2)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, .25)))

$LayoutPanel.Controls.Add($Campus_Dropdown, 1, 0)
$LayoutPanel.Controls.Add($Room_Dropdown, 1, 1)
$LayoutPanel.Controls.Add($PCC_TextBox, 1, 2)
$LayoutPanel.Controls.Add($Search_Button, 1, 3)
$LayoutPanel.Controls.Add($Option_Button, 2, 4)
$Form.controls.AddRange(@($LayoutPanel, $StatusBar))
#EndRegion

#region Option Popup
$Option_Popup = New-Object system.Windows.Forms.Form
$Option_Popup.Text = 'Options'
$Option_Popup.Backcolor = '#324e7a'
$Option_Popup.ForeColor = '#eeeeee' 
$Option_Popup.FormBorderStyle = "FixedDialog"
$Option_Popup.ClientSize = "$($Form.Size.Width),220"
$Option_Popup.TopMost = $true
$Option_Popup.StartPosition = 'Manual'
#$Option_Popup.ControlBox = $false
$Option_Popup.AutoSize = $true

$ScanLog_Button = New-Object system.Windows.Forms.Button
$ScanLog_Button.Text = "Open Scan Log"
$ScanLog_Button.Backcolor = '#616161'
$ScanLog_Button.ForeColor = '#eeeeee' 
$ScanLog_Button.Dock = 'Fill'
$ScanLog_Button.TabIndex = 3
$ScanLog_Button.Font = 'Segoe UI, 18pt'
$ScanLog_Button.FlatStyle = 1
$ScanLog_Button.FlatAppearance.BorderSize = 0

$ErrorLog_Button = New-Object system.Windows.Forms.Button
$ErrorLog_Button.Text = "Open Error Log"
$ErrorLog_Button.Backcolor = '#616161'
$ErrorLog_Button.ForeColor = '#eeeeee' 
$ErrorLog_Button.Dock = 'Fill'
$ErrorLog_Button.TabIndex = 3
$ErrorLog_Button.Font = 'Segoe UI, 18pt'
$ErrorLog_Button.FlatStyle = 1
$ErrorLog_Button.FlatAppearance.BorderSize = 0

$LayoutPanel_Options = New-Object System.Windows.Forms.TableLayoutPanel
$LayoutPanel_Options.Dock = "Fill"
$LayoutPanel_Options.ColumnCount = 4
$LayoutPanel_Options.RowCount = 4
#$LayoutPanel_Options.CellBorderStyle = 1
[void]$LayoutPanel_Options.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel_Options.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel_Options.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel_Options.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))

[void]$LayoutPanel_Options.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$LayoutPanel_Options.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$LayoutPanel_Options.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$LayoutPanel_Options.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))

$LayoutPanel_Options.Controls.Add($ScanLog_Button, 1, 0)
$LayoutPanel_Options.SetColumnSpan($ScanLog_Button, 2)
$LayoutPanel_Options.Controls.Add($ErrorLog_Button, 1, 1)
$LayoutPanel_Options.SetColumnSpan($ErrorLog_Button, 2)
$Option_Popup.controls.Add($LayoutPanel_Options)

#EndRegion

#region Asset Update Popup
$AssetUpdate_Popup = New-Object system.Windows.Forms.Form
$AssetUpdate_Popup.Text = 'Asset Update'
$AssetUpdate_Popup.Backcolor = '#324e7a'
$AssetUpdate_Popup.ForeColor = '#eeeeee' 
$AssetUpdate_Popup.FormBorderStyle = "FixedDialog"
$AssetUpdate_Popup.ClientSize = "$($Form.Size.Width),220"
$AssetUpdate_Popup.TopMost = $true
$AssetUpdate_Popup.StartPosition = 'Manual'
$AssetUpdate_Popup.ControlBox = $false
$AssetUpdate_Popup.AutoSize = $true

$Assigneduser_TextBox_Popup = New-Object system.Windows.Forms.TextBox
$Assigneduser_TextBox_Popup.multiline = $false
$Assigneduser_TextBox_Popup.Text = "Assigned User"
$Assigneduser_TextBox_Popup.Font = 'Segoe UI, 18pt'
$Assigneduser_TextBox_Popup.Backcolor = '#1b3666'
$Assigneduser_TextBox_Popup.ForeColor = '#a3a3a3' 
$Assigneduser_TextBox_Popup.Dock = 'Top'
$Assigneduser_TextBox_Popup.TabIndex = 1
$Assigneduser_TextBox_Popup.BorderStyle = 1
$Assigneduser_TextBox_Popup.Anchor = 'Left,Right'

$Status_Dropdown_Popup = New-Object System.Windows.Forms.ComboBox
$Status_Dropdown_Popup.DropDownStyle = 'DropDown'
$Status_Dropdown_Popup.Text = "Status"
$Status_Dropdown_Popup.Backcolor = '#1b3666'
$Status_Dropdown_Popup.ForeColor = '#eeeeee' 
$Status_Dropdown_Popup.AutoCompleteMode = 'SuggestAppend'
$Status_Dropdown_Popup.AutoCompleteSource = 'ListItems'
$Status_Dropdown_Popup.TabIndex = 2
$Status_Dropdown_Popup.Dock = "Fill"
$Status_Dropdown_Popup.FlatStyle = 0
$Status_Dropdown_Popup.Anchor = 'Top, Left, Right'
$Status_Dropdown_Popup.Font = 'Segoe UI, 18pt'

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
$AssetUpdate_Popup.AcceptButton.DialogResult = 'OK'

$Cancel_Button_Popup = New-Object system.Windows.Forms.Button
$Cancel_Button_Popup.Text = "Cancel"
$Cancel_Button_Popup.Font = 'Segoe UI, 18pt'
$Cancel_Button_Popup.Backcolor = '#1b3666'
$Cancel_Button_Popup.ForeColor = '#eeeeee' 
$Cancel_Button_Popup.Dock = 'Fill'
$Cancel_Button_Popup.TabIndex = 4
$Cancel_Button_Popup.FlatStyle = 1
$Cancel_Button_Popup.FlatAppearance.BorderSize = 0
$AssetUpdate_Popup.CancelButton = $Cancel_Button_Popup
$AssetUpdate_Popup.CancelButton.DialogResult = 'Cancel'

$LayoutPanel_Popup = New-Object System.Windows.Forms.TableLayoutPanel
$LayoutPanel_Popup.Dock = "Fill"
$LayoutPanel_Popup.ColumnCount = 4
$LayoutPanel_Popup.RowCount = 4
#$LayoutPanel_Popup.CellBorderStyle = 1
[void]$LayoutPanel_Popup.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel_Popup.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel_Popup.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel_Popup.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))

[void]$LayoutPanel_Popup.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 6)))
[void]$LayoutPanel_Popup.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 6)))
[void]$LayoutPanel_Popup.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 6)))
[void]$LayoutPanel_Popup.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 1)))

$LayoutPanel_Popup.Controls.Add($Assigneduser_TextBox_Popup, 1, 0)
$LayoutPanel_Popup.SetColumnSpan($Assigneduser_TextBox_Popup, 2)
$LayoutPanel_Popup.Controls.Add($Status_Dropdown_Popup, 1, 1)
$LayoutPanel_Popup.SetColumnSpan($Status_Dropdown_Popup, 2)
$LayoutPanel_Popup.Controls.Add($OK_Button_Popup, 1, 2)
$LayoutPanel_Popup.Controls.Add($Cancel_Button_Popup, 2, 2)
$AssetUpdate_Popup.controls.Add($LayoutPanel_Popup)
#EndRegion

#region Login Window
$Login_Form = New-Object system.Windows.Forms.Form
$Login_Form.Backcolor = '#324e7a'
$Login_Form.ForeColor = '#eeeeee' 
$Login_Form.FormBorderStyle = "FixedDialog"
$Login_Form.ClientSize = "400,220"
$Login_Form.TopMost = $true
$Login_Form.StartPosition = 'CenterScreen'
$Login_Form.ControlBox = $false
$Login_Form.AutoSize = $true

$Username_TextBox = New-Object system.Windows.Forms.TextBox
$Username_TextBox.multiline = $false
$Username_TextBox.Text = $env:USERNAME
$Username_TextBox.Font = 'Segoe UI, 18pt'
$Username_TextBox.Backcolor = '#1b3666'
$Username_TextBox.ForeColor = '#a3a3a3' 
$Username_TextBox.Dock = 'Top'
$Username_TextBox.TabIndex = 1
$Username_TextBox.BorderStyle = 1
$Username_TextBox.Anchor = 'Left,Right'

$Password_TextBox = New-Object system.Windows.Forms.TextBox
$Password_TextBox.multiline = $false
$Password_TextBox.Text = "PimaRocks"
$Password_TextBox.Font = 'Segoe UI, 18pt'
$Password_TextBox.Backcolor = '#1b3666'
$Password_TextBox.ForeColor = '#a3a3a3' 
$Password_TextBox.Dock = 'Top'
$Password_TextBox.TabIndex = 2
$Password_TextBox.BorderStyle = 1
$Password_TextBox.Anchor = 'Left,Right'
$Password_TextBox.PasswordChar = '*'
$Password_TextBox.Select()

$OK_Button_Login = New-Object system.Windows.Forms.Button
$OK_Button_Login.Text = "Login"
$OK_Button_Login.Backcolor = '#616161'
$OK_Button_Login.ForeColor = '#eeeeee' 
$OK_Button_Login.Dock = 'Fill'
$OK_Button_Login.TabIndex = 3
$OK_Button_Login.Font = 'Segoe UI, 18pt'
$OK_Button_Login.FlatStyle = 1
$OK_Button_Login.FlatAppearance.BorderSize = 0
$Login_Form.AcceptButton = $OK_Button_Login
$Login_Form.AcceptButton.DialogResult = 'OK'

$Cancel_Button_Login = New-Object system.Windows.Forms.Button
$Cancel_Button_Login.Text = "Cancel"
$Cancel_Button_Login.Font = 'Segoe UI, 18pt'
$Cancel_Button_Login.Backcolor = '#1b3666'
$Cancel_Button_Login.ForeColor = '#eeeeee' 
$Cancel_Button_Login.Dock = 'Fill'
$Cancel_Button_Login.TabIndex = 4
$Cancel_Button_Login.FlatStyle = 1
$Cancel_Button_Login.FlatAppearance.BorderSize = 0
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
# Copy above to main code^^^
# Code below for basic UI functions

#region Fun
$Inventoried = New-Object System.Media.SoundPlayer
$Inventoried.SoundLocation = "$PSScriptRoot\identify.wav"
#EndRegion

$screen = [System.Windows.Forms.Screen]::AllScreens
$Inventory = Start-SeFirefox -PrivateBrowsing -ImplicitWait 5 -Quiet
$Inventory.Manage().Window.Position = "0,0"
$Inventory.Manage().Window.Size = "$([math]::Round($screen[0].bounds.Width / 2.7)),$($screen[0].bounds.Height)"
$ITAM = Start-SeFirefox -PrivateBrowsing -ImplicitWait 5 -Quiet
$ITAM.Manage().Window.Position = "$($Inventory.Manage().Window.Size.Width - 12),0"
$ITAM.Manage().Window.Size = "$([math]::Round($screen[0].bounds.Width / 2.3)),$($screen[0].bounds.Height)"
$Inventory.Url = 'https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403'
$ITAM.Url = 'https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=402:26'

$Form.Location = "$($ITAM.Manage().Window.Size.Width + $ITAM.Manage().Window.Position.X - 15),100"
$AssetUpdate_Popup.Location = "$($Form.Location.X),$($PCC_Textbox.Location.Y)"

@('1', '2', '3', '4', '5', '6', '7', '8') | ForEach-Object { [void] $Campus_Dropdown.Items.Add($_) }
@('1', '2', '3', '4', '5') | ForEach-Object { [void] $Status_Dropdown_Popup.Items.Add($_) }

$Campus_Dropdown.add_SelectedIndexChanged( {
        $Room_Dropdown.Enabled = $false
        $Room_Dropdown.Text = 'Select Room'
        $Room_Dropdown.Items.Clear()
        @('1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15') | ForEach-Object { [void] $Room_Dropdown.Items.Add($_) }
        $Room_Dropdown.Enabled = $true
    })

$Search_Button.Add_MouseUp( {
        $AssetUpdate_Popup.ShowDialog()
        $Inventoried.playsync()
    })

$Option_Button.Add_MouseUp( {
        $Option_Popup.ShowDialog()
    })
$ScanLog_Button.Add_MouseUp( {
        $Option_Popup.DialogResult = 'OK'
    })
$ErrorLog_Button.Add_MouseUp( {
        $Option_Popup.DialogResult = 'OK'
    })
    
$OK_Button_Popup.Add_MouseUp( {
        $AssetUpdate_Popup.Close()
    })

$PCC_TextBox.Add_MouseDown( {
        $PCC_TextBox.clear()
        $PCC_TextBox.Forecolor = '#eeeeee'
    })

$Username_TextBox.Add_MouseDown( {
        $Username_TextBox.clear()
        $Username_TextBox.Forecolor = '#eeeeee'
    })

$Password_TextBox.Add_MouseDown( {
        $Password_TextBox.clear()
        $Password_TextBox.Forecolor = '#eeeeee'
    })
$OK_Button_Login.Add_MouseUp( {
        $Login_Form.DialogResult = 'OK'
    })
$Cancel_Button_Login.Add_MouseUp( {
        $Login_Form.DialogResult = 'Cancel'
    })

[void]$Login_Form.ShowDialog()

if ($Login_Form.DialogResult -eq 'OK') {
    [void]$Form.ShowDialog()
}
elseif ($Login_Form.DialogResult -eq 'Cancel') {

}
Stop-SeDriver $ITAM
Stop-SeDriver $Inventory

