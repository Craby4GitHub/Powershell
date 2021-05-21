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
$Form.Text = 'Inventory Helper Beta 0.3.1'
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

#region Functions

function Login-PimaSite ([Object[]]$Site) {

    Write-Log -Message "$($Credentials.UserName) attempting to login"
    
    try {
        $usernameElement = $Site.FindElementById('P101_USERNAME')
        $passwordElement = $Site.FindElementById('P101_PASSWORD')

        $usernameElement.Clear()
        $passwordElement.Clear()

        Send-SeKeys -Element $usernameElement -Keys $Credentials.UserName
        Send-SeKeys -Element $passwordElement -Keys $Credentials.GetNetworkCredential().Password
        $Site.FindElementById('P101_LOGIN').Click()
    }
    catch {
        Write-Log -Message "Had an issue with the login element on the site" -LogError $_.Exception.Message -Level FATAL
        exit
    }
}

Function Find-Asset($PCCNumber) {
    try {
        
        $InventoryTableAssets = $Inventory.FindElementByClassName('uReportStandard')

        if ($PCCNumber -match '\d{6}') {
            $PCCNumbers = $InventoryTableAssets.FindElementsByXPath("//td[@headers='WAITAMBAST_BARCODE']")
            for ($i = 0; $i -lt $PCCNumbers.Count; $i++) {
                if ($PCCNumbers[$i].text -eq $PCCNumber) {
                    return $i + 1
                    break
                }
            }    
        }
        else {
            $SerialNumbers = $InventoryTableAssets.FindElementsByXPath("//td[@headers='WAITAMBAST_SERIAL_NBR']")
            for ($i = 0; $i -lt $SerialNumbers.Count; $i++) {
                if ($SerialNumbers[$i].text -eq $PCCNumber) {
                    return $i + 1
                    break
                }
            } 
        }
    }
    catch {
        Write-Log -Message "Unable to get Inventory Table element for $($PCCNumber)." -LogError $_.Exception.Message -Level ERROR
    }
}

function Confirm-UIInput($UIInput, $RegEx, $ErrorMSG) {
    switch -regex ($UIInput.ToString()) {
        'System.Windows.Forms.TextBox' {  
            if ($UIInput.Text -match $RegEx) {
                $ErrorProvider.SetError($UIInput, '')
            }
            else {
                Write-Log -Message $ErrorMSG -Control $UIInput.Text
                $ErrorProvider.SetError($UIInput, $ErrorMSG)
                return $false
            }
        }
        'System.Windows.Forms.ComboBox' {
            if ($UIInput.Items -contains $UIInput.Text) {
                $ErrorProvider.SetError($UIInput, '')
                return $true
            }
            else {
                Write-Log -Message 'Invalid Dropdown Selection' -Control $UIInput.Text -Level ERROR
                $ErrorProvider.SetError($UIInput, $ErrorMSG)
                return $false
            }
        }
        Default {}
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
Function Write-Log {
    [CmdletBinding()]
    Param(
        [ValidateSet("INFO", "WARN", "ERROR", "FATAL", "DEBUG")]
        [String]
        $Level = "INFO",
        $Message,
        $LogError,
        $Control
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss:ff")
    $Line = "$Stamp,$Level,$($Credentials.UserName),$Control,$Message,$LogError"

    Add-Content $PSScriptRoot\ITAMScan_Errorlog.csv -Value $Line
}
#endregion

#region UI Actions

$Username_TextBox.Add_MouseDown( {
        $Username_TextBox.Clear()
        $Username_TextBox.Forecolor = '#eeeeee'
    })
    
$Password_TextBox.Add_MouseDown( {
        $Password_TextBox.clear()
        $Password_TextBox.Forecolor = '#eeeeee'
    })

$Search_Button.Add_MouseDown( {
        Confirm-UIInput -UIInput $Campus_Dropdown -ErrorMSG 'Invalid Campus'
        Confirm-UIInput -UIInput $Room_Dropdown -ErrorMSG 'Invalid Room'
        Confirm-UIInput -UIInput $PCC_TextBox -ErrorMSG 'Invalid Search Term' -RegEx '\w'
    })

$Search_Button.Add_MouseUp( {
        if (Confirm-NoError) {
            $StatusBar.Text = "Searching $($PCC_TextBox.Text) in $($Room_Dropdown.SelectedItem)"
            Write-Log -Message "Searching $($PCC_TextBox.Text) at $($Campus_Dropdown.SelectedItem): $($Room_Dropdown.SelectedItem)"

            $AssetIndex = Find-Asset $PCC_TextBox.Text
            if ($AssetIndex) {
                $StatusBar.Text = "$($PCC_TextBox.Text) Found!"
                try {

                    # Click Verify radio button
                    $Inventory.FindElementById("f02_$('{0:d4}' -f $AssetIndex)_0001").Click()

                    # Click Submit button
                    #$Inventory.FindElementById('B3258732422858420').Click()
                    $Inventory.ExecuteScript("apex.submit('SUBMIT')")
                    $Inventoried.playsync()
                            
                    $StatusBar.Text = "$($PCC_TextBox.Text) inventoried to $($Campus_Dropdown.SelectedItem): $($Room_Dropdown.SelectedItem)"
                    Write-Log -message "$($PCC_TextBox.Text) inventoried to $($Campus_Dropdown.SelectedItem): $($Room_Dropdown.SelectedItem)"
                    Add-Content $PSScriptRoot\ITAMScan_Scanlog.csv -Value "$($PCC_TextBox.Text),$($Campus_Dropdown.SelectedItem),$($Room_Dropdown.SelectedItem)"
                }
                catch {
                    Write-Log -Message 'Could not find/click Verify/Submit' -LogError $_.Exception.Message -Level ERROR
                }
            }
            else {
                $StatusBar.Text = "Unable to find $($PCC_TextBox.Text) in $($Room_Dropdown.SelectedItem), opening ITAM to edit"
                Write-Log -Message "Unable to find $($PCC_TextBox.Text) in $($Room_Dropdown.SelectedItem) at $($Campus_Dropdown.SelectedItem)"
        
                try {
                    # Click Magnifying glass button
                    $ITAM.FindElementById('R3070613760137337_column_search_root').Click()
        
                    if ($PCC_TextBox.Text -match '^\d{6}$') {
                        $ITAM.FindElementById('R3070613760137337_column_search_drop_2_c1i').Click()
                    }
                    else {
                        # Click Serial Number option
                        $ITAM.FindElementById('R3070613760137337_column_search_drop_2_c5i').Click()
                    }
            
                    # Entering PCCNumber into search bar
                    $ITAM.FindElementById('R3070613760137337_search_field').SendKeys($PCC_TextBox.Text)
        
                    # Click Go button
                    $ITAM.FindElementById('R3070613760137337_search_button').Click()
                }
                catch {
                    Write-Log -Message "Had an issue navigating ITAM to search for $($PCC_TextBox.Text)" -Level ERROR
                }
            
                try {
                    Write-Log -Message "Clicking edit for asset for $($PCC_TextBox.Text)"
                    #NOTE: Only clicks on the first entry, may need to load whole table in future to verify only 1 asset found
                    # Or find a better way of finding the asset in itam
                    $ITAM.FindElementByXPath('/html/body/form/div[5]/table/tbody/tr/td[1]/div/div[2]/div/div/div/div/div[2]/div[2]/div[6]/div[1]/table/tbody/tr[2]/td[1]/a').Click()
                }
                catch {
                    Write-Log -Message "Could not find or click edit option for $($PCC_TextBox.Text)" -Level ERROR
                    Add-Content $PSScriptRoot\ITAMScan_Scanlog.csv -Value "$($PCC_TextBox.Text),$($Campus_Dropdown.SelectedItem),$($Room_Dropdown.SelectedItem),'Not in ITAM'"
                    $StatusBar.Text = "Could not find $($PCC_TextBox.Text) in ITAM, saved data to log..."
                    
                    #Remove filter
                    $ITAM.FindElementByClassName('a-IRR-button--remove').Click()
                    
                    $PCC_TextBox.Clear()
                    $PCC_TextBox.Select()

                    return
                }
        
                try {
                    $AssetStatus_Element = $ITAM.FindElementById('P27_WAITAMBAST_STATUS')
                    $AssetStatusOptions_Element = Get-SeSelectionOption -Element $AssetStatus_Element -ListOptionText
                    $Status_Dropdown_Popup.Items.Clear()
                    $Status_Dropdown_Popup.Items.AddRange($AssetStatusOptions_Element)
                    $Status_Dropdown_Popup.Text = $AssetStatus_Element.getattribute('value')
        
                    $AssetAssignedUser_Element = $ITAM.FindElementById('P27_WAITAMBAST_ASSIGNED_USER')
                    $Assigneduser_TextBox_Popup.Text = $AssetAssignedUser_Element.getattribute('value')
                }
                catch {
                    Write-Log -Message 'Could not load asset information for Inventory Helper from ITAM' -LogError $_.Exception.Message -Level ERROR -Control $Status_Dropdown_Popup.SelectedItem
                    return
                }
                
                [void]$AssetUpdate_Popup.ShowDialog()
                if ($AssetUpdate_Popup.DialogResult -eq 'OK') {
        
                    Write-Log -Message "Updating $($PCC_TextBox.Text) to Campus: $($Campus_Dropdown.SelectedItem ) and Room: $($Room_Dropdown.SelectedItem)"
            
                    try {
                        # Clearing room on ITAM and enter room from UI
                        $AssetRoom_Element = $ITAM.FindElementById('P27_WAITAMBAST_ROOM')
                        $AssetRoom_Element.Clear()
                        Send-SeKeys -Element $AssetRoom_Element -Keys $Room_Dropdown.SelectedItem
        
                        # Setting status on ITAM from Inventory Helper
                        $ITAM.FindElementById('P27_WAITAMBAST_STATUS') | Get-SeSelectionOption -ByValue $Status_Dropdown_Popup.Text
        
                        Write-Log -Message "Setting Assigned User for $($PCC_TextBox.Text) in ITAM and entering: $($Assigneduser_TextBox_Popup.Text)"
                        $AssetAssignedUser_Element.Clear()
                        Send-SeKeys -Element $AssetAssignedUser_Element -Keys $Assigneduser_TextBox_Popup.Text
        
                        # Selecting ITAM campus from Inventory Helper
                        $ITAM.FindElementById('P27_WAITAMBAST_LOCATION') | Get-SeSelectionOption -ByValue $Campus_Dropdown.SelectedItem
        
                        # Clicking Apply Changes button
                        $ITAM.ExecuteScript("apex.submit('SAVE')")

                        $ITAM.Url = ("https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=402:26:$($ITAM.FindElementById('pInstance').getattribute('value')):::::")

                        #Remove filter
                        $ITAM.FindElementByClassName('a-IRR-button--remove').Click()
                    }
                    catch {
                        Write-Log -Message "Had issue updating $($PCC_TextBox.Text) to Campus: $($Campus_Dropdown.SelectedItem ) and Room: $($Room_Dropdown.SelectedItem)"  -LogError $_.Exception.Message -Level ERROR
                    }
                    
                    try {
                        $Inventory.Navigate().refresh()
                    }
                    catch {
                        Write-Log -Message "Could not refresh Inventory site after updating $($PCC_TextBox.Text)" -LogError $_.Exception.Message -Level FATAL
                    }

                    $Inventory.Navigate().refresh()
                    $AssetIndex = Find-Asset $PCC_TextBox.Text
                    if ($AssetIndex) {
                        $StatusBar.Text = "$($PCC_TextBox.Text) Found!"
                        try {
                            Write-Log -Message 'Clicking Verify and Submit button'
                            $Inventory.FindElementById("f02_$('{0:d4}' -f $AssetIndex)_0001").Click()

                            # Click Submit button
                            #$Inventory.FindElementById('B3258732422858420').Click()
                            $Inventory.ExecuteScript("apex.submit('SUBMIT')")
                            $Inventoried.playsync()
                            $StatusBar.Text = "$($PCC_TextBox.Text) has been inventoried to $($Campus_Dropdown.SelectedItem): $($Room_Dropdown.SelectedItem)"
                            Write-Log -message "$($PCC_TextBox.Text) has been inventoried to $($Campus_Dropdown.SelectedItem): $($Room_Dropdown.SelectedItem)"
                            Add-Content $PSScriptRoot\ITAMScan_Scanlog.csv -Value "$($PCC_TextBox.Text),$($Campus_Dropdown.SelectedItem),$($Room_Dropdown.SelectedItem)"
                        }
                        catch {
                            Write-Log -Message 'Could not find/click Verify/Submit' -LogError $_.Exception.Message -Level ERROR
                        }
                    }
                    else {
                        Write-Log -Message "Update to $($PCC_TextBox.Text) $($Campus_Dropdown.SelectedItem) : $($Room_Dropdown.SelectedItem) attempted in ITAM but not showing in Inventory"
                    }           
                }
                elseif ($AssetUpdate_Popup.DialogResult -eq 'Cancel') {
                    Write-Log -Message "Canceling Asset update for $($PCC_TextBox.Text) to Campus:$($Campus_Dropdown.SelectedItem ) and Room:$($Room_Dropdown.SelectedItem)"
    
                    $AssetUpdate_Popup.Close()
                    $ITAM.ExecuteScript("apex.navigation.redirect('f?p=402:26:$($ITAM.FindElementById('pInstance').getattribute('value'))::NO:::')")

                    #Remove filter
                    $ITAM.FindElementByClassName('a-IRR-button--remove').Click()
                    $StatusBar.Text = 'Ready'
                }
            }
            $PCC_TextBox.Clear()
            $PCC_TextBox.Select()
        }
    })

$Campus_Dropdown.add_SelectedIndexChanged( {
        $Room_Dropdown.Enabled = $false
        try {
            $Inventory.FindElementById('P1_WAITAMBAST_LOCATION') | Get-SeSelectionOption -ByValue $Campus_Dropdown.SelectedItem
            $RoomDropDown_Element = ($Inventory.FindElementById('P1_WAITAMBAST_ROOM')).text.split("`n").Trim()
            $Room_Dropdown.Text = 'Select Room'
            $Room_Dropdown.Items.Clear()
            $Room_Dropdown.Items.AddRange($RoomDropDown_Element)
            $Room_Dropdown.Enabled = $true
        }
        catch {
            Write-Log -Message 'Could not load campus/room for Inventory Helper dropdowns' -LogError $_.Exception.Message -Level ERROR
        }
    })

$Room_Dropdown.add_SelectedIndexChanged( {
        try {
            $Inventory.FindElementById('P1_WAITAMBAST_ROOM') | Get-SeSelectionOption -ByValue $Room_Dropdown.SelectedItem
        }
        catch {
            Write-Log -Message 'Could not load room for Inventory Helper dropdowns' -Level ERROR
        }
    })

$PCC_TextBox.Add_MouseDown( {
        $PCC_TextBox.Clear()
        $PCC_TextBox.Forecolor = '#eeeeee'
    })

$Assigneduser_TextBox_Popup.Add_MouseDown( {
        $Assigneduser_TextBox_Popup.Forecolor = '#eeeeee'
    })

$Option_Button.Add_MouseUp( {
        $Option_Popup.ShowDialog()
    })
$ScanLog_Button.Add_MouseUp( {
        $Option_Popup.DialogResult = 'OK'
        Invoke-Item $PSScriptRoot\ITAMScan_Scanlog.csv
    })
$ErrorLog_Button.Add_MouseUp( {
        $Option_Popup.DialogResult = 'OK'
        Invoke-Item $PSScriptRoot\ITAMScan_Errorlog.csv
    })
#endregion

#region Fun
$Inventoried = New-Object System.Media.SoundPlayer
$Inventoried.SoundLocation = "$PSScriptRoot\identify.wav"
#EndRegion

#region Main
$screen = [System.Windows.Forms.Screen]::AllScreens
$Inventory = Start-SeFirefox -PrivateBrowsing -ImplicitWait 5 -Quiet
$Inventory.Manage().Window.Position = "0,0"
$Inventory.Manage().Window.Size = "$([math]::Round($screen[0].bounds.Width / 2.7)),$($screen[0].bounds.Height)"
$ITAM = Start-SeFirefox -PrivateBrowsing -ImplicitWait 5 -Quiet
$ITAM.Manage().Window.Position = "$($Inventory.Manage().Window.Size.Width - 12),0"
$ITAM.Manage().Window.Size = "$([math]::Round($screen[0].bounds.Width / 2.3)),$($screen[0].bounds.Height)"
$Inventory.Url = 'https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403'
$ITAM.Url = 'https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=402:26'

$Form.Location = "$($ITAM.Manage().Window.Size.Width + $ITAM.Manage().Window.Position.X - 15),0"
$AssetUpdate_Popup.Location = "$($Form.Location.X),$($PCC_Textbox.Location.Y)"

[void]$Login_Form.ShowDialog()
if ($Login_Form.DialogResult -eq 'OK') {
    $Password = ConvertTo-SecureString $Password_TextBox.Text -AsPlainText -Force
    $global:Credentials = New-Object System.Management.Automation.PSCredential ($Username_TextBox.text, $Password)

    Login-PimaSite $ITAM
    Login-PimaSite $Inventory
try {
    $test = $ITAM.FindElementById('welcome')
    $test2 = $Inventory.FindElementByClassName('userBlock')
}
catch {
    
}
    

    if ( $test -and $test2) {
        try {
            $Inventory.ExecuteScript("apex.widget.tabular.paginate('R3257120268858381',{min:1,max:2000,fetched:2000})")
            $CampusDropDown_Element = ($Inventory.FindElementById('P1_WAITAMBAST_LOCATION')).text.split("`n").Trim()
            $Campus_Dropdown.Items.AddRange($CampusDropDown_Element)

            [void]$Form.ShowDialog()

            Write-Log -Message "Ending session for $($Credentials.UserName)"
            Stop-SeDriver $ITAM
            Stop-SeDriver $Inventory
            break
        }
        catch {
            Write-Log -Message "Could not verify login of $($Credentials.UserName)" -LogError $_.Exception.Message -Level FATAL
            exit
        }
    }
    else {
        $RelogChoice = [System.Windows.Forms.MessageBox]::Show("Login Failed, please relaunch.", 'Warning', 'RetryCancel', 'Warning')
        switch ($RelogChoice) {
            # Need to figure out logic to re-enter creds and what not
            'Retry' { 
                [void]$Login_Form.ShowDialog() 
                break
            }
            'Cancel' {
                Stop-SeDriver $ITAM
                Stop-SeDriver $Inventory
                Write-Log -Message "$($Credentials.UserName) failed to login" -Level WARN
                exit
            }
        }        
    }
}
elseif ($Login_Form.DialogResult -eq 'Cancel') {
    Stop-SeDriver $ITAM
    Stop-SeDriver $Inventory
}
#endregion