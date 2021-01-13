# https://www.powershellgallery.com/packages/Selenium/3.0.0

#Requires -Modules Selenium
#Install-Module -Name Selenium -RequiredVersion 3.0.0

#region UI
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

#https://colorhunt.co/palette/10792

$Form = New-Object system.Windows.Forms.Form
$Form.AutoScaleMode = 'Font'
$Form.StartPosition = 'CenterScreen'
$Form.text = "ITAM - Inventory Automation"
$Form.Font = 'Segoe UI, 18pt'
$Form.TopMost = $true
$Form.BackColor = '#303841'
$Form.ForeColor = '#eeeeee' 
$Form.FormBorderStyle = 'None'
$Form.ClientSize = New-Object System.Drawing.Point(378, 659)

$Campus_Dropdown = New-Object System.Windows.Forms.ComboBox
$Campus_Dropdown.DropDownStyle = 'DropDown'
#$Campus_Dropdown.DropDownHeight = $Campus_Dropdown.ItemHeight * 5
$Campus_Dropdown.ItemHeight = 3000
$Campus_Dropdown.Text = 'Select Campus'
#$Campus_Dropdown.Font = 'Segoe UI, 18pt'
$Campus_Dropdown.BackColor = '#3a4750'
$Campus_Dropdown.ForeColor = '#eeeeee' 
$Campus_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Campus_Dropdown.AutoCompleteSource = 'ListItems'
$Campus_Dropdown.TabIndex = 1
#$Campus_Dropdown.Dock = "Fill"
$Campus_Dropdown.FlatStyle = 0
$Campus_Dropdown.Anchor = 'left, Right'

$Room_Label = New-Object system.Windows.Forms.Label
$Room_Label.text = "Room:" 
#$Room_Label.Font = 'Segoe UI, 10pt, style=Bold'
$Room_Label.AutoSize = $true
$Room_Label.Dock = 'Bottom'
$Room_Label.Anchor = 'Left,Right,Bottom'

$Room_Dropdown = New-Object System.Windows.Forms.ComboBox
$Room_Dropdown.DropDownStyle = 'DropDown'
$Room_Dropdown.DropDownHeight = $Room_Dropdown.ItemHeight * 5
$Room_Dropdown.Text = 'Select Room'
#$Room_Dropdown.Font = 'Segoe UI, 18pt'
$Room_Dropdown.BackColor = '#3a4750'
$Room_Dropdown.ForeColor = '#eeeeee' 
$Room_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Room_Dropdown.AutoCompleteSource = 'ListItems'
$Room_Dropdown.TabIndex = 2
$Room_Dropdown.Dock = "Fill"
$Room_Dropdown.Enabled = $false
$Room_Dropdown.FlatStyle = 0
$Room_Dropdown.Anchor = 'Left,Right'

$PCC_Label = New-Object system.Windows.Forms.Label
$PCC_Label.text = "PCC Number:"
#$PCC_Label.Font = 'Segoe UI, 10pt, style=Bold'
$PCC_Label.AutoSize = $true
$PCC_Label.Dock = 'Bottom'
$PCC_Label.Anchor = 'Bottom'

$PCC_TextBox = New-Object system.Windows.Forms.TextBox
$PCC_TextBox.multiline = $false
#$PCC_TextBox.Font = 'Segoe UI, 15pt'
$PCC_TextBox.BackColor = '#3a4750'
$PCC_TextBox.ForeColor = '#eeeeee' 
$PCC_TextBox.Dock = 'Fill'
$PCC_TextBox.TabIndex = 3
$PCC_TextBox.BorderStyle = 1
$PCC_TextBox.Anchor = 'Left,Right'

$Search_Button = New-Object system.Windows.Forms.Button
$Search_Button.text = "Search"
$Search_Button.Dock = 'Fill'
$Search_Button.TabIndex = 4
$Search_Button.BackColor = '#00adb5'
$Search_Button.ForeColor = '#eeeeee' 
#$Search_Button.Font = 'Segoe UI, 18pt, style=Bold'
$Search_Button.FlatStyle = 1
$Search_Button.FlatAppearance.BorderSize = 0
$Form.AcceptButton = $Search_Button

$Close_Button = New-Object system.Windows.Forms.Button
$Close_Button.text = "X"
$Close_Button.Dock = 'Fill'
$Close_Button.BackColor = '#303841'
#$Close_Button.Font = 'Segoe UI, 8pt, style=Bold'
$Close_Button.FlatStyle = 0
$Close_Button.FlatAppearance.BorderSize = 1
$Close_Button.FlatAppearance.BorderColor = '#3a4750'

$StatusBar = New-Object System.Windows.Forms.StatusBar
$StatusBar.Text = "Ready"
$StatusBar.SizingGrip = $false
$StatusBar.Dock = 'Bottom'
$StatusBar.BackColor = '#3a4750'
#$StatusBar.ForeColor = '#eeeeee'
$StatusBar.Font = 'Segoe UI, 10pt, style=Bold'

#Region Panel
$LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$LayoutPanel.Dock = "Fill"
$LayoutPanel.ColumnCount = 3
$LayoutPanel.RowCount = 5
$LayoutPanel.CellBorderStyle = 1

[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 3)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))


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
$AssetUpdate_Popup.AutoSize = $true

$Status_DropdownLabel_Popup = New-Object system.Windows.Forms.Label
$Status_DropdownLabel_Popup.text = "Status"
$Status_DropdownLabel_Popup.Font = 'Segoe UI, 8pt'
$Status_DropdownLabel_Popup.ForeColor = '#eeeeee' 
$Status_DropdownLabel_Popup.AutoSize = $true
$Status_DropdownLabel_Popup.Dock = 'Bottom'

$Status_Dropdown_Popup = New-Object System.Windows.Forms.ComboBox
$Status_Dropdown_Popup.DropDownStyle = 'DropDown'
$Status_Dropdown_Popup.BackColor = '#3a4750'
$Status_Dropdown_Popup.ForeColor = '#eeeeee' 
$Status_Dropdown_Popup.AutoCompleteMode = 'SuggestAppend'
$Status_Dropdown_Popup.AutoCompleteSource = 'ListItems'
$Status_Dropdown_Popup.TabIndex = 1
$Status_Dropdown_Popup.Dock = "Fill"
$Status_Dropdown_Popup.FlatStyle = 0
$Status_Dropdown_Popup.Anchor = 'Left,Right'
$Status_Dropdown_Popup.Font = 'Segoe UI, 8pt'

$Assigneduser_TextBoxLabel_Popup = New-Object system.Windows.Forms.Label
$Assigneduser_TextBoxLabel_Popup.text = "Assigned User"
$Assigneduser_TextBoxLabel_Popup.Font = 'Segoe UI, 8pt'
$Assigneduser_TextBoxLabel_Popup.ForeColor = '#eeeeee' 
$Assigneduser_TextBoxLabel_Popup.AutoSize = $true
$Assigneduser_TextBoxLabel_Popup.Dock = 'Bottom'

$Assigneduser_TextBox_Popup = New-Object system.Windows.Forms.TextBox
$Assigneduser_TextBox_Popup.multiline = $false
$Assigneduser_TextBox_Popup.BackColor = '#3a4750'
$Assigneduser_TextBox_Popup.ForeColor = '#eeeeee' 
$Assigneduser_TextBox_Popup.Dock = 'Top'
$Assigneduser_TextBox_Popup.TabIndex = 2
$Assigneduser_TextBox_Popup.BorderStyle = 1
$Assigneduser_TextBox_Popup.Anchor = 'Left,Right'

$OK_Button_Popup = New-Object system.Windows.Forms.Button
$OK_Button_Popup.text = "OK"
$OK_Button_Popup.BackColor = '#00adb5'
$OK_Button_Popup.ForeColor = '#eeeeee' 
$OK_Button_Popup.Dock = 'Bottom'
$OK_Button_Popup.TabIndex = 3
$OK_Button_Popup.Font = 'Segoe UI, 10pt, style=Bold'
$OK_Button_Popup.FlatStyle = 1
$OK_Button_Popup.FlatAppearance.BorderSize = 0
$AssetUpdate_Popup.AcceptButton = $OK_Button_Popup

$Cancel_Button_Popup = New-Object system.Windows.Forms.Button
$Cancel_Button_Popup.text = "Cancel"
$Cancel_Button_Popup.Font = 'Segoe UI, 8pt, style=Bold'
$Cancel_Button_Popup.BackColor = '#3a4750'
$Cancel_Button_Popup.ForeColor = '#eeeeee' 
$Cancel_Button_Popup.Dock = 'Bottom'
$Cancel_Button_Popup.TabIndex = 4
$Cancel_Button_Popup.FlatStyle = 1
$Cancel_Button_Popup.FlatAppearance.BorderSize = 0

$LayoutPanel_Popup = New-Object System.Windows.Forms.TableLayoutPanel
$LayoutPanel_Popup.Dock = "Fill"
$LayoutPanel_Popup.ColumnCount = 2
$LayoutPanel_Popup.RowCount = 3
$LayoutPanel_Popup.CellBorderStyle = 1
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

$LayoutPanel.Controls.Add($Close_Button, 4, 0)
$LayoutPanel.Controls.Add($Campus_Dropdown, 1, 0)
#$LayoutPanel.Controls.Add($Room_Label, 3,0 )
$LayoutPanel.Controls.Add($Room_Dropdown, 1, 1)
$LayoutPanel.Controls.Add($PCC_Label, 1, 2)
#$LayoutPanel.Controls.Add($LayoutPanel_Popup, 1, 5)
$LayoutPanel.Controls.Add($PCC_TextBox, 1, 3)
#$LayoutPanel.SetColumnSpan($PCC_TextBox, 2)
$LayoutPanel.Controls.Add($Search_Button, 1, 4)
#$LayoutPanel.SetRowSpan($Search_Button, 2)
#$LayoutPanel.SetColumnSpan($Search_Button, 3)

$Form.controls.AddRange(@($LayoutPanel, $StatusBar))
#EndRegion 

#region Functions

function Login_ITAM {
    param (
        [bool]$FirstLogin
    )
    
    if ($FirstLogin) {
        $global:Credentials = Get-Credential
    }

    Write-Log -Message "$($Credentials.UserName) attempting to login"
    
    try {
        Write-Log -Message "Getting Username and Password elements"
        $usernameElement = Get-SeElement -Driver $Driver -Wait -Id 'P101_USERNAME'
        $passwordElement = Get-SeElement -Driver $Driver -Id 'P101_PASSWORD'
    }
    catch {
        Write-Log -Message "Unable to get Username and Password elements" -LogError $_.Exception.Message -Level FATAL
    }


    $usernameElement.Clear()
    $passwordElement.Clear()

    try {
        Write-Log -Message "Entering Username and Password into elements"
        Send-SeKeys -Element $usernameElement -Keys $Credentials.UserName
        Send-SeKeys -Element $passwordElement -Keys $Credentials.GetNetworkCredential().Password
        Get-SeElement -Driver $Driver -ID 'P101_LOGIN' | Invoke-SeClick
    }
    catch {
        Write-Log -Message "Could not enter credentials into website" -Level WARN
    }
         
    
}
Function Find-Asset {
    param (
        $PCCNumber,
        $Campus,
        $Room,
        $Page
    )

    try {
        Write-Log -Message "Getting Inventory Table element for $($PCCNumber). $($Campus): $($Room) Page: $($Page)"
        $InventoryTable = Get-SeElement -Driver $driver -XPath '/html/body/form/div/table/tbody/tr/td[1]/section[2]/div[2]/div/table/tbody[2]/tr/td/table/tbody'
        
        $InventoryTableAssests = $InventoryTable.FindElementsByTagName('tr')
        $PCCNumberFront_xpath = '/html/body/form/div/table/tbody/tr/td[1]/section[2]/div[2]/div/table/tbody[2]/tr/td/table/tbody/tr['
        $PCCNumberBack_xpath = ']/td[2]'
    
        for ($i = 1; $i -le $InventoryTableAssests.Count; $i++) {
            if ($InventoryTable.FindElementByXPath($PCCNumberFront_xpath + $i + $PCCNumberBack_xpath).text -eq $PCC_TextBox.Text) {
                return $i
                break
            }
        }
    }
    catch {
        Write-Log -Message "Unable to get Inventory Table element for $($PCCNumber). $($Campus): $($Room) Page: $($Page)" -LogError $_.Exception.Message -Level ERROR
    }
}
function Update-Asset {
    param (
        $PCCNumber,
        $RoomNumber,
        $Campus
    )
    Write-Log -Message "Updating $PCCNumber in ITAM"
    
    try {
        Write-Log -Message 'Opening ITAM site'
        Open-SeUrl -Driver $Driver -Url $ITAM_URL
        Login_ITAM
    }
    catch {
        Write-Log -Message 'Could not open or log into main ITAM site' -LogError $_.Exception.Message -Level FATAL
    }

    try {
        Write-Log -Message 'Click Magnifying glass button'
        Get-SeElement -Driver $Driver -Id 'R3070613760137337_column_search_root' | Invoke-SeClick
    }
    catch {
        Write-Log -Message 'Could not find or click Magnifying glass button' -LogError $_.Exception.Message -Level ERROR
    }
    
    try {
        Write-Log -Message 'Click Barcode option'
        Get-SeElement -Driver $Driver -Id 'R3070613760137337_column_search_drop_2_c1i' | Invoke-SeClick
    }
    catch {
        Write-Log -Message 'Could not find or click Barcode option' -LogError $_.Exception.Message -Level ERROR
    }

    try {
        Write-Log -Message "Entering $PCCNumber into search bar"
        Get-SeElement -Driver $Driver -Id 'R3070613760137337_search_field' | Send-SeKeys -Keys $PCCNumber
    }
    catch {
        Write-Log -Message "Could not enter $PCCNumber into search bar" -LogError $_.Exception.Message -Level ERROR
    }

    try {
        Write-Log -Message 'Clicking Go button'
        Get-SeElement -Driver $Driver -Id 'R3070613760137337_search_button' | Invoke-SeClick
    }
    catch {
        Write-Log -Message 'Could not click or find Go Button' -LogError $_.Exception.Message -Level ERROR
    }

    try {
        Write-Log -Message "Clicking edit for asset for $PCCNumber"
        #NOTE: Only clicks on the first entry, may need to load whole table in future to verify only 1 asset found
        # Also script will fail if it can not find the pcc number in ITAM
        Get-SeElement -Driver $Driver -xPath '/html/body/form/div[5]/table/tbody/tr/td[1]/div/div[2]/div/div/div/div/div[2]/div[2]/div[6]/div[1]/table/tbody/tr[2]/td[1]/a' | Invoke-SeClick
    }
    catch {
        Write-Log -Message 'Could not find or click edit option for asset' -LogError $_.Exception.Message -Level ERROR
        Add-Content $PSScriptRoot\ITAMScan_Scanlog.csv -Value "$PCCNumber,$Campus,$Room,'Not in ITAM'"
    }

    try {
        Write-Log -Message 'Populating UI status dropdown from ITAM and set UI selection to current asset status'
        $AssetStatus_Element = Get-SeElement -Driver $Driver -Id "P27_WAITAMBAST_STATUS"
        $AssetStatusOptions_Element = Get-SeSelectionOption -Element $AssetStatus_Element -ListOptionText
        $Status_Dropdown_Popup.Items.Clear()
        $Status_Dropdown_Popup.Items.AddRange($AssetStatusOptions_Element)
        $Status_Dropdown_Popup.Text = $AssetStatus_Element.getattribute('value')
    }
    catch {
        Write-Log -Message 'Could not find Asset Status element' -LogError $_.Exception.Message -Level ERROR -Control $Status_Dropdown_Popup.SelectedItem
    }

    try {
        Write-Log -Message 'Populating UI assigned User from ITAM and set UI selection to current assigned User'
        $AssetAssignedUser_Element = Get-SeElement -Driver $Driver -Id "P27_WAITAMBAST_ASSIGNED_USER"
        $Assigneduser_TextBox_Popup.Text = $AssetAssignedUser_Element.getattribute('value')
    }
    catch {
        Write-Log -Message 'Could not get Assets assigned user element' -LogError $_.Exception.Message -Level ERROR -Control $Assigneduser_TextBox_Popup.Text
    }
    $OK_Button_Popup.Add_MouseUp( {
        
            $Global:Cancelled = $false
            Write-Log -Message "Updating $($PCCNumber) to Campus: $($Campus) and Room: $($RoomNumber)"
      
            try {
                Write-Log -Message 'Clearing room on ITAM and enter room from UI'
                $AssetRoom_Element = Get-SeElement -Driver $Driver -Id 'P27_WAITAMBAST_ROOM'
                $AssetRoom_Element.Clear()
                Send-SeKeys -Element $AssetRoom_Element -Keys $Room_Dropdown.SelectedItem #$RoomNumber
            }
            catch {
                Write-Log -Message 'Could not get Asset room element or clear/enter room' -LogError $_.Exception.Message -Level ERROR
            }

            try {
                Write-Log -Message 'Selecting status on ITAM from UI data'
                $AssetStatus_Element = Get-SeElement -Driver $Driver -Id "P27_WAITAMBAST_STATUS"
                Get-SeSelectionOption -Element $AssetStatus_Element -ByValue $Status_Dropdown_Popup.Text
            }
            catch {
                Write-Log -Message 'Could not get Asset status element or set the status' -LogError $_.Exception.Message -Level ERROR
            }

            Write-Host $Assigneduser_TextBox_Popup.Text
            try {
                Write-Log -Message "Clearing Assigned User on ITAM and entering: $($Assigneduser_TextBox_Popup.Text)"
                $AssetAssignedUser_Element.Clear()
                Send-SeKeys -Element $AssetAssignedUser_Element -Keys $Assigneduser_TextBox_Popup.Text
            }
            catch {
                Write-Log -Message 'Could not clear or set Assigned user element' -LogError $_.Exception.Message -Level ERROR
            }

            try {
                Write-Log -Message 'Selecting ITAM campus from UI'
                $AssetAssignedLocation_Element = Get-SeElement -Driver $Driver -Id "P27_WAITAMBAST_LOCATION"
                Get-SeSelectionOption -Element $AssetAssignedLocation_Element -ByValue $Campus_Dropdown.SelectedItem
            }
            catch {
                Write-Log -Message 'Could not get assigned campus or set assigned campus' -LogError $_.Exception.Message -Level ERROR
            }

            try {
                Write-Log -Message 'Clicking Apply Changes button'
                Get-SeElement -Driver $Driver -xPath '/html/body/form/div[5]/table/tbody/tr/td[1]/div[1]/div[1]/div/div[2]/button[3]' | Invoke-SeClick
            }
            catch {
                Write-Log -Message 'Could not get or click Apply Changes button' -LogError $_.Exception.Message -Level ERROR
            }

            try {
                Write-Log -Message 'Opening Inventory site'
                Open-SeUrl -Driver $Driver -Url $Inventory_URL
                Login_ITAM
                $Driver.ExecuteScript("apex.widget.tabular.paginate('R3257120268858381',{min:1,max:10000,fetched:10000})")
            }
            catch {
                Write-Log -Message 'Could not open Inventory ITAM site' -LogError $_.Exception.Message -Level FATAL
            }

            try {
                Write-Log -Message 'Setting UI Campus and Room'
                $LocationDropDown_Element = Get-SeElement -Driver $Driver -Id "P1_WAITAMBAST_LOCATION" -Timeout 1
                Get-SeSelectionOption -Element $LocationDropDown_Element -ByValue $Campus_Dropdown.SelectedItem
                $RoomDropDown_Element = Get-SeElement -Driver $Driver -Id "P1_WAITAMBAST_ROOM"
                Get-SeSelectionOption -Element $RoomDropDown_Element -ByValue $Room_Dropdown.SelectedItem #$RoomNumber
            }
            catch {
                Write-Log -Message 'Issue with getting Campus/Room from site or setting UI campus/room' -LogError $_.Exception.Message -Level ERROR
            }
            $AssetUpdate_Popup.Close()
            $StatusBar.Text = 'Ready'
        })

    $Cancel_Button_Popup.Add_MouseUp( {
            Write-Log -Message "Canceling Asset update for $($PCCNumber) to Campus:$($Campus) and Room:$($RoomNumber)"

            try {
                Write-Log -Message 'Opening Inventory site'
                Open-SeUrl -Driver $Driver -Url $Inventory_URL
                Login_ITAM
                $Driver.ExecuteScript("apex.widget.tabular.paginate('R3257120268858381',{min:1,max:10000,fetched:10000})")
            }
            catch {
                Write-Log -Message 'Could not open Inventory ITAM site' -LogError $_.Exception.Message -Level FATAL
            }

            try {
                Write-Log -Message "Setting UI Campus and Room in UI"
                $LocationDropDown_Element = Get-SeElement -Driver $Driver -Id "P1_WAITAMBAST_LOCATION" -Timeout 1
                Get-SeSelectionOption -Element $LocationDropDown_Element -ByValue $Campus
                $RoomDropDown_Element = Get-SeElement -Driver $Driver -Id "P1_WAITAMBAST_ROOM"
                Get-SeSelectionOption -Element $RoomDropDown_Element -ByValue $RoomNumber
            }
            catch {
                Write-Log -Message "Issue with getting/setting Campus/Room for UI" -LogError $_.Exception.Message -Level ERROR
            }

            $Global:Cancelled = $true
            $AssetUpdate_Popup.Close()
            $StatusBar.Text = 'Ready'
        })
    [void]$AssetUpdate_Popup.ShowDialog()
}
function Confirm-Asset {
    param (
        $PCCNumber,
        $Campus,
        $Room
    )

    $StatusBar.Text = "Searching for $($PCC_TextBox.Text) in $($Room_Dropdown.SelectedItem)"
      
    $AssetIndex = Find-Asset $PCCNumber $Campus $Room $page
    if ($AssetIndex) {
        $StatusBar.Text = "$($PCC_TextBox.Text) Found!"
        try {
            Write-Log -Message 'Clicking Verify and Submit button'
            Get-SeElement -Driver $Driver -Id "f02_$('{0:d4}' -f $AssetIndex)_0001" | Invoke-SeClick
            Get-SeElement -Driver $Driver -Id 'B3258732422858420' | Invoke-SeClick

            $StatusBar.Text = "$($PCCNumber) has been inventoried to $($Campus): $($Room)"
            Write-Log -message "$($PCCNumber) has been inventoried to $($Campus): $($Room)"
            Add-Content $PSScriptRoot\ITAMScan_Scanlog.csv -Value "$PCCNumber,$Campus,$Room"
            $PCC_TextBox.Select()
            break
        }
        catch {
            Write-Log -Message 'Could not find/click Verify/Submit' -LogError $_.Exception.Message -Level ERROR
        }
    }

    else {
        #$StatusBar.Text = "Unable to find $PCCNumber in $Room, opening ITAM to edit"
        $StatusBar.Text = "Unable to find $PCCNumber in $Room, saving to file..."
        Write-Log -Message "Unable to find $($PCCNumber) in $($Room) at $($Campus)"
        Add-Content $PSScriptRoot\ITAMScan_Scanlog.csv -Value "$PCCNumber,$Campus,$Room,'Update'"
        $PCC_TextBox.Select()
        Update-Asset -PCCNumber $PCCNumber -RoomNumber $Room -Campus $Campus
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
                Write-Log -Message 'Invalid Dropdown Selection' -Control $UIInput.Text
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
        [Parameter(Mandatory = $False)]
        [ValidateSet("INFO", "WARN", "ERROR", "FATAL", "DEBUG")]
        [String]
        $Level = "INFO",

        [Parameter(Mandatory = $True)]
        [string]
        $Message,

        [Parameter(Mandatory = $false)]
        [string]
        $LogError,

        [Parameter(Mandatory = $false)]
        [string]
        $Control
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp,$Level,$($Credentials.UserName),$Control,$Message,$LogError"

    Add-Content $PSScriptRoot\ITAMScan_Errorlog.csv -Value $Line
}
#endregion

#region UI Actions

$Close_Button.Add_MouseUp( {
        $Form.Close()
    })

$Search_Button.Add_MouseDown( {
        Confirm-UIInput -UIInput $Campus_Dropdown -ErrorMSG 'Invalid Campus'
        Confirm-UIInput -UIInput $Room_Dropdown -ErrorMSG 'Invalid Room'
        Confirm-UIInput -UIInput $PCC_TextBox -RegEx '^\d{6}$' -ErrorMSG 'Invalid PCC Number'
    })

$Search_Button.Add_MouseUp( {
        if (Confirm-NoError) {
            $StatusBar.Text = 'Starting Search...'

            try {
                Write-Log -Message "Getting Room Dropdown element options"
                $RoomDropDown_Element = Get-SeElement -Driver $Driver -Id "P1_WAITAMBAST_ROOM"
                $RoomDropDownOptions_Element = Get-SeSelectionOption -Element $RoomDropDown_Element -ListOptionText
            }
            catch {
                Write-Log -Message 'Could not load room dropdown' -LogError $_.Exception.Message -Level ERROR
            }
            
            foreach ($room in $RoomDropDownOptions_Element) {
                if ($room -eq $Room_Dropdown.Text) {
                    try {
                        Write-Log -Message "Selecting room from dropdown"
                        Get-SeSelectionOption -Element $RoomDropDown_Element -ByValue $Room_Dropdown.Text
                    }
                    catch {
                        Write-Log -Message "Could not select room: $($Room_Dropdown.Text) from dropdown" -LogError $_.Exception.Message -Level ERROR
                    }
                    break
                }
            }

            try {
                Write-Log -Message 'Reloading Inventory page so data is always updated'
                #Open-SeUrl -Driver $Driver -Url ($Inventory_URL + ":1:$($loginInstance)::NO:RP::")
            }
            catch {
                Write-Log -Message 'Could not reload Inventory Page' -LogError $_.Exception.Message -Level ERROR
            }
            
            Confirm-Asset -PCCNumber $PCC_TextBox.Text -Campus $Campus_Dropdown.SelectedItem -Room $Room_Dropdown.SelectedItem

            $PCC_TextBox.Clear()
            $PCC_TextBox.Focused
        }
    })

$Campus_Dropdown_SelectedIndexChanged = {
    $Room_Dropdown.Enabled = $false
    try {
        $LocationDropDown_Element = Get-SeElement -Driver $Driver -Id "P1_WAITAMBAST_LOCATION"
        Get-SeSelectionOption -Element $LocationDropDown_Element -ByValue $Campus_Dropdown.SelectedItem
        $RoomDropDown_Element = (Get-SeElement -Driver $Driver -Id "P1_WAITAMBAST_ROOM").text.split("`n")
        #$RoomDropDownOptions_Element = Get-SeSelectionOption -Element $RoomDropDown_Element -ListOptionText
    }
    catch {
        Write-Log -Message 'Could not load campus/room for UI dropdowns' -LogError $_.Exception.Message -Level ERROR
    }

    $Room_Dropdown.Text = 'Select Room'
    $Room_Dropdown.Items.Clear()
    $Room_Dropdown.Items.AddRange($RoomDropDown_Element)
    $Room_Dropdown.Enabled = $true
}
$Campus_Dropdown.add_SelectedIndexChanged($Campus_Dropdown_SelectedIndexChanged)

#region Window Drag
$global:dragging = $false
$global:mouseDragX = 0
$global:mouseDragY = 0

$LayoutPanel.Add_MouseDown( { 
        $global:dragging = $true
        $global:mouseDragX = [System.Windows.Forms.Cursor]::Position.X - $Form.Left
        $global:mouseDragY = [System.Windows.Forms.Cursor]::Position.Y - $Form.Top
    })

$LayoutPanel.Add_MouseMove( { if ($global:dragging) {
            $screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
            $currentX = [System.Windows.Forms.Cursor]::Position.X
            $currentY = [System.Windows.Forms.Cursor]::Position.Y
            [int]$newX = [Math]::Min($currentX - $global:mouseDragX, $screen.Right - $Form.Width)
            [int]$newY = [Math]::Min($currentY - $global:mouseDragY, $screen.Bottom - $Form.Height)
            $form.Location = New-Object System.Drawing.Point($newX, $newY)
        } })

$LayoutPanel.Add_MouseUp( { $global:dragging = $false })
#endregion

#endregion

#region Main

$ITAM_URL = 'https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=402:26'
$Inventory_URL = 'https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403'

$Driver = Start-SeFirefox -PrivateBrowsing #-Headless

Open-SeUrl -Driver $Driver -Url $Inventory_URL

$global:Credentials = $null
Login_ITAM -FirstLogin $true

$Global:loginInstance = (Get-SeElement -Driver $Driver -Id 'pInstance').getattribute('value')
$Global:Cancelled = $false

try {
    Write-Log -Message 'Loading campus options for UI on first attempt'
    $LocationDropDown_Element = Get-SeElement -Driver $Driver -Id "P1_WAITAMBAST_LOCATION"
    $LocationDropDownOptions_Element = Get-SeSelectionOption -Element $LocationDropDown_Element -ListOptionText
}
catch {
    Write-Log -Message 'Could not load campus for UI on first attempt' -LogError $_.Exception.Message -Level ERROR
}
$Campus_Dropdown.Items.AddRange($LocationDropDownOptions_Element)
$Driver.ExecuteScript("apex.widget.tabular.paginate('R3257120268858381',{min:1,max:10000,fetched:10000})")
[void]$Form.ShowDialog()

Write-Log -Message "Ending session for $($Credentials.UserName)"

$Driver.close()
$Driver.quit()

#endregion