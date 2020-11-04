# https://www.powershellgallery.com/packages/Selenium/3.0.0

#Requires -Modules Selenium
#Install-Module -Name Selenium -RequiredVersion 3.0.0

#region UI
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

$Form = New-Object system.Windows.Forms.Form
$Form.FormBorderStyle = "FixedDialog"
$Form.ClientSize = "450,250"
$Form.text = "ITAM - Inventory Automation"
$Form.TopMost = $true
$Form.StartPosition = 'CenterScreen'
$Form.BackColor = '#546b94'
$Form.ForeColor = '#ffffff'

#$PimaIcon = New-Object System.Drawing.Icon ("$PSScriptRoot\favicon.ico")
#$Form.Icon = $PimaIcon

$Campus_Dropdown = New-Object System.Windows.Forms.ComboBox
$Campus_Dropdown.DropDownStyle = 'DropDown'
$Campus_Dropdown.DropDownHeight = $Campus_Dropdown.ItemHeight * 5
$Campus_Dropdown.Text = 'Select Campus'
$Campus_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Campus_Dropdown.AutoCompleteSource = 'ListItems'
$Campus_Dropdown.TabIndex = 1
$Campus_Dropdown.Dock = "Bottom"

$Room_Dropdown = New-Object System.Windows.Forms.ComboBox
$Room_Dropdown.DropDownStyle = 'DropDown'
$Room_Dropdown.DropDownHeight = $Room_Dropdown.ItemHeight * 5
$Room_Dropdown.Text = 'Select Room'
$Room_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Room_Dropdown.AutoCompleteSource = 'ListItems'
$Room_Dropdown.TabIndex = 2
$Room_Dropdown.Dock = "Bottom"

$PCC_Label = New-Object system.Windows.Forms.Label
$PCC_Label.text = "PCC Number:"
$PCC_Label.Font = 'Segoe UI, 10pt, style=Bold'
$PCC_Label.AutoSize = $true
$PCC_Label.Dock = 'Bottom'


$PCC_TextBox = New-Object system.Windows.Forms.TextBox
$PCC_TextBox.multiline = $false
$PCC_TextBox.Dock = 'Top'
$PCC_TextBox.TabIndex = 3

$Search_Button = New-Object system.Windows.Forms.Button
$Search_Button.text = "Search"
$Search_Button.Dock = 'Top'
$Search_Button.TabIndex = 4
$Search_Button.BackColor = '#a1adc4'
$Search_Button.Font = 'Microsoft Sans Serif, 8pt, style=Bold'
$Form.AcceptButton = $Search_Button

$StatusBar = New-Object System.Windows.Forms.StatusBar
$StatusBar.Text = "Ready"

#Region Panel
$panel = New-Object System.Windows.Forms.TableLayoutPanel
$panel.Dock = "Fill"
$panel.ColumnCount = 10
$panel.RowCount = 10
$panel.CellBorderStyle = 1
[void]$panel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))

$panel.Controls.Add($Campus_Dropdown, 2, 5)
$panel.SetColumnSpan($Campus_Dropdown, 3)
$panel.Controls.Add($Room_Dropdown, 5, 5)
$panel.SetColumnSpan($Room_Dropdown, 3)
$panel.Controls.Add($PCC_Label, 2, 7)
$panel.SetColumnSpan($PCC_Label, 3)
$panel.Controls.Add($PCC_TextBox, 2, 8)
$panel.SetColumnSpan($PCC_TextBox, 3)
$panel.Controls.Add($Search_Button, 5, 8)
$panel.SetColumnSpan($Search_Button, 3)

$Form.controls.AddRange(@($panel,$StatusBar))
#EndRegion 
#region Assest Update Popup
$AssetUpdate_Popup = New-Object system.Windows.Forms.Form
$AssetUpdate_Popup.Text = 'Asset Update'
$AssetUpdate_Popup.FormBorderStyle = "FixedDialog"
$AssetUpdate_Popup.ClientSize = "250,150"
$AssetUpdate_Popup.TopMost = $true
$AssetUpdate_Popup.StartPosition = 'CenterScreen'
$AssetUpdate_Popup.ControlBox = $false

$Assigneduser_TextBoxLabel_Popup = New-Object system.Windows.Forms.Label
$Assigneduser_TextBoxLabel_Popup.text = "Assigned User"
$Assigneduser_TextBoxLabel_Popup.AutoSize = $true
$Assigneduser_TextBoxLabel_Popup.Dock = 'Bottom'

$Assigneduser_TextBox_Popup = New-Object system.Windows.Forms.TextBox
$Assigneduser_TextBox_Popup.multiline = $false
$Assigneduser_TextBox_Popup.Dock = 'Top'
$Assigneduser_TextBox_Popup.TabIndex = 1

$Status_DropdownLabel_Popup = New-Object system.Windows.Forms.Label
$Status_DropdownLabel_Popup.text = "Status"
$Status_DropdownLabel_Popup.AutoSize = $true
$Status_DropdownLabel_Popup.Dock = 'Bottom'

$Status_Dropdown_Popup = New-Object System.Windows.Forms.ComboBox
$Status_Dropdown_Popup.DropDownStyle = 'DropDown'
$Status_Dropdown_Popup.AutoCompleteMode = 'SuggestAppend'
$Status_Dropdown_Popup.AutoCompleteSource = 'ListItems'
$Status_Dropdown_Popup.TabIndex = 2
$StatusList = @('Active', 'Disposed', 'On Hand', 'Missing', 'Stolen')
$Status_Dropdown_Popup.Items.AddRange(($StatusList | ForEach-Object { $($_) }))

$OK_Button_Popup = New-Object system.Windows.Forms.Button
$OK_Button_Popup.text = "OK"
$OK_Button_Popup.Dock = 'Bottom'
$OK_Button_Popup.TabIndex = 4

$Cancel_Button_Popup = New-Object system.Windows.Forms.Button
$Cancel_Button_Popup.text = "Cancel"
$Cancel_Button_Popup.Dock = 'Bottom'
$Cancel_Button_Popup.TabIndex = 4

$panelpopup = New-Object System.Windows.Forms.TableLayoutPanel
$panelpopup.Dock = "Fill"
$panelpopup.ColumnCount = 2
$panelpopup.RowCount = 3
$panelpopup.CellBorderStyle = 1
[void]$panelpopup.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$panelpopup.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$panelpopup.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$panelpopup.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$panelpopup.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20)))

$panelpopup.Controls.Add($Assigneduser_TextBox_Popup, 1, 1)
$panelpopup.Controls.Add($Status_Dropdown_Popup, 0, 1)
$panelpopup.Controls.Add($Assigneduser_TextBoxLabel_Popup, 1, 0)
$panelpopup.Controls.Add($Status_DropdownLabel_Popup, 0, 0)
$panelpopup.Controls.Add($OK_Button_Popup, 0, 2)
$panelpopup.Controls.Add($Cancel_Button_Popup, 1, 2)

$AssetUpdate_Popup.controls.Add($panelpopup)
#endregion

#endregion

#region Functions

function Login_ITAM {
    param (
        [bool]$FirstLogin
    )

    #do {
    if ($FirstLogin) {
        $global:Credentials = Get-Credential
    }
    

    $usernameElement = Find-SeElement -Driver $Driver -Wait -Timeout 10 -Id 'P101_USERNAME'
    $passwordElement = Find-SeElement -Driver $Driver -Id 'P101_PASSWORD'

    $usernameElement.Clear()
    $passwordElement.Clear()
    
    Send-SeKeys -Element $usernameElement -Keys $Credentials.UserName
    Send-SeKeys -Element $passwordElement -Keys $Credentials.GetNetworkCredential().Password
    Find-SeElement -Driver $Driver -ID 'P101_LOGIN' | Invoke-SeClick
    <#
        try {
            $LoginTimeOut = Find-SeElement -Driver $Driver -Id 'apex_login_throttle_sec'
        }
        catch { 
        }
        if ($LoginTimeOut.Text -gt 0) {
            Start-Sleep -Seconds $LoginTimeOut.Text
        } #>
    #} until (!$LoginTimeOut.Enabled)
    
}
Function Find-Asset {
    param (
        $PCCNumber
    )

    $InventoryTable = Find-SeElement -Driver $driver -XPath '/html/body/form/div/table/tbody/tr/td[1]/section[2]/div[2]/div/table/tbody[2]/tr/td/table/tbody'
    $InventoryTableAssests = $InventoryTable.FindElementsByTagName('tr')
    $PCCNumberFront_xpath = '/html/body/form/div/table/tbody/tr/td[1]/section[2]/div[2]/div/table/tbody[2]/tr/td/table/tbody/tr['
    $PCCNumberBack_xpath = ']/td[2]'

    for ($i = 1; $i -le $InventoryTableAssests.Count; $i++) {
        if ($InventoryTable.FindElementByXPath($PCCNumberFront_xpath + $i + $PCCNumberBack_xpath).text -eq $pccnumber) {
            return $i
            break
        }
    }
}

function Update-Asset {
    param (
        $PCCNumber,
        $RoomNumber,
        $Campus
    )
    $ITAMAsset = $ITAMAssests | Where-Object -Property 'Barcode #' -eq $PCCNumber
    if ($null -ne $ITAMAsset) {
        Open-SeUrl -Driver $Driver -Url "https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403:2:15764768460589::NO:RP:P2_WAITAMBAST_SEQ:$($ITAMAsset.'IT #')"
        Login_ITAM
        $AssetStatus_Element = Find-SeElement -Driver $Driver -Id "P2_WAITAMBAST_STATUS"
        $Status_Dropdown_Popup.Text = $AssetStatus_Element.getattribute('value')

        $AssetAssignedUser_Element = Find-SeElement -Driver $Driver -Id "P2_WAITAMBAST_ASSIGNED_USER"
        $Assigneduser_TextBox_Popup.Text = $AssetAssignedUser_Element.getattribute('value')

        $AssetAssignedLocation_Element = Find-SeElement -Driver $Driver -Id "P2_WAITAMBAST_LOCATION"

        $Status_Dropdown_Popup.Text = $AssetStatus_Element.Text

        $AssetRoom_Element = Find-SeElement -Driver $Driver -Id 'P2_WAITAMBAST_ROOM'

        $OK_Button_Popup.Add_MouseUp( {
                
                $AssetRoom_Element.Clear()
                Send-SeKeys -Element $AssetRoom_Element -Keys $RoomNumber

                $AssetStatus_Element.Clear()
                Send-SeKeys -Element $AssetStatus_Element -Keys $Status_Dropdown_Popup.Text

                $AssetAssignedUser_Element.Clear()
                Send-SeKeys -Element $AssetAssignedUser_Element -Keys $Assigneduser_TextBox_Popup.Text

                Get-SeSelectionOption -Element $AssetAssignedLocation_Element -ByValue $Campus

                #Enable when code is ready to deploy
                #Find-SeElement -Driver $Driver -Id 'B3263727731989509' | Invoke-SeClick

                Open-SeUrl -Driver $Driver -Url "https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403"
                Login_ITAM
                $LocationDropDown_Element = Find-SeElement -Driver $Driver -Id "P1_WAITAMBAST_LOCATION" -Timeout 1
                Get-SeSelectionOption -Element $LocationDropDown_Element -ByValue $Campus

                $RoomDropDown_Element = Find-SeElement -Driver $Driver -Id "P1_WAITAMBAST_ROOM"
                Get-SeSelectionOption -Element $RoomDropDown_Element -ByValue $RoomNumber

                $AssetUpdate_Popup.Close()
            })

        $Cancel_Button_Popup.Add_MouseUp( {
                Open-SeUrl -Driver $Driver -Url "https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403"
                Login_ITAM
                $LocationDropDown_Element = Find-SeElement -Driver $Driver -Id "P1_WAITAMBAST_LOCATION" -Timeout 1
                Get-SeSelectionOption -Element $LocationDropDown_Element -ByValue $Campus

                $RoomDropDown_Element = Find-SeElement -Driver $Driver -Id "P1_WAITAMBAST_ROOM"
                Get-SeSelectionOption -Element $RoomDropDown_Element -ByValue $RoomNumber
                $AssetUpdate_Popup.Close()
            })
        [void]$AssetUpdate_Popup.ShowDialog()
    }
}

function Confirm-Dropdown($Dropdown, $ErrorMSG) {
    if ($Dropdown.Items -contains $Dropdown.Text) {
        $ErrorProvider.SetError($Dropdown, '')
        return $true
    }
    else {
        Write-Log -Level INFO -Message $ErrorMSG -Element $Dropdown.Text
        $ErrorProvider.SetError($Dropdown, $ErrorMSG)
        return $false
    }     
}

function Confirm-NoError {
    $i = 0
    foreach ($control in $panel.controls) {
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
        $Element
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp,$Level,$env:COMPUTERNAME,$Message,$Element"

    Add-Content $PSScriptRoot\log.csv -Value $Line
}

Function Get-File($filePath) {   
    try {
        $file = Import-Csv -Path $filePath
    }
    catch {
        Write-Log -Level 'FATAL' -Message $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message, 'Critical Issue', 'OK', 'Error')
        exit
    }
    return $file
}
#endregion

#region UI Actions

$Search_Button.Add_MouseUp( {
        Confirm-Dropdown -Dropdown $Room_Dropdown -ErrorMSG 'Invalid Room'
    })


$Search_Button.Add_MouseUp( {
        if (Confirm-NoError) {
            $ErrorProvider.Clear()

            $RoomDropDown_Element = Find-SeElement -Driver $Driver -Id "P1_WAITAMBAST_ROOM"

            $RoomDropDownOptions_Element = Get-SeSelectionOption -Element $RoomDropDown_Element -ListOptionText
            foreach ($room in $RoomDropDownOptions_Element) {
                if ($room -eq $Room_Dropdown.Text) {
                    Get-SeSelectionOption -Element $RoomDropDown_Element -ByPartialText $Room_Dropdown.Text
                    break
                }
            }

            try {
                $PageDropdown = Find-SeElement -Driver $Driver -Id "X01_3257120268858381"
                $PageDropdownOptions = Get-SeSelectionOption -Element $PageDropdown -ListOptionText               
            }
            catch {
                Write-Host 'No extra pages'
            }

            if ($null -eq $PageDropdown) {
                $AssestIndex = Find-Asset -PCCNumber $PCC_TextBox.Text
                if ($AssestIndex) {
                    #Click Verify
                    Find-SeElement -Driver $Driver -Id "f02_000$($AssestIndex)_0001" | Invoke-SeClick
                }
                else {
                    Update-Asset -PCCNumber $PCC_TextBox.Text -RoomNumber $Room_Dropdown.Text -Campus $Campus_Dropdown.SelectedItem
                } 
            }
            else {
                for ($page = 0; $page -lt $PageDropdownOptions.Count; $page++) {
                    $PageDropdown = Find-SeElement -Driver $Driver -Id "X01_3257120268858381"
                    Get-SeSelectionOption -Element $PageDropdown -ByIndex $page

                    $AssestIndex = Find-Asset -PCCNumber $PCC_TextBox.Text
                    if ($AssestIndex) {
                        #Click Verify
                        Find-SeElement -Driver $Driver -Id "f02_000$($AssestIndex)_0001" | Invoke-SeClick
                        break
                    }
                    if ($page -eq $PageDropdownOptions.Count - 1) {
                        Update-Asset -PCCNumber $PCC_TextBox.Text -RoomNumber $Room_Dropdown.Text -Campus $Campus_Dropdown.SelectedItem
                    }
                }
            }
            $PCC_TextBox.Clear()
            $PCC_TextBox.Focused
        }
    })

$Campus_Dropdown_SelectedIndexChanged = {
    $LocationDropDown_Element = Find-SeElement -Driver $Driver -Id "P1_WAITAMBAST_LOCATION"
    Get-SeSelectionOption -Element $LocationDropDown_Element -ByPartialText $Campus_Dropdown.SelectedItem
    $RoomDropDown_Element = Find-SeElement -Driver $Driver -Id "P1_WAITAMBAST_ROOM"
    $RoomDropDownOptions_Element = Get-SeSelectionOption -Element $RoomDropDown_Element -ListOptionText
    $Room_Dropdown.Items.AddRange($RoomDropDownOptions_Element)
}
$Campus_Dropdown.add_SelectedIndexChanged($Campus_Dropdown_SelectedIndexChanged)
#endregion

$ITAMAssests = Get-File -filePath 'C:\Users\wrcrabtree\Downloads\inventory_report.csv'

$Driver = Start-SeFirefox -PrivateBrowsing
Open-SeUrl -Driver $Driver -Url "https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403"

Login_ITAM -FirstLogin $true

$LocationDropDown_Element = Find-SeElement -Driver $Driver -Id "P1_WAITAMBAST_LOCATION"
$LocationDropDownOptions_Element = Get-SeSelectionOption -Element $LocationDropDown_Element -ListOptionText
$Campus_Dropdown.Items.AddRange($LocationDropDownOptions_Element)

#Submit Button
#Find-SeElement -Driver $Driver -Id 'B3258732422858420' | Invoke-SeClick

#Edit object in ITAM Inventory
#https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403:2:15764768460589::NO:RP:P2_WAITAMBAST_SEQ:25271

#Reload page when no rows
#https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403:1:7876177828604::NO:RP::

[void]$Form.ShowDialog()
#$Driver.close()
#$Driver.quit()