# https://www.powershellgallery.com/packages/Selenium/3.0.0

#Requires -Modules Selenium
#Install-Module -Name Selenium -RequiredVersion 3.0.0

#region UI
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

$Form = New-Object system.Windows.Forms.Form
$Form.FormBorderStyle = "FixedDialog"
$Form.ClientSize = "450,350"
$Form.text = "Timesheet Automation"
$Form.TopMost = $true
$Form.StartPosition = 'CenterScreen'

$Room_Dropdown = New-Object System.Windows.Forms.ComboBox
$Room_Dropdown.DropDownStyle = 'DropDown'
$Room_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Room_Dropdown.AutoCompleteSource = 'ListItems'
$Room_Dropdown.TabIndex = 1

$PCC_Label = New-Object system.Windows.Forms.Label
$PCC_Label.text = "PCC Number:"
$PCC_Label.AutoSize = $true
$PCC_Label.Dock = 'Bottom'

$PCC_TextBox = New-Object system.Windows.Forms.TextBox
$PCC_TextBox.multiline = $false
$PCC_TextBox.Dock = 'Bottom'
$PCC_TextBox.TabIndex = 2

$Search_Button = New-Object system.Windows.Forms.Button
$Search_Button.text = "Search"
$Search_Button.Dock = 'Bottom'
$Search_Button.TabIndex = 3
$Form.AcceptButton = $Search_Button

$panel = New-Object System.Windows.Forms.TableLayoutPanel
$panel.Dock = "Fill"
$panel.ColumnCount = 2
$panel.RowCount = 4
$panel.CellBorderStyle = 1
[void]$panel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$panel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33)))

$panel.Controls.Add($Room_Dropdown, 0, 0)
$panel.Controls.Add($PCC_Label, 0, 1)
$panel.Controls.Add($PCC_TextBox, 0, 2)
$panel.Controls.Add($Search_Button, 0, 3)
$panel.SetColumnSpan($Submit_Button, 2)

$Form.controls.Add($panel)

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
function Find-Asset {
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
                    write-host $PCC_TextBox.Text 'Not found on single page'
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
                        write-host $PCC_TextBox.Text 'Not found on multi page'
                        $ITAMAsset = $ITAMAssests | Where-Object -Property 'Barcode #' -eq $PCC_TextBox.Text
                        if ($null -ne $ITAMAsset) {
                            Open-SeUrl -Driver $Driver -Url "https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403:2:15764768460589::NO:RP:P2_WAITAMBAST_SEQ:$($ITAMAsset.'IT #')"
                            Login_ITAM

                            $ChangeRoom = [System.Windows.Forms.MessageBox]::Show("Update $($PCC_TextBox.Text) to room: $($Room_Dropdown.Text)", 'Warning', 'OKCancel', 'Warning')
                            switch ($ChangeRoom) {
                                "OK" {
                                    $AssetRoom_Element = Find-SeElement -Driver $Driver -Id 'P2_WAITAMBAST_ROOM'

                                    $AssetRoom_Element.Clear()
            
                                    Send-SeKeys -Element $AssetRoom_Element -Keys $Room_Dropdown.Text

                                    Find-SeElement -Driver $Driver -Id 'B3263727731989509' | Invoke-SeClick
                                } 
                                "Cancel" {
                                    Open-SeUrl -Driver $Driver -Url "https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403"
                                    Login_ITAM
                                } 
                            }



                        } 
                    }
                }
            }
            #Work on what happens when it doesnt find
   
            #add code to edit this object because it was not found in this room
            #look up object initam to find its current room
            #Edit object in ITAM Inventory
            #https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403:2:15764768460589::NO:RP:P2_WAITAMBAST_SEQ:25271

            $PCC_TextBox.Clear()
            $PCC_TextBox.Focused
        }
    })
#endregion

$ITAMAssests = Get-File -filePath 'C:\Users\wrcrabtree\Downloads\inventory_report.csv'

$Driver = Start-SeFirefox -PrivateBrowsing
Open-SeUrl -Driver $Driver -Url "https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403"

Login_ITAM -FirstLogin $true

$LocationDropDown = Find-SeElement -Driver $Driver -Id "P1_WAITAMBAST_LOCATION"
Get-SeSelectionOption -Element $LocationDropDown -ByPartialText "West Campus"

$RoomDropDown_Element = Find-SeElement -Driver $Driver -Id "P1_WAITAMBAST_ROOM"
$RoomDropDownOptions_Element = Get-SeSelectionOption -Element $RoomDropDown_Element -ListOptionText
$Room_Dropdown.Items.AddRange($RoomDropDownOptions_Element)

#Submit Button
#Find-SeElement -Driver $Driver -Id 'B3258732422858420' | Invoke-SeClick

#Edit object in ITAM Inventory
#https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403:2:15764768460589::NO:RP:P2_WAITAMBAST_SEQ:25271

#Reload page when no rows
#https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403:1:7876177828604::NO:RP::

[void]$Form.ShowDialog()
$Driver.close()
$Driver.quit()