If (-not(Get-InstalledModule Selenium -ErrorAction silentlycontinue)) {
    Install-Module Selenium -Confirm:$False -Force -Scope CurrentUser
}

#region GUI
. (Join-Path $PSSCRIPTROOT "GUI.ps1")
#endregion

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
    foreach ($control in $Main_LayoutPanel.Controls) {
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

#region Sounds
$InventoriedSound = New-Object System.Media.SoundPlayer
$InventoriedSound.SoundLocation = "$PSScriptRoot\Sounds\identify.wav"

$ErrorSound = New-Object System.Media.SoundPlayer
$ErrorSound.SoundLocation = "$PSScriptRoot\Sounds\warcry1.wav"

$PortalCast = New-Object System.Media.SoundPlayer
$PortalCast.SoundLocation = "$PSScriptRoot\Sounds\portalcast.wav"

$PortalEnter = New-Object System.Media.SoundPlayer
$PortalEnter.SoundLocation = "$PSScriptRoot\Sounds\portalenter.wav"

$D2Button = New-Object System.Media.SoundPlayer
$D2Button.SoundLocation = "$PSScriptRoot\Sounds\button.wav"
#EndRegion

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
    $D2Button.play()
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
                    $InventoriedSound.play()
                            
                    $StatusBar.Text = "$($PCC_TextBox.Text) inventoried to $($Campus_Dropdown.SelectedItem): $($Room_Dropdown.SelectedItem)"
                    $StatusBar.Backcolor = 'Green'
                    Write-Log -message "$($PCC_TextBox.Text) inventoried to $($Campus_Dropdown.SelectedItem): $($Room_Dropdown.SelectedItem)"
                    Add-Content $PSScriptRoot\ITAMScan_Scanlog.csv -Value "$($PCC_TextBox.Text),$($Campus_Dropdown.SelectedItem),$($Room_Dropdown.SelectedItem)"
                }
                catch {
                    Write-Log -Message 'Could not find/click Verify/Submit' -LogError $_.Exception.Message -Level ERROR
                }
            }
            else {
                $ErrorSound.play()
                $StatusBar.Text = "Unable to find $($PCC_TextBox.Text) in $($Room_Dropdown.SelectedItem), opening ITAM to edit"
                $StatusBar.Backcolor = 'Orange'
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
                    Write-Log -Message "Clicking edit for asset $($PCC_TextBox.Text) in ITAM"
                    #NOTE: Only clicks on the first entry, may need to load whole table in future to verify only 1 asset found
                    # Or find a better way of finding the asset in itam
                    $ITAM.FindElementByXPath('/html/body/form/div[5]/table/tbody/tr/td[1]/div/div[2]/div/div/div/div/div[2]/div[2]/div[6]/div[1]/table/tbody/tr[2]/td[1]/a').Click()
                }
                catch {
                    Write-Log -Message "Could not find or click edit option for $($PCC_TextBox.Text)" -Level ERROR
                    Add-Content $PSScriptRoot\ITAMScan_Scanlog.csv -Value "$($PCC_TextBox.Text),$($Campus_Dropdown.SelectedItem),$($Room_Dropdown.SelectedItem),'Not in ITAM'"
                    $StatusBar.Text = "Could not find $($PCC_TextBox.Text) in ITAM, saved data to log..."
                    
                    #Remove filter
                    $ITAM.FindElementByCssSelector('button[title="Remove Filter"]').Click()
                                        
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
                        $ITAM.FindElementByCssSelector('button[title="Remove Filter"]').Click()
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
                            $InventoriedSound.play()
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
    
                    $ITAM.ExecuteScript("apex.navigation.redirect('f?p=402:26:$($ITAM.FindElementById('pInstance').getattribute('value'))::NO:::')")

                    #Remove filter
                    $ITAM.FindElementByClassName('icon-remove').Click()
                    $StatusBar.Text = 'Ready'
                }
            }
            $StatusBar.Backcolor = $Theme.StatusBar.Backcolor
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
        $D2Button.Play()
        $Option_Popup.ShowDialog()
    })
$ScanLog_Button.Add_MouseUp( {
        $D2Button.Play()
        $Option_Popup.DialogResult = 'OK'
        Invoke-Item $PSScriptRoot\ITAMScan_Scanlog.csv
    })
$ErrorLog_Button.Add_MouseUp( {
        $D2Button.Play()
        $Option_Popup.DialogResult = 'OK'
        Invoke-Item $PSScriptRoot\ITAMScan_Errorlog.csv
    })
#endregion

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
$Option_Popup.Location = "$($Form.Location.X),$($PCC_Textbox.Location.Y)"

$PortalCast.play()
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
    catch {}
   
    if ( $test -and $test2) {
        try {
            $Inventory.ExecuteScript("apex.widget.tabular.paginate('R3257120268858381',{min:1,max:2000,fetched:2000})")
            $CampusDropDown_Element = ($Inventory.FindElementById('P1_WAITAMBAST_LOCATION')).text.split("`n").Trim()
            $Campus_Dropdown.Items.AddRange($CampusDropDown_Element)

            $PortalEnter.play()
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