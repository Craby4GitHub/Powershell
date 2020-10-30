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
$panel.SetColumnSpan($Submit_Button,2)

$Form.controls.Add($panel)

#endregion

#region Functions

function Login_ITAM {

    #do {
        $Credentials = Get-Credential

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
    #} until (!$LoginCheck.Enabled)
    
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
            #Click Verify
            (Find-SeElement -Driver $Driver -Id "f02_000$($i)_0001").click()
            #return $true
            break
        }
        else {
            #return $false
        }
    }
    
}
#endregion

#region UI Actions
$Search_Button.Add_Click( {
    $RoomDropDown = Find-SeElement -Driver $Driver -Id "P1_WAITAMBAST_ROOM"

    $RoomDropDownOptions = Get-SeSelectionOption -Element $RoomDropDown -ListOptionText
        foreach ($room in $RoomDropDownOptions) {
            if ($room -eq $Room_Dropdown.Text) {
                Get-SeSelectionOption -Element $RoomDropDown -ByPartialText $Room_Dropdown.Text
                break
            }else {
                #$ErrorProvider.SetError($Room_Dropdown, 'Enter Valid Room')
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
            Find-Asset -PCCNumber $PCC_TextBox.Text
        }
        else {
            for ($page = 0; $page -lt $PageDropdownOptions.Count; $page++) {
                $PageDropdown = Find-SeElement -Driver $Driver -Id "X01_3257120268858381"
                Get-SeSelectionOption -Element $PageDropdown -ByIndex $page
                Find-Asset -PCCNumber $PCC_TextBox.Text
            }
        }
        #Work on what happens when it doesnt find
        if ($null -eq (Find-Asset -PCCNumber $PCC_TextBox.Text)) {
            #add code to edit this object because it was not found in this room
            write-host 'ehhhhh'
        }

        $PCC_TextBox.Clear()
        $PCC_TextBox.Focused
    })
#endregion
#$ITAMAssests = Import-Csv "C:\Users\wrcrabtree\Downloads\assets.csv"

$Driver = Start-SeFirefox -PrivateBrowsing
Open-SeUrl -Driver $Driver -Url "https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403"

Login_ITAM

$LocationDropDown = Find-SeElement -Driver $Driver -Id "P1_WAITAMBAST_LOCATION"
Get-SeSelectionOption -Element $LocationDropDown -ByPartialText "West Campus"

$RoomDropDown = Find-SeElement -Driver $Driver -Id "P1_WAITAMBAST_ROOM"
$RoomDropDownOptions = Get-SeSelectionOption -Element $RoomDropDown -ListOptionText
$Room_Dropdown.Items.AddRange($RoomDropDownOptions)

#Submit Button
#(Find-SeElement -Driver $Driver -Id 'B3258732422858420').click()

#Edit object in ITAM Inventory
#https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403:2:15764768460589::NO:RP:P2_WAITAMBAST_SEQ:25271

#Reload page when no rows
#https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=403:1:7876177828604::NO:RP::

[void]$Form.ShowDialog()
$Driver.close()
$Driver.quit()