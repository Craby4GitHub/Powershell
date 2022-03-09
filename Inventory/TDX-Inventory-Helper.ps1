If (-not(Get-InstalledModule Selenium -ErrorAction silentlycontinue)) {
    Install-Module Selenium -Confirm:$False -Force -Scope CurrentUser
}

#$screen = [System.Windows.Forms.Screen]::AllScreens
$ITAM = Start-SeFirefox -PrivateBrowsing -ImplicitWait 5 -Quiet
$ITAM.Manage().Window.Position = "0,0"
#$ITAM.Manage().Window.Size = "$([math]::Round($screen[0].bounds.Width / 2.3)),$($screen[0].bounds.Height)"

# TDX Sites
$tdxDesktopURL = 'https://service.pima.edu/SBTDNext/Home/Desktop/Default.aspx'
$tdxAssetDesktopURL = 'https://service.pima.edu/SBTDNext/Apps/1258/Assets/Default.aspx'

$ITAM.Navigate().GoToURL($tdxDesktopURL)


$Credentials = Get-Credential

# Find login fields for later and remove any text currently in them
$usernameElement = $ITAM.FindElementById('username')
$passwordElement = $ITAM.FindElementById('password')
$usernameElement.Clear()
$passwordElement.Clear()

# Login to site
Send-SeKeys -Element $usernameElement -Keys $Credentials.UserName
Send-SeKeys -Element $passwordElement -Keys $Credentials.GetNetworkCredential().Password
$ITAM.FindElementByName('_eventId_proceed').click()

# Go to Asset application in TDX
# TODO: Site loads to quickly and is opened in a new window, fix so that it opens in the orginal window
Start-Sleep -Seconds 3
$ITAM.Navigate().GoToURL($tdxAssetDesktopURL)



function Search-TDXAsset($SearchTerm) {
    # Find search box and enter in search term(pccnumber or SN)
    $tdxAssetSearchBox = $itam.FindElementByName('txtAssetSerial')
    Send-SeKeys -Element $tdxAssetSearchBox -Keys $SearchTerm

    # Find and click magnifying glass to search
    $tdxAssetSearchBox = $itam.FindElementByID('btnAssetLookup')
    $tdxAssetSearchBox.Click()
    
}

function Update-TDXAsset {

    # Switch to Asset Detail page
    $itam.SwitchTo().Window($itam.WindowHandles[1]) | Out-Null
    # Go to the Update page for the asset. This if nifty as the only difference in the url are these 2 strings
    $ITAM.Navigate().GoToURL($ITAM.Url.Replace('AssetDet', 'Update'))

    # Select and clear the current inventory date
    $tdxLastInventoryDate = $itam.FindElementByID('attribute126172')
    $tdxLastInventoryDate.clear()
    $tdxLastInventoryDate.click()

    # Set Last Inventory Date to current date
    Send-SeKeys -Element $tdxLastInventoryDate -Keys $(get-date -Format "MM/dd/yyyy")

    $tdxSubmitAssetUpdate = $itam.FindElementByID('btnSubmit')
    $tdxSubmitAssetUpdate.click()

    # Future addition: Check settings were saved
    $tdxAssetUpdateSuccess = $itam.FindElementByClassName('alert-success')
    if ($null -ne $tdxAssetUpdateSuccess.Enabled) {
        Write-Host 'Success!'
        $itam.Close()
        $itam.SwitchTo().Window($itam.WindowHandles[0]) | Out-Null
    }

}
