$appdata = $env:APPDATA
$oneXagentPath = "Avaya\one-X Agent\2.5"
$profilePath = "Profiles\default"

function Get-4DigitCode($message) {
    $code = Read-Host $message
    if ($code -match '\d{4}') {
        return $code
    }
    else {
        Write-Host "Invalid, enter a 4 digit code"
        Get-4DigitCode -message $message
    }
}


# Load Default Config xml file
[xml]$configXML = Get-Content "$appdata\$oneXagentPath\Config.xml"

# Get extension and set
$phoneExtension = Get-4DigitCode -message "Enter 4 digit phone extension"
$sipUserAccount = $configXML.ConfigData.parameter | Where-Object name -eq SipUserAccount
$sipUserAccount.value = [string]$phoneExtension

# Save Config xml file
Copy-Item -Path $configXML -Destination "$appdata\$oneXagentPath"




# Load Default Settings xml file
[xml]$profileXML = Get-Content "$appdata\$oneXagentPath\$profilePath\Settings.xml"

# Get acd/Extension and set
$acdLogin = Get-4DigitCode -message "Enter 4 digit ACD login code"
$profileXML.Settings.login.Agent.Login = $acdLogin
$profileXML.Settings.login.Telephony.User.Station = $phoneExtension

# Autologin, optional
#$profileXML.Settings.login.Agent.AutoLogin = 'true'
#$profileXML.Settings.login.Telephony.User.AutoLogin = 'true'

# Save Settings xml file
Copy-Item -Path $profileXML -Destination "$appdata\$oneXagentPath\$profilePath"

# Copy default Aux Codes
Copy-Item -Path $PSScriptRoot\AuxReasonCodes.xml -Destination "$appdata\$oneXagentPath\$profilePath"