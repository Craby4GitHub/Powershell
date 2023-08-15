<#
.SYNOPSIS
    Updates the settings of the Avaya one-X Agent application.

.DESCRIPTION
    This script modifies the Config.xml and Settings.xml files to update the phone extension,
    ACD login code, and optionally enable autologin for the Avaya one-X Agent application.
    Then sets the Aux Reason codes.
#>

function Get-FourDigitCode {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $message
    )
    do {
        $inputCode = Read-Host $message
        if ($inputCode -match '\d{4}') {
            return $inputCode
        }
        else {
            Write-Host "Invalid, enter a 4 digit code"
        }
    } while ($true)
}
function Set-OneXAgentConfiguration {
    [CmdletBinding()]
    param ()

    $oneXagentPath = "Avaya\one-X Agent\2.5"

    # Check if required XML files and directories exist
    $configXmlPath = "$env:APPDATA\$oneXagentPath\Config.xml"
    $settingsXmlPath = "$env:APPDATA\$oneXagentPath\Profiles\default\Settings.xml"

    # Copy default Aux Codes
    Copy-Item -Path "$PSScriptRoot\Config.xml" -Destination "$env:APPDATA\$oneXagentPath"
    Copy-Item -Path "$PSScriptRoot\Settings.xml" -Destination "$env:APPDATA\$oneXagentPath\Profiles\default"

    if (-not (Test-Path $configXmlPath) -or -not (Test-Path $settingsXmlPath)) {
        Write-Warning "Required XML files or directories not found."
        return
    }

    # Load Default Config xml file
    [xml]$configXmlFile = Get-Content -Path $configXmlPath

    # Get extension and set
    $phoneExtension = Get-FourDigitCode -message "Enter your 4 digit phone number"
    $sipUserAccount = $configXmlFile.ConfigData.parameter | Where-Object { $_.name -eq 'SipUserAccount' }
    $sipUserAccount.value = [string]$phoneExtension

    # Save Config xml file
    $configXmlFile.Save($configXmlPath)

    # Load Default Settings xml file
    [xml]$settingsXmlFile = Get-Content -Path $settingsXmlPath
 
    # Get acd/Extension and set
    $acdLogin = Get-FourDigitCode -message "Enter your 4 digit ACD login code"
    $settingsXmlFile.Settings.login.Agent.Login = $acdLogin.ToString()
    $settingsXmlFile.Settings.login.Telephony.User.Station = $phoneExtension.ToString()

    # Autologin, optional
    do {
        $autologinChoice = Read-Host "Do you want to enable autologin? (yes/no)"
        if ($autologinChoice.ToLower() -eq 'yes') {
            $settingsXmlFile.Settings.login.Agent.AutoLogin = 'true'
            $settingsXmlFile.Settings.login.Telephony.User.AutoLogin = 'true'
            break
        }
        elseif ($autologinChoice.ToLower() -eq 'no') {
            break
        }
        else {
            Write-Warning "Invalid input. Please enter 'yes' or 'no'."
        }
    } while ($true)

    # Save Settings xml file
    $settingsXmlFile.Save($settingsXmlPath)
    
    # Copy default Aux Codes
    Copy-Item -Path "$PSScriptRoot\AuxReasonCodes.xml" -Destination "$env:APPDATA\$oneXagentPath\Profiles\default"
}

# Execute the function to update the settings
Set-OneXAgentConfiguration