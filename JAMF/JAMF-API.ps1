# https://pccjamf.jamfcloud.com/api/doc/
# https://pccjamf.jamfcloud.com/classicapi/doc/

# Setting Log name
$logName = ($MyInvocation.MyCommand.Name -split '\.')[0] + ' log'
$logFile = "$PSScriptroot\$logName.csv"
. ((Get-Item $PSScriptRoot).Parent.FullName + '\Callable\Write-Log.ps1')

function Get-JamfAuthClassic {
    # Create header used in all JAMF Classic API calls
    $classicHeader = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $classicHeader.Add("Authorization", "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($jamfCredentials.UserName):$($jamfCredentials.GetNetworkCredential().Password)")))
    return $classicHeader
}
function Get-JamfAuthPro {
    # Create header and recieving a bearer token used in all JAMF Pro API calls
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($jamfCredentials.UserName):$($jamfCredentials.GetNetworkCredential().Password)")))
    $token = Invoke-RestMethod "https://pccjamf.jamfcloud.com/api/v1/auth/token" -Method 'POST' -Headers $Headers -ContentType application/json
    $proHeader = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $proHeader.Add("Authorization", "Bearer " + $token.token)
    return $proHeader
}

function Get-JamfAllComputers {
    return (Invoke-RestMethod -Method GET -Headers $jamfClassicHeader -Uri "https://pccjamf.jamfcloud.com/JSSResource/computers" -ContentType "application/json" -UseBasicParsing).computers.computer.Name
}

function Search-JamfComputers($serialNumber, $name) {
    # This is horrible logic... but it works(?)
    # Wishlist: Make better
    if ($null -eq $serialNumber) {
        try {
            return (Invoke-RestMethod -Method GET -Headers $jamfClassicHeader -Uri "https://pccjamf.jamfcloud.com/JSSResource/computers/name/$name" -ContentType "application/json" -UseBasicParsing).Computer
        }
        catch {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Searching for $serialNumber $name in JAMF has failed. Status Code - $($_.Exception.Response.StatusCode.value__)"
        }
    }
    else {
        try {
            return (Invoke-RestMethod -Method GET -Headers $jamfClassicHeader -Uri "https://pccjamf.jamfcloud.com/JSSResource/computers/serialnumber/$serialNumber" -ContentType "application/json" -UseBasicParsing).Computer
        }
        catch {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Searching for $serialNumber $name in JAMF has failed. Status Code - $($_.Exception.Response.StatusCode.value__)"
        }
    }
}

function Search-JamfMobileDevices($serialNumber, $name) {
    # This is horrible logic... but it works(?)
    # Wishlist: Make better
    if ($null -eq $serialNumber) {
        try {
            return (Invoke-RestMethod -Method GET -Headers $jamfClassicHeader -Uri "https://pccjamf.jamfcloud.com/JSSResource/mobiledevices/name/$name" -ContentType "application/json" -UseBasicParsing).mobile_device
        }
        catch {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Searching for $serialNumber $name in JAMF has failed. Status Code - $($_.Exception.Response.StatusCode.value__)"
        }
    }
    else {
        try {
            return (Invoke-RestMethod -Method GET -Headers $jamfClassicHeader -Uri "https://pccjamf.jamfcloud.com/JSSResource/mobiledevices/serialnumber/$serialNumber" -ContentType "application/json" -UseBasicParsing).mobile_device
        }
        catch {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Searching for $serialNumber $name in JAMF has failed. Status Code - $($_.Exception.Response.StatusCode.value__)"
        }
    }
}

function Get-JamfMobileDeviceEnrollmentProfiles {
    return (Invoke-RestMethod "https://pccjamf.jamfcloud.com/JSSResource/mobiledeviceenrollmentprofiles" -Method 'GET' -Headers $jamfClassicHeader -ContentType application/json).mobile_device_enrollment_profiles.mobile_device_enrollment_profile
}

function Get-JamfLicensedSoftware {
    return (Invoke-RestMethod "https://pccjamf.jamfcloud.com/JSSResource/licensedsoftware" -Method 'GET' -Headers $jamfClassicHeader -ContentType application/json).licensed_software.licensed_software
}

function Find-JamfLicensedSoftware($SearchTerm) {
    # Find by name or ID
    if ($SearchTerm -match '^\d{1,}$') {
        return (Invoke-RestMethod "https://pccjamf.jamfcloud.com/JSSResource/licensedsoftware/id/$SearchTerm" -Method 'GET' -Headers $jamfClassicHeader -ContentType application/json).licensed_software.licenses.license
    }
    else {
        return (Invoke-RestMethod "https://pccjamf.jamfcloud.com/JSSResource/licensedsoftware/name/$SearchTerm" -Method 'GET' -Headers $jamfClassicHeader -ContentType application/json).licensed_software.licenses.license
    }
}

function Remove-JamfComputer($ID) {
    # You can DELETE using the resource URLs with parameters of {name}, {udid}, {serial number}, or {macaddress}.
    Invoke-RestMethod "https://pccjamf.jamfcloud.com/JSSResource/computers/id/$ID" -Method 'DELETE' -Headers $jamfClassicHeader -ContentType application/json
}

function Get-JamfMobileGroups {
    (Invoke-RestMethod "https://pccjamf.jamfcloud.com/JSSResource/mobiledevicegroups" -Method 'GET' -Headers $jamfClassicHeader -ContentType application/json).mobile_device_groups.mobile_device_group
}

function Get-JamfMobileGroupsByID($ID) {
    (Invoke-RestMethod "https://pccjamf.jamfcloud.com/JSSResource/mobiledevicegroups/id/$ID" -Method 'GET' -Headers $jamfClassicHeader -ContentType application/json).mobile_device_group
}

function Update-JamfMobileGroups($ID, [array]$SerialNumbers) {
    $body = @"
    <mobile_device_additions>key</mobile_device_additions>
"@
    Invoke-RestMethod "https://pccjamf.jamfcloud.com/JSSResource/mobiledevicegroups/id/$ID" -Method 'PUT' -Headers $jamfClassicHeader -Body $body -ContentType application/xml
}

function Get-JamfMobileDevicePreStage {
    return (Invoke-RestMethod "https://pccjamf.jamfcloud.com/api/v2/mobile-device-prestages?page=0&page-size=1000&sort=id%3Adesc" -Method 'GET' -Headers $jamfProHeader -ContentType application/json).results
}

function Get-JamfMobileDevicePreStageByID($ID) {
    return Invoke-RestMethod "https://pccjamf.jamfcloud.com/api/v2/mobile-device-prestages/$ID" -Method 'GET' -Headers $jamfProHeader -ContentType application/json
}
function Get-JamfMobileDevicePreStageScopeByID($ID) {
    return (Invoke-RestMethod "https://pccjamf.jamfcloud.com/api/v2/mobile-device-prestages/$ID/scope" -Method 'GET' -Headers $jamfProHeader -ContentType application/json).assignments.serialnumber
}

function Get-JamfMobileDevicePreStageScope {
    return (Invoke-RestMethod "https://pccjamf.jamfcloud.com/api/v2/mobile-device-prestages/scope" -Method 'GET' -Headers $jamfProHeader -ContentType application/json).serialsByPrestageId
}

function Update-JamfMobileDeviceFromPreStageScope($ID, [array]$SerialNumbers) {
    $versionLock = (Get-JamfMobileDevicePreStageByID -ID $ID).versionLock
    $params = @{
        "serialNumbers" = $SerialNumbers;
        "versionLock"   = $versionLock;
    } | ConvertTo-Json
    Invoke-RestMethod "https://pccjamf.jamfcloud.com/api/v2/mobile-device-prestages/$ID/scope" -Method 'PUT' -Headers $jamfProHeader -Body $params -ContentType application/json
}

function Remove-JamfMobileDeviceFromPreStageScope($ID, [array]$SerialNumbers) {
    $versionLock = (Get-JamfMobileDevicePreStageByID -ID $ID).versionLock
    $params = @{
        "serialNumbers" = $SerialNumbers;
        "versionLock"   = $versionLock;
    } | ConvertTo-Json
    Invoke-RestMethod "https://pccjamf.jamfcloud.com/api/v2/mobile-device-prestages/$ID/scope/delete-multiple" -Method 'POST' -Headers $jamfProHeader -Body $params -ContentType application/json
}

function Add-JamfMobileDeviceFromPreStageScope($ID, [array]$SerialNumbers) {
    $versionLock = (Get-JamfMobileDevicePreStageByID -ID $ID).versionLock
    $params = @{
        "serialNumbers" = $SerialNumbers;
        "versionLock"   = $versionLock;
    } | ConvertTo-Json
    Invoke-RestMethod "https://pccjamf.jamfcloud.com/api/v2/mobile-device-prestages/$ID/scope" -Method 'POST' -Headers $jamfProHeader -Body $params -ContentType application/json
}

$jamfCredentialFile = Get-Content $PSScriptRoot\JAMFCreds.json | ConvertFrom-Json

[SecureString]$securePassword = $jamfCredentialFile.password | ConvertTo-SecureString 

[PSCredential]$jamfCredentials = New-Object System.Management.Automation.PSCredential -ArgumentList $jamfCredentialFile.username, $securePassword

$jamfClassicHeader = Get-JamfAuthClassic
$jamfProHeader = Get-JamfAuthPro