function Get-JamfAuthClassic {
    $Creds = Get-Credential
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($Creds.UserName):$($Creds.GetNetworkCredential().Password)")))
    return $headers
}
function Get-JamfAuthPro {
    $Creds = Get-Credential
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($Creds.UserName):$($Creds.GetNetworkCredential().Password)")))
    $token = Invoke-RestMethod "https://pccjamf.jamfcloud.com/api/v1/auth/token" -Method 'POST' -Headers $Headers -ContentType application/json
    $BearerHeader = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $BearerHeader.Add("Authorization", "Bearer " + $token.token)
    return $BearerHeader
}

function Get-JamfComputers {
    $Headers = Get-JamfAuthClassic
    return (Invoke-RestMethod "https://pccjamf.jamfcloud.com/JSSResource/computers" -Method 'GET' -Headers $Headers -ContentType application/json).computers.computer.Name
}

function Get-JamfMobileDeviceEnrollmentProfiles {
    $Headers = Get-JamfAuthClassic
    return (Invoke-RestMethod "https://pccjamf.jamfcloud.com/JSSResource/mobiledeviceenrollmentprofiles" -Method 'GET' -Headers $Headers -ContentType application/json).mobile_device_enrollment_profiles.mobile_device_enrollment_profile
}

function Get-JamfMobileDevicePreStage {
    $auth = Get-JamfAuthPro
    return (Invoke-RestMethod "https://pccjamf.jamfcloud.com/api/v2/mobile-device-prestages?page=0&page-size=1000&sort=id%3Adesc" -Method 'GET' -Headers $auth -ContentType application/json).results
}

function Get-JamfMobileDevicePreStageScopeByID($ID) {
    $auth = Get-JamfAuthPro
    return (Invoke-RestMethod "https://pccjamf.jamfcloud.com/api/v2/mobile-device-prestages/$ID/scope" -Method 'GET' -Headers $auth -ContentType application/json).assignments
}

function Get-JamfMobileDevicePreStageScope {
    $auth = Get-JamfAuthPro
    return Invoke-RestMethod "https://pccjamf.jamfcloud.com/api/v2/mobile-device-prestages/scope" -Method 'GET' -Headers $auth -ContentType application/json
}