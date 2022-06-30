. (Join-Path $PSSCRIPTROOT "JAMF-API.ps1")
$iPads = Get-Content ".\ipad.csv"
foreach ($ipad in $iPads) {
    try {
        $jamfID = Search-JamfMobileDevices -serialNumber $ipad
    }
    catch {
        Write-Host "Could not find $ipad" 
    }
    try {
        Post-JamfMobileDeviceCommandLostMode -deviceID $jamfID.general.id -message "This device has been checked out from the PCC library and is overdue. Please go to pima.edu/library to find a library nearest you to renew or turn in this device. If you have questions, please call the number below:" -phone "520-206-4900"

    }
    catch {
        Write-Host "Could not lock $ipad" 
    }
}