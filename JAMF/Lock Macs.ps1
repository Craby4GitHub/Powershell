. (Join-Path $PSSCRIPTROOT "JAMF-API.ps1")
$macs = Get-Content ".\mac.csv"
foreach ($mac in $macs) {
    try {
        $jamfObject = Search-JamfComputers -serialNumber $mac
    }
    catch {
        Write-Host "Could not find $mac" 
    }
    try {
        Post-JamfComputerCommandDeviceLock -id $jamfObject.general.id -passcode ENTERPASSCODE
        Start-Sleep -Seconds 1
    }
    catch {
        Write-Host "Could not lock $mac" 
    }
}