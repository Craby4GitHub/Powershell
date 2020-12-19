. (Join-Path $PSSCRIPTROOT "JAMF-API.ps1")

$SerialNumbers = @('C07VX4EYHNM4', 'CCQNF17VG22V', 'DMPCTNH2MF3V', 'DMPWLX54JF8M')

$JAMFMobileDevices = Get-JamfMobileDevicePreStageScope
foreach ($SerialNumber in $SerialNumbers) {
    if (($JAMFMobileDevices | Select-Object -ExpandProperty $SerialNumber) -eq $PrestageID) {
        
    } 
}