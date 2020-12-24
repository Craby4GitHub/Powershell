. (Join-Path $PSSCRIPTROOT "JAMF-API.ps1")

#$SerialNumbers = @('GG7CR66SMF3Q', 'test')
$SerialNumbers = Import-Csv -Path 'C:\users\wrcrabtree\Desktop\sn.csv'
function Move-JAMFMobileDevicePreStage($SourceID, $TargetID, $Serials) {
    $sourcePreStage = Get-JamfMobileDevicePreStageScopeByID -ID $SourceID
    $targetPreStage = Get-JamfMobileDevicePreStageScopeByID -ID $TargetID
    
    foreach ($SerialNumber in $SerialNumbers) {
        if ($sourcePreStage.Contains($SerialNumber)) {
            $sourcePreStage = $sourcePreStage | Where-Object { $_ -ne $SerialNumber }
            $targetPreStage += $SerialNumber
        }
    }
    Update-JamfMobileDeviceFromPreStageScope -ID $SourceID -SerialNumbers $sourcePreStage
    Update-JamfMobileDeviceFromPreStageScope -ID $TargetID -SerialNumbers $targetPreStage
}

Move-JAMFMobileDevicePreStage -SourceID 6 -TargetID 7 -Serials $SerialNumbers