. (Join-Path $PSSCRIPTROOT "Get-JamfComputers.ps1")
$computers = Get-JamfComputers
write-host $computers