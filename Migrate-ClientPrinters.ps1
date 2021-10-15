$oldServer = 'wc-vm-prtsvr'
$newServer = 'WC-Print'

$printers = Get-printer -Name '\\wc-vm-prtsvr*'

ForEach ($printer in $printers) {
    $newprinter = $printer.name -replace $oldServer, $newServer

    Write-Host "Removing $($printer.name)"
    Remove-printer $printer

    # add logic to tverify printer exists on new server
    Write-Host "Adding $($newprinter)"
    Add-printer -connectionname $newprinter
}

Write-Host "Completed, closing in 30 seconds"
Start-Sleep -Seconds 30