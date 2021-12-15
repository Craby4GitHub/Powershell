$oldServer = 'wc-vm-prtsvr'
$newServer = 'WC-Print'

$oldPrinterQueues = Get-printer -Name "\\$oldServer*"

ForEach ($printer in $oldPrinterQueues) {

    Write-Host "Removing $($printer.name)"
    Remove-printer $printer.name

    $newprinter = $printer.name -replace $oldServer, $newServer

    if (Get-printer -ComputerName $newServer -Name $newprinter.Split('\')[-1]) {
        Write-Host "Adding $($newprinter)"
        Add-printer -connectionname $newprinter
    }
    else {
        Write-Host "$newprinter does not exist on $newServer"
    } 
}

Write-Host "Completed, closing in 30 seconds"
Start-Sleep -Seconds 30