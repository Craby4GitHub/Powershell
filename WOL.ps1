. "$PSScriptRoot\Write-Log.ps1"
. "$PSScriptRoot\Send-WOL.ps1"

$Computer = Read-Host "Enter computer name:"

write-host "Looking for $Computer"
$computerDetails = (Get-WmiObject -Class SMS_R_SYSTEM -Namespace "root\sms\site_PCC" -computerName 'do-sccm.pcc-domain.pima.edu' | Where-Object { $_.Name -eq "$Computer" })

if ($null -eq $computerDetails) {
    write-host "$Computer not found in SCCM"
    exit
}

$MAC = $computerDetails.MACAddresses[0]
$IP = $computerDetails.IPADDRESSES[0]
write-host 'MAC:' $MAC
write-host 'IP:' $IP

$i = 1
do {
    Write-host "Sending Ping number $i to $Computer"
    Send-WOL -mac $mac -ip $IP
    
    if (Test-Connection -ComputerName $Computer -Count 3 -quiet) {
        Write-Log -status $true -Message "$Computer,$MAC,$IP" -logfile "$PSScriptRoot\log.csv"
        Write-host "im awake"
        exit
    }
    else {
        Write-Log -status $false -Message "$Computer,$MAC,$IP" -logfile "$PSScriptRoot\log.csv"
        write-host 'nope'
    }

    Start-Sleep -Seconds 30
    $i++
} until ($i -gt 4)