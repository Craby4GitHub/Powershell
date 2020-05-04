. "$PSScriptRoot\Write-Log.ps1"
function Send-WOL
{
<# 
  .SYNOPSIS  
    Send a WOL packet to a broadcast address
  .PARAMETER mac
   The MAC address of the device that need to wake up
  .PARAMETER ip
   The IP address where the WOL packet will be sent to
  .EXAMPLE 
   Send-WOL -mac 00:11:32:21:2D:11 -ip 192.168.8.255 
#>

[CmdletBinding()]
param(
[Parameter(Mandatory=$True,Position=1)]
[string]$mac,
$ip="255.255.255.255", 
[int]$port=9
)
$broadcast = [Net.IPAddress]::Parse($ip)
 
$mac=(($mac.replace(":","")).replace("-","")).replace(".","")
$target=0,2,4,6,8,10 | ForEach-Object {[convert]::ToByte($mac.substring($_,2),16)}
$packet = (,[byte]255 * 6) + ($target * 16)
 
$UDPclient = new-Object System.Net.Sockets.UdpClient
$UDPclient.Connect($broadcast,$port)
[void]$UDPclient.Send($packet, 102) 

}
$Computer = Read-Host "Enter computer name:"

$SiteCode = 'PCC'
$SiteServer = 'do-sccm.pcc-domain.pima.edu'
write-host "Looking for $Computer"
$computerDetails=(Get-WmiObject -Class SMS_R_SYSTEM -Namespace "root\sms\site_$SiteCode" -computerName $SiteServer | Where-Object {$_.Name -eq "$Computer"})
if ($null -eq $computerDetails) {
    write-host "$Computer not found in SCCM"
    exit
}

$MAC = $computerDetails.MACAddresses[0]# | Out-String
$IP = $computerDetails.IPADDRESSES[0]
write-host 'MAC:' $MAC
write-host 'IP:' $IP

$i = 1
do {
    Write-host "Sending Ping number $i to $Computer"
    Send-WOL -mac $mac -ip $IP
    
    if (Test-Connection -ComputerName $Computer -Count 3 -quiet) {
      Write-Log -status $true -Message "$Computer,$MAC,$IP" -logfile "$PSScriptRoot\log.csv"
      write-host "im awake"
      $i = 41
    }else{
        Write-Log -status $false -Message "$Computer,$MAC,$IP" -logfile "$PSScriptRoot\log.csv"
        write-host 'nope'
    }
    Start-Sleep -Seconds 30
  $i++
} until ($i -gt 4)