$name = (Get-WmiObject -query "Select * from Win32_SystemEnclosure").SMBiosAssetTag

$computername = (Get-WmiObject Win32_ComputerSystem)

$computername.Rename("EC-L146"+$name+"CN")
Write-Host "Computer renamed to EC-L146$name+CN"