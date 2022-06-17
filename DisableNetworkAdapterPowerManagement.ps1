# Relaunch as an elevated process:
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}
  
$networkAdapaters = Get-WmiObject win32_networkadapter
$powerMgmtSettings = Get-WmiObject MSPower_DeviceEnable -Namespace root\wmi
foreach ($powerMgmtSetting in $powerMgmtSettings) {
    foreach ($adapter in $networkAdapaters) {
        if ($adapter.AdapterType -like "Ethernet*" -and $adapter.name -match "Wireless|Wi(-)?Fi") {
            $powerMgmtSetting.enable = $False
            $powerMgmtSetting.psbase.put()
        }
    }
}