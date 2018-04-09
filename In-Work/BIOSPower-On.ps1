$biosSettings = Get-WmiObject -Class HP_BIOSSettingInterface -Namespace root/hp/instrumentedBIOS

$biosSettings.SetBIOSSetting('Num Lock State at Power-On','On')
$biosSettings.SetBIOSSetting('After Power Loss','On')
$biosSettings.SetBIOSSetting('Fast Boot',$status)
$biosSettings.SetBIOSSetting('Legacy Support',$status)
$biosSettings.SetBIOSSetting('Monday',$status)
$biosSettings.SetBIOSSetting('Tuesday',$status)
$biosSettings.SetBIOSSetting('Wednesday',$status)
$biosSettings.SetBIOSSetting('Thursday',$status)
$biosSettings.SetBIOSSetting('Friday',$status)
$biosSettings.SetBIOSSetting('Saturday',$status)
$biosSettings.SetBIOSSetting('Sunday',$status)