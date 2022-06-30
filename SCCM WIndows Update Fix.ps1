# Updates not showing in Software center and getting an 0x80004005 in C:\Windows\CCM\Logs\WUAHandler.log? This may fix it
# https://www.prajwaldesai.com/failed-to-add-update-source-for-wuagent-of-type-2-error-0x80004005/

# Run script as admin
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
  Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
  exit
}

# Backup the old file, to be safe
Rename-Item C:\Windows\System32\GroupPolicy\Machine\Registry.pol -NewName Registry.pol.old

# Restart SCCM service
Restart-Service -Name "SMS Agent Host" -Force

# Set up cycles that will be ran
$configCycles = @(
    '{00000000-0000-0000-0000-000000000114}', # Software Updates Deployment Evaluation Cycle
    '{00000000-0000-0000-0000-000000000113}'  # Software Update Scan Cycle
)

# Run each configuration Manager Action that is related to Windows Updates
foreach($clycle in $configCycles){
    Invoke-WMIMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule $clycle | Out-Null
}

# Skylar B. addition, fixes stuck installs
WMIC /Namespace:\\root\ccm path SMS_Client CALL ResetPolicy 1 /NOINTERACTIVE