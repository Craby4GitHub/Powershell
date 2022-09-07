# Import the ConfigurationManager.psd1 module 
if ($null -eq (Get-Module ConfigurationManager)) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
}

# Connect to the site's drive if it is not already present
if ($null -eq (Get-PSDrive -Name 'PCC' -PSProvider CMSite -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name 'PCC' -PSProvider CMSite -Root 'do-sccm.pcc-domain.pima.edu'
}

# Set the current location to be the site code.
Set-Location "PCC:\"

# Load TDX API functions
. (Join-Path $PSSCRIPTROOT "TDX-API.ps1")

# Load JAMF API functions
. ((Get-Item $PSScriptRoot).Parent.FullName + '\JAMF\JAMF-API.ps1')

# Setting Log name
$logName = ($MyInvocation.MyCommand.Name -split '\.')[0] + ' log'
$logFile = "$PSScriptroot\$logName.csv"
. ((Get-Item $PSScriptRoot).Parent.FullName + '\Callable\Write-Log.ps1')

# Pulling all devices from SCCM. Once for heartbeat and another for serial numbers
# Reason: There is a sudo ratelimit when calling from SCCM, so might as well load it all at one, one time
# Wishlist: Logic to filter by computer name? Reason: filter out VM's, servers(?)
$allSccmDevices = Get-CMDevice -CollectionName 'Agent Installed' | Select-Object Name, ResourceID, LastDDR
$allSccmSerialNumber = Get-WmiObject -Class SMS_G_system_SYSTEM_ENCLOSURE  -Namespace root\sms\site_PCC -ComputerName "do-sccm.pcc-domain.pima.edu" | Select-Object ResourceID, SerialNumber

Write-Log -level INFO -message "Loaded $($allSccmDevices.count) devices from SCCM"

# Pulling all TDX assets
# Wishlist: Filter for only computers. Currently also pulls printers, TVs, ect
$allTDXAssets = Search-TDXAssets -AppName ITAsset
Write-Log -level INFO -message "Loaded $($allTDXAssets.count) devices from TDX"

$allSccmDevices = Get-CMDevice -CollectionName 'Agent Installed' | Select-Object Name, ResourceID, LastDDR
# What other info is in object? Find somthing that will determine AD object
$allSccmSerialNumber = Get-WmiObject -Class SMS_G_system_SYSTEM_ENCLOSURE  -Namespace root\sms\site_PCC -ComputerName "do-sccm.pcc-domain.pima.edu" | Select-Object ResourceID, SerialNumber


#region Import from AD
Write-progress -Activity 'Getting computers from EDU and PCC Domain...'
$pccDomain = Get-ADComputer -Filter { (OperatingSystem -notlike '*windows*server*') } -Properties OperatingSystem -Server PCC-Domain.pima.edu
$eduDomain = Get-ADComputer -Filter { (OperatingSystem -notlike '*windows*server*') } -Properties OperatingSystem -Server EDU-Domain.pima.edu

# Combine the domains together
$adComputers = $pccDomain += $eduDomain
#endregion


foreach ($disposedAsset in $allTDXAssets |  Where-Object -Property StatusName -EQ 'Disposed') {
    if ($disposedAsset.ManufacturerName -ne 'Apple') {
        foreach ($ghostAsset in $allSccmSerialNumber | Where-Object -Property SerialNumber -EQ $disposedAsset.SerialNumber) {
            # Delete AD object
            if (!"Deleted AD object") {
                foreach ($adComputer in $adComputers | Where-Object -Property DNSHostName -Match "(?<Campus>[a-z]{2})(?<BldgAndRoom>.{5})(?<PCCNumber>$disposedAsset)[a-z]{2}$") {
                    # Delete AD Object
                }
            }
            # Delete SCCM Object
        }
    }
    else {
        # Delete JAMF object
    }   
}