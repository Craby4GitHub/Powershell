
#$initParams = @{}
  
# Import the ConfigurationManager.psd1 module 
if ($null -eq (Get-Module ConfigurationManager)) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" #@initParams 
}

# Connect to the site's drive if it is not already present
if ($null -eq (Get-PSDrive -Name 'PCC' -PSProvider CMSite -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name 'PCC' -PSProvider CMSite -Root 'do-sccm.pcc-domain.pima.edu' #@initParams
}

# Set the current location to be the site code.
Set-Location "PCC:\" #@initParams

# Load TDX API functions
. (Join-Path $PSSCRIPTROOT "TDX-API.ps1")

# Pulling all devices from the SCCM. Once for heartbeat and another for serial numbers 
# Wishlist: Logic to filter by computer name? Reason: filter out VM's, servers(?)
$allSccmDevices = Get-CMDevice -CollectionName 'Agent Installed' | Select-Object Name, ResourceID, LastDDR
$allSccmSerialNumber = Get-WmiObject -Class SMS_G_system_SYSTEM_ENCLOSURE  -Namespace root\sms\site_PCC -ComputerName "do-sccm.pcc-domain.pima.edu" | Select-Object ResourceID, SerialNumber
Write-Log -level INFO -string "Loaded $($allSccmDevices.count) devices from SCCM"

# Pulling all TDX assets
# Wishlist: Filter for only computers. Currently also pulls printers, TVs, ect
$allTDXAssets = Search-TDXAssets
Write-Log -level INFO -string "Loaded $($allTDXAssets.count) devices from TDX"

foreach ($tdxAsset in $allTDXAssets) {
    [int]$pct = ($allTDXAssets.IndexOf($tdxAsset)/$allTDXAssets.Count)*100
    Write-progress -Activity '...' -PercentComplete $pct -status "$pct% Complete"
    Write-Log -level INFO -string "Searching for $($tdxAsset.tag) in SCCM records"

    # Getting PCC and Serial number for current device
    # Wishlist: Fix logic if more than one entry is returned. Will get generic error if there are
    $sccmDeviceInfo = $allSccmDevices | Where-Object -Property name -like "*$($tdxAsset.Tag)*"
    $sccmDeviceSerialNumber = ($allSccmSerialNumber | Where-Object -Property ResourceID -EQ $sccmDeviceInfo.ResourceID).SerialNumber
    
    # Verifying TDX Serial Number data to SCCM data. Reason: SCCM data is search based on pcc number. A computer could be misnamed or there could be virtual machines
    if ($tdxAsset.SerialNumber -eq $sccmDeviceSerialNumber) {
        Write-Log -level INFO -string "Serial Numbers match"
        
        # Wishlist: add check for null date. If it is null, then the edit wont be able to change the date
        if ($null -eq ($tdxAsset.Attributes | Where-Object -Property Name -eq 'Last Inventory Date').Value) {
            $allCustomAttributes = [PSCustomObject]@{
                ID    = '126172';
                Name  = 'Last Inventory Date';
                Value = '';
            } 
        }
        else {
            $tdxAssetInventoryDate = Get-date (Get-TDXAssetAttributes -ID $tdxAsset.ID | Where-Object -Property Name -eq 'Last Inventory Date').Value -ErrorAction SilentlyContinue # erroraction for null dates 
        }

        # Check to see if TDX inventory date is atleast X days older than the last SCCM heartbeat
        if (($sccmDeviceInfo.LastDDR - $tdxAssetInventoryDate).Days -gt 1) {
            Write-Log -level INFO -string "Updating $($tdxAsset.tag) inventory date to $($sccmDeviceInfo.LastDDR)"
            Edit-TDXAsset -Asset $tdxAsset -sccmLastHardwareScan $sccmDeviceInfo.LastDDR
        }
        else {
            Write-Log -level INFO -string "TDX inventory date is most recent"
        }
    }
    else {
        Write-Log -level INFO -string "Serial Numbers do not match"
    }
}