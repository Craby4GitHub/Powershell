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

# Pulling all devices from SCCM. Once for heartbeat and another for serial numbers 
# Wishlist: Logic to filter by computer name? Reason: filter out VM's, servers(?)
$allSccmDevices = Get-CMDevice -CollectionName 'Agent Installed' | Select-Object Name, ResourceID, LastDDR
$allSccmSerialNumber = Get-WmiObject -Class SMS_G_system_SYSTEM_ENCLOSURE  -Namespace root\sms\site_PCC -ComputerName "do-sccm.pcc-domain.pima.edu" | Select-Object ResourceID, SerialNumber

Write-Log -level INFO -message "Loaded $($allSccmDevices.count) devices from SCCM"

# Pulling all TDX assets
# Wishlist: Filter for only computers. Currently also pulls printers, TVs, ect
$allTDXAssets = Search-TDXAssets
Write-Log -level INFO -message "Loaded $($allTDXAssets.count) devices from TDX"

foreach ($tdxAsset in $allTDXAssets) {
    # Proress bar to show how far along we are
    [int]$pct = ($allTDXAssets.IndexOf($tdxAsset) / $allTDXAssets.Count) * 100
    Write-progress -Activity '...' -PercentComplete $pct -status "$pct% Complete"
    
    switch ($tdxAsset.SupplierName) {
        'Apple' { 
            # Different API calls for mobile devices and computers
            # Wishlist: Also include iphones instead of just tablets?
            if ($tdxAsset.ProductType -eq 'Tablet') {
                Write-Log -level INFO -message "Searching for mobile asset in JAMF records" -assetSerialNumber $tdxAsset.SerialNumber
                $jamfInventoryDate = (Search-JamfMobileDevices -serialNumber $tdxAsset.SerialNumber).General.last_inventory_update_utc
            }
            else {
                Write-Log -level INFO -message "Searching for computer asset in JAMF records" -assetSerialNumber $tdxAsset.SerialNumber
                $jamfInventoryDate = (Search-JamfComputers -serialNumber $tdxAsset.SerialNumber).General.report_date_utc
            }

            if ($null -ne $mdmInventoryDate) {
                $mdmInventoryDate = $jamfInventoryDate
            }else {
                Write-Log -level WARN -message "Device not in JAMF" -assetSerialNumber $tdxAsset.SerialNumber
            }
        }
        Default {

            Write-Log -level INFO -message "Searching for asset in SCCM records" -assetSerialNumber $tdxAsset.SerialNumber
            # Getting resource ID from searching by Serial Number and getting last hardware scan
            $sccmResourceID = ($allSccmSerialNumber | Where-Object -Property SerialNumber -EQ $tdxAsset.SerialNumber).ResourceID
            $sccmDeviceInfo = $allSccmDevices | Where-Object -Property ResourceID -eq $sccmResourceID

            # Verify there is an SCCM Object
            if ($null -ne $sccmDeviceInfo) {
                Write-Log -level INFO -message "Found $($sccmDeviceInfo.Name) in SCCM" -assetSerialNumber $tdxAsset.SerialNumber

                # There shouldnt be any duplicates because we are searching SCCM based on serial number, so save this info to the log for manual clean up
                # Wishlist: Auto cleanup?
                if ($sccmDeviceInfo.Count -gt 1) {
                    foreach ($resource in $sccmDeviceInfo) {
                        Write-Log -level ERROR -message "Duplicate Serial Number" -assetSerialNumber $tdxAsset.SerialNumber
                    }
                }

                $mdmInventoryDate = $sccmDeviceInfo.LastDDR
            }
            else {
                Write-Log -level WARN -message "Device is not in SCCM" -assetSerialNumber $tdxAsset.SerialNumber
                break
            }
        }
    }
        
    # Get assets last inventory date data from TDX. If the asset has no inventory data, set a fake date
    if ($null -ne ($tdxAsset.Attributes | Where-Object -Property Name -eq 'Last Inventory Date').Value) {
        $tdxAssetInventoryDate = Get-date (Get-TDXAssetAttributes -ID $tdxAsset.ID | Where-Object -Property Name -eq 'Last Inventory Date').Value
    }
    else {
        $tdxAssetInventoryDate = get-date '05/03/1989' # Fake date
    }

    # Check to see if TDX inventory date is atleast X days older than the last SCCM heartbeat. If it is, edit the TDX asset
    if (($mdmInventoryDate - $tdxAssetInventoryDate).Days -gt 2) {
        Write-Log -level INFO -message "Updating inventory date to $($mdmInventoryDate)" -assetSerialNumber $tdxAsset.SerialNumber
        Edit-TDXAsset -Asset $tdxAsset -sccmLastHardwareScan $mdmInventoryDate
    }
    else {
        Write-Log -level INFO -message "TDX inventory date of $($tdxAssetInventoryDate) is more recent than $($mdmInventoryDate)" -assetSerialNumber $tdxAsset.SerialNumber
    }
}