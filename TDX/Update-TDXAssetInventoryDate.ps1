function Get-SCCMDevice($computerName) {

    # Site configuration
    $SiteCode = "PCC" # Site code 
    $ProviderMachineName = "do-sccm.pcc-domain.pima.edu" # SMS Provider machine name

    $initParams = @{}
  
    # Import the ConfigurationManager.psd1 module 
    if ($null -eq (Get-Module ConfigurationManager)) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
    }

    # Connect to the site's drive if it is not already present
    if ($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
    }

    # Set the current location to be the site code.
    Set-Location "$($SiteCode):\" @initParams


    $device = Get-CMDevice -Name $computerName -Fast

    return $device
}

. (Join-Path $PSSCRIPTROOT "TDX-API.ps1")
#. Join-Path (Get-ChildItem .\JAMF) "JAMF-API.ps1"

#$allSCCMDevices = Get-CMDevice -CollectionName 'Agent Installed' | Select-Object Name, ResourceID, LastHardwareScan
$allTDXAssets = Search-TDXAsset



foreach ($tdxAsset in $allTDXAssets) {
    write-host "Searching $($tdxasset.tag) in SCCM"
    $SCCM = Get-SCCMDevice -computerName $('*' + $tdxAsset.Tag + '*')

    # Fix logic if more than one entry. Will get generic error if not
    $sccmSerialNumber = (Get-WmiObject -Class SMS_G_system_SYSTEM_ENCLOSURE  -Namespace root\sms\site_PCC -ComputerName "do-sccm.pcc-domain.pima.edu" -Filter "ResourceID = $($SCCM.ResourceID)").Serialnumber

    # Verify the assets serial number match. Reason: Computer could be misnamed or there could be virtual machines as we only search SCCM on pcc number
    write-host "Verify: TDX SN: $($tdxasset.SerialNumber)-----SCCM SN:$sccmSerialNumber"
    if ($tdxAsset.SerialNumber -eq $sccmSerialNumber) {
        $assetAttributes = (Get-TDXAssetDetails -ID $tdxAsset.ID).Attributes

        # Check to see if TDX inventory date is newer than the last SCCM heartbeat
        # Wishlist: I would like to move away from using an index value on the attributes as more attributes may be added in the future
        write-host "Verify: TDX Date: $($assetAttributes[1].value)-----SCCM Date:$($SCCM.LastDDR)"
        if ($assetAttributes[1].value -gt $SCCM.LastDDR) {
            write-host "Found $($tdxAsset.tag)" -ForegroundColor Green
            Edit-TDXAsset -Asset $tdxAsset -sccmLastHardwareScan $SCCM.LastHardwareScan
        }
    }
}