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

$sccmquery = Get-CMDevice -CollectionName 'Agent Installed' | Select-Object Name, ResourceID, LastDDR

# Converting from SCCM object otherwise we will hit quota limit when doing sccm searches
$allSCCMDevices = [System.Collections.ArrayList]@()
foreach ($sccmDevice in $sccmquery) {
    $allSCCMDevices += [PSCustomObject]@{
        Name       = $sccmDevice.Name
        ResourceID = $sccmDevice.ResourceID
        LastDDR    = $sccmDevice.LastDDR
    }
}

$allTDXAssets = Search-TDXAsset
foreach ($tdxAsset in $allTDXAssets) {
    write-host "Searching $($tdxasset.tag) in SCCM"
    #$SCCM = Get-SCCMDevice -computerName $('*' + $tdxAsset.Tag + '*')
    $SCCM = $allSCCMDevices | Where-Object -Property name -like "*$($tdxAsset.Tag)*"

    # Fix logic if more than one entry. Will get generic error if not
    $sccmSerialNumber = (Get-WmiObject -Class SMS_G_system_SYSTEM_ENCLOSURE  -Namespace root\sms\site_PCC -ComputerName "do-sccm.pcc-domain.pima.edu" -Filter "ResourceID = $($SCCM.ResourceID)").Serialnumber

    # Verify the assets serial number match. Reason: Computer could be misnamed or there could be virtual machines as we only search SCCM on pcc number
    if ($tdxAsset.SerialNumber -eq $sccmSerialNumber) {
        write-host "Serial Numbers match" -ForegroundColor Green
        $assetAttributes = (Get-TDXAssetDetails -ID $tdxAsset.ID).Attributes

        # Check to see if TDX inventory date is newer than the last SCCM heartbeat
        # Wishlist: I would like to move away from using an index value on the attributes as more attributes may be added in the future
        if ($assetAttributes[1].value -lt $SCCM.LastDDR) {
            write-host "Updating $($tdxAsset.tag) inventory date to $($SCCM.LastDDR) " -ForegroundColor Green
            Edit-TDXAsset -Asset $tdxAsset -sccmLastHardwareScan $SCCM.LastDDR
        }
    }
}