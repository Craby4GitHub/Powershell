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
    $SCCM = Get-SCCMDevice -computerName $('*' + $tdxAsset.Tag + '*')
    $sccmSerialNumber = (Get-WmiObject -Class SMS_G_system_SYSTEM_ENCLOSURE  -Namespace root\sms\site_PCC -ComputerName "do-sccm.pcc-domain.pima.edu" -Filter "ResourceID = $($SCCM.ResourceID)").Serialnumber

    if($tdxAsset.SerialNumber -eq $sccmSerialNumber){
    #need to add check if the invetory date is older than what we are updating
    $assetAttributes = (Get-TDXAssetDetails -ID $asset.ID).Attributes
        
        write-host "Found $($tdxAsset.tag)" -ForegroundColor Green
        Edit-TDXAsset -Asset $tdxAsset -sccmLastHardwareScan $SCCM.LastHardwareScan
    }

}