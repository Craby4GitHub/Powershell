#Requires -Modules activedirectory
import-module activedirectory

# Load TDX API functions
. (Join-Path $PSSCRIPTROOT "TDX-API.ps1")

# Pulling all TDX assets
# Wishlist: Filter for only computers. Currently also pulls printers, TVs, ect
$allTDXAssets = Search-TDXAssets
Write-Log -level INFO -message "Loaded $($allTDXAssets.count) devices from TDX"

'Loading AD Computers'
#region Import from AD
#Write-progress -Activity 'Getting computers from EDU and PCC Domain...'
$PCCArray = Get-ADComputer -Filter { (OperatingSystem -notlike '*windows*server*') } -Properties OperatingSystem -Server PCC-Domain.pima.edu
  
$EDUArray = Get-ADComputer -Filter { (OperatingSystem -notlike '*windows*server*') } -Properties OperatingSystem -Server EDU-Domain.pima.edu

#endregion

foreach ($tdxAsset in $allTDXAssets) {
    if ($null -ne $tdxAsset.tag) {
        # Wishlist: Add regex groups because searching on tag will search the WHOLE name, including the room
        $matchedPccTag = @()
        $matchedPccTag += $PCCArray | Where-Object -Property Name -Match $tdxAsset.tag
        $matchedPccTag += $EDUArray | Where-Object -Property Name -Match $tdxAsset.tag

        foreach ($computerName in $matchedPccTag) {
            if ($computerName -notmatch $tdxAsset.LocationRoomName) { 
                Write-Log -level WARN -message $computerName.Name -assetSerialNumber $tdxAsset.SerialNumber
            }
            else {
                Write-Log -level INFO -message $computerName.Name -assetSerialNumber $tdxAsset.SerialNumber
            }
        }
    }
}