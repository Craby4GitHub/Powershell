#Requires -Modules activedirectory
import-module activedirectory

# Load TDX API functions
. (Join-Path $PSSCRIPTROOT "TDX-API.ps1")

# Setting Log name
$logName = ($MyInvocation.MyCommand.Name -split '\.')[0] + ' log'
$logFile = "$PSScriptroot\$logName.csv"
. ((Get-Item $PSScriptRoot).Parent.FullName + '\Callable\Write-Log.ps1')

# Pulling all TDX assets
# Wishlist: Filter for only computers. Currently also pulls printers, TVs, ect
$allTDXAssets = Search-TDXAssets
Write-Log -level INFO -message "Loaded $($allTDXAssets.count) devices from TDX"

#region Import from AD
#Write-progress -Activity 'Getting computers from EDU and PCC Domain...'
$pccDomain = Get-ADComputer -Filter { (OperatingSystem -notlike '*windows*server*') } -Properties OperatingSystem -Server PCC-Domain.pima.edu
  
$eduDomain = Get-ADComputer -Filter { (OperatingSystem -notlike '*windows*server*') } -Properties OperatingSystem -Server EDU-Domain.pima.edu

# Combine the domains together
$pccDomain += $eduDomain
#endregion

foreach ($tdxAsset in $allTDXAssets) {
    # Proress bar to show how far along we are
    [int]$pct = ($allTDXAssets.IndexOf($tdxAsset) / $allTDXAssets.Count) * 100
    Write-progress -Activity "Working on $($tdxAsset.Tag)" -PercentComplete $pct -status "$pct% Complete"

    if ($null -ne $tdxAsset.tag) {
        $pccDomain | Where-Object -Property Name -Match "(?<other>.*)(?<PCCNumber>$($tdxAsset.tag))[a-z]{2}$" | foreach-object { 
            if ($Matches['other'].StartsWith('DC')) {
                if ($Matches['PCCNumber'] -eq $tdxAsset.Tag) {
                    if ($Matches['other'].substring(2, 5) -ne $tdxAsset.LocationRoomName) {
                        Write-Log -level INFO -message "TDX Room:$($tdxAsset.LocationRoomName) does not match $($_.Name)" -assetSerialNumber $tdxAsset.SerialNumber
                    }
                }
            }
            else {
                $campus, $bldgAndRoom = $Matches['other'].split('-')
                if ($bldgAndRoom -ne $tdxAsset.LocationRoomName) {
                    Write-Log -level INFO -message "TDX Room:$($tdxAsset.LocationRoomName) does not match $($_.Name)" -assetSerialNumber $tdxAsset.SerialNumber
                }
            }
        }
    }
    else {
        Write-Log -level ERROR -message "No PCC Number in TDX" -assetSerialNumber $tdxAsset.SerialNumber
    }
}