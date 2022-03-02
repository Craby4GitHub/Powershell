#Requires -Modules activedirectory
import-module activedirectory

# Load TDX API functions
. (Join-Path $PSSCRIPTROOT "TDX-API.ps1")

# Pulling all TDX assets
# Wishlist: Filter for only computers. Currently also pulls printers, TVs, ect
$allTDXAssets = Search-TDXAssets

#region Import from AD
Write-progress -Activity 'Getting computers from EDU and PCC Domain...'
#$pccDomain = Get-ADComputer -Filter { (OperatingSystem -notlike '*windows*server*') } -Properties OperatingSystem -Server PCC-Domain.pima.edu
$pccDomain = Get-ADComputer -Filter { (name -like 'wc*') } -Properties OperatingSystem -Server PCC-Domain.pima.edu 
#$eduDomain = Get-ADComputer -Filter { (OperatingSystem -notlike '*windows*server*') } -Properties OperatingSystem -Server EDU-Domain.pima.edu
$eduDomain = Get-ADComputer -Filter { (name -like 'wc*') } -Properties OperatingSystem -Server EDU-Domain.pima.edu

# Combine the domains together
$pccDomain += $eduDomain
#endregion

$InconsistentArray = @()

foreach ($computer in $pccDomain) {
    $Matches = $null

    $matchedComputer = $computer | Where-Object -Property DNSHostName -Match "(?<other>.*)(?<PCCNumber>\d{6})[a-z]{2}\.(?<Domain>PCC|EDU)"

    # Verify there is a computer with a 6 digit number and 2 suffix characters
    if ($null -ne $matchedComputer) {

        $tdxAsset = $allTDXAssets | Where-Object -Property Tag -EQ $Matches['PCCNumber']
        
        # Verify TDX has an asset matching The PCC Number
        if ($null -ne $tdxAsset) {
            $campus, $bldgAndRoom = $Matches['other'].split('-')
            if ($bldgAndRoom -ne $tdxAsset.LocationRoomName.Split(' ')[0]) {
                $InconsistentArray += "$($tdxAsset.Tag),$($tdxAsset.SerialNumber),$($tdxAsset.LocationRoomName),$($computer.Name),$($Matches['Domain'])"
            }
        }
        else {
            $InconsistentArray += "$($Matches['PCCNumber']),Not in TDX,Not in TDX,$($computer.Name),$($Matches['Domain'])"
        }
    }

    
}

# Save to log
New-Item -Path "$PSScriptroot\Inconsistent.csv" -value "PCC Number,Serial Number,TDX Room Number,AD Computer Name,Domain`n" -Force | Out-Null
$InconsistentArray | Out-File -FilePath "$PSScriptroot\Inconsistent.csv" -Append -Encoding ASCII