#Requires -Modules activedirectory
import-module activedirectory

# Load TDX API functions
. (Join-Path $PSSCRIPTROOT "TDX-API.ps1")

# Pulling all TDX assets
# Wishlist: Filter for only computers. Currently also pulls printers, TVs, ect
Write-progress -Activity 'Getting computers from TDX...'
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
    [int]$pct = ($pccDomain.IndexOf($computer) / $pccDomain.count) * 100
    Write-progress -Activity "Comparing $($computer.Name) to TDX data..." -percentcomplete $pct -status "$pct% Complete"
    $Matches = $null

    # Verifying the computer name matches a relatively valid name
    # Not currently using the campus match
    $matchedComputer = $computer | Where-Object -Property DNSHostName -Match "(?<Campus>[a-z]{2})(?<BldgAndRoom>.{5})(?<PCCNumber>\d{6})[a-z]{2}$"

    # Verify there is a computer that matches
    if ($null -ne $matchedComputer) {

        # Find tdx asset that matches Active Directory PCC number
        $tdxAsset = $allTDXAssets | Where-Object -Property Tag -EQ $Matches.PCCNumber

        <#
        foreach ($tdxAsset in $allTDXAssets) {
            if ($tdxAsset.Tag -eq $Matches.PCCNumber) {
                if ($Matches.BldgAndRoom.Trim('-') -ne $tdxAsset.LocationRoomName.Split(' ')[0]) {
                    $InconsistentArray += "Bldg or Room does not match, $($tdxAsset.Tag),$($tdxAsset.SerialNumber),$($tdxAsset.LocationRoomName),$($computer.Name),$($Matches.Domain)"
                }#elseif($Matches.Campus -ne $tdxAsset.Location) {
                #}
                else {
                    # Building or room match TDX
                }
            }
            elseif ($allTDXAssets.IndexOf($tdxAsset) -eq $pccDomain.count) {
                $InconsistentArray += "Not found in TDX,$($Matches.PCCNumber),Unknown,Unknown,$($computer.Name),Unknown)"
            }
        }
        #>
        # Compare AD object to tdxAsset
        if ($null -ne $tdxAsset) {
            if ($Matches.BldgAndRoom.Trim('-') -ne $tdxAsset.LocationRoomName.Split(' ')[0]) {
                $InconsistentArray += "Bldg or Room does not match, $($tdxAsset.Tag),$($tdxAsset.SerialNumber),$($tdxAsset.LocationRoomName),$($computer.Name),$($Matches.Domain)"
            }#elseif($Matches.Campus -ne $tdxAsset.Location) {
            #}
            else {
                # Building or room match TDX
            }
        }
        else {
            $InconsistentArray += "Not found in TDX,$($Matches.PCCNumber),Unknown,Unknown,$($computer.Name),$($Matches.Domain)"
        }
    }
    else {
        $InconsistentArray += "Invalid AD name,Unknown,Unknown,Unknown,$($computer.Name),Unknown)"
    }  
}

# Save to log
Write-progress -Activity 'Saving Log...'
New-Item -Path "$PSScriptroot\Inconsistent.csv" -value "Issue,PCC Number,Serial Number,TDX Room Number,AD Computer Name,Domain`n" -Force | Out-Null
$InconsistentArray | Out-File -FilePath "$PSScriptroot\Inconsistent.csv" -Append -Encoding ASCII