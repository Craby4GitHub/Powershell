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

$pccDomain += $eduDomain
#$pimaDomains = @($pccDomain, $eduDomain)

#endregion

#region Regex
# 15 Characters
$15Characters = [regex]"\w{15}"

# 2 letter campus code - 1 or 2 Building letter 8-9 numbers(2-3 for room and 6 for PCC) 2 letter computer type
$NormalCampus = [regex]"^([a-z]{4}\d{2}|(([a-z]{2}|\d{2})-([a-z]|[a-z]{2})))\d{7,9}[a-z]{2}$"

# 2 letter campus code 2 Building letter 9 numbers(3 for room and 6 for PCC) 2 letter computer type
$DownTownCampus = [regex]"^[a-zA-Z]{4}\d{9}[a-zA-Z]{2}$"

# CARES
$CaresAct = [regex]"[a-z]{3}-[a-z]{3}\d{6}[a-z]{2}"

# VDI
$VDI = [regex]"^(VDI-)\w{0,}$"

# VM's Dont need em
$VM = [regex]"[vV]\d{0,}$"
#endregion


foreach ($tdxAsset in $allTDXAssets) {
    if ($null -ne $tdxAsset.tag) {
        # Wishlist: Add regex groups because searching on tag will search the WHOLE name, including the room
        $matchedPccTag = @()
        $matchedPccTag += $pccDomain | Where-Object -Property Name -Match "$($tdxAsset.tag)[a-z]{2}$"

        # Compare computer name to PCC Naming convention
        # https://docs.google.com/spreadsheets/d/1gLkgjxNlxwbNizH_EsQY_-ARmaStVOYra1pwcxIQvgM/edit#gid=0
        $matchesRegex = @()
        $notmatchRegEx = @()
        foreach ($computer in $matchedPccTag) {
            #Write-Log -level INFO -message "$($computer.Name) matched $($tdxAsset.tag)" -assetSerialNumber $tdxAsset.SerialNumber
            Switch -regex ($computer.Name) {
                !$15Characters {
                    Write-Log -level ERROR -message "$($computer.Name) too long" -assetSerialNumber $tdxAsset.SerialNumber
                    $notmatchRegEx += , @($computer.name)
                    break
                }
                $NormalCampus {
                    Write-Log -level INFO -message "$($computer.Name) passes verification of name" -assetSerialNumber $tdxAsset.SerialNumber
                    $matchesRegex += , @($computer)
                    break
                }
                $CaresAct {
                    break
                }
                $VDI {
                    break
                }
                $VM {
                    break
                }
                default {
                    Write-Log -level ERROR -message "$($computer.Name) name doesnt match regex logic" -assetSerialNumber $tdxAsset.SerialNumber
                    $notmatchRegEx += , @($computer.name)
                }
            }
        }

        # For the computers that pass the name verification, pull the PCC number and room and compare it to TDX
        foreach ($computer in $matchesRegex) {
            if ($computer.name.StartsWith('DC')) {
                $Campus = $computer.name.substring(0, 2)
                $Room = $computer.name.substring(2, 4)
                $PCCNumber = $computer.name.substring(7, 6)

                if ($PCCNumber -eq $tdxAsset.Tag) {
                    #Write-Log -level info -message "TDX:$($tdxAsset.tag) match AD:$PCCNumber for computer $($computer.Name)" -assetSerialNumber $tdxAsset.SerialNumber
                    if ($Room -ne $tdxAsset.LocationRoomName) {
                        Write-Log -level WARN -message "TDX Room:$($tdxAsset.LocationRoomName) does not match $($computer.Name)" -assetSerialNumber $tdxAsset.SerialNumber
                    }
                    else {
                        #Write-Log -level INFO -message "TDX:$($tdxAsset.LocationRoomName) does match $($computer.Name)" -assetSerialNumber $tdxAsset.SerialNumber
                    }
                }
                else {
                    Write-Log -level INFO -message "TDX PCC:$($tdxAsset.Tag) does not match computer $($computer.Name)" -assetSerialNumber $tdxAsset.SerialNumber
                }
            }
            else {
                $Campus = $computer.name.substring(0, 2)
                $Room = $computer.name.substring(3, 4)
                $PCCNumber = $computer.name.substring(7, 6)

                if ($PCCNumber -eq $tdxAsset.Tag) {
                    #Write-Log -level info -message "TDX:$($tdxAsset.tag) match AD:$PCCNumber for computer $($computer.Name)" -assetSerialNumber $tdxAsset.SerialNumber
                    if ($Room -ne $tdxAsset.LocationRoomName) {
                        Write-Log -level WARN -message "TDX:$($tdxAsset.LocationRoomName) does not match $($computer.Name)" -assetSerialNumber $tdxAsset.SerialNumber
                    }
                    else {
                        #Write-Log -level INFO -message "TDX Room:$($tdxAsset.LocationRoomName) does match $($computer.Name)" -assetSerialNumber $tdxAsset.SerialNumber
                    }
                }
                else {
                    Write-Log -level INFO -message "TDX PCC:$($tdxAsset.Tag) does not match computer $($computer.Name)" -assetSerialNumber $tdxAsset.SerialNumber
                }
          
            }
        }
    }
    else {
        Write-Log -level ERROR -message "No PCC Number in TDX" -assetSerialNumber $tdxAsset.SerialNumber
    }
}