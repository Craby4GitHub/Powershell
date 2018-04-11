Clear-Host

#Requires -Modules activedirectory
import-module activedirectory

Function New-InventoryObject() {   
    param([string]$PCCNumber, [string]$Room, [string]$Campus)
    return [pscustomobject] @{     
        'PCC Number' = $PCCNumber
        'Room'       = $Room
        'Campus'     = $Campus
    }
}

Function Get-File {
    Do {
        $filePath = Get-FileName $PSScriptroot

        $correctFile = read-host 'Is' $filePath "the correct file? (Y/N)"
        if ($correctFile -eq 'Y' -and $filePath -ne $null) {
            return $inputFile = Import-Csv $filePath           
        }
        else {
            write-host "Your selection is empty or does not exist"
        }
    }until($correctFile -eq 'Y' -and $inputFile -ne $null)
}

Function Get-FileName($initialDirectory) {
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    [void]$OpenFileDialog.ShowDialog()
    $OpenFileDialog.FileName
}

$matchesRegex = @()
$notmatchRegEx = @()
$PCCNumberArray = @()
$AssetHash = @{}

'Loading selected file into an array'
#region Parse ITAM list to array
ForEach ($object in Get-File) {
    Write-progress -Activity 'Working...' -currentoperation 'Loading selected file into an array...'
    if (($object.'Asset Type' -eq 'CPU') -or
        ($object.'Asset Type' -eq 'Laptop') -or
       (($object.'Asset Type' -eq 'Tablet') -and ($object.'Manufacturer' -eq 'Microsoft'))) {
        try {
            $AssetHash.Add($object.'Barcode #', $object.'Room')
        }
        Catch {
            Write-output "$($object.'IT #'), $($object.'Location'), $($_.Exception.Message)" | Sort-Object | Out-File -FilePath ($PSScriptroot + '\Error.csv') -Append
        }
    }	
}
#endregion

'Loading AD Computers'
#region Import from AD
Write-progress -Activity 'Working...' -status 'Getting computers from EDU and PCC Domain...'
$PCCArray = (Get-ADComputer -Filter {(OperatingSystem -notlike '*windows*server*')} -Properties OperatingSystem -Server PCC-Domain.pima.edu).Name
  
$EDUArray = (Get-ADComputer -Filter {(OperatingSystem -notlike '*windows*server*')} -Properties OperatingSystem -Server EDU-Domain.pima.edu).Name

#endregion

#region Regex
# 15 Characters
$15Characters = [regex]"\w{15}"

# 2 letter campus code - 1 or 2 Building letter 8-9 numbers(2-3 for room and 6 for PCC) 2 letter computer type
$NormalCampus = [regex]"^([a-zA-Z]{4}|([a-zA-Z]{2}-([a-zA-Z]|[a-zA-Z]{2})))(\d{8}|\d{9})[a-zA-Z]{2}$"

# 2 letter campus code 2 Building letter 9 numbers(3 for room and 6 for PCC) 2 letter computer type
$DownTownCampus = [regex]"^[a-zA-Z]{4}\d{9}[a-zA-Z]{2}$"
# VDI
$VDI = [regex]"^(VDI-)\w{0,}$"

# VM's Dont need em
$VM = [regex]"[vV]\d{0,}$"
#endregion

'Comparing AD Computers to PCC naming convention standards'
#region Compare all computer names to PCC Standard
$i = 0
foreach ($singleComputer in ($EDUArray + $PCCArray)) {
    [int]$pct = ($i/($EDUArray + $PCCArray).count)*100
    Write-progress -Activity 'Working...' -PercentComplete $pct -currentoperation "$pct% Complete" -status 'Loading selected file into an array'
    Switch -regex ($singleComputer) {
        !$15Characters {
            $notmatchRegEx += , @($singleComputer)
            break
        }
        $NormalCampus {
            $matchesRegex += , @($singleComputer)
            break
        }
        $DownTownCampus {
            $matchesRegex += , @($singleComputer)
            break
        }
        $VDI {
            break
        }
        $VM {
            break
        }
        default {
            $notmatchRegEx += , @($singleComputer)
        }
    }
    $i++
}
#endregion

'Extracting Campus, Room number and PCC number from computer names'
#region Pull PCC Number and Room Number
$i = 0
foreach ($singleComputer in $matchesRegex) {
    [int]$pct = ($i/($matchesRegex).count)*100
    Write-progress -Activity 'Working...' -PercentComplete $pct -currentoperation "$pct% Complete" -status 'Pulling Campus, Room number and PCC number from AD computer name'
    Switch -regex ($singleComputer) {
        $NormalCampus {
            $Campus = $singlecomputer.substring(0,2)

            $Room = $singlecomputer.substring(3,4)

            $PCCNumber = $singlecomputer.substring(7,6)

            $PCCNumberArray += New-InventoryObject -PCCNumber $PCCNumber -Room $Room -Campus $Campus

            break
        }        
        $DownTownCampus {
            $Campus = $singlecomputer.substring(0,2)

            $Room = $singlecomputer.substring(2,4)

            $PCCNumber = $singlecomputer.substring(7,6)

            $PCCNumberArray += New-InventoryObject -PCCNumber $PCCNumber -Room $Room -Campus $Campus

            break
        }
        default {
        }
    }
    $i++
}
#endregion

'Finally, comparing ITAMS and AD PCC Number to the room number on hand...'
#region Main meat
For ($i = 0; $i -le ($PCCNumberArray.count - 1); $i++) {

    [int]$pct = ($i/$PCCNumberArray.count)*100
    Write-progress -Activity 'Working...' -percentcomplete $pct -currentoperation "$pct% Complete" -status 'Comparing ITAM items to AD items'
    
    if ($AssetHash[$PCCNumberArray[$i].'PCC Number'] -notmatch $PCCNumberArray[$i].'Room') {
        Write-output "$($PCCNumberArray[$i].'PCC Number'), $($PCCNumberArray[$i].'Room'), $($PCCNumberArray[$i].'Campus')" | Sort-Object | Out-File -FilePath ($PSScriptroot + '\Inconsistent.csv')-Append
    }
}
#endregion
$notmatchRegEx | Sort-Object | Out-File -FilePath ($PSScriptroot + '\NonStandard Name.csv') -Force