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


#region Parse ITAM list to array
ForEach ($object in Get-File) {
    if (($object.'Asset Type' -eq 'CPU') -or
        ($object.'Asset Type' -eq 'Laptop') -or
        (($object.'Asset Type' -eq 'Tablet') -and ($object.'Manufacturer' -eq 'Microsoft'))) {
        try {
            $AssetHash.Add($object.'Barcode #', $object.'Room')
        }
        Catch {
            Write-output "$($object.'IT #'), $($object.'Location')" | Sort-Object | Out-File -FilePath 'C:\users\wrcrabtree\downloads\Error.csv' -Append
        }
    }	
}
#endregion

#region Import from AD

$PCCArray = (Get-ADComputer -Filter {(OperatingSystem -notlike '*windows*server*')} -Properties OperatingSystem -Server PCC-Domain.pima.edu).Name
  
$EDUArray = (Get-ADComputer -Filter {(OperatingSystem -notlike '*windows*server*')} -Properties OperatingSystem -Server EDU-Domain.pima.edu).Name

#endregion

#region Regex
# 15 Characters
$15Characters = [regex]"\w{15}"

# 2 letter campus code - 1 Building letter 9 numbers(3 for room and 6 for PCC) 2 letter computer type
$NormalCampus = [regex]"^[a-zA-Z]{2}-[a-zA-Z]\d{9}[a-zA-Z]{2}$"

# 2 letter campus code - 2 Building letter 8 numbers(2 for room and 6 for PCC) 2 letter computer type
$WestCampus = [regex]"^[a-zA-Z]{2}-[a-zA-Z]{2}\d{8}[a-zA-Z]{2}$"

# 2 letter campus code 2 Building letter 9 numbers(3 for room and 6 for PCC) 2 letter computer type
$DownTownCampus = [regex]"^[a-zA-Z]{4}\d{9}[a-zA-Z]{2}$"

# VDI
$VDI = [regex]"^(VDI-)\w{0,}$"

# VM's Dont need em
$VM = [regex]"[vV]\d{0,}$"
#endregion

#region Compare all computer names to PCC Standard
foreach ($singleComputer in ($EDUArray + $PCCArray)) {
    Switch -regex ($singleComputer) {
        !$15Characters {
            $notmatchRegEx += , @($singleComputer)
            break
        }
        $NormalCampus {
            $matchesRegex += , @($singleComputer)
            break
        }
        $WestCampus {
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
}
#endregion

#region Pull PCC Number and Room Number
foreach ($singleComputer in $matchesRegex) {
    Switch -regex ($singleComputer) {
        $NormalCampus {
            $firstPass = $singleComputer -creplace "^[a-zA-Z]{2}-[a-zA-Z]\d{3}"
            $PCCNumber = $firstPass -creplace "[a-zA-Z]{2}|[vV]\d{0,}$"

            $secondPass = $singleComputer -creplace "^[a-zA-Z]{2}-"
            $Room = $secondPass -creplace "\d{6}[a-zA-Z]{2}$"

            $Campus = $singleComputer -creplace "-\w{12}$"

            $PCCNumberArray += New-InventoryObject -PCCNumber $PCCNumber -Room $Room -Campus $Campus

            break
        }
        $WestCampus {
            $firstPass = $singleComputer -creplace "^[a-zA-Z]{2}-[a-zA-Z]{2}\d{2}"
            $PCCNumber = $firstPass -creplace "[a-zA-Z]{2}$"

            $secondPass = $singleComputer -creplace "^[a-zA-Z]{2}-"
            $Room = $secondPass -creplace "\d{6}[a-zA-Z]{2}$"

            $Campus = $singleComputer -creplace "-\w{12}$"

            $PCCNumberArray += New-InventoryObject -PCCNumber $PCCNumber -Room $Room -Campus $Campus

            break
        }
        $DownTownCampus {
            $firstPass = $singleComputer -creplace "^[a-zA-Z]{4}\d{3}"
            $PCCNumber = $firstPass -creplace "[a-zA-Z]{2}$"

            $secondPass = $singleComputer -creplace "^[a-zA-Z]{2}"
            $Room = $secondPass -creplace "\d{6}[a-zA-Z]{2}$"

            $Campus = $singleComputer -creplace "-\w{12}$"

            $PCCNumberArray += New-InventoryObject -PCCNumber $PCCNumber -Room $Room -Campus $Campus

            break
        }
        default {
        }
    }
}
#endregion

'Finally, comparing ITAMS and AD PCC Number to the room number on hand...'
    
For ($i = 0; $i -le ($PCCNumberArray.count - 1); $i++) {

    [int]$pct = ($i/$PCCNumberArray.count)*100
    Write-progress -Activity 'Working...' -percentcomplete $pct -currentoperation "$pct% Complete" -status 'Please Wait'
    
    if ($AssetHash[$PCCNumberArray[$i].'PCC Number'] -notmatch $PCCNumberArray[$i].'Room') {
        Write-output "$($PCCNumberArray[$i].'PCC Number'), $($PCCNumberArray[$i].'Room')" | Sort-Object | Out-File -FilePath 'C:\users\wrcrabtree\downloads\Inconsistent.csv' -Append
    }
}

$notmatchRegEx | Sort-Object | Out-File -FilePath 'C:\users\wrcrabtree\downloads\NotMatchRegEx.csv'

