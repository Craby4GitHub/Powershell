<#
Will Crabtree
Version: I dunno how to version things. :D

Compares ITAMs to Active Directry to find computers with incorrect names.

Download the ITAMs CSV from here: https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=402:1
    Login -> Assets -> Actions -> Download -> CSV

When ran, this will save 'Error.csv', 'NonStandard Name.csv' and 'Inconsistent.csv' to the users desktop.

'Error.csv' will contain mainly computers that have no PCC number.

'NonStandard Name.csv' will contain computer names that are not following the PCC Naming Convention

'Inconsistent.csv' will contain the computer found to be inconsistent between ITAMs and AD
    Headers:
        PCC Number: Its the PCC number. Yup
        AD Room: The room that is set in AD
        ITAMs Room: The room that is set in ITAMs
        Campus: The campus set in AD
#>

Clear-Host

#Requires -Modules activedirectory
import-module activedirectory

Function Main {
$matchesRegex = @()
$notmatchRegEx = @()
$PCCNumberArray = @()
$InconsistentArray = @()
$AssetHash = @{}

'Loading selected file into an array'
#region Parse ITAM list to array
ForEach ($object in Get-File) {
    Write-progress -Activity 'Loading selected file into an array...'
    if (($object.'Asset Type' -eq 'CPU') -or
        ($object.'Asset Type' -eq 'Laptop') -or
       (($object.'Asset Type' -eq 'Tablet') -and ($object.'Manufacturer' -eq 'Microsoft'))) {
        try {
            $AssetHash.Add($object.'Barcode #', $object.'Room')
        }
        Catch {
            Write-output "$($object.'IT #'), $($object.'Location'), $($_.Exception.Message)" | Sort-Object | Out-File -FilePath "$($env:USERPROFILE)\desktop\Error.csv" -Append -Encoding ASCII
        }
    }	
}
#endregion

'Loading AD Computers'
#region Import from AD
Write-progress -Activity 'Getting computers from EDU and PCC Domain...'
$PCCArray = (Get-ADComputer -Filter {(OperatingSystem -notlike '*windows*server*')} -Properties OperatingSystem -Server PCC-Domain.pima.edu).Name
  
$EDUArray = (Get-ADComputer -Filter {(OperatingSystem -notlike '*windows*server*')} -Properties OperatingSystem -Server EDU-Domain.pima.edu).Name

#endregion

#region Regex
# 15 Characters
$15Characters = [regex]"\w{15}"

# 2 letter campus code - 1 or 2 Building letter 8-9 numbers(2-3 for room and 6 for PCC) 2 letter computer type
# new ^([a-z]{4}|(([a-z]{2}|\d{2})-[a-z]{1,2}))\d{8,9}[a-z]{2}$
# gotta update region 'Pull PCC Number and Room Number' before using this ^
$NormalCampus = [regex]"^([a-z]{4}|(([a-z]{2}|\d{2})-([a-z]|[a-z]{2})))(\d{8}|\d{9})[a-z]{2}$"

# 2 letter campus code 2 Building letter 9 numbers(3 for room and 6 for PCC) 2 letter computer type
$DownTownCampus = [regex]"^[a-z]{4}\d{9}[a-z]{2}$"

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
    # Naming convention
    # https://docs.google.com/spreadsheets/d/1gLkgjxNlxwbNizH_EsQY_-ARmaStVOYra1pwcxIQvgM/edit#gid=0
    Write-progress -Activity 'Comparing computer names to PCC naming convention...' -PercentComplete $pct -status "$pct% Complete"
    Switch -regex ($singleComputer) {
        !$15Characters {
            $notmatchRegEx += , @($singleComputer)
            break
        }
        $NormalCampus {
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
    Write-progress -Activity 'Extracting Campus, Room number and PCC number from the AD computer names...' -PercentComplete $pct -status "$pct% Complete"
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
    Write-progress -Activity 'Comparing AD computer Room and PCC number to corresponding PCC number in ITAMS...' -percentcomplete $pct -status "$pct% Complete"
    
    if ($AssetHash[$PCCNumberArray[$i].'PCC Number'] -notmatch $PCCNumberArray[$i].'Room') {
        $InconsistentArray += "$($PCCNumberArray[$i].'PCC Number'), $($PCCNumberArray[$i].'Room'), $($AssetHash[$PCCNumberArray[$i].'PCC Number']), $($PCCNumberArray[$i].'Campus')"
    }
}
#endregion
New-Item -Path "$($env:USERPROFILE)\desktop\Inconsistent.csv" -value "PCC Number,AD Room Number,ITAMs Room Number,Campus`n" -Force | Out-Null
$InconsistentArray | Out-File -FilePath "$($env:USERPROFILE)\desktop\Inconsistent.csv" -Append -Encoding ASCII
$notmatchRegEx | Sort-Object | Out-File -FilePath "$($env:USERPROFILE)\desktop\NonStandard Name.csv" -Force
}


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

main