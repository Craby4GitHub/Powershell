
import-module activedirectory
#Requires -Modules activedirectory



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

Function Get-FileName($initialDirectory){
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    [void]$OpenFileDialog.ShowDialog()
    $OpenFileDialog.FileName
    $loadFileTextBox.text = $OpenFileDialog.FileName
}

Function Create-Object(){   
	param([string]$PCCNumber, [string]$Room,[string]$Campus)
    return [pscustomobject] @{     
        'PCC Number' = $PCCNumber
        'Room' = $Room
        'Campus' = $Campus
    }
}


	$matchesRegex = @()
    $notmatchRegEx = @()
    $PCCNumberArray = @()
    $AssetHash = @{}
	$ErrorComputer = @()




    

	$filePath = Get-FileName $PSScriptroot

	$correctFile = read-host 'Is' $filePath "the correct file? (Y/N)"
        
	if($correctFile -eq 'Y' -and $filePath -ne $null){
		return $inputFile = Import-Csv $filePath           
	}
	else{
		write-host "Your selection is empty or does not exist"
	}
    until(
		$correctFile -eq 'Y' -and $inputFile -ne $null
	)

	$Assets = Import-Csv -path $filePath


	$i=0
    #region Parse ITAM list to array
    ForEach($object in $Assets){
		$i++
		[int]$pct = ($i/$Assets.count)*100
		$progressbar1.Value = $pct
		if(($object.'Asset Type' -eq 'CPU') -or
		   ($object.'Asset Type' -eq 'Laptop') -or
		  (($object.'Asset Type' -eq 'Tablet') -and ($object.'Manufacturer' -eq 'Microsoft'))){
			try{
				$AssetHash.Add($object.'Barcode #', $object.'Room')
			}
			  Catch{
					Write-output "$($object.'IT #'), $($object.'Location')" | Sort-Object | Out-File -FilePath 'C:\users\wrcrabtree\downloads\Error.csv' -Append
				}
			}	
}
#endregion

    #region Import from AD

    $progressBarLabel.text = 'Pulling computers from PCC Domain...'
    

    #start-sleep -seconds 1
    $PCCArray = (Get-ADComputer -Filter {(OperatingSystem -notlike '*windows*server*')} -Properties OperatingSystem -Server PCC-Domain.pima.edu).Name

    $progressBarLabel.text = 'Pulling computers from EDU Domain...'
    $ProgressBar1.value = 50

    #start-sleep -seconds 1
    $EDUArray = (Get-ADComputer -Filter {(OperatingSystem -notlike '*windows*server*')} -Properties OperatingSystem -Server EDU-Domain.pima.edu).Name
    $ProgressBar1.value = 100
    #endregion

    $progressBarLabel.text = 'Checking AD Objects to see if they match standards...'
    $ProgressBar1.value = 0

	Function Regex-Compare {
    param([array]$Array)

    foreach ($singleComputer in $Array){
        Switch -regex ($singleComputer){
			!$15Characters {
                $Script:notmatchRegEx += ,@($singleComputer)
                break
            }
			$NormalCampus {
                $Script:matchesRegex += ,@($singleComputer)
                break
            }
            $WestCampus {
                $Script:matchesRegex += ,@($singleComputer)
                break
            }
            $DownTownCampus {
                $Script:matchesRegex += ,@($singleComputer)
                break
            }
            $VDI {
                break
            }
            $VM {
                break
            }
            default {
                $Script:notmatchRegEx += ,@($singleComputer)
				}
			}
		}
	}

    Regex-Compare -Array $EDUArray

    Regex-Compare -Array $PCCArray
    $ProgressBar1.value = 100

    #region Pull PCC Number and Room Number
    foreach ($singleComputer in $matchesRegex){
        Switch -regex ($singleComputer){
            $NormalCampus {
                $firstPass = $singleComputer -creplace "^[a-zA-Z]{2}-[a-zA-Z]\d{3}"
                $PCCNumber = $firstPass -creplace "[a-zA-Z]{2}|[vV]\d{0,}$"

                $secondPass = $singleComputer -creplace "^[a-zA-Z]{2}-"
                $Room = $secondPass -creplace "\d{6}[a-zA-Z]{2}$"

                $Campus = $singleComputer -creplace "-\w{12}$"

                

                break
            }
            $WestCampus {
                $firstPass = $singleComputer -creplace "^[a-zA-Z]{2}-[a-zA-Z]{2}\d{2}"
                $PCCNumber = $firstPass -creplace "[a-zA-Z]{2}$"

                $secondPass = $singleComputer -creplace "^[a-zA-Z]{2}-"
                $Room = $secondPass -creplace "\d{6}[a-zA-Z]{2}$"

                $Campus = $singleComputer -creplace "-\w{12}$"

				$PCCNumberArray += Create-Object -PCCNumber $PCCNumber -Room $Room -Campus $Campus

                break
            }
            $DownTownCampus {
                $firstPass = $singleComputer -creplace "^[a-zA-Z]{4}\d{3}"
                $PCCNumber = $firstPass -creplace "[a-zA-Z]{2}$"

                $secondPass = $singleComputer -creplace "^[a-zA-Z]{2}"
                $Room = $secondPass -creplace "\d{6}[a-zA-Z]{2}$"

                $Campus = $singleComputer -creplace "-\w{12}$"

				$PCCNumberArray += Create-Object -PCCNumber $PCCNumber -Room $Room -Campus $Campus

                break
            }
            default {
            }
        }
    }
#endregion

    $progressBarLabel.text = 'Finally, comparing ITAMS and AD PCC Number to the room number on hand...'
    $ProgressBar1.value = 0

    
    For($i = 0; $i -le ($PCCNumberArray.count - 1); $i++){
        $ProgressBar1.value = $i/$PCCNumberArray.count
        $progressBarLabel.text = 'test...'
		Write-Host "I got here"
        if($AssetHash[$PCCNumberArray[$i].'PCC Number'] -notmatch $PCCNumberArray[$i].'Room'){
            Write-output "$($PCCNumberArray[$i].'PCC Number'), $($PCCNumberArray[$i].'Room')" | Sort-Object | Out-File -FilePath 'C:\users\wrcrabtree\downloads\Inconsistent.csv' -Append
        }
    }

    $notmatchRegEx | Sort-Object | Out-File -FilePath 'C:\users\wrcrabtree\downloads\NotMatchRegEx.csv'
		Write-Host "I got here"
	$progressBarLabel.text = 'Done'
