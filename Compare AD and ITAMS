clear

# Import Active Directory module
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

$matchesRegex = @()
$notmatchRegEx = @()
$PCCNumberArray = @()
$AssetHash = @{}

$Assets = Import-Csv 'C:\assets (3).csv'


#region Parse ITAM list to array
ForEach($object in $Assets){
    if(($object.'Asset Type' -eq 'CPU') -or 
       ($object.'Asset Type' -eq 'Laptop') -or 
      (($object.'Asset Type' -eq 'Tablet') -and ($object.'Manufacturer' -eq 'Microsoft'))){
        
        $AssetHash.Add($object.'Barcode #', $object.'Room')

        }
}
#endregion

#region Import from AD
#$EDUArray = (Get-ADComputer -Filter {(OperatingSystem -notlike "*windows*server*")} -Properties OperatingSystem -Server EDU-Domain.pima.edu).Name
#$PCCArray = (Get-ADComputer -Filter {(OperatingSystem -notlike "*windows*server*")} -Properties OperatingSystem -Server PCC-Domain.pima.edu).Name
#endregion

Function Regex-Compare {
    param([array]$Array)

    foreach ($singleComputer in $Array){
        Switch -regex ($singleComputer){
            !$15Characters {
                $Global:notmatchRegEx += ,@($singleComputer)
                break
            }
            $NormalCampus {
                $Global:matchesRegex += ,@($singleComputer)
                break
            }
            $WestCampus {
                $Global:matchesRegex += ,@($singleComputer)
                break
            }
            $DownTownCampus {
                $Global:matchesRegex += ,@($singleComputer)
                break
            }
            $VDI {
                break
            }
            $VM {
                break
            }
            default {
                $Global:notmatchRegEx += ,@($singleComputer)
            }
        }
    }
}


Regex-Compare -Array $EDUArray

Regex-Compare -Array $PCCArray


foreach ($singleComputer in $matchesRegex){
        Switch -regex ($singleComputer){
            $NormalCampus {
                $firstPass = $singleComputer -creplace "^[a-zA-Z]{2}-[a-zA-Z]\d{3}"
                $PCCNumber = $firstPass -creplace "[a-zA-Z]{2}|[vV]\d{0,}$"

                $secondPass = $singleComputer -creplace "^[a-zA-Z]{2}-"
                $Room = $secondPass -creplace "\d{6}[a-zA-Z]{2}$"


                $PCCNumberArray += New-Object PsObject -Property @{
                    'PCC Number' = $PCCNumber
                    'Room' = $Room
                }

                break
            }
            $WestCampus {
                $firstPass = $singleComputer -creplace "^[a-zA-Z]{2}-[a-zA-Z]{2}\d{2}"
                $PCCNumber = $firstPass -creplace "[a-zA-Z]{2}$"

                $secondPass = $singleComputer -creplace "^[a-zA-Z]{2}-"
                $Room = $secondPass -creplace "\d{6}[a-zA-Z]{2}$"
                
                $PCCNumberArray += New-Object PsObject -Property @{
                    'PCC Number' = $PCCNumber
                    'Room' = $Room
                }
                
                break
            }
            $DownTownCampus {
                $firstPass = $singleComputer -creplace "^[a-zA-Z]{4}\d{3}"
                $PCCNumber = $firstPass -creplace "[a-zA-Z]{2}$"

                $secondPass = $singleComputer -creplace "^[a-zA-Z]{2}"
                $Room = $secondPass -creplace "\d{6}[a-zA-Z]{2}$"
                
                $PCCNumberArray += New-Object PsObject -Property @{
                    'PCC Number' = $PCCNumber
                    'Room' = $Room
                }
                
                break
            }
            default {
            }
        }
    }



For($i = 0; $i -le ($PCCNumberArray.count - 1); $i++){
    Write-Progress -Activity 'Compare status' -percentComplete ($i / $PCCNumberArray.count * 100)
    if($AssetHash[$PCCNumberArray[$i].'PCC Number'] -notmatch $PCCNumberArray[$i].'Room'){
        Write-output "$($PCCNumberArray[$i].'PCC Number'), $($PCCNumberArray[$i].'Room')" | Sort-Object | Out-File -FilePath 'C:\Inconsistent.csv' -Append
    }
}
$notmatchRegEx | Sort-Object | Out-File -FilePath 'C:\NotMatchRegEx.csv'
