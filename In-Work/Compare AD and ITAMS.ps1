Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region Begin GUI{ 

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '370,309'
$Form.text                       = "PCC-IT: AD and ITAMS compare"
$Form.TopMost                    = $false

$startButton                     = New-Object system.Windows.Forms.Button
$startButton.text                = "Start"
$startButton.width               = 66
$startButton.height              = 30
$startButton.location            = New-Object System.Drawing.Point(9,102)
$startButton.Font                = 'Microsoft Sans Serif,10'

$loadFileButton                  = New-Object system.Windows.Forms.Button
$loadFileButton.text             = "Load File"
$loadFileButton.width            = 69
$loadFileButton.height           = 30
$loadFileButton.location         = New-Object System.Drawing.Point(7,24)
$loadFileButton.Font             = 'Microsoft Sans Serif,10'

$loadFileTextBox                 = New-Object system.Windows.Forms.TextBox
$loadFileTextBox.multiline       = $false
$loadFileTextBox.width           = 239
$loadFileTextBox.height          = 20
$loadFileTextBox.location        = New-Object System.Drawing.Point(96,28)
$loadFileTextBox.Font            = 'Microsoft Sans Serif,10'

$loadFileBox                     = New-Object system.Windows.Forms.Groupbox
$loadFileBox.height              = 66
$loadFileBox.width               = 347
$loadFileBox.text                = "Load File"
$loadFileBox.location            = New-Object System.Drawing.Point(9,24)

$WinForm1                        = New-Object system.Windows.Forms.Form
$WinForm1.ClientSize             = '546,206'
$WinForm1.text                   = "PCC-IT: AD and ITAMS compare"
$WinForm1.TopMost                = $false

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 290
$ProgressBar1.height             = 60
$ProgressBar1.location           = New-Object System.Drawing.Point(24,60)

$progressBarLabel                = New-Object system.Windows.Forms.Label
$progressBarLabel.text           = "Progress:"
$progressBarLabel.AutoSize       = $true
$progressBarLabel.width          = 25
$progressBarLabel.height         = 10
$progressBarLabel.location       = New-Object System.Drawing.Point(23,29)
$progressBarLabel.Font           = 'Microsoft Sans Serif,10'

$Groupbox1                       = New-Object system.Windows.Forms.Groupbox
$Groupbox1.height                = 137
$Groupbox1.width                 = 348
$Groupbox1.text                  = "Current Work, I dunno..."
$Groupbox1.location              = New-Object System.Drawing.Point(9,161)

$Form.controls.AddRange(@($startButton,$loadFileBox,$Groupbox1))
$loadFileBox.controls.AddRange(@($loadFileButton,$loadFileTextBox))
$Groupbox1.controls.AddRange(@($ProgressBar1,$progressBarLabel))

#region gui events {
$loadFileButton.Add_MouseUp({ Get-FileName })
$startButton.Add_MouseUp({ Start-Code })
#endregion events }

#endregion GUI }



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

Function Start-Code(){

    $matchesRegex = @()
    $notmatchRegEx = @()
    $PCCNumberArray = @()
    $AssetHash = @{}

    $Assets = Import-Csv -path $loadFileTextBox.Text
    $progressBarLabel.text = 'Converting file to array...'
    $ProgressBar1.value = 0

    #region Parse ITAM list to array
    ForEach($object in $Assets){
    #For($i=0;$i -le $Assets.Count;$i++){
    if(($object.'Asset Type' -eq 'CPU') -or 
       ($object.'Asset Type' -eq 'Laptop') -or 
      (($object.'Asset Type' -eq 'Tablet') -and ($object.'Manufacturer' -eq 'Microsoft'))){
        
        $AssetHash.Add($object.'Barcode #', $object.'Room')

        }
    $ProgressBar1.value = 100
}
#endregion
    
    #region Import from AD

    $progressBarLabel.text = 'Pulling computers from PCC Domain...'
    $ProgressBar1.value = 0
    
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

                $PCCNumberArray += New-Object PsObject -Property @{
                    'PCC Number' = $PCCNumber
                    'Room' = $Room
                    'Campus' = $Campus
                }

                break
            }
            $WestCampus {
                $firstPass = $singleComputer -creplace "^[a-zA-Z]{2}-[a-zA-Z]{2}\d{2}"
                $PCCNumber = $firstPass -creplace "[a-zA-Z]{2}$"

                $secondPass = $singleComputer -creplace "^[a-zA-Z]{2}-"
                $Room = $secondPass -creplace "\d{6}[a-zA-Z]{2}$"
                
                $Campus = $singleComputer -creplace "-\w{12}$"

                $PCCNumberArray += New-Object PsObject -Property @{
                    'PCC Number' = $PCCNumber
                    'Room' = $Room
                    'Campus' = $Campus
                }
                
                break
            }
            $DownTownCampus {
                $firstPass = $singleComputer -creplace "^[a-zA-Z]{4}\d{3}"
                $PCCNumber = $firstPass -creplace "[a-zA-Z]{2}$"

                $secondPass = $singleComputer -creplace "^[a-zA-Z]{2}"
                $Room = $secondPass -creplace "\d{6}[a-zA-Z]{2}$"
                
                $Campus = $singleComputer -creplace "-\w{12}$"

                $PCCNumberArray += New-Object PsObject -Property @{
                    'PCC Number' = $PCCNumber
                    'Room' = $Room
                    'Campus' = $Campus
                }
                
                break
            }
            default {
            }
        }
    }
#endregion

    $progressBarLabel.text = 'Finally, comparing ITAMS and AD PCC Number to the room number on hand...'
    $ProgressBar1.value = 0

    #ForEach($PCCValue in $PCCNumberArray){

    For($i = 0; $i -le ($PCCNumberArray.count - 1); $i++){
        $ProgressBar1.value = $i/$PCCNumberArray.count
        $progressBarLabel.text = 'test...'
        if($AssetHash[$PCCNumberArray[$i].'PCC Number'] -notmatch $PCCNumberArray[$i].'Room'){
            Write-output "$($PCCNumberArray[$i].'PCC Number'), $($PCCNumberArray[$i].'Room')" | Sort-Object | Out-File -FilePath 'C:\Inconsistent.csv' -Append
        }
    }

    $notmatchRegEx | Sort-Object | Out-File -FilePath 'C:\NotMatchRegEx.csv'
}

[void]$Form.ShowDialog()