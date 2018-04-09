$VerbosePreference = "SilentlyContinue"
$logfilepath = "$home\Desktop\TasksLog.csv"
$ErrorActionPreference = "SilentlyContinue"

import-module activedirectory

$list = (Get-ADComputer -filter {Name -like "WC-C344*PC"}).name
$apps = get-content -Path \\ec-nas\EC-IT\Scripts\StartMenuItems.txt
$StartupPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\"


foreach ($computer in $list)
    {
    #"ComputerName: $computer"
    foreach ($app in $apps){
    Invoke-Command -computername $computer {
    #if(test-path $StartupPath$app){
        #Write-Host "Found $app on $computer"
        Remove-Item -Path $StartupPath$app -Recurse}
        }
}
#}