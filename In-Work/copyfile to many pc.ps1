$VerbosePreference = "continue"
$list = (Get-ADComputer -Filter {Name -like "WC-H306139972CC"}).name
Write-Verbose  -Message "Found $($list.count) computers"
$ErrorActionPreference = "SilentlyContinue"

New-PSDrive -Name P -PSProvider FileSystem -Root \\ec-nas\EC-IT\IntelWin10\
$i = 0
foreach ($computername in $list){

    Write-Progress -id 1 -Activity 'Status' -percentComplete ($i / $list.count * 100)
    if(Get-WmiObject -Class win32_networkadapter -ComputerName $computername | where {($_.Name -like "Intel(R) Ethernet Connection I217-LM") -or ($_.Name -like "Intel(R) 82579LM Gigabit Network Connection")}){
    
        #Copy-Item -Recurse -Filter *.* -Path \'\ec-nas\EC-IT\IntelWin10' "\\$computername\c$\temp\IntelDrivers" -force

        Write-Host "Copied to $computername"	
        Invoke-Command -ComputerName $computername -ScriptBlock{P:\APPS\PROSETDX\Winx64\DxSetup.exe BD=1 DMIX=1 ANS=1}
        #Invoke-Command -ComputerName $computername -ScriptBlock{C:\temp\IntelDrivers\APPS\PROSETDX\Winx64\DxSetup.exe BD=1 DMIX=1 ANS=1 /q}
    }
    $i++
    
}
Remove-PSDrive P
sleep -Seconds 5