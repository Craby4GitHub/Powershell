import-module activedirectory
$list = (Get-ADComputer -filter {Name -like "ec-e618*"}).Name
Write-Verbose  -Message "Trying to query $($list.count) computers found in AD"


foreach ($computername in $list)  {  
([WMI]'').ConvertToDateTime((Get-WmiObject Win32_OperatingSystem).InstallDate)
(Get-WmiObject -Class Win32_Product -Filter {name LIKE "CC 2015"}).Uninstall()

if ($? -eq $true){

    Write-Verbose -Message "Uninstalled program for $computername"

    }
    else{

    Write-Verbose -Message "Had an issue with $computername"

    }

}