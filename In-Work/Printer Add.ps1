#$param ([string[]]$printer = $null)

$Printer = Read-Host "Please list printers seperated by a comma"

##Map Printer
foreach($Printer in $Printers){
	RUNDLL32 PRINTUI.DLL,PrintUIEntry /ga /z /n\\EC-Print\$Printer
	}