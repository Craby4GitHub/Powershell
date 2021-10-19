Import-Module ActiveDirectory
$userName = ''
$computerOU = "DC=edu-domain,DC=pima,DC=edu"

$computers = (Get-ADComputer -Filter * -SearchBase $computerOU).name
$computersWithComma = $computers -join ","
Set-ADUser -Identity $userName -LogonWorkstations $computersWithComma