Import-Module ActiveDirectory
$userName = ''
$computerOU = @('')
$computerList = @()

foreach ($ou in $computerOU) {
    $computerList += (Get-ADComputer -Filter * -SearchBase $ou).name
}

Set-ADUser -Identity $userName -LogonWorkstations $($computerList -join ",")