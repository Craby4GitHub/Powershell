Import-Module activedirectory

$stagingOU = 'CN=Computers,DC=edu-domain,DC=pima,DC=edu'
$laptopOU = 'OU=COVID-19 Laptops,OU=Computers,OU=IT Services,OU=West,OU=PCC,DC=PCC-Domain,DC=pima,DC=edu'


Get-ADComputer -SearchBase $laptopOU -Filter {Name -like "WC-R016*SN"} | ForEach-Object {
    Add-ADGroupMember -Identity 'WC-MBAM-Laptop-Computers' -Members $_.DistinguishedName -WhatIf
}

