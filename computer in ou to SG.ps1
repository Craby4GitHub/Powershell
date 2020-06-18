Import-Module activedirectory

$stagingOU = 'OU=West,OU=Staging,DC=PCC-Domain,DC=pima,DC=edu'
$laptopOU = 'OU=COVID-19 Laptops,OU=Computers,OU=IT Services,OU=West,OU=PCC,DC=PCC-Domain,DC=pima,DC=edu'

Get-ADComputer -SearchBase $stagingOU -Filter {Name -like "WC-R016*SN"} | ForEach-Object {Move-ADObject -Identity $_.DistinguishedName -TargetPath $laptopOU}

$mbamSecurityGroup = Get-ADGroup -Filter {Name -like "WC-MBAM-Laptop-Computers"}

Get-ADComputer -SearchBase $laptopOU -Filter * | ForEach-Object {Add-ADGroupMember -Identity $mbamSecurityGroup -Members $_.DistinguishedName}