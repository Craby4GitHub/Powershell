$VerbosePreference = "continue"
$list = (Get-ADComputer -filter {Name -like "EC-L142*LC"}).Name
Write-Verbose  -Message "We found $($list.count) computers!"
$ErrorActionPreference = "SilentlyContinue"

foreach ($computername in $list)
{

$app = Get-WmiObject -class win32_product -Filter "name like '%wepa%'"


$app.Uninstall()



}