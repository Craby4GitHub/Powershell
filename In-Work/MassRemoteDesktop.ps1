Import-Module activedirectory
$VerbosePreference = "SilentlyContinue"
$list = (Get-ADComputer -filter {Name -like "wc-d206*ln"}).Name
Write-Verbose  -Message "Found $($list.count) computers!"
$logfilepath = "$home\Desktop\TasksLog.csv"
$ErrorActionPreference = "SilentlyContinue"

foreach ($computername in $list){
    mstsc /v:$computername
}
