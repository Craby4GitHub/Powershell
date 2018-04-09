Import-Module activedirectory
$VerbosePreference = "SilentlyContinue"
$list = (Get-ADComputer -filter {Name -like "EC-l142*pc"}).Name
Write-Verbose  -Message "Found $($list.count) computers!"
$logfilepath = "$home\Desktop\TasksLog.csv"
$ErrorActionPreference = "SilentlyContinue"

foreach ($computername in $list)
{
    Restart-Computer -computername $computername -Force
}
