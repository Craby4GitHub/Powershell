Import-Module activedirectory

$list = (Get-ADComputer -filter {Name -like "wc-d206*ln"}).Name
Write-Verbose  -Message "Found $($list.count) computers!"

foreach ($computername in $list){
    mstsc /v:$computername
}
