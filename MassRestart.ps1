Import-Module activedirectory

$list = (Get-ADComputer -filter { Name -like "EC-l142*pc" }).Name
Write-Verbose  -Message "Found $($list.count) computers!"


function Restart-Computers($ComputerNames) {
    foreach ($computername in $list) {
        Restart-Computer -computername $computername -Force
    }
}