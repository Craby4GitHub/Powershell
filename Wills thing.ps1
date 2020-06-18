$list = (Get-ADComputer -filter {Name -like "EC-L116*"}).Name
Write-Verbose  -Message "Trying to query $($list.count) computers found in AD"

foreach ($computername in $list){
    Get-ScheduledTask -taskname *adobe* | select-object state | Disable-ScheduledTask

    if ($? -eq $true){
        Write-Verbose -Message "I found $($?.count) tasks for $computername"
    }
    
    else{
        Disable-ScheduledTask
    }
}