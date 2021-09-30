$computers = Get-ADComputer -Filter { Name -like "EC*C" }
foreach ($computer in $computers) {
    # Proress bar to show how far along we are
    [int]$pct = ($computers.IndexOf($computer) / $computers.Count) * 100
    Write-progress -Activity '...' -PercentComplete $pct -status "$pct% Complete"
    try {
        if (Test-Connection -ComputerName $computer.name -Count 3 -quiet) {
            if ($null -eq (Get-HotFix -ComputerName $computer.name -ErrorAction SilentlyContinue | Where-Object { ($_.HotFixID -eq 'KB5005566') -or ($_.HotFixID -eq 'KB5005565') })) {
                Write-Host $computer.name
                "Fail,$($computer.name)" | Out-File -filePath "$PSScriptRoot\computers.csv" -Append -Encoding ascii
            }
            else {
                "Pass,$($computer.name)" | Out-File -filePath "$PSScriptRoot\computers.csv" -Append -Encoding ascii
            }
        }
        else {
            "Offline,$($computer.name)" | Out-File -filePath "$PSScriptRoot\computers.csv" -Append -Encoding ascii
        }
    }
    catch {
        "Issue,$($computer.name)" | Out-File -filePath "$PSScriptRoot\computers.csv" -Append -Encoding ascii
    } 
}