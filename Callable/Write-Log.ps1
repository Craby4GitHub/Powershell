function Write-Log {
    
    param (
        [ValidateSet('ERROR', 'INFO', 'VERBOSE', 'WARN')]
        [Parameter(Mandatory = $true)]
        [string]$level,

        [Parameter(Mandatory = $true)]
        [string]$message
    )
    	
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $logString = "$timeStamp, $level, $message"

    Add-Content -Path $logFile -Value $logString -Force

    if ($level -eq 'WARN' -or 'ERROR') {
        Out-Host -InputObject "$logString"
    }
}