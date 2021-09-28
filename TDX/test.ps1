# Load TDX API functions
. "$((get-item $PSScriptRoot).Parent.FullName)\JAMF\JAMF-API.ps1"
#Get-ChildItem -Filter '*.ps1' | Foreach { . $_.FullName }
$test = "$((get-item $PSScriptRoot).Parent.FullName)\JAMF\JAMF-API.ps1"
write-host $test


Function Test-SQLServer {
    Param(
        $Server
    )

    Try { 
        $ReturnedInfo = Cmdlet-ThatChecksstuff -ErrorAction Stop
        $Ports = $ReturnedInfo.Ports
        $Notes = $ReturnedInfo.ExtraInfo
    }
    Catch {
        $Ports = "Error retrieving ports"
        $Notes = "Inspect this host again"
    }
    [PSCustomObject]@{
        Server = $Server
        Ports  = $Ports
        Notes  = $Notes
    }
}

#$computers = Get-CMDevice -CollectionName 'Agent Installed' #| Select-Object Name, ResourceID, LastDDR