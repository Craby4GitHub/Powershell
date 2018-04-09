
$VerbosePreference = "continue"
$list = (Get-ADComputer -filter {Name -like "EC-L123*LC"}).Name
Write-Verbose  -Message "Trying to query $($list.count) computers found in AD"
$ErrorActionPreference = "SilentlyContinue"
foreach ($computername in $list)
{
    Remove-Item \\$computername\c$\windows\ccmcache\ -Recurse -Force
    if ($? -eq $true)
    {
    Write-Verbose -Message "Deleting tasks for $computername"
    }
    else
    {

    Write-Verbose -Message "Had an issue with $computername"

    }

}