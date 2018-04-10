workflow UninstallPrograms
{
#   its a parameter. Yup.
param
(
[Parameter (Mandatory = $true)]
[string]$computer
)

#   Disable this if you wna see errors. I have bad error handling so... I would leave it.
$ErrorActionPreference = "SilentlyContinue"

$VerbosePreference = "continue"

#   List of programs to uninstall. Could make this reference a file if this list gets too big.
$uninstalls = @("Search App by Ask", "Shopping App by Ask", "Avery Toolbar", "Google Toolbar for Internet Explorer", "Avery Teoma Search App", "Ask Toolbar", "eiPower Saver Agent", "Ask Shopping Toolbar")

#   Get the computer(s) and display how many it found in AD
$list = (Get-ADComputer -Filter {Name -like $computer}).Name
Write-Verbose  -Message "Trying to query $($list.count) computers found in AD"


    #   For each computer found in the list...
    foreach -parallel ($computername in $list)  
    {
    
        #   For each element in the uninstall list...
        foreach ($element in $uninstalls)
        {
            #   Uninstall the current program
            inlinescript
            {
                $program = Get-WmiObject -ComputerName $Using:computername -Class win32_product | where {$_.name -like $Using:element}
                if($program -ne $null)
                {
                    $program.uninstall()
                    Write-Verbose -Message "Uninstalled $Using:element on $Using:computername"
                }

                else
                {
                    Write-Verbose -Message "$Using:element wasnt found on $Using:computername"
                }
            }
               
        }
    }
}

#   Ask user for computer name.
UninstallPrograms -computer (Read-Host "Enter computer name. Wild card * is accepted. EX: EC-E513*C")
