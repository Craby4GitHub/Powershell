workflow UninstallRemote{
    param(
    [Parameter (Mandatory = $true)]
    [string]$computer
    )

#   Disable this if you wna see errors. I have bad error handling so... I would leave it.
$ErrorActionPreference = "SilentlyContinue"

$VerbosePreference = "continue"

#   List of programs to uninstall. Could make this reference a file if this list gets too big.
$uninstalls = @('VirtualBox')

#   Get the computer(s) and display how many it found in AD
$computers = (Get-ADComputer -Filter {Name -like $computer}).Name
Write-Verbose  -Message "Trying to query $($list.count) computers found in AD"


    #   For each computer found in the list...
    foreach -parallel ($computer in $computers)  {
        if(Test-Connection -ComputerName $computer -Count 1 -Quiet){
            #   For each element in the uninstall list...
            foreach ($element in $uninstalls){
                #   Uninstall the current program
                inlinescript{
                    $program = Get-WmiObject -ComputerName $Using:computer -Class win32_product | where {$_.name -match $Using:element}
                    #$program = Get-WmiObject -ComputerName $Using:computername -Class Win32Reg_AddRemovePrograms | where {$_.displayname -match $Using:element}
                    if($program -ne $null){
                        $program.uninstall()
                        Write-Verbose -Message "Uninstalled $Using:element on $Using:computer"
                    }

                    else{
                        Write-Verbose -Message "$Using:element wasnt found on $Using:computer"
                    }
                }
                
            }
        }else{"Can't connect to $($Computer)"}
    }
}

#   Ask user for computer name.
UninstallRemote -computer (Read-Host "Enter computer name. Wildcards * are accepted. EX: EC-E513*C")
