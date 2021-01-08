function Get-ADComputers($ComputerName) {
    Import-Module activedirectory
    return (Get-ADComputer -filter { Name -like $ComputerName }).Name
}
function Get-MappedDrives($ComputerName) {
    #Ping remote machine, continue if available
    if (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
        #Get remote explorer session to identify current user
        $explorer = Get-WmiObject -ComputerName $ComputerName -Class win32_process | ? { $_.name -eq "explorer.exe" }
      
        #If a session was returned check HKEY_USERS for Network drives under their SID
        if ($explorer) {
            $Hive = [long]$HIVE_HKU = 2147483651
            $sid = ($explorer.GetOwnerSid()).sid
            $owner = $explorer.GetOwner()
            $RegProv = get-WmiObject -List -Namespace "root\default" -ComputerName $ComputerName | Where-Object { $_.Name -eq "StdRegProv" }
            $DriveList = $RegProv.EnumKey($Hive, "$($sid)\Network")
        
            #If the SID network has mapped drives iterate and report on said drives
            if ($DriveList.sNames.count -gt 0) {
                "$($owner.Domain)\$($owner.user) on $($ComputerName)"
                foreach ($drive in $DriveList.sNames) {
                    "$($drive)`t$(($RegProv.GetStringValue($Hive, "$($sid)\Network\$($drive)", "RemotePath")).sValue)"
                }
            }
            else { "No mapped drives on $($ComputerName)" }
        }
        else { "explorer.exe not running on $($ComputerName)" }
    }
    else { "Can't connect to $($ComputerName)" }
}

#In Work
function Add-ComputersToSecurityGroupFromOU($SecurityGroup) {
    Import-Module activedirectory

    $stagingOU = 'OU=West,OU=Staging,DC=PCC-Domain,DC=pima,DC=edu'
    $laptopOU = 'OU=COVID-19 Laptops,OU=Computers,OU=IT Services,OU=West,OU=PCC,DC=PCC-Domain,DC=pima,DC=edu'
    $SecurityGroup
    
    Get-ADComputer -SearchBase $stagingOU -Filter { Name -like "WC-R016*SN" } | ForEach-Object { Move-ADObject -Identity $_.DistinguishedName -TargetPath $laptopOU }
    
    $mbamSecurityGroup = Get-ADGroup -Filter { Name -like "WC-MBAM-Laptop-Computers" }
    
    Get-ADComputer -SearchBase $laptopOU -Filter * | ForEach-Object { Add-ADGroupMember -Identity $mbamSecurityGroup -Members $_.DistinguishedName }
    
}