import-module activedirectory
$list = (Get-ADComputer -filter {Name -like "MS-A18*C"}).Name
Write-Verbose  -Message "Trying to query $($list.count) computers found in AD"

function Get-MappedDrives($ComputerName){
    #Ping remote machine, continue if available
    if(Test-Connection -ComputerName $ComputerName -Count 1 -Quiet){
      #Get remote explorer session to identify current user
      $explorer = Get-WmiObject -ComputerName $ComputerName -Class win32_process | ?{$_.name -eq "explorer.exe"}
      
      #If a session was returned check HKEY_USERS for Network drives under their SID
      if($explorer){
        $Hive = [long]$HIVE_HKU = 2147483651
        $sid = ($explorer.GetOwnerSid()).sid
        $owner  = $explorer.GetOwner()
        $RegProv = get-WmiObject -List -Namespace "root\default" -ComputerName $ComputerName | Where-Object {$_.Name -eq "StdRegProv"}
        $DriveList = $RegProv.EnumKey($Hive, "$($sid)\Network")
        
        #If the SID network has mapped drives iterate and report on said drives
        if($DriveList.sNames.count -gt 0){
          "$($owner.Domain)\$($owner.user) on $($ComputerName)"
          foreach($drive in $DriveList.sNames){
            "$($drive)`t$(($RegProv.GetStringValue($Hive, "$($sid)\Network\$($drive)", "RemotePath")).sValue)"
          }
        }else{"No mapped drives on $($ComputerName)"}
      }else{"explorer.exe not running on $($ComputerName)"}
    }else{"Can't connect to $($ComputerName)"}
  }



foreach ($computername in $list)  { 
    Get-MappedDrives $computername
}