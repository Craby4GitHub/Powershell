Import-Module activedirectory
cls
$list = (Get-ADComputer -filter {Name -like "WC-H305*CC"}).Name


foreach ($computer in $list){
    Invoke-Command -ComputerName $computer -ScriptBlock{get-childitem "C:\users\*\desktop" -Exclude "*.lnk" | remove-item -Recurse -Force -Confirm:$false}
}
