$dir = get-childitem $(Split-Path $PSScriptRoot -Parent) | Get-Item | Select-Object -ExpandProperty Name
write-host $dir 












# select random staff sccm group across multiple. use the staff collections in client console
# create tickets to techincain with the properties: computer, primary user, ect
# assign computer to sccm update test group
# pull reports on how the test computer are doing with sccm reportsa , ticket updates(?)

$computers = Get-CMDevice -CollectionName 'Agent Installed' #| Select-Object Name, ResourceID, LastDDR