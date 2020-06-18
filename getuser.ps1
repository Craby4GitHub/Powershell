
$pullSIDs = Get-WmiObject -computername ec-l116137256lc Win32_UserProfile | Select-Object -ExpandProperty "sid"
foreach ($pulledSID in $pullSIDs){
$objSID = New-Object System.Security.Principal.SecurityIdentifier `
    ($pulledSID)
$objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
$objUser.Value

}
