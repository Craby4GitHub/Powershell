. (Join-Path $PSSCRIPTROOT "JAMF-API.ps1")

$SerialNumbers = get-content -Path 'C:\users\wrcrabtree\Desktop\sn.csv'

Update-JamfMobileGroups -ID 134 -SerialNumbers $SerialNumbers
