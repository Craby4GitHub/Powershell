. (Join-Path $PSSCRIPTROOT "Get-JamfAuth.ps1")
function Get-JamfComputers {
    Get-JamfAuth
    return (((Invoke-RestMethod "https://pccjamf.jamfcloud.com/JSSResource/computers" -Method 'GET' -Headers $headers -ContentType application/json).computers).computer).Name
}