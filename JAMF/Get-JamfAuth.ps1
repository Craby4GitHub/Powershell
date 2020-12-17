function Get-JamfAuth {
    $Creds = Get-Credential
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Authorization", "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($Creds.UserName):$($Creds.GetNetworkCredential().Password)")))
}