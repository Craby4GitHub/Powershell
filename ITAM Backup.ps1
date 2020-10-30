# https://www.powershellgallery.com/packages/Selenium/3.0.0

$Credentials = Get-Credential

$Driver = Start-SeFirefox -PrivateBrowsing
Enter-SeUrl -Driver $Driver -Url "https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=402:26:4099767476960:CSV::::"

#region Login to ITAM
$usernameElement = Find-SeElement -Driver $Driver -Wait -Timeout 10 -Id 'P101_USERNAME'
$passwordElement = Find-SeElement -Driver $Driver -Id 'P101_PASSWORD'
$loginButtonElement = Find-SeElement -Driver $Driver -Id 'P101_LOGIN'

Send-SeKeys -Element $usernameElement -Keys $Credentials.UserName
Send-SeKeys -Element $passwordElement -Keys $Credentials.GetNetworkCredential().Password
Invoke-SeClick -Element $loginButtonElement
#endregion
#Stop-SeDriver -Driver $Driver