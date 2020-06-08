$Credentials = Get-Credential

$Driver = Start-SeChrome -Incognito
Enter-SeUrl -Driver $Driver -Url "https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=402:1"

# Login to ITAM
$usernameElement = Find-SeElement -Driver $Driver -Wait -Timeout 10 -Id 'P101_USERNAME'
$passwordElement = Find-SeElement -Driver $Driver -Id 'P101_PASSWORD'
$loginButtonElement = Find-SeElement -Driver $Driver -Id 'P101_LOGIN'

Send-SeKeys -Element $usernameElement -Keys $Credentials.UserName
Send-SeKeys -Element $passwordElement -Keys $Credentials.GetNetworkCredential().Password
Invoke-SeClick -Element $loginButtonElement

# Navigate to Upload Assests Page
$uploadAssestsXPath = '/html/body/form/div[1]/div[4]/div/div[3]/div/a'
Find-SeElement -XPath $uploadAssestsXPath -Driver $Driver | Invoke-SeClick -Driver $driver

# Upload File
$importFromUploadRadioButton = Find-SeElement -Driver $Driver -Id 'P10_IMPORT_FROM_0'
$importFromUploadFileSelectionButton = Find-SeElement -Driver $Driver -Wait -Timeout 10 -Id 'P10_FILE_NAME'

Invoke-SeClick -Element $importFromUploadRadioButton -Driver $driver

$filePath = "C:\Users\Wrcrabtree\Desktop\test.csv"
Send-SeKeys -Element $importFromUploadFileSelectionButton -Keys $filePath

# Remove the quotes
(Find-SeElement -Driver $Driver -Id 'P10_ENCLOSED_BY').clear() 

$nextButtonXPath = '/html/body/form/div[5]/table/tbody/tr/td[1]/div[2]/div[1]/div/div[2]/button[2]'
Find-SeElement -XPath $nextButtonXPath -Driver $Driver | Invoke-SeClick -Driver $driver

for ($i = 1; $i -lt 13; $i++) {
    $uploadedColumn = Find-SeElement -Driver $Driver -Id "id3_$i"
    $ITAMColumn = Find-SeElement -Driver $Driver -Id "id1_$i"
    Get-SeSelectionOption -Element $ITAMColumn -ByPartialText $uploadedColumn.Text
}
#Start-Sleep -Seconds 10
#Stop-SeDriver -Driver $Driver