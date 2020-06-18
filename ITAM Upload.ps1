# https://www.powershellgallery.com/packages/Selenium/3.0.0

# To Add:
    # remove xpath for clicking next

    $Credentials = Get-Credential

    $Driver = Start-SeChrome -Incognito
    Enter-SeUrl -Driver $Driver -Url "https://pimaapps.pima.edu/pls/htmldb_pdat/f?p=402:1"

    #region Login to ITAM
    $usernameElement = Find-SeElement -Driver $Driver -Wait -Timeout 10 -Id 'P101_USERNAME'
    $passwordElement = Find-SeElement -Driver $Driver -Id 'P101_PASSWORD'
    $loginButtonElement = Find-SeElement -Driver $Driver -Id 'P101_LOGIN'

    Send-SeKeys -Element $usernameElement -Keys $Credentials.UserName
    Send-SeKeys -Element $passwordElement -Keys $Credentials.GetNetworkCredential().Password
    Invoke-SeClick -Element $loginButtonElement

    #ADD LOGIN SUCCESS/FAILURE CHECK

    # Click Next
    Find-SeElement -XPath '/html/body/form/div[1]/div[4]/div/div[3]/div/a' -Driver $Driver | Invoke-SeClick -Driver $Driver
    #endregion

    #region Upload File
    $importFromUploadRadioButton = Find-SeElement -Driver $Driver -Id 'P10_IMPORT_FROM_0'
    $importFromUploadFileSelectionButton = Find-SeElement -Driver $Driver -Wait -Timeout 10 -Id 'P10_FILE_NAME'
    Invoke-SeClick -Element $importFromUploadRadioButton -Driver $driver
    $filePath = "C:\Users\Wrcrabtree\Downloads\ITAM Upload - Sheet1.csv"
    Send-SeKeys -Element $importFromUploadFileSelectionButton -Keys $filePath

    # Remove the quotes
    (Find-SeElement -Driver $Driver -Id 'P10_ENCLOSED_BY').clear() 

    # Click Next
    Find-SeElement -XPath '/html/body/form/div[5]/table/tbody/tr/td[1]/div[2]/div[1]/div/div[2]/button[2]' -Driver $Driver | Invoke-SeClick -Driver $Driver
    #endregion

    #region Connect uploaded CSV headers to ITAM headers
    # Current Issue: ByPartialText selects first 'like' column. Mis-selects Type-> MODEL_SUBTYPE, Status -> ENCRYPTION_STATUS
    for ($i = 1; $i -lt 14; $i++) {
        $uploadedColumn = Find-SeElement -Driver $Driver -Id "id3_$i"
        $ITAMColumn = Find-SeElement -Driver $Driver -Id "id1_$i"
        Get-SeSelectionOption -Element $ITAMColumn -ByPartialText $uploadedColumn.Text
    }

    # Click Next
    Find-SeElement -XPath '/html/body/form/div[5]/table/tbody/tr/td[1]/div[2]/div[1]/div/div[2]/button[3]' -Driver $Driver | Invoke-SeClick -Driver $Driver
    #endregion

    # Manually verify upload then run backup script
    #Stop-SeDriver -Driver $Driver
