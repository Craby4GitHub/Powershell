function ApiAuthenticateAndBuildAuthHeaders {
	
    # Set the admin authentication URI and create an authentication JSON body.
    $authUri = $apiBaseUri + "api/auth/loginadmin"
    $authBody = [PSCustomObject]@{
        BEID           = $apiWSBeid;
        WebServicesKey = $apiWSKey;
    } | ConvertTo-Json
	
    # Call the admin login API method and store the returned token.
    # If this part fails, display errors and exit the entire script.
    # We cannot proceed without authentication.
    $authToken = try {
        Invoke-RestMethod -Method Post -Uri $authUri -Body $authBody -ContentType "application/json"
    }
    catch {

        # Display errors and exit script.
        Write-Log -level ERROR -string "API authentication failed:"
        Write-Log -level ERROR -string ("Status Code - " + $_.Exception.Response.StatusCode.value__)
        Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
        Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
        Write-Log -level INFO -string " "
        Write-Log -level ERROR -string "The import cannot proceed when API authentication fails. Please check your authentication settings and try again."
        Write-Log -level INFO -string " "
        Write-Log -level INFO -string "Exiting."
        Write-Log -level INFO -string $processingLoopSeparator
        Exit(1)
		
    }

    # Create an API header object containing an Authorization header with a
    # value of "Bearer {tokenReturnedFromAuthCall}".
    $apiHeadersInternal = @{"Authorization" = "Bearer " + $authToken }

    # Return the API headers.
    return $apiHeadersInternal
	
}

function Search-Asset($serialNumber) {
    
    $getAssetSearchUri = $apiBaseUri + "api/$($appID)/assets/search"
    
    # Filtering options
    # https://api.teamdynamix.com/TDWebApi/Home/type/TeamDynamix.Api.Assets.AssetSearch
    $assetBody = [PSCustomObject]@{
        SerialLike = $serialNumber;
    } | ConvertTo-Json

    $resp = Invoke-RestMethod -Method POST -Headers $apiHeaders -Uri $getAssetSearchUri -Body $assetBody -ContentType "application/json" -UseBasicParsing
    return $resp
}

function Edit-Asset {
    param (
        [Parameter(Mandatory=$true)]
        [int]$ID,
        [string]$Name,
        [string]$SerialNumber,
        [string]$Tag,
        [int]$OwningCustomerID,
        [datetime]$AcquisitionDate,
        [int]$LocationID,
        [int]$LocationRoomID,
        [Parameter(Mandatory=$true)]
        [int]$StatusID,
        [int]$ProductModelID,
        [int]$SupplierID

    )

    $assetBody = [PSCustomObject]@{
        SerialLike = $serialNumber;
    } | ConvertTo-Json

    $getAssetSearchUri = $apiBaseUri + "api/$($appID)/assets/$($ID)"
    $resp = Invoke-RestMethod -Method POST -Headers $apiHeaders -Uri $getAssetSearchUri -Body $assetBody -ContentType "application/json" -UseBasicParsing

}

function SCCmCrap($computernName) {

    Get-CMDevice -Name $computerName -Fast

    <#
    
    LastActiveTime                           : 9/1/2021 10:38:51 PM
    LastClientCheckTime                      : 8/15/2021 10:10:11 AM
    LastDDR                                  : 8/31/2021 3:29:56 PM
    LastHardwareScan                         : 8/31/2021 4:15:34 PM
    LastLogonUser                            : wrcrabtree
    last check in and internal ip
    #>
}

# Site configuration
$SiteCode = "PCC" # Site code 
$ProviderMachineName = "do-sccm.pcc-domain.pima.edu" # SMS Provider machine name

$initParams = @{}

# Import the ConfigurationManager.psd1 module 
if($null -eq (Get-Module ConfigurationManager)) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

# Connect to the site's drive if it is not already present
if($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams

$apiBaseUri = 'https://service.pima.edu/SBTDWebApi/'

$TDXCreds = Get-Content .\TDXCreds.json | ConvertFrom-Json  

$apiWSBeid = $TDXCreds.BEID
$apiWSKey = $TDXCreds.Key
$appID = 1258
$global:apiHeaders = ApiAuthenticateAndBuildAuthHeaders

Search-Asset -serialNumber 140759