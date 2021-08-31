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

function SearchAsset($serialNumber) {
    $appID = 1257
    $getAssetSearchUri = $apiBaseUri + "/api/$($appID)/assets/search"
    
    # Filtering options
    # https://api.teamdynamix.com/TDWebApi/Home/type/TeamDynamix.Api.Assets.AssetSearch
    $assetBody = [PSCustomObject]@{
        SerialLike = $serialNumber;
    } | ConvertTo-Json

    $resp = Invoke-RestMethod -Method POST -Headers $apiHeaders -Uri $getAssetSearchUri -Body $assetBody -ContentType "application/json" -UseBasicParsing
    return $resp
}


$apiBaseUri = 'https://service.pima.edu/SBTDWebApi/'
    
$apiWSBeid = ''

$apiWSKey = ''

$global:apiHeaders = ApiAuthenticateAndBuildAuthHeaders

SearchAsset -serialNumber 140759