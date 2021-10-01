# https://service.pima.edu/SBTDWebApi/
#region Helper functions
function Get-TdxApiRateLimit($apiCallResponse) {
    # Get the rate limit period reset.
    # Be sure to convert the reset date back to universal time because PS conversions will go to machine local.
    $rateLimitReset = ([DateTime]$apiCallResponse.Headers["X-RateLimit-Reset"]).ToUniversalTime()

    # Calculate the actual rate limit period in milliseconds.
    # Add 5 seconds to the period for clock skew just to be safe.
    $duration = New-TimeSpan -Start ((Get-Date).ToUniversalTime()) -End $rateLimitReset
    $rateLimitMsPeriod = $duration.TotalMilliseconds + 5000

    return $rateLimitMsPeriod
}

# Setting name outside of funtion as $MyInvocation is scopped based and would pull the function name
$scriptName = ($MyInvocation.MyCommand.Name -split '\.')[0] + ' log'
function Write-Log {
    
    param (
        [ValidateSet('ERROR', 'INFO', 'VERBOSE', 'WARN')]
        [Parameter(Mandatory = $true)]
        [string]$level,

        [Parameter(Mandatory = $true)]
        [string]$message,

        $assetSerialNumber
    )

    $scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition
    $logFile = "$PSScriptroot\$scriptName.csv"
    	
    $timeStamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $logString = "$timeStamp, $level, $assetSerialNumber, $message"
    #$logString | Export-Csv -Path $logFile -Append -Delimiter ','
    Add-Content -Path $logFile -Value $logString -Force
    #[System.IO.File]::AppendAllText($logFile, $logString + "`r`n")

    Out-Host -InputObject "$logString"
}
#endregion

#region API functions
function Get-TDXAuth($beid, $key) {
    # https://service.pima.edu/SBTDWebApi/Home/section/Auth#POSTapi/auth/loginadmin
    $uri = $apiBaseUri + "auth/loginadmin"

    # Creating body for post to TDX
    $body = [PSCustomObject]@{
        BEID           = $beid;
        WebServicesKey = $key;
    } | ConvertTo-Json

    # Attempt the API call, exit script because we cant go further with out authorization
    $authToken = try {
        Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType "application/json"
    }
    catch {
        Write-Log -level ERROR -message "API authentication failed, see the following log messages for more details."
        Write-Log -level ERROR -message ("Status Code - " + $_.Exception.Response.StatusCode.value__)
        Write-Log -level ERROR -message ("Status Description - " + $_.Exception.Response.StatusDescription)
        Write-Log -level ERROR -message ("Error Message - " + $_.ErrorDetails.Message)
        Exit(1)
    }

    # Create bearer token used as the header in all TDX API calls
    $apiBearerToken = @{"Authorization" = "Bearer " + $authToken }
    return $apiBearerToken
}

function Search-TDXAssets($serialNumber) {
    # Finds all assets or searches based on a criteria
    
    # https://service.pima.edu/SBTDWebApi/Home/section/Assets#POSTapi/{appId}/assets/search
    $uri = $apiBaseUri + "$($appID)/assets/search"
    
    # Currently only using the serial number to filter. More options can be added later. Link below for more options
    # https://api.teamdynamix.com/TDWebApi/Home/type/TeamDynamix.Api.Assets.AssetSearch

    # Creating body for post to TDX
    $body = [PSCustomObject]@{
        SerialLike = $serialNumber;
    } | ConvertTo-Json

    try {
        $response = Invoke-RestMethod -Method POST -Headers $apiHeaders -Uri $uri -Body $body -ContentType "application/json" -UseBasicParsing
        return $response
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 429) {

            # Get the amount of time we need to wait to retry in milliseconds.
            $resetWaitInMs = Get-TdxApiRateLimit -apiCallResponse $_.Exception.Response
            Write-Log -level WARN -message "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

            Start-Sleep -Milliseconds $resetWaitInMs

            Write-Log -level WARN -message "Retrying API call"
            Search-Assets -serialNumber $serialNumber
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Searching for assets failed, see the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -message ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -message ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
}

function Get-TDXAssetAttributes($ID) {
    # Useful for getting atrributes and attachments

    # https://service.pima.edu/SBTDWebApi/Home/section/Assets#GETapi/{appId}/assets/{id}
    $uri = $apiBaseUri + "$($appID)/assets/$($ID)"

    try {
        return (Invoke-RestMethod -Method GET -Headers $apiHeaders -Uri $uri -ContentType "application/json" -UseBasicParsing).Attributes
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 429) {

            # Get the amount of time we need to wait to retry in milliseconds.
            $resetWaitInMs = Get-TdxApiRateLimit -apiCallResponse $_.Exception.Response
            Write-Log -level WARN -message "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

            Start-Sleep -Milliseconds $resetWaitInMs

            Write-Log -level WARN -message "Retrying API call to retrieve all custom asset attributes" -assetSerialNumber $ID
            Get-TDXAssetAttributes -ID $ID
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Getting details on TDX ID $ID has failed. See the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -message ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -message ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
}

function Get-TDXAssetStatuses {
    # Unused, for the moment
    $statuses = @()
    # https://service.pima.edu/SBTDWebApi/Home/section/AssetStatuses#GETapi/{appId}/assets/statuses
    $uri = $apiBaseUri + "$($appID)/assets/statuses"

    $response = $status = Invoke-RestMethod -Method GET -Headers $apiHeaders -Uri $uri -ContentType "application/json" -UseBasicParsing
    
    # Find every active status
    foreach ($status in $response) {
        if ($status.IsActive) {
            $statuses += [PSCustomObject]@{
                Name = $status.Name;
                ID   = $status.ID;
            }
        }
    }
    
    return $statuses
}

function Edit-TDXAsset {
    param (
        [Parameter(Mandatory = $true)]
        [object]$asset,
        $sccmLastHardwareScan
    )
    
    # Load all custom attributes and set the asset up with those values as the input asset liekly doesnt have this info
    $allAttributes = @()
    $assetAttributes = Get-TDXAssetAttributes -ID $asset.ID
    foreach ($attribute in $assetAttributes) {
        $allAttributes += [PSCustomObject]@{
            ID    = $attribute.ID;
            Value = $attribute.Value;
        }
    }

    # Check for a SCCM hardware scan. Then check to see if asset has an inventory date. If it does, update that value. Otherwise create the attribute obcject and apply value.
    # Wishlist: Loop through all attributes and apply parameter value instead of only inventory date
    if ($null -ne $sccmLastHardwareScan) {
        if ($null -ne ($allAttributes | Where-Object -Property ID -eq '126172').Value) {
            ($allAttributes | Where-Object -Property ID -eq '126172').Value = $sccmLastHardwareScan.ToString("o") #formating for TDX date/time format
        }
        else {
            $allAttributes += [PSCustomObject]@{
                ID    = "126172";
                Value = $sccmLastHardwareScan.ToString("o"); #formating for TDX date/time format
            }
        }        
    }

    # TDX is all or nothing, so the body needs every editable attribute, otherwise null values are clear production asset value
    # https://pima.teamdynamix.com/SBTDWebApi/Home/type/TeamDynamix.Api.Assets.Asset#properties
    $body = [PSCustomObject]@{
        FormID                  = $Asset.FormID;
        ProductModelID          = $Asset.ProductModelID;
        SupplierID              = $Asset.SupplierID;
        StatusID                = $Asset.StatusID;
        LocationID              = $Asset.LocationID;
        LocationRoomID          = $Asset.LocationRoomID;
        Tag                     = $Asset.Tag;
        SerialNumber            = $Asset.SerialNumber;
        Name                    = $Asset.Name;
        PurchaseCost            = $Asset.PurchaseCost;
        AcquisitionDate         = $Asset.AcquisitionDate;
        ExpectedReplacementDate = $Asset.ExpectedReplacementDate;
        RequestingCustomerID    = $Asset.RequestingCustomerID;
        RequestingDepartmentID  = $Asset.RequestingDepartmentID;
        OwningCustomerID        = $Asset.OwningCustomerID;
        OwningDepartmentID      = $Asset.OwningDepartmentID;
        ParentID                = $Asset.ParentID;
        MaintenanceScheduleID   = $Asset.MaintenanceScheduleID;
        ExternalID              = $Asset.ExternalID;
        ExternalSourceID        = $Asset.ExternalSourceID;
        Attributes              = @($allAttributes);
    } | ConvertTo-Json
    
    # https://service.pima.edu/SBTDWebApi/Home/section/Assets#POSTapi/{appId}/assets/{id}
    $uri = $apiBaseUri + "$($appID)/assets/$($Asset.ID)"

    try {
        # Wishlist: Create logic to verify edit. Will need to use Invoke-Webrequest in order to get header info if it isnt an error
        $response = Invoke-RestMethod -Method POST -Headers $apiHeaders -Uri $uri -Body $body -ContentType "application/json" -UseBasicParsing
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 429) {

            # Get the amount of time we need to wait to retry in milliseconds.
            $resetWaitInMs = Get-TdxApiRateLimit -apiCallResponse $_.Exception.Response
            Write-Log -level WARN -message "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

            Start-Sleep -Milliseconds $resetWaitInMs

            Write-Log -level WARN -message "Retrying API call to edit the asset $($Asset.Tag)"
            Get-TDXAssetAttributes -ID $ID
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Editing the asset PCC Number $($Asset.Tag) has failed, see the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -message ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -message ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
}
#endregion

# Get creds and create the base uri and header for all API calls
$TDXCreds = Get-Content $PSScriptRoot\TDXCreds.json | ConvertFrom-Json
$apiBaseUri = 'https://service.pima.edu/SBTDWebApi/api/'
$appID = 1258
$apiHeaders = Get-TDXAuth -beid $TDXCreds.BEID -key $TDXCreds.Key