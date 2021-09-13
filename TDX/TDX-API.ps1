function Get-TDXAuth($beid, $key) {
	
    $uri = $apiBaseUri + "auth/loginadmin"
    $body = [PSCustomObject]@{
        BEID           = $beid;
        WebServicesKey = $key;
    } | ConvertTo-Json

    $authToken = try {
        Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType "application/json"
    }
    catch {
        # Display errors and exit script.
        Write-host "API authentication failed:"
        Exit(1)
    }

    $apiHeaders = @{"Authorization" = "Bearer " + $authToken }

    return $apiHeaders
}

function GetRateLimitWaitPeriodMs($apiCallResponse) {
    # Get the rate limit period reset.
    # Be sure to convert the reset date back to universal time because PS conversions will go to machine local.
    $rateLimitReset = ([DateTime]$apiCallResponse.Headers["X-RateLimit-Reset"]).ToUniversalTime()

    # Calculate the actual rate limit period in milliseconds.
    # Add 5 seconds to the period for clock skew just to be safe.
    $duration = New-TimeSpan -Start ((Get-Date).ToUniversalTime()) -End $rateLimitReset
    $rateLimitMsPeriod = $duration.TotalMilliseconds + 5000

    return $rateLimitMsPeriod
}

function Search-TDXAsset($serialNumber) {
    
    $uri = $apiBaseUri + "$($appID)/assets/search"
    
    # Filtering options
    # https://api.teamdynamix.com/TDWebApi/Home/type/TeamDynamix.Api.Assets.AssetSearch
    $Body = [PSCustomObject]@{
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
            $resetWaitInMs = GetRateLimitWaitPeriodMs -apiCallResponse $_.Exception.Response
            Write-host "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

            Start-Sleep -Milliseconds $resetWaitInMs

            Write-host "Retrying API call to retrieve all location custom attribute data for the organization."
            Search-Asset -serialNumber $serialNumber
        }
        else {
            # Display errors and exit script.
            Exit(1)
        }
    }
}

function Get-TDXAssetDetails($ID) {
    # Useful for getting atrributes and attachments
    $uri = $apiBaseUri + "$($appID)/assets/$($ID)"

    $response = Invoke-RestMethod -Method GET -Headers $apiHeaders -Uri $uri -ContentType "application/json" -UseBasicParsing
    return $response
}

function Get-TDXAssetStatuses {
    $statuses = @()
    $uri = $apiBaseUri + "$($appID)/assets/statuses"

    $response = $status = Invoke-RestMethod -Method GET -Headers $apiHeaders -Uri $uri -ContentType "application/json" -UseBasicParsing
    
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
    
    $allAttributes = @()

    $assetAttributes = (Get-TDXAssetDetails -ID $asset.ID).Attributes
    foreach ($attribute in $assetAttributes) {
        switch ($attribute.ID) {
            126172 {# Last Inventory Date
                if ($null -ne $sccmLastHardwareScan) {
                    $allAttributes += [PSCustomObject]@{
                        ID    = "126172";
                        Value = $sccmLastHardwareScan.ToString("o"); #formating for TDX date/time format
                    }
                }
            }
            Default {# All Others
                $allAttributes += [PSCustomObject]@{
                    ID    = $attribute.ID;
                    Value = $attribute.Value;
                }
            }
        }      
    }
    # Asset properties.
    # https://pima.teamdynamix.com/SBTDWebApi/Home/type/TeamDynamix.Api.Assets.Asset#properties
    $assetBody = [PSCustomObject]@{
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
    
    $uri = $apiBaseUri + "$($appID)/assets/$($Asset.ID)"
    $response = Invoke-RestMethod -Method POST -Headers $apiHeaders -Uri $uri -Body $assetBody -ContentType "application/json" -UseBasicParsing
    return $response
}

$TDXCreds = Get-Content $PSScriptRoot\TDXCreds.json | ConvertFrom-Json

$apiBaseUri = 'https://service.pima.edu/SBTDWebApi/api/'

$appID = 1258

$global:apiHeaders = Get-TDXAuth -beid $TDXCreds.BEID -key $TDXCreds.Key