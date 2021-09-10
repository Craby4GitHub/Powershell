function ApiAuthenticateAndBuildAuthHeaders {
	
    $authUri = $apiBaseUri + "api/auth/loginadmin"
    $authBody = [PSCustomObject]@{
        BEID           = $apiWSBeid;
        WebServicesKey = $apiWSKey;
    } | ConvertTo-Json

    $authToken = try {
        Invoke-RestMethod -Method Post -Uri $authUri -Body $authBody -ContentType "application/json"
    }
    catch {
        # Display errors and exit script.
        Write-host "API authentication failed:"
        Exit(1)
    }

    $apiHeadersInternal = @{"Authorization" = "Bearer " + $authToken }

    return $apiHeadersInternal
}

function GetRateLimitWaitPeriodMs {
    param (
        $apiCallResponse
    )

    # Get the rate limit period reset.
    # Be sure to convert the reset date back to universal time because PS conversions will go to machine local.
    $rateLimitReset = ([DateTime]$apiCallResponse.Headers["X-RateLimit-Reset"]).ToUniversalTime()

    # Calculate the actual rate limit period in milliseconds.
    # Add 5 seconds to the period for clock skew just to be safe.
    $duration = New-TimeSpan -Start ((Get-Date).ToUniversalTime()) -End $rateLimitReset
    $rateLimitMsPeriod = $duration.TotalMilliseconds + 5000

    # Return the millisecond rate limit wait.    
    return $rateLimitMsPeriod

}

function Search-Asset($serialNumber) {
    
    $uri = $apiBaseUri + "api/$($appID)/assets/search"
    
    # Filtering options
    # https://api.teamdynamix.com/TDWebApi/Home/type/TeamDynamix.Api.Assets.AssetSearch
    $assetBody = [PSCustomObject]@{
        SerialLike = $serialNumber;
    } | ConvertTo-Json


    try {

        $resp = Invoke-RestMethod -Method POST -Headers $apiHeaders -Uri $uri -Body $assetBody -ContentType "application/json" -UseBasicParsing
        return $resp

    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 429) {

            # Get the amount of time we need to wait to retry in milliseconds.
            $resetWaitInMs = GetRateLimitWaitPeriodMs -apiCallResponse $_.Exception.Response
            Write-host "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

            # Wait to retry now.
            Start-Sleep -Milliseconds $resetWaitInMs

            # Now retry.
            Write-host "Retrying API call to retrieve all location custom attribute data for the organization."
            Search-Asset -serialNumber $serialNumber
        }
        else {
            # Display errors and exit script.

            Exit(1)
        
        }
    }
}

function Get-Asset($ID) {
    
    $uri = $apiBaseUri + "api/$($appID)/assets/$($ID)"

    $resp = Invoke-RestMethod -Method GET -Headers $apiHeaders -Uri $uri  -ContentType "application/json" -UseBasicParsing
    return $resp
}

# Need to get array of array for statues with Name and ID
function Get-AllAssetStatus {
    $statuses = @()
    $uri = $apiBaseUri + "api/$($appID)/assets/statuses"

    $resp = $status = Invoke-RestMethod -Method GET -Headers $apiHeaders -Uri $uri -ContentType "application/json" -UseBasicParsing
    foreach ($status in $resp) {
        if ($status.IsActive) {
            $statuses += @($status.Name, $status.ID)
        }
    }
    return $statuses
}

function Edit-Asset {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Asset
    )
    $allAttributes = @()

    $assetAttributes = (get-asset -ID $asset.ID).Attributes
    foreach ($attribute in $assetAttributes) {
        if ($attribute.ID -ne '126172') {
            $allAttributes += [PSCustomObject]@{
                ID    = $attribute.ID;
                Value = $attribute.Value;
            }
        }        
    }

    $SCCM = Get-SCCMDevice -computerName $('*' + $Asset.Tag + '*')
    if ($SCCM.LastHardwareScan -ne $null) {
        #Pull SCCM Last Hardware scan to TDX Last Inventory Date
        $allAttributes += [PSCustomObject]@{
            ID    = "126172";
            Value = $SCCM.LastHardwareScan.ToString("o");
        }
    }
    
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
   
    
    $Uri = $apiBaseUri + "api/$($appID)/assets/$($Asset.ID)"
    $resp = Invoke-RestMethod -Method POST -Headers $apiHeaders -Uri $uri -Body $assetBody -ContentType "application/json" -UseBasicParsing
    return $resp
}

function Get-SCCMDevice($computerName) {

    # Site configuration
    $SiteCode = "PCC" # Site code 
    $ProviderMachineName = "do-sccm.pcc-domain.pima.edu" # SMS Provider machine name

    $initParams = @{}
  
    # Import the ConfigurationManager.psd1 module 
    if ($null -eq (Get-Module ConfigurationManager)) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
    }

    # Connect to the site's drive if it is not already present
    if ($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
    }

    # Set the current location to be the site code.
    Set-Location "$($SiteCode):\" @initParams


    $device = Get-CMDevice -Name $computerName -Fast

    return $device
}

$TDXCreds = Get-Content $PSScriptRoot\TDXCreds.json | ConvertFrom-Json


$apiBaseUri = 'https://service.pima.edu/SBTDWebApi/'


$apiWSBeid = $TDXCreds.BEID
$apiWSKey = $TDXCreds.Key
$appID = 1258
$global:apiHeaders = ApiAuthenticateAndBuildAuthHeaders

$asset = Search-Asset -serialNumber MXL824145Z


Edit-Asset -Asset $asset

#Get-AllAssetStatus