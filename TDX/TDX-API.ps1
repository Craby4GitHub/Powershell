# Global Variables
$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition
$scriptName = ($MyInvocation.MyCommand.Name -split '\.')[0]
$logFile = "$scriptPath\$scriptName.log"

#region 
function Write-Log {
    
    param (
        [ValidateSet('ERROR', 'INFO', 'VERBOSE', 'WARN')]
        [Parameter(Mandatory = $true)]
        [string]$level,

        [Parameter(Mandatory = $true)]
        [string]$string
    )
    
    # First roll log if over 10MB.
    if ($logFile | Test-Path) {
		
        # Get the full info for the current log file.
        $currentLog = Get-Item $logFile
		
        # Get log size in MB (from B)
        $currentLogLength = ([double]$currentLog.Length / 1024.00 / 1024.00)
		
        # If the log file exceeds 10mb, roll it by renaming with a timestamped name.
        if ($currentLogLength -ge 10) {
			
            $newLogFileName = (Get-Date).ToString("yyyy-MM-dd HHmmssfff") + " " + $currentLog.Name
            Rename-Item -Path $currentLog.FullName -NewName $newLogFileName -Force
        }
    }
	
    # Next remove old log files if there are more than 10.
    Get-ChildItem -Path $scriptPath -Filter *.log* | `
        Sort CreationTime -Desc | `
        Select -Skip 10 | `
        Remove-Item -Force

    $logString = (Get-Date).toString("yyyy-MM-dd HH:mm:ss") + " [$level] $string"
    #Add-Content -Path $logFile -Value $logString -Force
    [System.IO.File]::AppendAllText($logFile, $logString + "`r`n")
	
    $foregroundColor = $host.ui.RawUI.ForegroundColor
    $backgroundColor = $host.ui.RawUI.BackgroundColor

    Switch ($level) {
        { $_ -eq 'VERBOSE' -or $_ -eq 'INFO' } {
            
            Out-Host -InputObject "$logString"
            
        }

        { $_ -eq 'ERROR' } {

            $host.ui.RawUI.ForegroundColor = "Red"
            $host.ui.RawUI.BackgroundColor = "Black"

            Out-Host -InputObject "$logString"
    
            $host.ui.RawUI.ForegroundColor = $foregroundColor
            $host.UI.RawUI.BackgroundColor = $backgroundColor
        }

        { $_ -eq 'WARN' } {
    
            $host.ui.RawUI.ForegroundColor = "Yellow"
            $host.ui.RawUI.BackgroundColor = "Black"

            Out-Host -InputObject "$logString"
    
            $host.ui.RawUI.ForegroundColor = $foregroundColor
            $host.UI.RawUI.BackgroundColor = $backgroundColor

        }
    }
}


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
        Write-Log -level ERROR -string "API authentication failed, see the following log messages for more details."
        Write-Log -level ERROR -string ("Status Code - " + $_.Exception.Response.StatusCode.value__)
        Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
        Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
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

function Search-TDXAssets($serialNumber) {
    
    $uri = $apiBaseUri + "$($appID)/assets/search"
    
    # Using the serial number to filter. More options can be added later
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
            Write-Log -level WARN -string "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

            Start-Sleep -Milliseconds $resetWaitInMs

            Write-Log -level WARN -string "Retrying API call"
            Search-Asset -serialNumber $serialNumber
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -string "Searching for assets failed, see the following log messages for more details."
            Write-Log -level ERROR -string ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
}

function Get-TDXAssetAttributes($ID) {
    # Useful for getting atrributes and attachments
    $uri = $apiBaseUri + "$($appID)/assets/$($ID)"

    try {
        return (Invoke-RestMethod -Method GET -Headers $apiHeaders -Uri $uri -ContentType "application/json" -UseBasicParsing).Attributes
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 429) {

            # Get the amount of time we need to wait to retry in milliseconds.
            $resetWaitInMs = GetRateLimitWaitPeriodMs -apiCallResponse $_.Exception.Response
            Write-Log -level WARN -string "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

            Start-Sleep -Milliseconds $resetWaitInMs

            Write-Log -level WARN -string "Retrying API call to retrieve all location custom attribute data for the organization."
            Get-TDXAssetAttributes -ID $ID
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -string "Getting details on TDX ID $ID has failed, see the following log messages for more details."
            Write-Log -level ERROR -string ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
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

    # Wishlist: Need logic for when asset has no inventory date because we are currently searching for the attribute based on ID
    $assetAttributes = Get-TDXAssetAttributes -ID $asset.ID
    foreach ($attribute in $assetAttributes) {
        switch ($attribute.ID) {
            126172 {
                # Last Inventory Date
                if ($null -ne $sccmLastHardwareScan) {
                    $allAttributes += [PSCustomObject]@{
                        ID    = "126172";
                        Value = $sccmLastHardwareScan.ToString("o"); #formating for TDX date/time format
                    }
                }
            }
            Default {
                # All Others
                $allAttributes += [PSCustomObject]@{
                    ID    = $attribute.ID;
                    Value = $attribute.Value;
                }
            }
        }      
    }
    # TDX is all or nothing, so gotta upload everything
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

    try {
        # Wishlist: Create logic to verify edit
        $response = Invoke-RestMethod -Method POST -Headers $apiHeaders -Uri $uri -Body $assetBody -ContentType "application/json" -UseBasicParsing
        
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 429) {

            # Get the amount of time we need to wait to retry in milliseconds.
            $resetWaitInMs = GetRateLimitWaitPeriodMs -apiCallResponse $_.Exception.Response
            Write-Log -level WARN -string "Waiting $(($resetWaitInMs / 1000.0).ToString("N2")) seconds to rety API call due to rate-limiting."

            Start-Sleep -Milliseconds $resetWaitInMs

            Write-Log -level WARN -string "Retrying API call to retrieve all location custom attribute data for the organization."
            Get-TDXAssetAttributes -ID $ID
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -string "Editing the asset PCC Number $($Asset.Tag) has failed, see the following log messages for more details."
            Write-Log -level ERROR -string ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -string ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -string ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
}

$TDXCreds = Get-Content $PSScriptRoot\TDXCreds.json | ConvertFrom-Json

$apiBaseUri = 'https://service.pima.edu/SBTDWebApi/api/'

$appID = 1258

$global:apiHeaders = Get-TDXAuth -beid $TDXCreds.BEID -key $TDXCreds.Key