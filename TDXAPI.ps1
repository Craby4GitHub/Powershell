function ApiAuthenticateAndBuildAuthHeaders {
	
    $authUri = $apiBaseUri + "api/auth/loginadmin"
    $authBody = [PSCustomObject]@{
        BEID           = $apiWSBeid;
        WebServicesKey = $apiWSKey;
    } | ConvertTo-Json

    $authToken = Invoke-RestMethod -Method Post -Uri $authUri -Body $authBody -ContentType "application/json"
    
    $apiHeadersInternal = @{"Authorization" = "Bearer " + $authToken }

    return $apiHeadersInternal
}

function Search-Asset($serialNumber) {
    
    $uri = $apiBaseUri + "api/$($appID)/assets/search"
    
    # Filtering options
    # https://api.teamdynamix.com/TDWebApi/Home/type/TeamDynamix.Api.Assets.AssetSearch
    $assetBody = [PSCustomObject]@{
        SerialLike = $serialNumber;
    } | ConvertTo-Json

    $resp = Invoke-RestMethod -Method POST -Headers $apiHeaders -Uri $uri -Body $assetBody -ContentType "application/json" -UseBasicParsing
    return $resp
}

# Need to get array of array for statues with Name and ID
function Get-AllAssetStatus {
    $statuses = @()
    $uri = $apiBaseUri + "api/$($appID)/assets/statuses"

    $resp = $status = Invoke-RestMethod -Method GET -Headers $apiHeaders -Uri $uri -ContentType "application/json" -UseBasicParsing
    foreach ($status in $resp) {
        if ($status.IsActive) {
            $statuses += @($status.Name,$status.ID)
        }
    }
    return $statuses
}

function Edit-Asset {
    param (
        [Parameter(Mandatory = $true)]
        [int]$ID,
        [Parameter(Mandatory = $true)]
        [string]$SerialNumber,
        [Parameter(Mandatory = $true)]
        $StatusID,
        $OwningCustomerID
    )


<#

Changed Status from "Active" to "Disposed".
Changed Supplier from "HP" to "".
Changed Product Model from "HP - Z230 TWR" to "".
Changed Location from "West Campus" to "Nothing".
Changed Owner from "Will Crabtree" to "Aakash Gupta".
Changed Service Tag from "140759" to "".
Changed Acquisition Date from "Wed 5/21/14" to "".
Changed Last Inventory Date from "03/19/2021" to "Nothing".
Changed PO # from "P1421343" to "Nothing".
Changed Role from "Employee Device" to "Nothing".
Changed Warranty End Date from "05/30/2017" to "Nothing".


#>






    $assetBody = [PSCustomObject]@{
        SerialNumber = $SerialNumber;
        StatusID = $StatusID;
        OwningCustomerID = $OwningCustomerID;
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

$TDXCreds = Get-Content $PSScriptRoot\TDXCreds.json | ConvertFrom-Json

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

$apiBaseUri = 'https://service.pima.edu/SBTDWebApi/'


$apiWSBeid = $TDXCreds.BEID
$apiWSKey = $TDXCreds.Key
$appID = 1258
$global:apiHeaders = ApiAuthenticateAndBuildAuthHeaders

$asset = Search-Asset -serialNumber 140759

Get-AllAssetStatus