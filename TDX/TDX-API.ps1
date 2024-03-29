# https://service.pima.edu/SBTDWebApi/
# Powershell to TDX API functions.
# Not all api endpoints are listed, only ones that I use

# Setting Log name
$logName = ($MyInvocation.MyCommand.Name -split '\.')[0] + ' log'
$logFile = "$PSScriptroot\$logName.csv"
. ((Get-Item $PSScriptRoot).Parent.FullName + '\Callable\Write-Log.ps1')

#region API functions
function Get-TDXAuth($beid, $key) {
    # https://service.pima.edu/SBTDWebApi/Home/section/Auth#POSTapi/auth/loginadmin
    $uri = $baseURI + "auth/loginadmin"
    #$uri = $baseURI + "auth/login"

    # Creating body for post to TDX
    $body = [PSCustomObject]@{
        BEID           = $beid;
        WebServicesKey = $key;
    } | ConvertTo-Json

    # Attempt the API call, exit script because we cant go further with out authorization
    $authToken = try {
        Invoke-RestMethod -Method Post -Uri $uri -Body $body -ContentType "application/json" -ErrorVariable apiError
    }
    catch {
        Write-Log -level ERROR -message "API authentication failed, see the following log messages for more details."
        Write-Log -level ERROR -message ("Status Code - " + $apiError.ErrorRecord.Exception.Response.StatusCode)
        Write-Log -level ERROR -message ("Status Description - " + $apiError.ErrorRecord.Exception.Response.StatusDescription)
        Exit(1)
    }

    # Create bearer token used as the header in all TDX API calls
    $apiBearerToken = @{"Authorization" = "Bearer " + $authToken }
    return $apiBearerToken
}

#region Assets
function Search-TDXAssets {
    # https://service.pima.edu/SBTDWebApi/Home/section/Assets#POSTapi/{appId}/assets/search
    Param(
        # https://service.pima.edu/TDWebApi/Home/type/TeamDynamix.Api.Assets.AssetSearch

        [Parameter(ParameterSetName = "Basic")]
        [int]$MaxResults,

        # Identifiers
        [Parameter(ParameterSetName = "Identifiers")]
        [string[]]$SerialLike,
        [Parameter(ParameterSetName = "Identifiers")]
        [string]$SearchText,
        [Parameter(ParameterSetName = "Identifiers")]
        [int]$SavedSearchID,
        [Parameter(ParameterSetName = "Identifiers")]
        [int[]]$StatusIDs,
        [Parameter(ParameterSetName = "Identifiers")]
        [string[]]$ExternalIDs,
        [Parameter(ParameterSetName = "Identifiers")]
        [bool]$IsInService,
        [Parameter(ParameterSetName = "Identifiers")]
        [int[]]$StatusIDsPast,

        # Supplier and manufacturer
        [Parameter(ParameterSetName = "Supplier and manufacturer")]
        [int[]]$SupplierIDs,
        [Parameter(ParameterSetName = "Supplier and manufacturer")]
        [int[]]$ManufacturerIDs,

        # Location and room
        [Parameter(ParameterSetName = "Location and room")]
        [int[]]$LocationIDs,
        [Parameter(ParameterSetName = "Location and room")]
        [int]$RoomID,

        # Parent and contract
        [Parameter(ParameterSetName = "Parent and contract")]
        [int[]]$ParentIDs,
        [Parameter(ParameterSetName = "Parent and contract")]
        [int[]]$ContractIDs,
        [Parameter(ParameterSetName = "Parent and contract")]
        [int[]]$ExcludeContractIDs,

        # Tickets and forms
        [Parameter(ParameterSetName = "Tickets and forms")]
        [int[]]$TicketIDs,
        [Parameter(ParameterSetName = "Tickets and forms")]
        [int[]]$ExcludeTicketIDs,
        [Parameter(ParameterSetName = "Tickets and forms")]
        [int[]]$FormIDs,

        # Product and maintenance
        [Parameter(ParameterSetName = "Product and maintenance")]
        [int[]]$ProductModelIDs,
        [Parameter(ParameterSetName = "Product and maintenance")]
        [int[]]$MaintenanceScheduleIDs,

        # Departments
        [Parameter(ParameterSetName = "Departments")]
        [int[]]$UsingDepartmentIDs,
        [Parameter(ParameterSetName = "Departments")]
        [int[]]$RequestingDepartmentIDs,
        [Parameter(ParameterSetName = "Departments")]
        [int[]]$OwningDepartmentIDs,
        [Parameter(ParameterSetName = "Departments")]
        [int[]]$OwningDepartmentIDsPast,

        # Customers
        [Parameter(ParameterSetName = "Customers")]
        [guid[]]$UsingCustomerIDs,
        [Parameter(ParameterSetName = "Customers")]
        [guid[]]$RequestingCustomerIDs,
        [Parameter(ParameterSetName = "Customers")]
        [guid[]]$OwningCustomerIDs,
        [Parameter(ParameterSetName = "Customers")]
        [guid[]]$OwningCustomerIDsPast,

        # Custom attributes
        [Parameter(ParameterSetName = "Custom attributes")]
        [hashtable]$CustomAttributes,

        # Purchase and acquisition
        [Parameter(ParameterSetName = "Purchase and acquisition")]
        [double]$PurchaseCostFrom,
        [Parameter(ParameterSetName = "Purchase and acquisition")]
        [double]$PurchaseCostTo,
        [Parameter(ParameterSetName = "Purchase and acquisition")]
        [int]$ContractProviderID,
        [Parameter(ParameterSetName = "Purchase and acquisition")]
        [datetime]$AcquisitionDateFrom,
        [Parameter(ParameterSetName = "Purchase and acquisition")]
        [datetime]$AcquisitionDateTo,
        [Parameter(ParameterSetName = "Purchase and acquisition")]
        [datetime]$ExpectedReplacementDateFrom,
        [Parameter(ParameterSetName = "Purchase and acquisition")]
        [datetime]$ExpectedReplacementDateTo,
        [Parameter(ParameterSetName = "Purchase and acquisition")]
        [datetime]$ContractEndDateFrom,
        [Parameter(ParameterSetName = "Purchase and acquisition")]
        [datetime]$ContractEndDateTo,
        [Parameter(ParameterSetName = "Purchase and acquisition")]
        [bool]$OnlyParentAssets,

        [Parameter(Mandatory)]
        [Parameter(ParameterSetName = "TDX")]
        [ValidateSet("ITTicket", "ITAsset", "D2LTicket")]
        $AppName
    )

    # Finds all assets or searches based on a criteria. Attachments and Attributes are not in included in the results


    $appID = Get-TDXAppID -AppName $AppName
    $uri = $baseURI + $appID + '/assets/search'
    
    # Creating body for post to TDX
    $body = [PSCustomObject]@{}
    foreach ($param in $PSBoundParameters.GetEnumerator()) {
        if ($param.Key.ParameterSetName -ne 'TDX') {
            $body[$param.Key] = $param.Value
        } 
    }

    $json = $body | ConvertTo-Json

    try {
        return Invoke-RestMethod -Method POST -Headers $tdxAPIAuth -Uri $uri -Body $json -ContentType "application/json" -UseBasicParsing -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Search-TDXAssets API call"
            Search-TDXAssets $PSBoundParameters
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Searching for assets failed, see the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $apiError.ErrorRecord.Exception.Response.StatusCode)
            Write-Log -level ERROR -message ("Status Description - " + $apiError.ErrorRecord.Exception.Response.StatusDescription)
            Exit(1)
        }
    }
}

function Search-TDXAssetsBySavedSearch($searchID) {
    # Gets a page of assets matching the provided saved search and pagination options.


    
    # https://service.pima.edu/SBTDWebApi/api/{appId}/assets/searches/{searchId}/results
    $uri = $baseURI + $appIDAsset + "/assets/searches/$searchID/results"

    # Creating body for post to TDX
    $body = [PSCustomObject]@{
        SearchText	= ""	#String	This field is nullable.	The search text to filter on. If specified, this will override any search text that may have been part of the saved search.
        Page       = @{     #This field is required.	TeamDynamix.Api.RequestPage		The page of data being requested for the saved search.
            PageIndex = 0	#This field is required.	Int32	The 0-based page index to request.
            PageSize  = 1	#This field is required.	Int32	The size of each page being requested, ranging 1-200 (inclusive).
        }	
    } | ConvertTo-Json

    try {
        return Invoke-RestMethod -Method POST -Headers $tdxAPIAuth -Uri $uri -Body $body -ContentType "application/json" -UseBasicParsing -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Search-TDXAssetsBySavedSearch API call"
            Search-TDXAssetsBySavedSearch -$searchID
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Search-TDXAssetsBySavedSearch failed, see the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -message ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -message ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
}

function Get-TDXAssetProductModels {
    # https://service.pima.edu/SBTDWebApi/Home/section/Assets#GETapi/{appId}/assets/{id}
    $uri = $baseURI + $appIDAsset + "/assets/models"
    
    try {
        return Invoke-RestMethod -Method GET -Headers $tdxAPIAuth -Uri $uri -ContentType "application/json" -UseBasicParsing -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Get-TDXAssetProductModels API call" -assetSerialNumber $ID
            Get-TDXAssetProductModels
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Getting details on TDX Product ID# $ID has failed. See the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -message ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -message ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
    
}

function Edit-TDXAssetProductModel {
    param (
        [Parameter(Mandatory = $true)]
        $ID,
        [Parameter(Mandatory = $true)]
        $Name,
        [Parameter(Mandatory = $true)]
        $ManufacturerID,
        [Parameter(Mandatory = $true)]
        $ProductTypeID, $Description, $IsActive, $PartNumber, $Attributes
    )
    # https://service.pima.edu/SBTDWebApi/Home/section/Assets#GETapi/{appId}/assets/{id}
    $uri = $baseURI + $appIDAsset + "/assets/models/$ID"

    $body = [PSCustomObject]@{

        Name           =	$Name	#	String		The name of the product model.
        Description    =	$Description	#	String	This field is nullable.	The description of the product model.
        IsActive       =	$IsActive		#Boolean		The active status of the product model.
        ManufacturerID	= $ManufacturerID #	This field is required.	Int32		The ID of the manufacturer associated with the product model.
        ProductTypeID  =	$ProductTypeID#	This field is required.	Int32		The ID of the product type associated with the product model.
        PartNumber     =	$PartNumber	# String	This field is nullable.	The part number of the product model.
        Attributes     =	$Attributes# TeamDynamix.Api.CustomAttributes.CustomAttribute[]	This field is nullable.	The custom attributes associated with the product model.
    } | ConvertTo-Json
    
    try {
        return Invoke-RestMethod -Method PUT -Headers $tdxAPIAuth -Uri $uri -ContentType "application/json" -UseBasicParsing -Body $body -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Edit-TDXAssetProductModel API call" -assetSerialNumber $ID
            Edit-TDXAssetProductModel
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Getting details on TDX Product ID# $ID has failed. See the following log messages for more details."
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
    $uri = $baseURI + $appIDAsset + "/assets/$ID"
    
    try {
        return (Invoke-RestMethod -Method GET -Headers $tdxAPIAuth -Uri $uri -ContentType "application/json" -UseBasicParsing -ErrorVariable apiError).Attributes
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        #$statusCode = $_.Exception.Response.StatusCode.value__
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
                        
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Get-TDXAssetAttributes API call to retrieve all custom asset attributes"
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
    $uri = $baseURI + $appIDAsset + "/assets/statuses"
    $response = $status = Invoke-RestMethod -Method GET -Headers $tdxAPIAuth -Uri $uri -ContentType "application/json" -UseBasicParsing
    
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

function Search-TDXAssetFeed($ID) {
    # Useful for getting atrributes and attachments

    # GET https://service.pima.edu/SBTDWebApi/api/{appId}/assets/{id}/feed
    $uri = $baseURI + $appIDAsset + "/assets/$id/feed"
    
    try {
        return Invoke-RestMethod -Method GET -Headers $apiHeaders -Uri $uri -ContentType "application/json" -UseBasicParsing -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Search-TDXAssetFeed API call" -assetSerialNumber $ID
            Search-TDXAssetFeed
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Getting asset feed on TDX ID $ID has failed. See the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -message ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -message ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
}

function Edit-TDXAsset {
    param (
        [Parameter(Mandatory = $true)]
        [object]$asset,
        $lastHardwareScan
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
    if ($null -ne $lastHardwareScan) {
        if ($null -ne ($allAttributes | Where-Object -Property ID -eq '126172').Value) {
            ($allAttributes | Where-Object -Property ID -eq '126172').Value = $lastHardwareScan.ToString("o") #formating for TDX date/time format
        }
        else {
            $allAttributes += [PSCustomObject]@{
                ID    = "126172";
                Value = (Get-Date $lastHardwareScan).ToString("o"); #formating for TDX date/time format
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
    $uri = $baseURI + $appIDAsset + "/assets/$($Asset.ID)"

    try {
        # Wishlist: Create logic to verify edit. Will need to use Invoke-Webrequest in order to get header info if it isnt an error
        $response = Invoke-RestMethod -Method POST -Headers $tdxAPIAuth -Uri $uri -Body $body -ContentType "application/json" -UseBasicParsing -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Edit-TDXAsset API call on PCC# $($Asset.Tag)"
            Edit-TDXAsset -asset $asset -lastHardwareScan $lastHardwareScan
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

function Import-TDXAssets($assets) {
    # need to figure out what is needed for a importData object
    #https://pima.teamdynamix.com/SBTDWebApi/api/{appId}/assets/import
    $body = [PSCustomObject]@{
        importdata = @($assets)
        Settings   = @{
            UpdateItems = $true
            CreateItems = $true
            Mappings    = @(
                @{
                    FieldIdentifier   = $null	# This field is nullable.	The identifier of the field.     
                    CustomAttributeID =	0 # The ID of the associated custom attribute, or 0 if the field is not a custom attribute.
                    DefaultValue      = $null	# This field is nullable.	The default value of the field.
                    ClearOnNull       =	$false # Whether the field on the API object should be cleared when a value has not been provided for it.
                }
            )
        }   
    } | ConvertTo-Json #-Depth 4
    
    #https://pima.teamdynamix.com/SBTDWebApi/api/{appId}/assets/import
    $uri = $baseURI + $appIDAsset + "/assets/import"

    try {
        # Wishlist: Create logic to verify edit. Will need to use Invoke-Webrequest in order to get header info if it isnt an error
        $response = Invoke-RestMethod -Method POST -Headers $tdxAPIAuth -Uri $uri -Body $body -ContentType "application/json" -UseBasicParsing -ErrorVariable apiError
        Write-Host $response
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Import-TDXAssets API call."
            Import-TDXAssets
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Import-TDXAssets API call has failed, see the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -message ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -message ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
}
#endregion
#region Ticket
function Submit-TDXTicket {
    param (
        [Parameter(Mandatory = $true)]
        [int32]$AccountID, # The ID of the account/department associated with the ticket.
        [Parameter(Mandatory = $true)]
        [int32]$PriorityID, # The ID of the priority associated with the ticket.
        $RequestorUid, # The UID of the requestor associated with the ticket.
        [Parameter(Mandatory = $true)]
        [int32]$StatusID, # The ID of the ticket status associated with the ticket.
        [Parameter(Mandatory = $true)]
        [string]$Title, # The title of the ticket.
        [Parameter(Mandatory = $true)]
        [int32]$TypeID, # The ID of the ticket type associated with the ticket.
        [Int32]$ArticleID, # The ID of the Knowledge Base article associated with the ticket.
        [Int32]$ArticleShortcutID, # The ID of the shortcut that is used when viewing the ticket's Knowledge Base article. This is set when the ticket is associated with a cross client portal article shortcut.
        $Attributes, # The custom attributes associated with the ticket.
        [String]$Description, # The description of the ticket.
        [DateTime]$EndDate, # The end date of the ticket.
        [Int32]$EstimatedMinutes, # The estimated minutes of the ticket.
        [Double]$ExpensesBudget, # The expense budget of the ticket.
        [Int32]$FormID, # The ID of the form associated with the ticket.
        [DateTime]$GoesOffHoldDate, # The date the ticket goes off hold.
        [Int32]$ImpactID, # The ID of the impact associated with the ticket.
        [Int32]$LocationID, # The ID of the location associated with the ticket.
        [Int32]$LocationRoomID, # The ID of the location room associated with the ticket.
        [Int32]$ResponsibleGroupID, # The ID of the responsible group associated with the ticket.
        $ResponsibleUid, # The UID of the responsible user associated with the ticket.
        [Int32]$ServiceID, # The ID of the service associated with the ticket.
        [Int32]$ServiceOfferingID, # The ID of the service offering associated with the ticket.
        [Int32]$SourceID, # The ID of the ticket source associated with the ticket.
        [DateTime]$StartDate, # The start date of the ticket.
        [Double]$TimeBudget, # The time budget of the ticket.
        [Int32]$UrgencyID, # The ID of the urgency associated with the ticket.
        $ClassificationID, # The classification associated with the ticket.
        [bool]$IsRichHtml, #Indicates if the ticket description is rich-text or plain-text.
        [String]$RequestorName, #					The full name of the requestor associated with the ticket.
        [String]$RequestorFirstName	, #				The first name of the requestor associated with the ticket.
        [String]$RequestorLastName	, #				The last name of the requestor associated with the ticket.
        [String]$RequestorEmail		, #			The email address of the requestor associated with the ticket.
        [String]$RequestorPhone	, #				The phone number of the requestor associated with the ticket.
        $AppName
    )

    # Body
    # https://service.pima.edu/SBTDWebApi/Home/type/TeamDynamix.Api.Tickets.Ticket

    $ticketAttributes = [PSCustomObject]@{
        AccountID          = $AccountID # The ID of the account/department associated with the ticket.
        PriorityID         = $PriorityID # The ID of the priority associated with the ticket.
        RequestorUid       = $RequestorUid # The UID of the requestor associated with the ticket.
        StatusID           = $StatusID # The ID of the ticket status associated with the ticket.
        Title              = $Title # The title of the ticket.
        TypeID             = $TypeID # The ID of the ticket type associated with the ticket.
        ArticleID          = $ArticleID # The ID of the Knowledge Base article associated with the ticket.
        ArticleShortcutID  = $ArticleShortcutID # The ID of the shortcut that is used when viewing the ticket's Knowledge Base article. This is set when the ticket is associated with a cross client portal article shortcut.
        Attributes         = @($Attributes); # The custom attributes associated with the ticket.
        Description        = $Description # The description of the ticket.
        EndDate            = $EndDate # The end date of the ticket.
        EstimatedMinutes   = $EstimatedMinutes # The estimated minutes of the ticket.
        ExpensesBudget     = $ExpensesBudget # The expense budget of the ticket.
        FormID             = $FormID # The ID of the form associated with the ticket.
        GoesOffHoldDate    = $GoesOffHoldDate # The date the ticket goes off hold.
        ImpactID           = $ImpactID # The ID of the impact associated with the ticket.
        LocationID         = $LocationID # The ID of the location associated with the ticket.
        LocationRoomID     = $LocationRoomID # The ID of the location room associated with the ticket.
        ResponsibleGroupID = $ResponsibleGroupID # The ID of the responsible group associated with the ticket.
        ResponsibleUid     = $ResponsibleUid # The UID of the responsible user associated with the ticket.
        ServiceID          = $ServiceID # The ID of the service associated with the ticket.
        ServiceOfferingID  = $ServiceOfferingID # The ID of the service offering associated with the ticket.
        SourceID           = $SourceID # The ID of the ticket source associated with the ticket.
        StartDate          = $StartDate # The start date of the ticket.
        TimeBudget         = $TimeBudget # The time budget of the ticket.
        UrgencyID          = $UrgencyID # The ID of the urgency associated with the ticket.
        Classification     = $ClassificationID
        IsRichHtml         = $IsRichHtml #Indicates if the ticket description is rich-text or plain-text.
        RequestorName      = $RequestorName #					The full name of the requestor associated with the ticket.
        RequestorFirstName = $RequestorFirstName	 #				The first name of the requestor associated with the ticket.
        RequestorLastName  = $RequestorLastName	 #				The last name of the requestor associated with the ticket.
        RequestorEmail     = $RequestorEmail		 #			The email address of the requestor associated with the ticket.
        RequestorPhone     = $RequestorPhone	 #				The phone number of the requestor associated with the ticket.
    }

    $body = $ticketAttributes | ConvertTo-Json
    # https://service.pima.edu/SBTDWebApi/Home/type/TeamDynamix.Api.Tickets.TicketCreateOptions
    $appID = Get-TDXAppID -AppName $AppName
    $uri = $baseURI + $appID + "/tickets?EnableNotifyReviewer=$false&NotifyRequestor=$true&NotifyResponsible=$false&AllowRequestorCreation=$false"
    
    try {
        # Wishlist: Create logic to verify edit. Will need to use Invoke-Webrequest in order to get header info if it isnt an error
        return Invoke-RestMethod -Method POST -Headers $tdxAPIAuth -Uri $uri -Body $body -ContentType "application/json; charset=utf-8" -UseBasicParsing -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying API call to create the ticket: $Title"
            return Submit-TDXTicket @ticketAttributes
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Creating a ticket for $RequestorUid has failed, see the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -message ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -message ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
}

function Edit-TDXTicket($ticketID, $TypeID, $AccountID, $StatusID, $PriorityID, $RequestorUid, $ServiceID) {
    #POST https://service.pima.edu/SBTDWebApi/api/{appId}/tickets/{id}?notifyNewResponsible={notifyNewResponsible}
    #Edits an existing ticket.

    $uri = $baseURI + $appIDTicket + "/tickets/$ticketID"



    $body = @{
        op    = 'add' 
        path  = '/ServiceID'
        value = '32656'
    } | ConvertTo-Json
    <#$body = [PSCustomObject]@{

        [Int32]$TypeID      = $TypeID	# The ID of the ticket type associated with the ticket.
        #[Int32]$FormID=$FormID	# The ID of the form associated with the ticket.
        #[String]$Title=$Title	# The title of the ticket.
        # [String]	$Description	# The description of the ticket.
        [Int32]	$AccountID  = $AccountID	# The ID of the account/department associated with the ticket.
        #[Int32]	SourceID	# The ID of the ticket source associated with the ticket.
        [Int32]	$StatusID   = $StatusID	# The ID of the ticket status associated with the ticket.
        #[Int32]	ImpactID	# The ID of the impact associated with the ticket.
        #[Int32]	UrgencyID	# The ID of the urgency associated with the ticket.
        [Int32]	$PriorityID = $PriorityID	# The ID of the priority associated with the ticket.
        #DateTime	GoesOffHoldDate	The date the ticket goes off hold.
        $RequestorUid       = $RequestorUid	#The UID of the requestor associated with the ticket.
        #[Int32]	EstimatedMinutes	# The estimated minutes of the ticket.
        # [DateTime]StartDate	# The start date of the ticket.
        #DateTime	EndDate	The end date of the ticket.
        #Guid	ResponsibleUid	The UID of the responsible user associated with the ticket.
        # [Int32]	ResponsibleGroupID	# The ID of the responsible group associated with the ticket.
        # [Int32]	LocationID	# The ID of the location associated with the ticket.
        # [Int32]	LocationRoomID	# The ID of the location room associated with the ticket.
        [Int32]	$ServiceID  = $ServiceID	# The ID of the service associated with the ticket.
        #[Int32]	ServiceOfferingID	# The ID of the service offering associated with the ticket.
        # [Int32]	ArticleID	# The ID of the Knowledge Base article associated with the ticket.
        #[Int32]	ArticleShortcutID	# The ID of the shortcut that is used when viewing the ticket's Knowledge Base article. This is set when the ticket is associated with a cross client portal article shortcut.
        #TeamDynamix.Api.CustomAttributes.CustomAttribute[]	Attributes	The custom attributes associated with the ticket.
    } | ConvertTo-Json
#>

    try {
        return Invoke-RestMethod -Method PATCH -Headers $tdxAPIAuth -Uri $uri -ContentType "application/json" -UseBasicParsing -Body $body -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Get-TDXTicket API call on ticket $ticketID"
            Get-TDXTicket -ticketID $ticketID
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Editing the ticket $ticketID has failed, see the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -message ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -message ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
}

function Set-TDXAttachment($ticketID, $AppName, $Attachment) {
    # https://service.pima.edu/SBTDWebApi/Home/section/Tickets#POSTapi/{appId}/tickets/{id}/attachments
    $appID = Get-TDXAppID -AppName $AppName
    $uri = $baseURI + $appID + "/tickets/$ticketID/attachments"
    <#
    $form = @{test = Get-Content "C:\Users\wrcrabtree\Downloads\d3dba960-3397-11ed-a175-2f0ed4c829bb.mp3"}
$1 = "`r`n-----------------------------313681257239897303243066788965"
$2 = "`r`nContent-Disposition: form-data; name='d3dba960-3397-11ed-a175-2f0ed4c829bb.mp3'; filename='d3dba960-3397-11ed-a175-2f0ed4c829bb.mp3'"
$3 = "`r`nContent-Type: audio/mpeg`r`n`n"
$4 = "`r`n-----------------------------313681257239897303243066788965--`r`n"

#>


    $1 = "--CHANGEME"
    $2 = "`nContent-Disposition: form-data; name=`"aeneid.txt`"; filename=`"aeneid.txt`""
    $3 = "`nContent-Type: audio/octet-stream`n`n"
    $4 = "`n--CHANGEME--"
    $5 = 'FORSAN ET HAEC OLIM MEMINISSE IUVABIT'
    $body = $1 + $2 + $3 + $5 + $4
    #$body = $1 + $2 + $3 + $($form.values) + $4
    try {
        Invoke-WebRequest -Method POST -Headers $tdxAPIAuth -Uri $uri -Form $body -ErrorVariable test -ContentType "multipart/form-data; boundary=CHANGEME"
        return Invoke-RestMethod -Method POST -Headers $tdxAPIAuth -Uri $uri -body $body -ContentType "multipart/form-data;boundary=CHANGEME" -UseBasicParsing -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Set-TDXAttachment API call on ticket $ticketID"
            Get-TDXTicket -ticketID $ticketID
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Setting the attaxhment in ticket $ticketID has failed, see the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -message ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -message ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
}

function Get-TDXTicket($ticketID) {
    # GET https://service.pima.edu/SBTDWebApi/api/{appId}/tickets/{id}
    $uri = $baseURI + $appIDTicket + "/tickets/$ticketID"

    try {
        return Invoke-RestMethod -Method GET -Headers $tdxAPIAuth -Uri $uri -ContentType "application/json" -UseBasicParsing -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Get-TDXTicket API call on ticket $ticketID"
            Get-TDXTicket -ticketID $ticketID
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Getting the ticket $ticketID has failed, see the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -message ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -message ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
}

function Search-TDXTicket {
    # https://service.pima.edu/SBTDWebApi/Home/section/Tickets#POSTapi/{appId}/tickets/search

    [CmdletBinding(DefaultParameterSetName = "Basic")]
    param(
        # https://service.pima.edu/SBTDWebApi/Home/type/TeamDynamix.Api.Tickets.TicketSearch
        # Basic search parameters
        [Parameter(ParameterSetName = "Basic")]
        [string]$TicketID,
        [Parameter(ParameterSetName = "Basic")]
        [string]$ParentTicketID,
        [Parameter(ParameterSetName = "Basic")]
        [string]$SearchText,
        [Parameter(ParameterSetName = "Basic")]
        [int]$MaxResults,

        # Status parameters
        [Parameter(ParameterSetName = "Status")]
        [int[]]$StatusIDs,
        [Parameter(ParameterSetName = "Status")]
        [int[]]$PastStatusIDs,
        [Parameter(ParameterSetName = "Status")]
        [int[]]$StatusClassIDs,

        # Priority and urgency parameters
        [Parameter(ParameterSetName = "PriorityAndUrgency")]
        [int[]]$PriorityIDs,
        [Parameter(ParameterSetName = "PriorityAndUrgency")]
        [int[]]$UrgencyIDs,
        [Parameter(ParameterSetName = "PriorityAndUrgency")]
        [int[]]$ImpactIDs,

        # Account and type parameters
        [Parameter(ParameterSetName = "AccountAndType")]
        [int[]]$AccountIDs,
        [Parameter(ParameterSetName = "AccountAndType")]
        [int[]]$TypeIDs,
        [Parameter(ParameterSetName = "AccountAndType")]
        [int[]]$SourceIDs,

        # Date range parameters
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$CreatedDateFrom,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$CreatedDateTo,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$UpdatedDateFrom,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$UpdatedDateTo,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$ModifiedDateFrom,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$ModifiedDateTo,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$StartDateFrom,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$StartDateTo,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$EndDateFrom,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$EndDateTo,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$RespondedDateFrom,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$RespondedDateTo,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$RespondByDateFrom,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$RespondByDateTo,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$ClosedDateFrom,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$ClosedDateTo,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$CloseByDateFrom,
        [Parameter(ParameterSetName = "DateRanges")]
        [datetime]$CloseByDateTo,
        [Parameter(ParameterSetName = "DateRanges")]
        [int]$DaysOldFrom,
        [Parameter(ParameterSetName = "DateRanges")]
        [int]$DaysOldTo,

        # Responsibility parameters
        [Parameter(ParameterSetName = "Responsibility")]
        [guid[]]$ResponsibilityUids,
        [Parameter(ParameterSetName = "Responsibility")]
        [int[]]$ResponsibilityGroupIDs,
        [Parameter(ParameterSetName = "Responsibility")]
        [bool]$CompletedTaskResponsibilityFilter,
        [Parameter(ParameterSetName = "Responsibility")]
        [guid[]]$PrimaryResponsibilityUids,
        [Parameter(ParameterSetName = "Responsibility")]
        [int[]]$PrimaryResponsibilityGroupIDs,

        # SLA parameters
        [Parameter(ParameterSetName = "SLA")]
        [int[]]$SlaIDs,
        [Parameter(ParameterSetName = "SLA")]
        [bool]$SlaViolationStatus,
        [Parameter(ParameterSetName = "SLA")]
        [string[]]$SlaUnmetConstraints,
       
        # KB article parameters
        [Parameter(ParameterSetName = "KBArticle")]
        [int[]]$KBArticleIDs,

        # Requestor parameters
        [Parameter(ParameterSetName = "Requestor")]
        [guid[]]$RequestorUids,
        [Parameter(ParameterSetName = "Requestor")]
        [string]$RequestorNameSearch,
        [Parameter(ParameterSetName = "Requestor")]
        [string]$RequestorEmailSearch,
        [Parameter(ParameterSetName = "Requestor")]
        [string]$RequestorPhoneSearch,

        # Parameters for filtering by user IDs
        [Parameter(ParameterSetName = 'Uid')]
        [guid]$UpdatedByUid,
        [Parameter(ParameterSetName = 'Uid')]
        [guid]$ModifiedByUid,
        [Parameter(ParameterSetName = 'Uid')]
        [guid]$RespondedByUid,
        [Parameter(ParameterSetName = 'Uid')]
        [guid]$ClosedByUid,
        [Parameter(ParameterSetName = 'Uid')]
        [guid]$CreatedByUid,
        [Parameter(ParameterSetName = 'Uid')]
        [guid]$ReviewerUid,

        # Parameters for filtering by hold status
        [Parameter(ParameterSetName = 'Hold')]
        [bool]$IsOnHold,
        [Parameter(ParameterSetName = 'Hold')]
        [DateTime]$GoesOffHoldFrom,
        [Parameter(ParameterSetName = 'Hold')]
        [DateTime]$GoesOffHoldTo,

        # Parameters for filtering by location IDs
        [Parameter(ParameterSetName = 'Location')]
        [int[]]$LocationIDs,
        [Parameter(ParameterSetName = 'Location')]
        [int[]]$LocationRoomIDs,

        # Others
        [Parameter(ParameterSetName = 'Other')]
        [string[]]$TicketClassification,
        [Parameter(ParameterSetName = 'Other')]
        [bool[]]$AssignmentStatus,
        [Parameter(ParameterSetName = 'Other')]
        [bool]$ConvertedToTask,
        [Parameter(ParameterSetName = 'Other')]
        [int32[]]$ConfigurationItemIDs,
        [Parameter(ParameterSetName = 'Other')]
        [int32[]]$ExcludeConfigurationItemIDs,
        [Parameter(ParameterSetName = 'Other')]
        [int32[]]$ServiceIDs,
        [Parameter(ParameterSetName = 'Other')]
        [string[]]$CustomAttributes,
        [Parameter(ParameterSetName = 'Other')]
        [bool]$HasReferenceCode,

        [Parameter(Mandatory)]
        [Parameter(ParameterSetName = 'TDX')]
        [ValidateSet("ITTicket", "ITAsset", "D2LTicket")]
        $AppName
    )


    # Finds all tickets based on a criteria. Not all properties will be provided for each ticket. 
    # For example, ticket descriptions and custom attributes will not be returned. To retrieve such information, you must load a ticket individually.

    $appID = Get-TDXAppID -AppName $AppName
    $uri = $baseURI + $appID + '/tickets/search'
    
    # Creating body for post to TDX
    $body = [PSCustomObject]@{}
    foreach ($param in $PSBoundParameters.GetEnumerator()) {
        if ($param.Key.ParameterSetName -ne 'TDX') {
            $body[$param.Key] = $param.Value
        } 
    }

    $json = $body | ConvertTo-Json

    try {
        return Invoke-RestMethod -Method POST -Headers $tdxAPIAuth -Uri $uri -Body $json -ContentType "application/json" -UseBasicParsing -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Search-TDXTickets API call"
            Search-TDXTicket $PSBoundParameters
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Searching for ticket with $($term) for $($value) failed, see the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $apiError.ErrorRecord.Exception.Response.StatusCode)
            Write-Log -level ERROR -message ("Status Description - " + $apiError.ErrorRecord.Exception.Response.StatusDescription)
            Exit(1)
        }
    }
}

function Update-TDXTicket($ticketID, $StatusID, $Comment, $NotifyEmail, $IsPrivate, $IsRichHtml, $AppName) {
    # POST https://service.pima.edu/SBTDWebApi/api/{appId}/tickets/{id}/feed
    $appID = Get-TDXAppID -AppName $AppName
    $uri = $baseURI + $appID + "/tickets/$ticketID/feed"
    $body = [PSCustomObject]@{
        NewStatusID	= $StatusID     #Int32	This field is nullable.	The ID of the new status for the ticket. Leave null or 0 to not change the status.
        Comments    = $Comment      #String		The comments of the feed entry.
        Notify      = $NotifyEmail  #String[]	This field is nullable.	The email addresses to notify associated with the feed entry.
        IsPrivate   = $IsPrivate    #Boolean		The private status of the feed entry.
        IsRichHtml  = $IsRichHtml   #Boolean		Indicates if the feed entry is rich-text or plain-text.     
    } | ConvertTo-Json
    try {
        return Invoke-RestMethod -Method POST -Headers $tdxAPIAuth -Uri $uri -Body $body -ContentType "application/json; charset=utf-8" -UseBasicParsing -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Update-TDXTicket API call on ticket $ticketID"
            Update-TDXTicket -ticketID $ticketID -StatusID $StatusID -Comment $Comment -NotifyEmail $NotifyEmail -IsPrivate $IsPrivate -IsRichHtml $IsRichHtml -AppName $AppName
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Updating the ticket $ticketID has failed, see the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -message ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -message ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
}

function Edit-TDXTicketAddAsset($ticketID, $assetID, $AppName) {
    # POST https://service.pima.edu/SBTDWebApi/Home/section/Tickets#POSTapi/{appId}/tickets/{id}/assets/{assetId}
    # Adds an asset to a ticket.
    $appID = Get-TDXAppID -AppName $AppName
    $uri = $baseURI + $appID + "/tickets/$ticketID/assets/$assetID"
    
    try {
        return Invoke-RestMethod -Method POST -Headers $tdxAPIAuth -Uri $uri -ContentType "application/json" -UseBasicParsing -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Edit-TDXTicketAddAsset API call on ticket $ticketID"
            Edit-TDXTicketAddAsset -ticketID $ticketID -assetID $assetID -AppName $AppName
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Editing the ticket $ticketID with asset $assetID has failed, see the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $apiError.ErrorRecord.Exception.Response.StatusCode)
            Write-Log -level ERROR -message ("Status Description - " + $apiError.ErrorRecord.Exception.Response.StatusDescription)
            Exit(1)
        }
    }
        
}

function Set-TDXTicketContact {
    param (
        [Parameter(Mandatory)]
        [int]$ID,
        [Parameter(Mandatory)]
        [guid]$Contact
    )
    
    $appID = Get-TDXAppID -AppName $AppName
    $uri = $baseURI + $appID + "/tickets/$ID/contacts/$Contact"

    try {
        return Invoke-RestMethod -Method POST -Headers $tdxAPIAuth -Uri $uri -ContentType "application/json" -UseBasicParsing -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Set-TDXTicketContact API call"
            Set-TDXTicketContact -ID $ID -Contact $Contact
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Setting contact for $ID with $Contact failed, see the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $apiError.ErrorRecord.Exception.Response.StatusCode)
            Write-Log -level ERROR -message ("Status Description - " + $apiError.ErrorRecord.Exception.Response.StatusDescription)
            Exit(1)
        }
    }

}

function Get-TDXTicketWorkflow($ticketID) {
    #GET https://service.pima.edu/SBTDWebApi/api/{appId}/tickets/{id}/workflow
    $uri = $baseURI + $appIDTicket + "/tickets/$ticketID/workflow"

    try {
        return Invoke-RestMethod -Method GET -Headers $apiHeaders -Uri $uri -ContentType "application/json" -UseBasicParsing -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Get-TDXTicketWorkflow API call on ticket $ticketID"
            Get-TDXTicketWorkflow -ticketID $ticketID
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Getting the workflow on ticket $ticketID has failed, see the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -message ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -message ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
}
#endregion
#region People
function Search-TDXPeople([string]$SearchString, [int]$MaxResults) {
    #GET 
    # https://pima.teamdynamix.com/SBTDWebApi/api/people/lookup?searchText={searchText}&maxResults={maxResults}
    $uri = $baseURI + "people/lookup?searchText=$SearchString&maxResults=$MaxResults"

    try {
        # Wishlist: Create logic to verify edit. Will need to use Invoke-Webrequest in order to get header info if it isnt an error
        return Invoke-RestMethod -Method GET -Headers $tdxAPIAuth -Uri $uri -ContentType "application/json" -UseBasicParsing -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying API call to edit the asset $($Asset.Tag)"
            Search-TDXPeople -SearchString $SearchString -MaxResults $MaxResults
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

function Get-TDXPersonDetails($UID) {
    # https://pima.teamdynamix.com/SBTDWebApi/api/people/{uid}
    $uri = $baseURI + "people/$UID"

    try {
        # Wishlist: Create logic to verify edit. Will need to use Invoke-Webrequest in order to get header info if it isnt an error
        return Invoke-RestMethod -Method GET -Headers $tdxAPIAuth -Uri $uri -ContentType "application/json" -UseBasicParsing -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError            
   
            Write-Log -level WARN -message "Retrying Get-TDXPersonDetails API call"
            Get-TDXPersonDetails -UID $UID
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

#region Reports
function Get-TDXAssetReport($ID) {
    # https://service.pima.edu/SBTDWebApi/Home/section/AssetReports#POSTapi/{appId}/assets/searches/{searchId}/results
    $uri = $baseURI + $appIDAsset + "/assets/searches/$ID/results"

    $body = [PSCustomObject]@{
        SearchText =	$null #	String	This field is nullable.	The search text to filter on. If specified, this will override any search text that may have been part of the saved search.
        Page       =	@{ #This field is required.	TeamDynamix.Api.RequestPage		The page of data being requested for the saved search.
            PageIndex = 0	#This field is required.	Int32	The 0-based page index to request.
            PageSize  =	200 #This field is required.	Int32	The size of each page being requested, ranging 1-200 (inclusive).
        }         
    } | ConvertTo-Json

    try {
        return Invoke-RestMethod -Method POST -Headers $tdxAPIAuth -Uri $uri -ContentType "application/json" -UseBasicParsing  -ErrorVariable apiError -Body $body
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError            
   
            # Recursively call the function
            Write-Log -level WARN -message "Retrying API call for report $ID"
            Get-TDXAssetReport -ID $ID
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Getting report ID $ID has failed, see the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -message ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -message ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    } 
}
#endregion

#region Attributes
function Get-TDXAttribute($ID, $AppName) {
    $appID = Get-TDXAppID -AppName $AppName
    $uri = $baseURI + "/attributes/$ID/choices"

    try {
        return Invoke-RestMethod -Method GET -Headers $tdxAPIAuth -Uri $uri -ContentType "application/json; charset=utf-8" -UseBasicParsing -ErrorVariable apiError
    }
    catch {
        # If we got rate limited, try again after waiting for the reset period to pass.
        if ($apiError.ErrorRecord.Exception.Response.StatusCode -eq 429) {

            # Sleep based on the rate limit time
            Get-TdxApiRateLimit -apiCallResponse $apiError
            
            # Recursively call the function
            Write-Log -level WARN -message "Retrying Get-TDXAttribute API call on ticket $ID"
            Get-TDXAttribute -ID $ID -AppName $AppName
        }
        else {
            # Display errors and exit script.
            Write-Log -level ERROR -message "Getting attribute info for $ID has failed, see the following log messages for more details."
            Write-Log -level ERROR -message ("Status Code - " + $_.Exception.Response.StatusCode.value__)
            Write-Log -level ERROR -message ("Status Description - " + $_.Exception.Response.StatusDescription)
            Write-Log -level ERROR -message ("Error Message - " + $_.ErrorDetails.Message)
            Exit(1)
        }
    }
}
#endregion

#region Helpers
function Get-TdxApiRateLimit($apiCallResponse) {
    <#
    # Get the total wait API limit
    $apiResetSeconds = $apiCallResponse.ErrorRecord.Exception.Response.GetResponseHeader("X-RateLimit-limit")
    Write-Log -level WARN -message "Waiting $apiResetSeconds seconds to rety API call due to rate-limiting."

    # Show progress bar on total wait
    $time = [int]$apiResetSeconds + 5
    foreach ($i in (1..$time)) {
        $percentage = $i / $time
        $remaining = New-TimeSpan -Seconds ($time - $i)
        $message = "{0:p0} complete, remaining time {1}" -f $percentage, $remaining
        Write-Progress -Activity $message -PercentComplete ($percentage * 100)
        Start-Sleep 1
    }
    Write-progress -Activity 'Done...' -Completed
    #>
    Write-Log -level WARN -message "Waiting 60 seconds to rety API call due to rate-limiting."
    Start-Sleep -Seconds 60
}

function Get-TDXAppID($AppName) {
    switch ($AppName) {
        ITTicket { '1257' }
        ITAsset { '1258' }
        D2LTicket { '1755' }
        Default {
            # Display errors and exit script.
            Write-Log -level ERROR -message "$AppName does not exist"
            Exit(1)
        }
    } 
}
#endregion
#endregion

# Get creds and create the base uri and header for all API calls
$appIDTicket = '1257'
$appIDAsset = '1258'
#$baseURI = "https://service.pima.edu/SBTDWebApi/api/"
$baseURI = "https://service.pima.edu/TDWebApi/api/"

#$tdxCreds = Get-Credential
#$tdxAPIAuth = Get-TDXAuth -beid $tdxCreds.UserName -key $tdxCreds.GetNetworkCredential().Password

$tdxCreds = Get-Content $PSScriptRoot\tdx.json | ConvertFrom-Json
$tdxAPIAuth = Get-TDXAuth -beid $tdxCreds.BEID -key $tdxCreds.Key