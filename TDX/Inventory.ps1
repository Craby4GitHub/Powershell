. (Join-Path $PSSCRIPTROOT "TDX-API.ps1")

# Get all TDX Assets
$allTDXAssets = Search-TDXAssets

# Find all assets with an Owner
$assetOwners = $allTDXAssets | Where-Object -Property OwningCustomerName -NE 'None'

# Set up an empty array for users already worked on
$coveredUser = New-Object System.Collections.Generic.List[System.Object]

# Go through each Asset Owner
foreach ($assetOwner in $assetOwners) {
    [int]$pct = ($assetOwners.IndexOf($assetOwner) / $assetOwners.count) * 100
    Write-progress -Activity "Working on $($assetOwner.OwningCustomerName)" -percentcomplete $pct -status "$pct% Complete"

    # See if user already worked on and skip them if so
    if ($coveredUser.IndexOf($assetOwner.OwningCustomerName) -ge 0) {
        # Skip
    }
    else {
        # Add to covered list
        $coveredUser.Add($assetOwner.OwningCustomerName) 

        # Get all assets assigned to user
        $allUserAssets = $assetOwners | Where-Object -Property OwningCustomerName -EQ $assetOwner.OwningCustomerName

        # Prepare ticket description
        $Description = "Hello $($assetOwner.OwningCustomerName),`nLink to your Profile: https://service.pima.edu/TDClient/1920/Portal/People/Details?ID=$($assetOwner.OwningCustomerID)`nThe following assets are assigned to you:`n`n"

        # Adding users assets to Description
        foreach ($userAsset in $allUserAssets) {
            $Description += "PCC Number: $($userAsset.Tag)`nAsset Model: $($userAsset.ManufacturerName) $($userAsset.ProductModelName)`nCampus and Room: $($userAsset.LocationName) $($userAsset.LocationRoomName)`n`n" 
        }

        # Get the users home location and set the responsible group based on that
        $userDetails = Get-TDXPersonDetails -UID $assetOwner.OwningCustomerID
        switch ($userDetails.LocationName) {
            "29th St LC" { $responsibleGroup = 11412 } # SD
            "Ajo" { $responsibleGroup = 11412 } # SD
            "Arizona State Prison - Douglas" { $responsibleGroup = 11683 } # EC
            "Arizona State Prison - Wilmot" { $responsibleGroup = 11683 } # EC
            "Aviation Technology" { $responsibleGroup = 11682 } # DV
            "Davis Monthan AFB" { $responsibleGroup = 11683 } # EC
            "Desert Vista Campus" { $responsibleGroup = 11682 } # DV
            "District Office" { $responsibleGroup = 11684 } # DO
            "Downtown Campus" { $responsibleGroup = 11681 } # DC
            "East Campus" { $responsibleGroup = 11683 } # EC
            "El Pueblo LC" { $responsibleGroup = 11412 } # SD
            "El Rio LC" { $responsibleGroup = 11412 } # SD
            "Maintenance and Security" { $responsibleGroup = 11684 } # DO
            "Northwest Campus" { $responsibleGroup = 11680 } # NW
            "Not Campus Specific" { $responsibleGroup = 11412 } # SD
            "Pima County Jail" { $responsibleGroup = 11412 } # SD
            "Santa Cruz Center" { $responsibleGroup = 11412 } # SD
            "Truck Driver Training" { $responsibleGroup = 11684 } # DO
            "West Campus" { $responsibleGroup = 11679 } # WC
            Default { $responsibleGroup = 11412 } # SD
        }

        # Setup the ticket attributes
        $ticketOptions = @{
            AccountID          = 75673 # Campus Staff(CAMSTF)
            PriorityID         = 4537 # Normal
            RequestorUid       = $assetOwner.OwningCustomerID 
            ResponsibleGroupID = $responsibleGroup # Based on the users home site
            StatusID           = 52423 # In Process
            Title              = 'IT Asset Inventory 2022' 
            TypeID             = 41029 # Asset Management
            ClassificationID = '46' # Service Request
            Description        = $Description 
            ServiceID          = 51998 # Asset Management
            #FormID             = 57683 # Asset Service
            LocationID        = $userDetails.LocationID
        }

        # Create the ticket
        $ticket = Submit-TDXTicket @ticketOptions

        # Add assets directly to the ticket
        foreach ($userAsset in $allUserAssets) {
            Edit-TDXTicketAddAsset -ticketID $ticket.ID -assetID $userAsset.ID        
        }
    }
}
Write-progress -Activity 'Done...' -Completed