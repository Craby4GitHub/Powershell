. (Join-Path $PSSCRIPTROOT "TDX-API.ps1")

# Get all TDX Assets
$allTDXAssets = Search-TDXAssets

# Find all assets with an Owner
$assetOwners = $allTDXAssets | Where-Object -Property OwningCustomerName -NE 'None'

# Set up an empty array for users already worked on
$coveredUser = New-Object System.Collections.Generic.List[System.Object]

# Go through each Asset Owner
foreach ($assetOwner in $assetOwners) {

    # See if user already worked on and skip them if so
    if ($coveredUser.IndexOf($assetOwner.OwningCustomerName) -ge 0) {
        Write-host 'User already covered'
    }
    else {
        # Add to covered list
        $coveredUser.Add($assetOwner.OwningCustomerName) 

        # Get all assets assigned to user
        $allUserAssets = $assetOwners | Where-Object -Property OwningCustomerName -EQ $assetOwner.OwningCustomerName

        # Prepare ticket description
        $Description = "Hello $($assetOwner.OwningCustomerName),`nThe following assets are assigned to you:`n"

        # Adding users assets to Description
        foreach ($userAsset in $allUserAssets) {
            $Description += "PCC Number: $($userAsset.Tag)`nAsset Model: $($userAsset.ManufacturerName) $($userAsset.ProductModelName)`nCampus and Room: $($userAsset.LocationName) $($userAsset.LocationRoomName)`n`n" 
        }

        # Create the ticket
        # Add: change responsible group to ???
        # issue: Sets the ticket as an incident instead of service
        $ticket = Submit-TDXTicket -AccountID 75424 -PriorityID 4537 -RequestorUid $assetOwner.OwningCustomerID -StatusID 52423 -Title 'IT Asset Inventory 2022' -TypeID 41029 -Description $Description -ServiceID 51998
        
        # Add assets directly to the ticket
        foreach ($userAsset in $allUserAssets) {
            Edit-TDXTicketAddAsset -ticketID $ticket.ID -assetID $userAsset.ID        
        }
    }
}