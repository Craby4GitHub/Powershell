. (Join-Path $PSSCRIPTROOT "TDX-API.ps1")

# Get all TDX Assets
$allTDXAssets = Search-TDXAssets -AppName ITAsset

# Find all assets with an Owner that arent disposed
$assetOwners = $allTDXAssets | Where-Object { ($_.OwningCustomerName -NE 'None') -and ($_.StatusName -ne 'Disposed') -and ($_.StatusName -ne 'Missing') }

# Set up an empty array for users already worked on
$coveredUser = New-Object System.Collections.Generic.List[System.Object]

# Set oldest allowed inventory date
$inventoryDateCutOff = Get-Date '04/01/2021'

# Go through each Asset Owner
foreach ($assetOwner in $assetOwners) {
    [int]$pct = ($assetOwners.IndexOf($assetOwner) / $assetOwners.count) * 100
    Write-progress -Activity "$($assetOwner.OwningCustomerName)" -percentcomplete $pct -status "$pct% Complete"

    # See if user has been already worked on and skip them if so
    if ($coveredUser.IndexOf($assetOwner.OwningCustomerName) -ge 0) {
        # Skip
    }
    else {
        # Add to covered list
        $coveredUser.Add($assetOwner.OwningCustomerName) 

        # Get all assets assigned to user
        $allUserAssets = $assetOwners | Where-Object -Property OwningCustomerName -EQ $assetOwner.OwningCustomerName

        # Go through each asset
        :inner foreach ($userAsset in $allUserAssets) {

            # Get asset's custom attributes
            $assetCustomAttributes = Get-TDXAssetAttributes -ID $userAsset.ID

            # Select the last inventory date
            $lastInventoryDate = $assetCustomAttributes | Where-Object -Property Name -eq 'Last Inventory Date' | Select-Object -ExpandProperty value 

            # Check to see if the current inventory date is older than the Cut Off date
            if ($inventoryDateCutOff -gt $lastInventoryDate) {
                # Prepare ticket description
                $Description = "Hello $($assetOwner.OwningCustomerName),`nProperty Control is performing an annual inventory of IT assets. Please review the assets below and reply to this email/ticket indicating if the asset information is correct. If incorrect, please list any discrepancies.`nIf you have other IT assets that are not listed below, please provide the PCC number and a brief description of the device in your response. We will then update our records.`n`n"

                # Adding all of user's assets to Description
                foreach ($userAsset in $allUserAssets) {
                    $Description += "PCC Number: $($userAsset.Tag)`nAsset Model: $($userAsset.ManufacturerName) $($userAsset.ProductModelName)`nCampus and Room: $($userAsset.LocationName) $($userAsset.LocationRoomName)`n`n" 
                }

                $Description += "`nIf you have any questions, please feel free to respond to this email or call our Service Desk at 206-4900 and reference this ticket."

                # Get user details
                $userDetails = Get-TDXPersonDetails -UID $assetOwner.OwningCustomerID
                <#  Cant use LocationName because this isnt loaded into TDX from Banner   :[    Save for later if this maybe gets fixed?
                    # Get the users home location and set the responsible group based on that
                    switch ($userDetails.LocationName) {
                        "29th St LC" { $responsibleGroup = 11684 } # DO
                        "Ajo" { $responsibleGroup = 11684 } # DO
                        "Arizona State Prison - Douglas" { $responsibleGroup = 11683 } # EC
                        "Arizona State Prison - Wilmot" { $responsibleGroup = 11683 } # EC
                        "Aviation Technology" { $responsibleGroup = 11682 } # DV
                        "Davis Monthan AFB" { $responsibleGroup = 11683 } # EC
                        "Desert Vista Campus" { $responsibleGroup = 11682 } # DV
                        "District Office" { $responsibleGroup = 11684 } # DO
                        "Downtown Campus" { $responsibleGroup = 11681 } # DC
                        "East Campus" { $responsibleGroup = 11683 } # EC
                        "El Pueblo LC" { $responsibleGroup = 11684 } # DO
                        "El Rio LC" { $responsibleGroup = 11684 } # DO
                        "Maintenance and Security" { $responsibleGroup = 11684 } # DO
                        "Northwest Campus" { $responsibleGroup = 11680 } # NW
                        "Not Campus Specific" { $responsibleGroup = 11684 } # DO
                        "Pima County Jail" { $responsibleGroup = 11684 } # DO
                        "Santa Cruz Center" { $responsibleGroup = 11682 } # DV
                        "Truck Driver Training" { $responsibleGroup = 11684 } # DO
                        "West Campus" { $responsibleGroup = 11679 } # WC
                        Default { $responsibleGroup = 11412 } # SD
                    }
                    #>
                # Get the user's home location based on the work address and set the responsible group
                switch ($userDetails.WorkAddress.Split(' ')[0]) {
                    "29" { $responsibleGroup = 11684 } # DO
                    "CC" { $responsibleGroup = 11684 } # DO
                    "49" { $responsibleGroup = 11684 } # DO
                    "DV" { $responsibleGroup = 11682 } # DV
                    "DO" { $responsibleGroup = 11684 } # DO
                    "DC" { $responsibleGroup = 11681 } # DC
                    "East" { $responsibleGroup = 11683 } # EC
                    "EC" { $responsibleGroup = 11683 } # EC
                    "El" { $responsibleGroup = 11684 } # DO
                    "M" { $responsibleGroup = 11684 } # DO - M&S
                    "NW" { $responsibleGroup = 11680 } # NW
                    "Santa" { $responsibleGroup = 11682 } # DV
                    "WC" { $responsibleGroup = 11679 } # WC
                    Default { $responsibleGroup = 11412 } # SD
                }

                # Setup the ticket attributes
                $ticketAttributes = @{
                    AccountID          = 75673 # Campus Staff(CAMSTF)
                    PriorityID         = 4537 # Normal
                    RequestorUid       = $assetOwner.OwningCustomerID 
                    ResponsibleGroupID = $responsibleGroup # Based on the users home site
                    StatusID           = 52423 # In Process
                    Title              = 'IT Asset Inventory 2022' 
                    TypeID             = 30386 # USS - Service Desk
                    ClassificationID   = '46' # Service Request
                    Description        = $Description 
                    ServiceID          = 43361 # Campus IT Maintenance Tasks
                    FormID             = 37861 # Service Request
                    AppName = ITTicket
                    #LocationID         = $userDetails.LocationID
                }
                    
                # Create the ticket
                $ticket = Submit-TDXTicket @ticketAttributes

                # Add assets directly to the ticket
                foreach ($userAsset in $allUserAssets) {
                    Edit-TDXTicketAddAsset -ticketID $ticket.ID -assetID $userAsset.ID        
                }
                    
                Write-Host "Ticket created for $($userAsset.OwningCustomerName), $($ticket.ID)"
            }
            # Atleast one asset assigned to the user needs to be inventoried. Once one matches, dont need to continue checking
            break :inner
        }
    }
}
Write-progress -Activity 'Done!' -Completed