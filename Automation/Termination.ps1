
Function Get-FileName($initialDirectory) {  
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "Excel Files (*.xlsx)|*.xlsx"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

$importedTerminations = Import-Excel Get-FileName -initialDirectory 'c:\'


# Setup the ticket attributes based on TDX application
$ticketAttributes = @{
    AccountID        = 94478 # Students(PCCSTU)
    PriorityID       = 4537 # Normal
    StatusID         = 52421 # New
    Title            = "IT Asset Out Processing " 
    TypeID           = 30386 # USS - Service Desk
    ClassificationID = '46' # Service Request
    ServiceID        = 32655 # General IT Support
    FormID           = 54995 # BlackBoard Ticket Import
    SourceID         = 2517 # Blackboard
    IsRichHtml       = $true
    AppName          = 'ITTicket'
    Attribute        = @(@{
            ID    = 124304  # Blackboard Ticket Number
            Value = $ticketNumber
        }, @{
            ID    = 124429  # Phone Number
            Value = $userPhoneNumber
        }
    )
}

# Go through each termination
foreach ($termination in $importedTerminations) {

    $termDate = [DateTime]::FromOADate($termination.TERM_DATE)
    # Only want to work on valid users
    if ($termination.EMPLOYEE_ID -notmatch "A\d{8}") {
        Write-host "Non valid A-Number for $($termination.EMPLOYEE_NAME), skipping." 
    }
    else {
        # Find the user in TDX
        $tdxUser = Search-TDXPeople -SearchString $termination.EMPLOYEE_ID -MaxResults 1
        if ($null -eq $tdxUser.UID) {
            Write-Host "$($termination.EMPLOYEE_NAME), $($termination.EMPLOYEE_ID) is not in TDX"
        }
        else {
            #
            $Description = "Hello $($tdxUser.FullName),`nAs you may be aware, you are required to return all IT equipment issued to you upon your departure from Pima CC. This includes the following equipment:`n"
            
            # Look up users assigned assets
            $assignedAssets = Search-TDXAssets -Term OwningCustomerIDs -Value @($tdxUser.UID) -AppName ITAsset
            if ($assignedAssets.count -le 0) {
                write-host "No IT assets for $($tdxUser.FullName), $($tdxUser.PrimaryEmail)"
                break
            }
            else {
                # Adding all of user's assets to Description
                foreach ($asset in $assignedAssets) {
                    # Note: add product type? ie laptop, desktop
                    $Description += "
                        PCC Number: $($asset.Tag)`n
                        Asset Model: $($asset.ManufacturerName) $($asset.ProductModelName)`n
                        Campus and Room: $($asset.LocationName) $($asset.LocationRoomName)`n
                        "  
                }
                $Description += "
                Please make sure to fully power down before returning it to the IT department. If you are unable to return the equipment to your local IT shop, please let us know and we will arrange for it to be picked up.`n
                If you have any questions or concerns about returning your IT equipment, or if you believe there is an error in the equipment listed above, please reply to this email with any updates or corrections.`n
                Thank you for your cooperation.`n
                Best regards,`n
                PCC IT
                "
    
                # Get user details
                $userDetails = Get-TDXPersonDetails -UID $tdxUser.UID

                # Get the user's home location based on the work address
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

                # Setting important ticket values
                $ticketAttributes.Description = $Description
                $ticketAttributes.ResponsibleGroupID = $responsibleGroup
                $ticketAttributes.RequestorUid = $tdxUser.UID 

                # Create the ticket
                $tdxTicket = $null
                $tdxTicket = Submit-TDXTicket @ticketAttributes

                if ($null -eq $tdxTicket.id) {
                    Write-Host "Failed to create ticket for $($tdxUser.FullName), $($tdxUser.PrimaryEmail)"                
                }
                else {
                
                    # Add assets directly to the ticket
                    foreach ($asset in $assignedAssets) {
                        Edit-TDXTicketAddAsset -ticketID $tdxTicket.ID -assetID $asset.ID -AppName 'ITTicket'
                    }
                    
                    Write-Host "Ticket created for $($tdxUser.FullName), $($tdxTicket.ID)"
                }
            }
        }
    }
}