
Function Get-FileName($initialDirectory) {  
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "Excel Files (*.xlsx)|*.xlsx"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

$importedTerminations = Import-Excel "C:\Users\wrcrabtree\Downloads\_ce26f0bedbca3b941f6.xlsx"


# Setup the ticket attributes based on TDX application
$ticketAttributes = @{
    AccountID          = 94478 # Students(PCCSTU)
    PriorityID         = 4537 # Normal
    ResponsibleGroupID = 11412 # USS - Service Desk
    StatusID           = 52421 # New
    Title              = "IT Asset Out Processing " 
    TypeID             = 30386 # USS - Service Desk
    ClassificationID   = '46' # Service Request
    ServiceID          = 32655 # General IT Support
    FormID             = 54995 # BlackBoard Ticket Import
    SourceID           = 2517 # Blackboard
    IsRichHtml         = $true
    AppName            = 'ITTicket'
    Attribute          = @(@{
            ID    = 124304  # Blackboard Ticket Number
            Value = $ticketNumber
        }, @{
            ID    = 124429  # Phone Number
            Value = $userPhoneNumber
        }
    )
}



foreach ($termination in $importedTerminations) {

    $termDate = [DateTime]::FromOADate($termination.TERM_DATE)

    if ($termination.EMPLOYEE_ID -match "A\d{8}") {
        
        $tdxUser = Search-TDXPeople -SearchString $termination.EMPLOYEE_ID
        if ($null -ne $tdxUser.UID) {
            $ticketAttributes.RequestorUid = $tdxUser.UID 

            $assignedAssets = Search-TDXAssets -Term OwningCustomerIDs -Value @($tdxUser.UID) -AppName ITAsset

            if ($assignedAssets.count -ge 1) {
                <# Action to perform if the condition is true #>
            }else{
                write-host "No IT assets for $($termination.EMPLOYEE_NAME)"
            }
        }
        else {
            $ticketAttributes.RequestorUid = $null
            $ticketAttributes.RequestorName = $termination.EMPLOYEE_NAME                         # The full name of the requestor associated with the ticket.
            $ticketAttributes.RequestorFirstName = "$($termination.EMPLOYEE_NAME.split(', ')[1])" # The first name of the requestor associated with the ticket.
            $ticketAttributes.RequestorLastName = "$($termination.EMPLOYEE_NAME.split(', ')[0])"	 # The last name of the requestor associated with the ticket.
        }

        # Create the ticket
        $tdxTicket = $null
        $tdxTicket = Submit-TDXTicket @ticketAttributes

    }
    else {
        Write-host "No A Number for $($termination.EMPLOYEE_NAME)"
    }

