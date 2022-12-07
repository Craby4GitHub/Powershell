# Load TDX API functions
. ((Get-Item $PSScriptRoot).Parent.FullName + '\TDX\TDX-API.ps1')


$baseURL = 'https://servicedesk.edusupportcenter.com/api/v1'

$blackboardCredentials = Get-Credential

$body = [PSCustomObject]@{
    username = $blackboardCredentials.UserName
    password = $blackboardCredentials.GetNetworkCredential().Password
} | ConvertTo-Json

$authResponse = Invoke-RestMethod -Method POST -Uri "$baseURL/login" -body $body -ErrorVariable apiError -ContentType "application/json"
if ($authResponse.error_msg -eq 'login success') {
    # Set Bearer token, will be uased for all other APi calls
    $token = $authResponse.return_body.token

    # Get ticket queue and loop through them
    $ticketQueue = $null

    $ticketQueue = (Invoke-RestMethod -Method GET -Uri "$baseURL/2028/case?_queue_=976&token=$token&_pageSize_=10&_startPage_=1" -ErrorVariable apiError -ContentType "application/json").return_body.items
    
    write-host "$($ticketQueue.count) tickets for ITTicket"
    foreach ($request in $ticketQueue) {
        # Get BB ticket details
        $ticketNumber = $request.ticket_number.Split('-')[1]
        $ticket = (Invoke-RestMethod -Method GET -Uri "$baseURL/2028/case/$($ticketNumber)?token=$token&_replaceRuntime_=true&_history_=true&_source_=sd&_domain_=https://servicedesk.edusupportcenter.com" -ErrorVariable apiError -ContentType "application/json").return_body
        write-host $ticket.summary
        write-host $ticketNumber

        # Get user phone number by searching for Name and A Number
        $userANumber = $ticket.values | Where-Object -Property label -eq 'A Number' | Select-Object -ExpandProperty value 
            
        $body = [PSCustomObject]@{
            email         = ""
            extra_user_id = ""
            fields        = @()
            firstName     = ""
            first_name    = "$($request.customer.name.split(' ')[0])"
            lastName      = "$($request.customer.name.split(' ')[1])"
            last_name     = ""
            primary_phone = ""
            userId        = ""
            userName      = ""
            user_id       = ""
            user_name     = "$userANumber"
        } | ConvertTo-Json

        $searchContacts = Invoke-RestMethod -Method POST -Uri "$baseURL/2028/searchContacts?token=$token&_pageSize_=5&_startPage_=1" -ErrorVariable apiError -ContentType "application/json" -body $body
    
        # Verify only 1 result and set the phone number/email
        if ($searchContacts.return_body.count -eq 1) {
            $userPhoneNumber = $searchContacts.return_body.items.primary_phone
            $userEmail = $searchContacts.return_body.items.email
        }
        else {
            $userPhoneNumber = 'None'
            $userEmail = 'None'
        }

        # Setup the ticket attributes based on TDX application
        $ticketAttributes = @{
            AccountID          = 94478 # Students(PCCSTU)
            PriorityID         = 4537 # Normal
            ResponsibleGroupID = 11412 # USS - Service Desk
            StatusID           = 52421 # New
            Title              = $ticket.summary 
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

        # Get user TDX user profile to have them as the requestor in the ticket
        $tdxUser = Search-TDXPeople -SearchString $userANumber
        if ($null -ne $tdxUser.UID) {
            $ticketAttributes.RequestorUid = $tdxUser.UID 
        }
        else {
            $ticketAttributes.RequestorUid = $null
            $ticketAttributes.RequestorName = $request.customer.name                         # The full name of the requestor associated with the ticket.
            $ticketAttributes.RequestorFirstName = "$($request.customer.name.split(' ')[0])" # The first name of the requestor associated with the ticket.
            $ticketAttributes.RequestorLastName = "$($request.customer.name.split(' ')[1])"	 # The last name of the requestor associated with the ticket.
            $ticketAttributes.RequestorEmail = $userEmail		                             # The email address of the requestor associated with the ticket.
            $ticketAttributes.RequestorPhone = $userPhoneNumber	                             # The phone number of the requestor associated with the ticket.
        }   
        
        # Get details to be used in ticket description
        $caseDetails = $ticket.values | Where-Object -Property label -eq 'Case Details' | Select-Object -ExpandProperty value
        $caseDetailsConverted = $caseDetails -replace '&nbsp;', ' '
        $caseDetailsConverted = $caseDetailsConverted -replace '<[^>]+>', ''

        $ticketAttributes.Description = "Hello $($request.customer.name.split(' ')[0]),
        <p>The IT Service Desk at Pima Community College has received your case from our Tier 1 support and we will be contacting you shortly.</p>
        <p>Case Details: <p><b>$caseDetailsConverted</b></p></p>
        If this is an urgent matter, please call us at (520) 206-4900 and reference this ticket.
        " 
               
        # Create the ticket
        $tdxTicket = $null
        $tdxTicket = Submit-TDXTicket @ticketAttributes
    
        # Verify the ticket was created
        if ($null -ne $tdxTicket.id) {
            write-host $tdxTicket.id

            # Add BB ticket history to TDX ticket
            if ($ticket.history.count -ge 1) {
                # Reversing history as it would put the newest comment at the bottom
                [array]::Reverse($ticket.history)

                foreach ($history in $ticket.history) {
                    # This will select the most recent email comment and post it as a comment
                    if ($history.action_type -eq 'EMAIL_RESPONSE') {
                        $HTML = New-Object -Com "HTMLFile"
                        $src = [System.Text.Encoding]::Unicode.GetBytes($history.comment)
                        $html.write($src)
                        $history.comment = $html.all.tags("div") | ForEach-Object innertext | Select-Object -first 1
                    }
                    
                    # Make sure there is a comment made
                    if ($history.comment.Length -gt 0) {
                        $ticketHistory = Update-TDXTicket -ticketID $tdxTicket.ID -Comment $history.comment -IsPrivate $true -IsRichHtml $true -AppName 'ITTicket'
                    }
                }     
            }

            # Check if there are chat transcripts and add as a comment to TDX ticket
            # TODO: have it so it outputs to tdx in a more clean manner
            if ($ticket.chatTranscriptAttachment.count -ge 1) {
                write-host "External chat log found, adding to the ticket" -ForegroundColor Yellow
                foreach ($chat in $ticket.chatTranscriptAttachment) {
                    $chatLog = Invoke-RestMethod -Method GET -Uri "$baseURL/2028/file/$($chat.ID)?token=$token"
                    #$convertedChatLog = $Chatlog.split("`n")  
                    #$convertedChatLog[8..($convertedChatLog.getupperbound(0) - 1)]                   
                    $tdxChatLog = Update-TDXTicket -ticketID $tdxTicket.ID -Comment $chatLog -IsPrivate $true -IsRichHtml $false -AppName 'ITTicket'
                }
            }

            # Check if there are attachments and add to TX ticket
            # Current: does not attach to the ticket as I dont have the API command setup to add attachments
            # Ideas: check for .txt | .html and add as commnt to ticket?
            if ($ticket.Attachment.count -ge 1) {
                #$ticketattachment = Update-TDXTicket -ticketID $tdxTicket.ID -Comment "There are attachments related to this ticket. Please navigate to this ticket in Blackboard to download." -IsPrivate $true -IsRichHtml $true -AppName 'ITTicket'
                foreach ($attachment in $ticket.Attachment) {
                    write-host "attachments found" -ForegroundColor Yellow
                    #Invoke-webrequest -Uri "$baseURL/2028/file/$($attachment.ID)?token=$token" -UseBasicParsing -OutFile "c:\users\wrcrabtree\downloads\$($attachment.id)+$($attachment.name)"   
                }
            }

            # Create a blank object which will be used for an API that assigns the ticket to us in BB
            $body = [PSCustomObject]@{
                attachment           = @()
                comment              = ""
                csr                  = ""
                csrCCList            = ""
                customerCCList       = ""
                queue                = ""
                selectedCsr          = ""
                suppressExtComm      = ""
                transferInstitution  = ""
                transferTemplateType = ""
                useKbs               = @()
            } | ConvertTo-Json

            $grabTicket = Invoke-RestMethod -Method POST -Uri "$baseURL/2028/case/$($ticketNumber)/action/31745?token=$token" -ErrorVariable apiError -ContentType "application/json" -Body $body
            write-host $grabTicket

            # Create a object with a comment which will be used for an API that closes the ticket
            $body = [PSCustomObject]@{
                attachment           = @()
                comment              = "Ticket has been escalated to the PCCs IT ticketing system. Ticket: https://service.pima.edu/TDClient/1920/Portal/Requests/TicketRequests/TicketDet?TicketID=$($tdxTicket.id)"
                csr                  = ""
                csrCCList            = ""
                customerCCList       = ""
                queue                = ""
                selectedCsr          = ""
                suppressExtComm      = ""
                transferInstitution  = ""
                transferTemplateType = ""
                useKbs               = @()
            } | ConvertTo-Json

            $closeTicket = Invoke-RestMethod -Method POST -Uri "$baseURL/2028/case/$($ticketNumber)/action/31734?token=$token" -ErrorVariable apiError -ContentType "application/json" -Body $body
            write-host $closeTicket
        }
        else {
            Write-host "Failed to Create ticket for $ticketNumber"
        }
    }   
}
else {
    Write-Host "BB Auth failed"
    Start-Sleep -Seconds 5
}