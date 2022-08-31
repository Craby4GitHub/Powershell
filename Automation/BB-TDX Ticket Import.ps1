# Load TDX API functions
. ((Get-Item $PSScriptRoot).Parent.FullName + '\TDX\TDX-API.ps1')


$baseURL = 'https://servicedesk.edusupportcenter.com/api/v1'

#$blackboardCredentials = Get-Credential

$body = [PSCustomObject]@{
    username = $blackboardCredentials.UserName
    password = $blackboardCredentials.GetNetworkCredential().Password
} | ConvertTo-Json

$authResponse = Invoke-RestMethod -Method POST -Uri "$baseURL/login" -body $body -ErrorVariable apiError -ContentType "application/json"
if ($authResponse.error_msg -eq 'login success') {

    # Set Bearer token, will be uased for all other APi calls
    $token = $authResponse.return_body.token

    # Get ticket queue and loop through them
    $queueResponse = Invoke-RestMethod -Method GET -Uri "$baseURL/2028/case?_queue_=976&token=$token&_pageSize_=100&_startPage_=1" -ErrorVariable apiError -ContentType "application/json"
    $ticketQueue = $queueResponse.return_body.items   
    write-host "$($ticketQueue.count) tickets"
    foreach ($request in $ticketQueue) {

        # Get BB ticket details
        $ticketNumber = $request.ticket_number.Split('-')[1]
        $ticket = (Invoke-RestMethod -Method GET -Uri "$baseURL/2028/case/$($ticketNumber)?token=$token&_replaceRuntime_=true&_history_=true&_source_=sd&_domain_=https://servicedesk.edusupportcenter.com" -ErrorVariable apiError -ContentType "application/json").return_body
        write-host $ticket.summary
        write-host $ticketNumber

        # Get user phone number
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

        $searchContacts = Invoke-RestMethod -Method POST -Uri "$baseURL/2028/searchContacts?token=$token&_pageSize_=10&_startPage_=1" -ErrorVariable apiError -ContentType "application/json" -body $body
    
        # Verify only 1 result and set the phone number
        if ($searchContacts.return_body.count -eq 1) {
            $userPhoneNumber = $searchContacts.return_body.items.primary_phone
        }
        else {
            $userPhoneNumber = 'None'
        }
       

        $caseDetails = $ticket.values | Where-Object -Property label -eq 'Case Details' | Select-Object -ExpandProperty value

        $Description = "Hello $($request.customer.name.split(' ')[0]),`n
        The IT Service Desk at Pima Community College has received your case from our Tier 1 support. We will be contacting you soon via the contact information listed below.`n
            Case Details: $($caseDetails -replace '<[^>]+>','')
            Phone Number: $userPhoneNumber
            Email: $($searchContacts.return_body.items.email) `n
            If this is an urgent matter, please call us at (520) 206-4900 and reference this ticket.`n
            "

        # Get user TDX user profile
        $userANumber = $ticket.values | Where-Object -Property label -eq 'A Number' | Select-Object -ExpandProperty value 
        $tdxUserUID = '05a7e4ad-3409-ed11-bd6e-0003ff5063cd' # Inactive Student
        if ($userANumber -match 'a\d{8}') {
            $tdxUser = Search-TDXPeople -SearchString $userANumber
            if ($null -ne $tdxUser.UID) {
                $tdxUserUID = $tdxUser.UID
            }
        }


        # Setup the ticket attributes
        $ticketAttributes = @{
            AccountID          = 75673 # Campus Staff(CAMSTF)
            PriorityID         = 4537 # Normal
            RequestorUid       = $tdxUserUID 
            ResponsibleGroupID = 11412 # USS - Service Desk
            StatusID           = 52421 # New
            Title              = $ticket.summary 
            TypeID             = 30386 # USS - Service Desk
            ClassificationID   = '46' # Service Request
            Description        = $Description 
            ServiceID          = 32655 # General IT Support
            FormID             = 54995 # BlackBoard Ticket Import
            SourceID           = 2517 # Blackboard
            IsRichHtml         = $true
            Attribute          = @(@{
                    ID    = 124304  # Blackboard Ticket Number
                    Value = $ticketNumber
                }, @{
                    ID    = 124429  # Phone Number
                    Value = $userPhoneNumber
                }
            )
        }
                            
        # Create the ticket
        $tdxTicket = $null
        $tdxTicket = Submit-TDXTicket @ticketAttributes
    
        # Verify the ticket was created
        if ($null -ne $tdxTicket.id) {
            write-host $tdxTicket.id

            # Add ticket history to TDX ticket
            if ($ticket.history.count -ge 1) {
                [array]::Reverse($ticket.history)
                foreach ($history in $ticket.history) {
                    # converting unix time...
                    $UnixDate = [int64]$history.action_date / 1000
                    $time = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($UnixDate))

                    $comment = "Date: $time

                        Action: $($history.action)

                        Performed By: $($history.performed_by)

                        Comment: $($history.comment)"

                    $ticketHistory = Update-TDXTicket -ticketID $tdxTicket.ID -Comment $comment -IsPrivate $true -IsRichHtml $true
                }
            }

            # Check if there are chat transcripts and add as a comment to TDX ticket
            # TODO: have it so it outputs to tdx in a more clean manner
            if ($ticket.chatTranscriptAttachment.count -ge 1) {
                write-host "chat log found" -ForegroundColor Yellow
                foreach ($chat in $ticket.chatTranscriptAttachment) {
                    $chatLog = Invoke-RestMethod -Method GET -Uri "$baseURL/2028/file/$($chat.ID)?token=$token"
                    $ticketChat = Update-TDXTicket -ticketID $tdxTicket.ID -Comment $chatLog -IsPrivate $true -IsRichHtml $false
                }
            }

            # Check if there are attachments and add to TX ticket
            # Current: does not attach to the ticket
            # Ideas: check for .txt | .html and add as commnt to ticket?
            if ($ticket.Attachment.count -ge 1) {
                $ticketattachment = Update-TDXTicket -ticketID $tdxTicket.ID -Comment "There are attachments related to this ticket. Please navigate to this ticket in Blackboard to download." -IsPrivate $true -IsRichHtml $true
                foreach ($attachment in $ticket.Attachment) {
                    write-host "attachments found" -ForegroundColor Yellow
                    #Invoke-webrequest -Uri "$baseURL/2028/file/$($attachment.ID)?token=$token" -UseBasicParsing -OutFile "c:\users\wrcrabtree\downloads\$($attachment.id)+$($attachment.name)"   
                }
            }


            # L3 Grab ticket

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
            # close BB ticket

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