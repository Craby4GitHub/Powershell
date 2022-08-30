# Load TDX API functions
. ((Get-Item $PSScriptRoot).Parent.FullName + '\TDX\TDX-API.ps1')


$baseURL = 'https://servicedesk.edusupportcenter.com/api/v1'

#$blackboardCredentials = Get-Credential

$body = [PSCustomObject]@{
    username = $blackboardCredentials.UserName
    password = $blackboardCredentials.GetNetworkCredential().Password
} | ConvertTo-Json

$authResponse = Invoke-RestMethod -Method POST -Uri "$baseURL/login" -body $body -ErrorVariable apiError -ContentType "application/json"

$token = $authResponse.return_body.token

# Get queue
$queueResponse = Invoke-RestMethod -Method GET -Uri "$baseURL/2028/case?_queue_=976&token=$token&_pageSize_=1&_startPage_=1" -ErrorVariable apiError -ContentType "application/json"
$ticketQueue = $queueResponse.return_body.items   

write-host "$($ticketQueue.count) tickets"
foreach ($request in $ticketQueue) {
    # get ticket details
    $ticketNumber = $request.ticket_number.Split('-')[1]
    $ticket = (Invoke-RestMethod -Method GET -Uri "$baseURL/2028/case/$($ticketNumber)?token=$token&_replaceRuntime_=true&_history_=true&_source_=sd&_domain_=https://servicedesk.edusupportcenter.com" -ErrorVariable apiError -ContentType "application/json").return_body
    write-host $ticket.summary

    # Get user information
    $userANumber = $ticket.values | Where-Object -Property label -eq 'A Number' | Select-Object -ExpandProperty value 
    $tdxUser = Search-TDXPeople -SearchString $userANumber
    if ($tdxUser -match 'a\d{8}') {
        $tdxUserUID = $tdxUser.UID
    }
    else {
        #$tdxUserUID = '05a7e4ad-3409-ed11-bd6e-0003ff5063cd' # Inactive Student
        $tdxUserUID = '0344064a-8906-e911-a964-000d3a137856' # Inactive Student
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
        ServiceID          = 43361 # Campus IT Maintenance Tasks
        FormID             = 54995 # BlackBoard Ticket Import
        SourceID           = 2517 # Blackboard
    }
                        
    # Create the ticket
    $tdxTicket = Submit-TDXTicket @ticketAttributes

write-host $tdxTicket.id
    # Ticket history
    if ($ticket.history.count -ge 1) {
        foreach ($history in $ticket.history) {
            # converting time...
            $UnixDate = [int64]$ticket.history[0].action_date / 1000
            $time = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($UnixDate))

            $comment = "Action: $($history.action)
                        Date: $time
                        Performed By: $($history.performed_by)
                        Comment: $($history.comment)"

            Update-TDXTicket -ticketID $tdxTicket.ID -Comment $comment -IsPrivate $true -IsRichHtml $true
        }
    }

    # Check if there are chat transcripts
    if ($ticket.chatTranscriptAttachment.count -ge 1) {
        write-host "chat log found" -ForegroundColor Yellow
        foreach ($chat in $ticket.chatTranscriptAttachment) {
            $chatLog = Invoke-RestMethod -Method GET -Uri "$baseURL/2028/file/$($chat.ID)?token=$token"
            Update-TDXTicket -ticketID $tdxTicket.ID -Comment $chatLog -IsPrivate $true -IsRichHtml $true
        }
    }
    # Check if there are attachments
    if ($ticket.Attachment.count -ge 1) {
        foreach ($attachment in $ticket.Attachment) {
            write-host "attachments found" -ForegroundColor Yellow
            Invoke-webrequest -Uri "$baseURL/2028/file/$($attachment.ID)?token=$token" -UseBasicParsing -OutFile "c:\users\wrcrabtree\downloads\$($attachment.id)+$($attachment.name)"   
        }
    }
} 