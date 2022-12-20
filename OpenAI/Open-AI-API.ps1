# Setting Log name
$logName = ($MyInvocation.MyCommand.Name -split '\.')[0] + ' log'
$logFile = "$PSScriptroot\$logName.csv"
. ((Get-Item $PSScriptRoot).Parent.FullName + '\Callable\Write-Log.ps1')




function New-OpenAI-Headers() {    
    $authToken = "token"
    $orgID = "org-wOZxqpHMARUCkAMGe1niIwt3"
    $headers = @{
        Authorization         = "Bearer " + $authToken
        'OpenAI-Organization' = $orgID
    } 
    return $headers
}

function Get-OpenAIModels() {
    # https://beta.openai.com/docs/models/overview
    $response = try {
        Invoke-RestMethod -Method GET -Headers $headers -Uri "https://api.openai.com/v1/models" -ContentType "application/json" -ErrorVariable apiError
    }
    catch {
        Write-Log -level ERROR -message "API authentication failed: $apiError."
    }
    return $response.data
}

function Invoke-OpenAI-Completion($model, $prompt, $responseLength) {
    $body = @{
        model       = $model
        prompt      = $prompt
        temperature = 0
        max_tokens  = $responseLength
    } | ConvertTo-Json

    $response = try {
        Invoke-RestMethod -Method POST -Headers $headers -Uri "https://api.openai.com/v1/completions" -body $body -ContentType "application/json" -ErrorVariable apiError
    }
    catch {
        Write-Log -level ERROR -message "API authentication failed: $apiError."
    }
    return $response
}

function Invoke-OpenAI-Edit($model, $prompt, $instruction) {

    switch ($model) {
        code { $model = 'text-davinci-edit-001' }
        text { $model = 'text-davinci-edit-001' }
        Default { $model = 'text-davinci-edit-001' }
    }
    $body = @{
        model       = $model
        input       = $prompt
        temperature = 0
        instruction = $instruction
    } | ConvertTo-Json

    $response = try {
        Invoke-RestMethod -Method POST -Headers $headers -Uri "https://api.openai.com/v1/edits" -body $body -ContentType "application/json" -ErrorVariable apiError
    }
    catch {
        Write-Log -level ERROR -message "API authentication failed: $apiError."
    }
    return $response
}

# Headers used for all calls
$headers = New-OpenAI-Headers


# Example
#Invoke-OpenAI-Completion -model 'text-davinci-003' -prompt "I get no sound from my speakers" -responseLength 1000