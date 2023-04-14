function Get-PimaCommunityCollegeClasses {
    param(
        [string]$Combined = 'summer2023',
        [ValidatePattern('^([01]\d|2[0-3]):?([0-5]\d)$')]
        [string]$MeetEndTime = '',
        [ValidatePattern('^([01]\d|2[0-3]):?([0-5]\d)$')]
        [string]$MeetStartTime = '',
        [ValidateSet('Hybrid', 'InPerson', 'Online', 'SelfPacedInPerson', 'SelfPacedIndependent', 'Virtual', '')]
        [Alias('HY', 'IP', 'ON', 'SI', 'SD', 'VI')]
        [string]$Method = '',
        [ValidateSet('Desert Vista Campus', 'Downtown Campus', 'East Campus', 'Northwest Campus', 'West Campus', '29th Coalition Center', 'DMAFB', 'Nogales AZ- Santa Cruz County', 'PCC Aviation Technology Center', '')]
        [string]$Site = ''
    )

    $MethodMap = @{
        'Hybrid'             = 'HY'
        'InPerson'           = 'IP'
        'Online'             = 'ON'
        'SelfPacedInPerson'  = 'SI'
        'SelfPacedIndependent'= 'SD'
        'Virtual'            = 'VI'
    }

    $MappedMethod = ''
    if ($Method -ne '') {
        $MappedMethod = $MethodMap[$Method]
    }

    $ApiUrl = "https://do-mobile.pima.edu/getsoc/"
    $QueryString = @{
        'combined'      = $Combined
        'meetEndTime'   = $MeetEndTime
        'meetStartTime' = $MeetStartTime
    }

    if ($MappedMethod -ne '') {
        $QueryString['method'] = $MappedMethod
    }

    if ($Site -ne '') {
        $QueryString['site'] = $Site
    }

    try {
        $Response = Invoke-WebRequest -Uri $ApiUrl -Method Get -ContentType "application/json" -UseBasicParsing -Body $QueryString
        if ($Response.StatusCode -eq 200) {
            return $Response.Content | ConvertFrom-Json
        } else {
            Write-Error "API call failed with status code $($Response.StatusCode)"
        }
    }
    catch {
        Write-Error "Error occurred during API call: $_"
    }
}
