<#
.SYNOPSIS
   Retrieves a list of classes from Pima Community College based on the specified filters.

.DESCRIPTION
   This function makes an API call to Pima Community College's class search to retrieve a list of classes
   based on the provided filters, such as class schedule, meeting times, method, and campus location.

.PARAMETER Combined
   The class schedule, such as 'summer2023'. Default value is 'summer2023'.

.PARAMETER MeetEndTime
   The class meeting end time in 24-hour format (e.g., '2345'). Default value is an empty string.

.PARAMETER MeetStartTime
   The class meeting start time in 24-hour format (e.g., '1700'). Default value is an empty string.

.PARAMETER Method
   The class delivery method. Can be one of the following values: Hybrid, InPerson, Online, SelfPacedInPerson,
   SelfPacedIndependent, or Virtual. Default value is an empty string, which means 'Any'.

.PARAMETER Site
   The campus location of the class. Can be one of the following values: Desert Vista Campus, Downtown Campus,
   East Campus, Northwest Campus, West Campus, 29th Coalition Center, DMAFB, Nogales AZ- Santa Cruz County, or
   PCC Aviation Technology Center. Default value is an empty string, which means 'All Locations'.

.EXAMPLE
   Get-PimaCommunityCollegeClasses -Combined 'fall2023' -Method 'Online' -Site 'East Campus'

   Retrieves a list of online classes at East Campus for the Fall 2023 schedule.

#>
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

    # Map the method parameter to its 2-letter code
    $MethodMap = @{
        'Hybrid'               = 'HY'
        'InPerson'             = 'IP'
        'Online'               = 'ON'
        'SelfPacedInPerson'    = 'SI'
        'SelfPacedIndependent' = 'SD'
        'Virtual'              = 'VI'
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
        # Make an API call with the specified query parameters
        $Response = Invoke-WebRequest -Uri $ApiUrl -Method Get -ContentType "application/json" -UseBasicParsing -Body $QueryString
        if ($Response.StatusCode -eq 200) {
            return $Response.Content | ConvertFrom-Json
        }
        else {
            Write-Error "API call failed with status code $($Response.StatusCode)"
        }
    }
    catch {
        Write-Error "Error occurred during API call: $_"
    }
}