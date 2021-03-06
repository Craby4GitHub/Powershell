$ErrorActionPreference = "Continue"
#Requires -Modules activedirectory
workflow Get-ImageData {
    #   its a parameter. Yup.
    param([string]$computer)

    #   Get the temp path of the user runnin this script
    $TempPath = Join-Path -Path $([System.IO.Path]::GetTempPath()) -ChildPath "ImageData"
    
    #   Create sub folder
    New-Item -Path $TempPath -ItemType Directory 
    
    #   Get list of computers based on filter query and get thier names
    $computerArray = (Get-ADComputer -Filter { Name -like "$computer" }).Name
    
    #    For each single computer in the computer list... parallel
    foreach -parallel ($singleComputer in $computerArray) {
        #   Making a random name for the temp file
        $TempFileName = [System.Guid]::NewGuid().Guid + ".csv"
        
        #   Geting temp file path
        $FullTempFilePath = Join-Path -Path $TempPath -ChildPath $TempFileName
        
        #   Where the meat is, test connection so we don't waste time on computers that are offline.
        if (Test-Connection -ComputerName $singleComputer -Count 1 -ErrorAction SilentlyContinue) {
            #   Get when the computer was imaged, converting it to a readable format
            $imageDate = ([WMI]'').ConvertToDateTime((Get-WmiObject -ComputerName $singleComputer win32_operatingsystem).installdate)
            
            #   Write out to file and screen
            "$singleComputer : $imageDate" | tee-object -filePath $FullTempFilePath #Write out to the random file
        }

        else {
            #   Write out to file and screen
            #   Want better error handling
            "$singleComputer : Unknown" | tee-object -filePath $FullTempFilePath #Write out to the random file
        }
    }

    $TempFiles = Get-ChildItem -Path $TempPath
    
    #   Concatenate all the files
    foreach ($TempFile in $TempFiles) {
        Get-Content -Path $TempFile.FullName | Out-File $PSScriptRoot\ImageDate.csv -Append
    }
    #   Delete the temp files
    Remove-Item -Path $TempPath -Force -Recurse #clean up
}