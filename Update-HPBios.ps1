<#  Creator @gwblok - GARYTOWN.COM
    Used to download BIOS Updates from HP, then Extract the bin file.
    It then checks and suspends bitlocker and runs the upgrade.  It does NOT reboot the machine, you can modify the command line to do that, or have your deployment method call the reboot.
    The download and extract areas on in the TEMP folder, as well as the log.
    
    REQUIREMENTS:  HP Client Management Script Library
    Download / Installer: https://ftp.hp.com/pub/caps-softpaq/cmit/hp-cmsl.html  - This will download version 1.1.1. and install if needed
    Docs: https://developers.hp.com/hp-client-management/doc/client-management-script-library-0
    This Script was created using version 1.1.1

    Updates: 2019.03.14
        Replaced [Decimal] with [version].  Hopefully will fix issues caused by machines that had more than one decimal point in version.
        Orginally had [Decimal] in code to remove leading "0" on BIOS version reported by HP. Example: Local Version said 1.45, from HP site as 01.45
        Modified Bitlocker Detection.  Technically, probably don't even need the suspend bitlocker code, as HP's upgrade util is supposed to do it.
        Added logic for both HPBIOSUPDREC64.exe & HPFirmwareUpdRec64.exe updaters
#>


# Relaunch as an elevated process:
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}
  
$OS = "Win10"
$Category = "bios"
$HPContent = $ENV:TEMP
  
  
$DownloadDir = "$($HPContent)\Downloads"
if (-not (Test-Path $DownloadDir)) { New-Item $DownloadDir -ItemType Directory }
$ExtractedDir = "$($HPContent)\Extracted"
if (-not (Test-Path $ExtractedDir)) { New-Item $ExtractedDir -ItemType Directory }
$ProductCode = (Get-WmiObject -Class Win32_BaseBoard).Product
$Model = (Get-WmiObject -Class Win32_ComputerSystem).Model
  
$PoshURL = "https://hpia.hpcloud.hp.com/downloads/cmsl/hp-cmsl-1.6.5.exe"
  
try {
    Get-HPBiosVersion
    Write-Output "HP Module Installed"
}
  
catch {
    Write-Output "HP Module Not Loaded, Loading.... Now"
    Invoke-WebRequest -Uri $PoshURL -OutFile "$($DownloadDir)\HPCM.exe"
    Start-Process -FilePath "$($DownloadDir)\HPCM.exe" -ArgumentList "/verysilent" -Wait
    Write-Output "Finished Downloading and Installing HP Module"
}
  

$CurrentBIOS = Get-HPBiosVersion
Write-Output "Current Installed BIOS Version: $($CurrentBIOS)"
Write-Output "Checking Product Code $($ProductCode) for BIOS Updates"
$BIOS = Get-SoftpaqList -platform $ProductCode -os $OS -category $Category
$MostRecent = ($Bios | Measure-Object -Property "ReleaseDate" -Maximum).Maximum
$BIOS = $BIOS | Where-Object "ReleaseDate" -eq "$MostRecent"
  
if ([version]$CurrentBIOS -ne [version]$Bios.Version) {
    Write-Output "Updated BIOS available, Version: $([version]$BIOS.Version)"
    $DownloadPath = "$($DownloadDir)\$($Model)\$($BIOS.Version)"
    if (-not (Test-Path $DownloadPath)) { New-Item $DownloadPath -ItemType Directory }
    $ExtractedPath = "$($ExtractedDir)\$($Model)\$($BIOS.Version)"
    if (-not (Test-Path $ExtractedPath)) { New-Item $ExtractedPath -ItemType Directory }
  
    Write-Output "Downloading BIOS Update for: $($Model) aka $($ProductCode)"
    Get-Softpaq -number $BIOS.ID -saveAs "$($DownloadPath)\$($BIOS.id).exe" -Verbose
  
    Write-Output "Creating Readme file with BIOS Info HERE: $($DownloadPath)\$($Bios.ReleaseDate).txt"
    $BIOS | Out-File -FilePath "$($DownloadPath)\$($Bios.ReleaseDate).txt"
    $BiosFileName = Get-ChildItem -Path "$($DownloadPath)\*.exe" | Select-Object -ExpandProperty "Name"
    
    Write-Output "Extracting Downloaded BIOS File to: $($ExtractedPath)"
    $argList = @{
        FilePath     = "$($DownloadPath)\$($BiosFileName)"
        ArgumentList = @(
            '-e'
            '-s'
            "-f `"$ExtractedPath`""
        )
    }
    Start-Process @argList

    if ((Get-BitLockerVolume -MountPoint c:).VolumeStatus -eq "FullyDecrypted") {
        Write-Output "Bitlocker Not Present"
    }
    Else {
        Write-Output "Suspending Bitlocker"
        Suspend-BitLocker -MountPoint "C:" -RebootCount 1
    }
  
  
    $argList = @{
        FilePath     = ""
        ArgumentList = @(
            '-r'
            '-s'
            '-b'
        )
    }
  
    if (Get-HPBIOSSetupPasswordIsSet) {
        $argList.ArgumentList += "-p:$PSScriptRoot\password.bin"
    }
    else {
        #Set-HPBiosSetupPassword -NewPassword 
        Write-Host "Bios Password not set"
    }
    if (Test-Path "$($ExtractedPath)\HPBIOSUPDREC64.exe") {
        $argList.FilePath = "$($ExtractedPath)\HPBIOSUPDREC64.exe"
        Write-Output "Using HPBIOSUpdRec64.exe to Flash BIOS"
        Start-Process @argList
    }
    if (Test-Path "$($ExtractedPath)\HPFirmwareUpdRec64.exe") {
        $argList.FilePath = "$($ExtractedPath)\HPFirmwareUpdRec64.exe"
        Write-Output "Using HPFirmwareUpdRec64.exe to Flash BIOS"
        #Start-Process "$($ExtractedPath)\HPFirmwareUpdRec64.exe" -ArgumentList $args -wait
    }
    Write-Output "HP BIOS update Applied, Will Install after next reboot"
    Start-Sleep -Seconds 10
}
ELSE {
    Write-Output 'BIOS already Current'
    Start-Sleep -Seconds 10
}