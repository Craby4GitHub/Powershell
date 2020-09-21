# For SCCM Client Version 5.00.8913.1032
# 
# & \\$strSimpleDomainName\SYSVOL\$strFullDomainName\scripts\SCCM_1910\ccmsetup.exe /USEPKICERT /NOCRLCHECK SMSMP=do-sccm.pcc-domain.pima.edu CCMHOSTNAME=sccm.pima.edu SMSSITECODE=PCC 
# & \\$strSimpleDomainName\SYSVOL\$strFullDomainName\scripts\SCCM_1910\ccmsetup.exe
#
# IMPORTANT:  If the Config Mgr Client Certificate Template is modified Windows Certificate Services, then $strPCCOID and/or $strEDUOID will need to be updated with the new OID string found in each template.
#

$strPCCOID = "1.3.6.1.4.1.311.21.8.1551373.1795746.8867729.6336595.5940394.173.16347132.1013110"
$strEDUOID = "1.3.6.1.4.1.311.21.8.16319599.11845440.1393050.13379328.9083809.115.15869700.12764145"


$strWindowsPath = Get-Content env:windir
$strLogPath = "$strWindowsPath\Temp\_SCCM_Agent_Deployment.log"
$strFullDomainName = Get-WMIObject Win32_ComputerSystem | Select -Expand Domain
$strSimpleDomainName = [Regex]::Matches($strFullDomainName,".+(?=\.pima\.edu$)") | Select -Expand Value
$strOSArchitecture = Get-WMIObject Win32_OperatingSystem | Select -Expand OSArchitecture

If ($strOSArchitecture -eq "64-bit")

    {

     $strOSA = "x64"

    }

Else

    {

     $strOSA = "i386"

    }

$boolHasWindowsTemp = (Test-Path C:\Windows\Temp) -eq $True
$boolServiceFound = (Get-Service CCMExec -EA SilentlyContinue) -ne $Null
$boolNewestAgentInstalled = (Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\ | Where {$_.PSChildName -like "{6343A6B8-D881-4B6C-AC85-2384F0B839BD}" -or $_.PSChildName -like "{61938454-8C30-47D3-80A1-51B67BB8A4EC}"}) -ne $null
$boolCertInstalled = (Get-ChildItem Cert:\LocalMachine\My | ForEach {$_ | Select @{N="Template";E={($_.Extensions | Where {$_.OID.FriendlyName -match "Certificate Template information"}).Format(0) -replace "(.+)?=(.+)\((.+)?", '$2'}} | Where {$_.Template -like "*$strPCCOID*" -or $_.Template -like "*$strEDUOID*" -or $_.Template -eq "Config Mgr Client Certificate"}}) -ne $Null
$boolHasCertFlag = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\CCMSetup -Name LastSuccessfulInstallParams | Where {$_.LastSuccessfulInstallParams -like "*PKI*"}) -ne $Null

If ($boolHasWindowsTemp -eq $False)

    {

     New-Item -Path $strWindowsPath -Name Temp -ItemType Directory

    }

Add-Content $strLogPath "::::::::::::::::::::::::::::::: START :::::::::::::::::::::::::::::::::::::::::"

$strDate = Get-Date
Add-Content $strLogPath "$strDate :::: Windows temporary directory found ::::: $boolHasWindowsTemp"
Add-Content $strLogPath "$strDate :::: CCMExec service found ::::::::::::::::: $boolServiceFound"
Add-Content $strLogPath "$strDate :::: Newest agent installed :::::::::::::::: $boolNewestAgentInstalled"
Add-Content $strLogPath "$strDate :::: Machine certificate found ::::::::::::: $boolCertInstalled"
Add-Content $strLogPath "$strDate :::: Previous install used cert flag ::::::: $boolHasCertFlag"

If ($boolServiceFound -eq $True)

    {
     
     If ($boolNewestAgentInstalled -eq $True)

        {

         If ($boolCertInstalled -eq $True)

            {

             If ($boolHasCertFlag -eq $False)

                {

                 $strDate = Get-Date
                 Add-Content $strLogPath "$strDate :::: Attempting to re-install using PKI flags"
                 Add-Content $strLogPath "::::::::::::::::::::::::::::::: END :::::::::::::::::::::::::::::::::::::::::::"

                 & \\$strSimpleDomainName\SYSVOL\$strFullDomainName\scripts\SCCM_1910\ccmsetup.exe /USEPKICERT /NOCRLCHECK SMSMP=do-sccm.pcc-domain.pima.edu CCMHOSTNAME=sccm.pima.edu SMSSITECODE=PCC

                }

             Else

                {
 
                 $strDate = Get-Date
                 Add-Content $strLogPath "$strDate :::: Doing nothing"
                 Add-Content $strLogPath "::::::::::::::::::::::::::::::: END :::::::::::::::::::::::::::::::::::::::::::"

                }

            }

         Else

            {

             $strDate = Get-Date
             Add-Content $strLogPath "$strDate :::: Doing nothing"
             Add-Content $strLogPath "::::::::::::::::::::::::::::::: END :::::::::::::::::::::::::::::::::::::::::::"

            }


        }

     Else

        {

         If ($boolCertInstalled -eq $True)

            {

             $strDate = Get-Date
             Add-Content $strLogPath "$strDate :::: Installing with PKI flags"
             Add-Content $strLogPath "::::::::::::::::::::::::::::::: END :::::::::::::::::::::::::::::::::::::::::::"

             & \\$strSimpleDomainName\SYSVOL\$strFullDomainName\scripts\SCCM_1910\ccmsetup.exe /USEPKICERT /NOCRLCHECK SMSMP=do-sccm.pcc-domain.pima.edu CCMHOSTNAME=sccm.pima.edu SMSSITECODE=PCC

            }

         Else

            {

             $strDate = Get-Date
             Add-Content $strLogPath "$strDate :::: Installing without PKI flags"
             Add-Content $strLogPath "::::::::::::::::::::::::::::::: END :::::::::::::::::::::::::::::::::::::::::::"

             & \\$strSimpleDomainName\SYSVOL\$strFullDomainName\scripts\SCCM_1910\ccmsetup.exe SMSMP=do-sccm.pcc-domain.pima.edu SMSSITECODE=PCC

            }

        }

    }

Else

    {

     If ($boolCertInstalled -eq $True)

        {

         $strDate = Get-Date
         Add-Content $strLogPath "$strDate :::: Installing with PKI Flags"
         Add-Content $strLogPath "::::::::::::::::::::::::::::::: END :::::::::::::::::::::::::::::::::::::::::::"

         & \\$strSimpleDomainName\SYSVOL\$strFullDomainName\scripts\SCCM_1910\ccmsetup.exe /USEPKICERT /NOCRLCHECK SMSMP=do-sccm.pcc-domain.pima.edu CCMHOSTNAME=sccm.pima.edu SMSSITECODE=PCC
         
        }

     Else

        {

         $strDate = Get-Date
         Add-Content $strLogPath "$strDate :::: Installing without PKI Flags"
         Add-Content $strLogPath "::::::::::::::::::::::::::::::: END :::::::::::::::::::::::::::::::::::::::::::"

         & \\$strSimpleDomainName\SYSVOL\$strFullDomainName\scripts\SCCM_1910\ccmsetup.exe SMSMP=do-sccm.pcc-domain.pima.edu SMSSITECODE=PCC

        }

    }