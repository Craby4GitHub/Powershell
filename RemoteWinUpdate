

$ErrorActionPreference = "Continue"
workflow Add-WindowsUpdateRemote
{
  #   its a parameter. Yup.
  param([string]$computer)

  #   Get list of computers based on filter query and get thier names
  $computerArray = (Get-ADComputer -Filter {Name -like "$computer"}).Name

  #    For each single computer in the computer list... parallel
  foreach -parallel ($singleComputer in $computerArray)
  {


  $OSArch = (Get-WmiObject -Class Win32_Operatingsystem -ComputerName $singlecomputer).OSArchitecture

  if($OSArch -eq "64-bit")
  {
      $updatePath = "C:\Updates\x64"
  }

  else
  {
      $updatePath = "C:\Updates\x86"
  }


  Add-WindowsPackage -NoRestart -Online -PackagePath $updatePath


  }
}

Add-WindowsUpdateRemote -computer (Read-Host "Enter computer name. Wild card * is accepted. EX: EC-E513*C")



