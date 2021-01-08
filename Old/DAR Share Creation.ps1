Clear-Host

# Import Active Directory module
import-module activedirectory

#Requires -Modules activedirectory
$networkPath = "\\WC-STORAGENAS\classes\"
#$networkPath = "\\Wc-storagenas\Aztec_Press\"
$ITSecurityGroup = "WC-IT-DAR_RW"
$darOULocation = 'OU=DAR_Classroom_Users,OU=West Campus,OU=EDU_Groups,DC=edu-domain,DC=pima,DC=edu'

Function Main {
	
    $courseList = @()
    $instructorsList = @()
    $sudentList = @()
    
    #region File to Array
    ForEach ($object in Get-File) {
        $courseName = $object.CRSEDESC + '-' + $object.SFRSTCR_CRN
        $Role = $object.CLASS_ROLE
        $UserName = ($object.EMAIL).Split("@")

        if ($Role -eq 'instructor') {
            $instructorsList += , @($courseName, $UserName[0])
        }   

        if ($Role -eq 'student') {
            $sudentList += , @($courseName, $UserName[0])
        }  

        if ($courseList -notcontains $courseName) {
            $courseList += $courseName
        }
    }
    #endregion

    #region Create each course folder, remove inheritance, create groups and set permissions on folders

    Write-Host 'Doing all the things...'  -ForegroundColor Green
    Write-Host '------------------------------------------'
    ForEach ($course in $courseList) {

        $courseFolder = $networkPath + $course
        $courseFolderDropbox = $networkPath + $course + "\dropbox"
        New-Folder -Course $course

        Remove-Inheritance -folder $courseFolder
        Remove-Inheritance -folder $courseFolderDropbox


    #region Create EDU-Domain groups

        # Query AD to see if group exists already
        $courseToString = [string]$course
        $testGroup = Get-ADGroup -Filter {SamAccountName -eq $courseToString}
    
        if ($testGroup -eq $null) {
               
            Write-Host $course 'security group doesnt exist...creating it now' -Foregroundcolor Yellow

            #$path = 'OU=Aztec_Press_Classes,OU=_Security Groups,OU=IT_Services,OU=West,OU=_EDU,DC=edu-domain,DC=pima,DC=edu'
            $desc = 'This is for students of ' + $courseToString

            New-ADGroup -Name $courseToString -SamAccountName $courseToString -GroupCategory Security -GroupScope Global -DisplayName $courseToString -Path $darOULocation -Description $desc | Out-Null
         
            Write-Host "Pausing script to allow AD to sync new group" -ForegroundColor Yellow
            Start-Sleep -s 60      

        }
        else {
            Write-Host $course 'Group exisits' -Foregroundcolor Cyan
        }
    #endregion

    #region Add security groups to folders
        Set-Permission -Folder $courseFolder -SecurityGroup $courseToString -Permission "ReadAndExecute"
    
        Set-Permission -Folder $courseFolderDropbox -SecurityGroup $courseToString -Permission "Write"

        Set-Permission -Folder $courseFolder -SecurityGroup $ITSecurityGroup -Permission "FullControl"

        Set-Permission -Folder $courseFolderDropbox -SecurityGroup $ITSecurityGroup -Permission "FullControl"

    #endregion

    }

    Write-Host '*********************************************'
    #endregion

    #region Add instructor to the course folders

    Write-Host 'Adding instructors to the course folders...'  -ForegroundColor Green
    Write-Host '------------------------------------------'
    ForEach ($instructorCourse in $instructorsList) {
        
        $courseFolder = $networkPath + $instructorCourse[0]
        $courseFolderDropbox = $networkPath + $instructorCourse[0] + "\dropbox"

        Set-Permission -Folder $courseFolder -SecurityGroup $instructorCourse[1] -Permission "Modify"
        Set-Permission -Folder $courseFolderDropbox -SecurityGroup $instructorCourse[1] -Permission "Modify"
    }
    Write-Host '*********************************************'
    #endregion

    #region Add students to groups

    Write-Host 'Adding students to security groups...' -ForegroundColor Green
    Write-Host '------------------------------------------'
    ForEach ($student in $sudentList) {
        Try {
            Get-AdGroup -Identity $student[0] | Out-Null
        }
        Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            Write-host "Warning, course group $student[0] doesnt exist in AD." -ForegroundColor Red
        }

        $testStudent = Get-ADGroupMember -Identity $student[0] | Where-Object -Property SamAccountName -eq $student[1]
        if ($testStudent -eq $null) {
            Write-Host "Adding $student[1] to $student[0]" -ForegroundColor Yellow
            # Add current user to security group for that course
            Try {
                Add-ADGroupMember -Identity $student[0] -Members $student[1] | Out-Null
            }
            Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                Write-host "Warning, user $student[1] doesnt exist in AD, couldnt be added to $student[0]" -ForegroundColor Red
            }
        }
        else {
            Write-Host $student[1] 'is already in' $student[0] -ForegroundColor Cyan
        }
    }
    #endregion

    #region Add all DAR groups to single group
    
    $AllDARGroups = Get-ADGroup -SearchBase $darOULocation -filter {GroupCategory -eq "Security" -and Name -ne "WC-AllDar"}
    foreach ($Group in $AllDARGroups) {
        Write-Host 'Adding' $Group.name 'to WC-AllDar'
        Add-ADGroupMember 'WC-AllDAR' $Group
    }
    #endregion

}
#region Functions
Function New-Folder {
    param([string]$Course)

    $courseFolder = $networkPath + $Course
    $courseFolderDropbox = $networkPath + $Course + "\dropbox"
     
    Write-Host $course':' -ForegroundColor Gray

    # Create course root folder
    if (Test-Path $courseFolder) {
        Write-Host $courseFolder 'already exists' -ForegroundColor Yellow
    }
    else {
        Write-Host "Creating $courseFolder course folder..." -ForegroundColor Cyan
        New-Item -ItemType Directory -Path $courseFolder -Force | Out-Null
    }

    # Create course dropbox folder
    if (Test-Path $courseFolderDropbox) {
        Write-Host $courseFolderDropbox 'already exists' -ForegroundColor Yellow
    }
    else {
        Write-Host "Creating $courseFolderDropbox folder..." -ForegroundColor Cyan
        New-Item -ItemType Directory -Path $courseFolderDropbox -Force | Out-Null
    }
}

Function Remove-Inheritance {
    param([string]$Folder)
    if (Test-Path $Folder) {
        $acl = Get-ACL -Path $Folder
        $testInheritance = $acl | Select-Object -ExpandProperty access | Where-Object {$_.isinherited -eq $true}
        $acl.SetAccessRuleProtection($True, $False)
       # if($testInheritance.count -gt 0){
            Write-Host "Removing inheritance for $Folder folder..." -ForegroundColor Cyan
            (Get-Item $Folder).SetAccessControl($acl)
      #  }
    }
    else {
        Write-Host $Folder 'does not exist.' -ForegroundColor Red
    }
}

Function Remove-DomainUsers {
    param([string]$Folder, [string]$Domain)

    
    if (Test-Path $Folder) {
        $acl = (Get-Item $Folder).GetAccessControl('Access')
        $removeAccounts = $acl.Access | Where-Object { $_.IsInherited -eq $false -and $_.IdentityReference -eq $Domain + '-domain\Domain Users' }
        try {
            Write-Host 'Removing Domain Users from' $Folder 'folder...' -ForegroundColor Cyan
            $acl.RemoveAccessRuleAll($removeAccounts)
        }
        Catch [System.Management.Automation.MethodInvocationException] {
            Write-Host $domain 'domain user is not in this folder.' -ForegroundColor Red
        }
        (Get-Item $Folder).SetAccessControl($acl)
    }
    else {
        Write-Host $Folder 'does not exist.' -ForegroundColor Red
    }
}

Function Set-Permission {
    param([string]$Folder, [string]$SecurityGroup, [string]$Permission)

    if (Test-Path $Folder) {
    
        $acl = (Get-Item $Folder).GetAccessControl("Access")
        $testforaccount = $acl.Access | Where-Object { $_.IsInherited -eq $false -and $_.IdentityReference -like "*$SecurityGroup*" }
        if ($testforaccount -eq $null) {
            $ar = New-Object System.Security.AccessControl.FileSystemAccessRule($SecurityGroup, $Permission, "ContainerInherit,ObjectInherit", "None", "Allow")
            $acl.SetAccessRule($ar)

            Write-Host 'Adding' $SecurityGroup 'to' $Folder 'with' $Permission 'permissions.' -Foregroundcolor Cyan
            (Get-Item $Folder).SetAccessControl($acl)
        }
    }
    else {
        Write-Host $Folder 'does not exist.' -ForegroundColor Red
    }
}

Function Get-File {
    Do {
        $filePath = Find-File $PSScriptroot

        $correctFile = read-host 'Is' $filePath 'the correct file? (Y/N)'
        
        if ($correctFile -eq 'Y' -and $filePath -ne $null) {
            return $inputFile = Import-Csv $filePath           
        }
        else {
            write-host 'Your selection is empty or does not exist'
        }
    }until($correctFile -eq 'Y' -and $inputFile -ne $null)
   
}

Function Find-File($initialDirectory) {
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    [void]$OpenFileDialog.ShowDialog()
    $OpenFileDialog.FileName
}
#endregion

main