Clear-Host

# Import Active Directory module
import-module activedirectory

#Requires -Modules activedirectory
$networkPath = "\\WC-STORAGENAS\classes\"
#$networkPath = "\\Wc-storagenas\Aztec_Press\"
$ITSecurityGroup = "WC-IT-DAR_RW"
$darOULocation = 'OU=DAR_Classroom_Users,OU=West Campus,OU=EDU_Groups,DC=edu-domain,DC=pima,DC=edu'

Function Main{
	
	$courseList = @()
	$instructorsList = @()
	$sudentList = @()
    
    #region File to Array
    ForEach($object in Get-File){
        $courseName = $object.CRSEDESC + '-' + $object.SFRSTCR_CRN
        $Role = $object.CLASS_ROLE
        $UserName = ($object.EMAIL).Split("@")

        if($Role -eq 'instructor'){
            $instructorsList += ,@($courseName, $UserName[0])
        }   

        if($Role -eq 'student'){
            $sudentList += ,@($courseName, $UserName[0])
        }  

        if($courseList -notcontains $courseName){
            $courseList += $courseName
        }
    }
    #endregion

    #region Create each course folder and remove inheritance

    Write-Host 'Creating course folders and removing inheritance...'  -ForegroundColor Green
    Write-Host '------------------------------------------'
    ForEach($course in $courseList){
        New-Folder -Course $course
    }
    Write-Host '*********************************************'
    #endregion

    #region Add security group and instructor to the course folders

    Write-Host 'Adding Security group and instructor to the course folders...'  -ForegroundColor Green
    Write-Host '------------------------------------------'
    ForEach($course in $instructorsList){
        
        $courseFolder = $networkPath + $course[0]
        $courseFolderDropbox = $networkPath + $course[0] + "\dropbox"

        Write-Host 'Adding' $ITSecurityGroup 'to'$Course[0] -Foregroundcolor Cyan
        Set-Permission -Folder $courseFolder -SecurityGroup $ITSecurityGroup -Permission "FullControl"

        Write-Host 'Adding' $course[1] 'to' $course[0] -ForegroundColor Cyan
        Set-Permission -Folder $courseFolder -SecurityGroup $course[1] -Permission "Modify"

        Write-Host 'Adding' $course[1] 'to' $courseFolderDropbox  -ForegroundColor Cyan
        Set-Permission -Folder $courseFolderDropbox -SecurityGroup $course[1] -Permission "Modify"
    }
    Write-Host '*********************************************'
    #endregion

    #region Create EDU-Domain groups
    Write-Host 'Making student groups if they dont exist...'  -ForegroundColor Green
    Write-Host '------------------------------------------'
    ForEach($course in $courseList){

        $courseFolder = $networkPath + $course
        $courseFolderDropbox = $networkPath + $course + "\dropbox"
    
        # Query AD to see if group exists already
        $courseToString = [string]$course
        $testGroup = Get-ADGroup -Filter {SamAccountName -eq $courseToString}
    
        if ($testGroup -eq $null){
               
            Write-Host $course 'security group doesnt exist...creating it now' -Foregroundcolor Yellow

            #$path = 'OU=Aztec_Press_Classes,OU=_Security Groups,OU=IT_Services,OU=West,OU=_EDU,DC=edu-domain,DC=pima,DC=edu'
            $desc = 'This is for students of ' + $courseToString

            New-ADGroup -Name $courseToString -SamAccountName $courseToString -GroupCategory Security -GroupScope Global -DisplayName $courseToString -Path $darOULocation -Description $desc | Out-Null
         
            Write-Host "Pausing script to allow AD to sync new group" -ForegroundColor Yellow
            Start-Sleep -s 90      

        }else{
        Write-Host $course 'Group exisits' -Foregroundcolor Cyan
        }

        Write-Host "Adding $courseToString Read permissions on $courseFolder" -Foregroundcolor Yellow
        Set-Permission -Folder $courseFolder -SecurityGroup $courseToString -Permission "ReadAndExecute"
         
        Write-Host "Adding $courseToString Write permissions on $courseFolderDropbox" -Foregroundcolor Yellow
        Set-Permission -Folder $courseFolderDropbox -SecurityGroup $courseToString -Permission "Write"

    }

    Write-Host '*********************************************'
    #endregion

    #region Add students to groups

    Write-Host 'Adding students to security groups...' -ForegroundColor Green
    Write-Host '------------------------------------------'
    ForEach($student in $sudentList){
        Try{
            Get-AdGroup -Identity $student[0] | Out-Null
        }
        Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]{
            Write-host 'Warning, course group ' $student[0] 'doesnt exist in AD.' -ForegroundColor Red
        }

        $testStudent = Get-ADGroupMember -Identity $student[0] | Where-Object -Property SamAccountName -eq $student[1]
        if($testStudent -eq $null){
            Write-Host 'Adding' $student[1] 'to' $student[0] -ForegroundColor Yellow
            # Add current user to security group for that course
            Try{
                Add-ADGroupMember -Identity $student[0] -Members $student[1] | Out-Null
            }
            Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]{
                Write-host 'Warning, user' $student[1] 'doesnt exist in AD, couldnt be added to' $student[0] -ForegroundColor Red
            }
       }else{
            Write-Host $student[1] 'is already in' $student[0] -ForegroundColor Cyan
       }
    }
    #endregion

    #region Add all DAR groups to single group
    $AllDARGroups = Get-ADGroup -SearchBase $darOULocation -filter {GroupCategory -eq "Security" -and Name -ne "WC-AllDar"}
    foreach($Group in $AllDARGroups){
        Write-Host 'Adding' $Group.name 'to WC-AllDar'
        Add-ADGroupMember 'WC-AllDAR' $Group
    }
    #endregion
}
#region Functions
Function New-Folder{
    param([string]$Course)

    $courseFolder = $networkPath + $Course
    $courseFolderDropbox = $networkPath + $Course + "\dropbox"
     
    Write-Host $course':' -ForegroundColor Gray

    # Create course root folder
    Write-Host 'Checking folders...' -ForegroundColor Cyan
    if(Test-Path $courseFolder){
        Write-Host 'Root folder already exists' -ForegroundColor Yellow
    }else{
        Write-Host 'Creating root course folder...' -ForegroundColor Cyan
        New-Item -ItemType Directory -Path $courseFolder -Force | Out-Null
    }

    # Create course dropbox folder
    if(Test-Path $courseFolderDropbox){
        Write-Host 'Dropbox folder already exists' -ForegroundColor Yellow
    }else{
        Write-Host 'Creating dropbox folder...' -ForegroundColor Cyan
        New-Item -ItemType Directory -Path $courseFolderDropbox -Force | Out-Null
    }
    
    # Remove inheritance on newly created course folder
    Remove-Inheritance -folder $courseFolder
    Remove-Inheritance -folder $courseFolderDropbox 

    Remove-DomainUsers -folder $courseFolder -Domain 'EDU'
    Remove-DomainUsers -folder $courseFolder -Domain 'PCC'

}

Function Remove-Inheritance {
    param([string]$folder)

    Write-Host 'Removing inheritance for' $folder 'folder...' -ForegroundColor Cyan
    $acl = Get-ACL -Path $folder
    $acl.SetAccessRuleProtection($True, $True)
    (Get-Item $Folder).SetAccessControl($acl)
}

Function Remove-DomainUsers {
    param([string]$Folder, [string]$Domain)

    Write-Host 'Removing Domain Users from' $Folder 'folder...' -ForegroundColor Cyan
    if(Test-Path $Folder){
        $acl = (Get-Item $Folder).GetAccessControl('Access')
        $removeAccounts = $acl.Access | Where-Object{ $_.IsInherited -eq $false -and $_.IdentityReference -eq $Domain + '-domain\Domain Users' }
        $acl.RemoveAccessRuleAll($removeAccounts)
        (Get-Item $Folder).SetAccessControl($acl)
    }else{
        Write-Host 'Folder does not exist' -ForegroundColor Yellow
    }
}

Function Set-Permission{
    param([string]$Folder, [string]$SecurityGroup,[string]$Permission)

    $acl = (Get-Item $Folder).GetAccessControl("Access")
    $ar = New-Object System.Security.AccessControl.FileSystemAccessRule($SecurityGroup,$Permission,"ContainerInherit,ObjectInherit","None","Allow")
    $acl.SetAccessRule($ar)
    (Get-Item $Folder).SetAccessControl($acl)
}

Function Get-File{
    Do{
        $filePath = Find-File $PSScriptroot

        $correctFile = read-host 'Is' $filePath "the correct file? (Y/N)"
        
        if($correctFile -eq 'Y' -and $filePath -ne $null){
            return $inputFile = Import-Csv $filePath           
        }else{
            write-host "Your selection is empty or does not exist"
        }
    }until($correctFile -eq 'Y' -and $inputFile -ne $null)
   
}

Function Find-File($initialDirectory){
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    [void]$OpenFileDialog.ShowDialog()
    $OpenFileDialog.FileName
}
#endregion

main