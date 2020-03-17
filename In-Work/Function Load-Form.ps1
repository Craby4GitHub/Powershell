Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                       = New-Object System.Windows.Forms.Form    
$Form.Size                  = New-Object System.Drawing.Size(300,150)  
$Form.FormBorderStyle       = "FixedDialog"
$Form.StartPosition         = "CenterScreen"
$Form.Text                  = "Enter Computer Name"
$Form.ControlBox            = $false
$Form.TopMost               = $true

$TBComputerName             = New-Object System.Windows.Forms.TextBox
$TBComputerName.Location    = New-Object System.Drawing.Size(25,20)
$TBComputerName.Size        = New-Object System.Drawing.Size(215,50)
 
$GBComputerName             = New-Object System.Windows.Forms.GroupBox
$GBComputerName.Location    = New-Object System.Drawing.Size(30,10)
$GBComputerName.Size        = New-Object System.Drawing.Size(225,50)
$GBComputerName.Text        = "Computer name:"
 
$ButtonOK                   = New-Object System.Windows.Forms.Button
$ButtonOK.Location          = New-Object System.Drawing.Size(195,70)
$ButtonOK.Size              = New-Object System.Drawing.Size(50,20)
$ButtonOK.Text              = "OK"
$Form.AcceptButton          = $ButtonOK

$Form.controls.AddRange(@($TBComputerName,$GBComputerName,$ButtonOK))

$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

Function Set-OSDComputerName {

    $ErrorProvider.Clear()

    if ($TBComputerName.Text.Length -eq 0) {
        $ErrorProvider.SetError($GBComputerName, "Please enter a computer name.")
    }

    elseif ($TBComputerName.Text.Length -gt 15) {
        $ErrorProvider.SetError($GBComputerName, "Computer name cannot be more than 15 characters.")
    }

    #Validation Rule for computer names.
    elseif ($TBComputerName.Text -match "^[-_]|[^a-zA-Z0-9-_]"){
        $ErrorProvider.SetError($GBComputerName, "Computer name invalid, please correct the computer name.")
    }

    else{
        $OSDComputerName = $TBComputerName.Text.ToUpper()
        $TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment
        $TSEnv.Value("OSDComputerName") = "$($OSDComputerName)"
        $Form.Close()
    }
}

$ButtonOK.Add_Click({Set-OSDComputerName})

[void]$Form.ShowDialog()