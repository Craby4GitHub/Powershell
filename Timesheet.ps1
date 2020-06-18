# https://www.powershellgallery.com/packages/Selenium/3.0.0

$Credentials = Get-Credential

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object system.Windows.Forms.Form
$Form.FormBorderStyle = "FixedDialog"
$Form.ClientSize = "400,350"
$Form.text = "Timesheet Automation"
$Form.TopMost = $true
$Form.StartPosition = 'CenterScreen'

$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Location = New-Object System.Drawing.Size(7,63)  
$groupBox.text = "Unfilled Timesheets" 
$Form.controls.Add($groupBox)

$Driver = Start-SeChrome -Incognito
Open-SeUrl -Driver $Driver -Url "https://ban8sso.pima.edu/ssomanager/c/SSB?pkg=bwpktais.P_SelectTimeSheetRoll"


#region Login to MyPima
$usernameElement = Find-SeElement -Driver $Driver -Wait -Timeout 10 -Id 'username'
$passwordElement = Find-SeElement -Driver $Driver -Id 'password'

Send-SeKeys -Element $usernameElement -Keys $Credentials.UserName
Send-SeKeys -Element $passwordElement -Keys $Credentials.GetNetworkCredential().Password
Find-SeElement -Driver $Driver -ClassName 'btn-submit' | Invoke-SeClick 
#endregion

# Select pay period
$PayPeriodDropdown = Find-SeElement -Driver $Driver -Id "period_1_id"
$payPeriods = @()
# https://devblogs.microsoft.com/powershell/parsing-text-with-powershell-1-3/
# Status regex needs fixing to match whole status, not just first word
# gets all the pay periods so we can go through as many as we need
$PayPeriodDropdown.Text -split "\n" |
    Select-String -Pattern "(?<startMonth>\w{3}) (?<startDay>\d{2}), (?<startYear>\d{4}) to (?<endMonth>\w{3}) (?<endDay>\d{2}), (?<endYear>\d{4})   (?<status>\w+)" |
        Foreach-Object {
            $startMonth, $startDay, $startYear, $endMonth, $endDay, $endYear, $status = $_.Matches[0].Groups['startMonth', 'startDay','startYear','endMonth','endDay','endYear', 'status'].Value
            $payPeriods += [PSCustomObject] @{
                startDay = $startDay
                startMonth = $startMonth
                startYear = $startYear
                endDay = $endDay
                endMonth = $endMonth
                endYear = $endYear
                status = $status
            }
        }
$Checkboxes = @()
$y = 20
foreach ($payPeriod in $payPeriods) {
    if ($payPeriod.status -eq 'in' -or 'not') {

        $CheckBox                       = New-Object system.Windows.Forms.CheckBox
        $CheckBox.text                  = "$($payPeriod.startDay)-$($payPeriod.startMonth)-$($payPeriod.startYear)"
        $CheckBox.AutoSize              = $false
        $CheckBox.width                 = 95
        $CheckBox.height                = 20
        $CheckBox.location              = New-Object System.Drawing.Point(9,$y)
        $y += 30
        $CheckBox.Font                  = 'Microsoft Sans Serif,10'

        $groupBox.Controls.Add($CheckBox)
        $Checkboxes += $Checkbox
       }
}
$groupBox.size = New-Object System.Drawing.Size(200,(40*$checkboxes.Count)) 
[void]$Form.ShowDialog()

 #Get-SeSelectionOption -Element $PayPeriodDropdown -ByPartialText "$($payPeriod.startMonth) $($payPeriod.startDay), $($payPeriod.startYear) to $($payPeriod.endMonth) $($payPeriod.endDay), $($payPeriod.endYear)"
        

        # Click Time Sheet button
        #Find-SeElement -XPath '/html/body/div[3]/form/table[2]/tbody/tr/td/input' -Driver $Driver | Invoke-SeClick -Driver $Driver

        #$JobsSeqNo = Find-SeElement -Name 'JobsSeqNo' -Driver $Driver | Get-SeElementAttribute -Attribute 'Value'
       <#
        $EarnCode = @(
            ('1HR','Hourly Pay'),
            ('LAN','Annual Leave Taken'),
            ('LSK','Sick Leave Taken'),
            ('OCP','On Call Pay'),
            ('CBP','Call Back Worked'),
            ('LES','Personal Leave charged to Sick'),
            ('LEA','Personal Leave charged to Ann'),
            ('CTT','Comp Time Taken'),
            ('LJD','Jury Duty Pay'),
            ('LBR','Bereavement Leave'),
            ('LHR','Holiday Pay'),
            ('LRC','Recess Pay'),
            ('LUP','Unpaid Leave'),
            ('FMS','FMLA Paid by Sick Leave'),
            ('FMA','FMLA Paid by Annual Leave'),
            ('FMU','FMLA Unpaid Leave'),
            ('WCP','Emergency Treatment Leave'),
            ('WCC','Workers Comp Paid by Sick 1/3'),
            ('WCA','Workers Comp Paid by Ann 1/3'),
            ('WCU','Workers Comp Unpaid Leave 2/3'),
            ('LML','Military Leave Pay'),
            ('LAP','Administrative Leave Paid'),
            ('LAU','Administrative Leave Unpaid'),
            ('LPC','College Paid Closure'),
            ('LMU','Military Leave Unpaid')
        )
        #>
        #Open-SeUrl -url "https://bannerweb.pima.edu/pls/pccp/bwpkteci.P_TimeInOut?JobsSeqNo=$($JobsSeqNo)&LastDate=0&EarnCode=$($EarnCode[0][0])&DateSelected=$($payPeriod.startDay)-$($payPeriod.startMonth)-$($payPeriod.startYear)&LineNumber=5"
    


#$Monday = Find-SeElement -ClassName 'dbheader' -Driver $Driver | Get-SeElementAttribute -Attribute 'Text'
#https://bannerweb.pima.edu/pls/pccp/bwpkteci.P_TimeInOut?JobsSeqNo=315598&LastDate=0&EarnCode=1HR&DateSelected=04-JUL-2020&LineNumber=5


# Timesheet

#Stop-SeDriver -Driver $Driver
