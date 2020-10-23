# https://www.powershellgallery.com/packages/Selenium/3.0.0

$Credentials = Get-Credential

#Install-Module -Name Selenium -RequiredVersion 3.0.0

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object system.Windows.Forms.Form
$Form.FormBorderStyle = "FixedDialog"
$Form.ClientSize = "400,350"
$Form.text = "Timesheet Automation"
$Form.TopMost = $true
$Form.StartPosition = 'CenterScreen'

$Submit_Button = New-Object system.Windows.Forms.Button
$Submit_Button.text = "Submit"
$Submit_Button.width = 60
$Submit_Button.height = 30
$Submit_Button.location = New-Object System.Drawing.Point(150, 300)

$TimeSheets = New-Object system.Windows.Forms.ListView
$TimeSheets.text = "listView"
$TimeSheets.width = 360
$TimeSheets.height = 70
$TimeSheets.location = New-Object System.Drawing.Point(5, 10)
$TimeSheets.CheckBoxes = $true

$Week = New-Object system.Windows.Forms.ListView
$Week.text = "listView"
$Week.width = 360
$Week.height = 70
$Week.location = New-Object System.Drawing.Point(5, 100)
$Week.CheckBoxes = $true
$Week.Items.AddRange(@('Saturday', 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'))

$TimeIn_TextBox = New-Object system.Windows.Forms.TextBox
$TimeIn_TextBox.multiline = $false
$TimeIn_TextBox.width = 100
$TimeIn_TextBox.height = 20
$TimeIn_TextBox.location = New-Object System.Drawing.Point(20, 257)
$TimeIn_TextBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$TimeIn_Label = New-Object system.Windows.Forms.Label
$TimeIn_Label.text = "Time In"
$TimeIn_Label.AutoSize = $true
$TimeIn_Label.width = 25
$TimeIn_Label.height = 10
$TimeIn_Label.location = New-Object System.Drawing.Point($($TimeIn_TextBox.Location.X + ($TimeIn_TextBox.Width - $TimeIn_Label.Width) * .25), $($TimeIn_TextBox.Location.Y - 20))
$TimeIn_Label.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$Lunch_TextBox = New-Object system.Windows.Forms.TextBox
$Lunch_TextBox.multiline = $false
$Lunch_TextBox.width = 100
$Lunch_TextBox.height = 20
$Lunch_TextBox.location = New-Object System.Drawing.Point(120, 257)
$Lunch_TextBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$Lunch_Label = New-Object system.Windows.Forms.Label
$Lunch_Label.text = "Lunch Start"
$Lunch_Label.AutoSize = $true
$Lunch_Label.width = 25
$Lunch_Label.height = 10
$Lunch_Label.location = New-Object System.Drawing.Point($($Lunch_TextBox.Location.X + ($Lunch_TextBox.Width - $Lunch_Label.Width) * .25), $($Lunch_TextBox.Location.Y - 20))
$Lunch_Label.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$LunchTimeAmount_CheckBox = New-Object system.Windows.Forms.CheckBox
$LunchTimeAmount_CheckBox.text = "Half Hour Lunch"
$LunchTimeAmount_CheckBox.AutoSize = $false
$LunchTimeAmount_CheckBox.width = 95
$LunchTimeAmount_CheckBox.height = 20
$LunchTimeAmount_CheckBox.location = New-Object System.Drawing.Point(100, 300)
$LunchTimeAmount_CheckBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$TimeOut_TextBox = New-Object system.Windows.Forms.TextBox
$TimeOut_TextBox.multiline = $false
$TimeOut_TextBox.width = 100
$TimeOut_TextBox.height = 20
$TimeOut_TextBox.location = New-Object System.Drawing.Point(237, 257)
$TimeOut_TextBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$TimeOut_Label = New-Object system.Windows.Forms.Label
$TimeOut_Label.text = "Time Out"
$TimeOut_Label.AutoSize = $true
$TimeOut_Label.width = 25
$TimeOut_Label.height = 10
$TimeOut_Label.location = New-Object System.Drawing.Point($($TimeOut_TextBox.Location.X + ($TimeOut_Label.Width - $ComputerName_Campus_Label.Width) * .25), $($TimeOut_TextBox.Location.Y - 20))
$TimeOut_Label.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$Form.controls.AddRange(@($Submit_Button, $TimeSheets, $Week,$TimeIn_Label,$TimeIn_TextBox,$Lunch_Label,$Lunch_TextBox,$TimeOut_Label,$TimeOut_TextBox))
$Driver = Start-SeFirefox -PrivateBrowsing
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
# gets all the pay periods so we can go through as many as we need
$PayPeriodDropdown.Text -split "\n" |
Select-String -Pattern "(?<startMonth>\w{3}) (?<startDay>\d{2}), (?<startYear>\d{4}) to (?<endMonth>\w{3}) (?<endDay>\d{2}), (?<endYear>\d{4}) (?<status>(\w+\s{1}\w+)|(\w+))" |
ForEach-Object {
    $startMonth, $startDay, $startYear, $endMonth, $endDay, $endYear, $status = $_.Matches[0].Groups['startMonth', 'startDay', 'startYear', 'endMonth', 'endDay', 'endYear', 'status'].Value
    $payPeriods += [PSCustomObject] @{
        startDate = [DateTime]::ParseExact("$($startDay)$($startMonth)$($startYear)", 'ddMMMyyyy', $null)
        endDate   = [DateTime]::ParseExact("$($endDay)$($endMonth)$($endYear)", 'ddMMMyyyy', $null)
        status    = $status
    }
}

# Adding checkboxes for each pay period not started or in progress
$i = 0
foreach ($payPeriod in $payPeriods) {
    if (($payPeriod.status -eq 'Not Started') -or ($payPeriod.status -eq 'In Progress')) {
        [void]$TimeSheets.Items.Add($($([cultureinfo]::InvariantCulture.DateTimeFormat.GetAbbreviatedMonthName($payPeriod.startdate.month)) + ' ' + $($payPeriod.startDate.tostring('dd')) + ', ' + $($payPeriod.startDate.Year)))
        $TimeSheets.items[$i].Tag = $payPeriod
    }
    $i++
}



$EarnCodes = @(
    ('1HR', 'Hourly Pay'),
    ('LAN', 'Annual Leave Taken'),
    ('LSK', 'Sick Leave Taken'),
    ('OCP', 'On Call Pay'),
    ('CBP', 'Call Back Worked'),
    ('LES', 'Personal Leave charged to Sick'),
    ('LEA', 'Personal Leave charged to Ann'),
    ('CTT', 'Comp Time Taken'),
    ('LJD', 'Jury Duty Pay'),
    ('LBR', 'Bereavement Leave'),
    ('LHR', 'Holiday Pay'),
    ('LRC', 'Recess Pay'),
    ('LUP', 'Unpaid Leave'),
    ('FMS', 'FMLA Paid by Sick Leave'),
    ('FMA', 'FMLA Paid by Annual Leave'),
    ('FMU', 'FMLA Unpaid Leave'),
    ('WCP', 'Emergency Treatment Leave'),
    ('WCC', 'Workers Comp Paid by Sick 1/3'),
    ('WCA', 'Workers Comp Paid by Ann 1/3'),
    ('WCU', 'Workers Comp Unpaid Leave 2/3'),
    ('LML', 'Military Leave Pay'),
    ('LAP', 'Administrative Leave Paid'),
    ('LAU', 'Administrative Leave Unpaid'),
    ('LPC', 'College Paid Closure'),
    ('LMU', 'Military Leave Unpaid')
)

$Submit_Button.Add_Click( {
        foreach ($x in $TimeSheets.CheckedItems) {
            Get-SeSelectionOption -Element $PayPeriodDropdown -ByPartialText "$($x.Text)"
            
            # Click Time Sheet button
            Find-SeElement -XPath '/html/body/div[3]/form/table[2]/tbody/tr/td/input' -Driver $Driver | Invoke-SeClick -Driver $Driver

            $JobsSeqNo = Find-SeElement -Name 'JobsSeqNo' -Driver $Driver | Get-SeElementAttribute -Attribute 'Value'

            $TimeSheetTable = Get-SeElement -Driver $Driver -XPath '/html/body/div[3]/table[1]/tbody/tr[5]/td/form/table[1]'

            $TimeSheetTableRows = $TimeSheetTable.FindElementsByTagName('tr')

            $EnteredHours = @()

            #Ignore Total Hours/Unit Rows
            foreach($Row in $TimeSheetTableRows[0..($TimeSheetTableRows.Count-3)]){
                $dataSeperated = $Row.text -split '\n'
                if ($dataSeperated[0] -ne 'Earning Shift Default') {
                    for ($i = 0; $i -lt $EarnCodes.Count; $i++) {
                        if ($EarnCodes[$i][1] -match $dataSeperated[0]) {
                            $EarnCode = $EarnCodes[$i][1]
                        }
                    }
                }else {
                    
                }
                $EarnCode = $null

                
                foreach($day in $dataSeperated[-7..-1]){
                    if ($day -notmatch 'Enter Hours') {
                        
                    }
                }
                $EnteredHours += [PSCustomObject] @{
                    Earning = $EarnCode
                    days   = $null

                }
            }

            for ($i = 0; $x.tag.startDate.AddDays($i) -lt $x.tag.endDate.AddDays(1); $i++) {
                foreach ($day in $Week.Items) {
                    if ($x.tag.startDate.addDays($i).dayofweek -eq $day.text) {
                        $day.Tag += @($x.tag.startDate.addDays($i))
                    }
                }
            }

            foreach ($checkedDay in $Week.CheckedItems.Tag) {
                Open-SeUrl -Target $Driver -url "https://bannerweb.pima.edu/pls/pccp/bwpkteci.P_TimeInOut?JobsSeqNo=$($JobsSeqNo)&LastDate=0&EarnCode=$($EarnCodes[0][0])&DateSelected=$($checkedDay.toString("dd-MMM-yyyy"))&LineNumber=5"
                

                $TimeIn = Find-SeElement -Id 'timein_input_id' -Driver $Driver
                $TimeOut = Find-SeElement -Id 'timeout_input_id' -Driver $Driver
                $AmPmDropDownIn = Find-SeElement -Name 'TimeInAm' -Driver $Driver
                $AmPmDropDownOut = Find-SeElement -Name 'TimeOutAm' -Driver $Driver

                [DateTime]$SubmitTimeIn = $TimeIn_TextBox.Text
                Send-SeKeys -Element $TimeIn[0] -Keys $SubmitTimeIn.ToString('hh:mm')
                Get-SeSelectionOption -Element $AmPmDropDownIn[0] -ByValue $SubmitTimeIn.ToString('tt')

                if ($null -ne $Lunch_TextBox.Text) {
                    [DateTime]$SubmitLunchTime = $Lunch_TextBox.Text
                    Send-SeKeys -Element $TimeOut[0] -Keys $SubmitLunchTime.ToString('hh:mm')
                    Get-SeSelectionOption -Element $AmPmDropDownOut[0] -ByValue $SubmitLunchTime.ToString('tt')

                    if (!$LunchTimeAmount_CheckBox.Checked) {
                        $SubmitLunchTime = $SubmitLunchTime.AddMinutes(60)
                    }
                    else {
                        $SubmitLunchTime = $SubmitLunchTime.AddMinutes(30)
                    }

                    Send-SeKeys -Element $TimeIn[1] -Keys $SubmitLunchTime.ToString('hh:mm')
                    Get-SeSelectionOption -Element $AmPmDropDownIn[1] -ByValue $SubmitLunchTime.ToString('tt')

                    [DateTime]$SubmitTimeOut = $TimeOut_TextBox.Text
                    Send-SeKeys -Element $TimeOut[1] -Keys $SubmitTimeOut.ToString('hh:mm')
                    Get-SeSelectionOption -Element $AmPmDropDownOut[1] -ByValue $SubmitTimeOut.ToString('tt')
                }
                else {
                    [DateTime]$SubmitTimeOut = $TimeOut_TextBox.Text
                    Send-SeKeys -Element $TimeOut[0] -Keys $SubmitTimeOut.ToString('hh:mm')
                    Get-SeSelectionOption -Element $AmPmDropDownOut[0] -ByValue $SubmitTimeOut.ToString('tt')
                }

                # Click Save button
                Find-SeElement -XPath '/html/body/div[3]/form/table[3]/tbody/tr[2]/td/input[2]' -Driver $Driver | Invoke-SeClick -Driver $Driver
            
            }
            # Click Time Sheet button
            Find-SeElement -XPath '/html/body/div[3]/form/table[3]/tbody/tr[1]/td/input[1]' -Driver $Driver | Invoke-SeClick -Driver $Driver
        }
    })


[void]$Form.ShowDialog()

#Stop-SeDriver -Driver $Driver