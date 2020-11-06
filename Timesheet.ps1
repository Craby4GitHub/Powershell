# https://www.powershellgallery.com/packages/Selenium/3.0.0

$Credentials = Get-Credential

#Install-Module -Name Selenium -RequiredVersion 3.0.0

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

$Form = New-Object system.Windows.Forms.Form
$Form.FormBorderStyle = "FixedDialog"
$Form.ClientSize = "450,350"
$Form.text = "Timesheet Automation"
$Form.TopMost = $true
$Form.StartPosition = 'CenterScreen'

$panel = New-Object System.Windows.Forms.TableLayoutPanel
$panel.Dock = "Fill"
$panel.ColumnCount = 2
$panel.RowCount = 3
$panel.CellBorderStyle = "none"
[void]$panel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 60)))
[void]$panel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25)))

$TimeSheets = New-Object system.Windows.Forms.ListView
$TimeSheets.text = "listView"
$TimeSheets.width = 360 
$TimeSheets.height = 70
$TimeSheets.Location = New-Object System.Drawing.Point($($Form.Width * .05), 10)
$TimeSheets.CheckBoxes = $true
$TimeSheets.Dock = 'Fill'

$WeekList = New-Object system.Windows.Forms.ListView
$WeekList.text = "listView"
$WeekList.Location = New-Object System.Drawing.Point(0, $($($TimeSheets.Location.Y + $TimeSheets.Size.Height + 5)))
$WeekList.CheckBoxes = $true
$WeekList.Items.AddRange(@('Sat', 'Sun', 'Mon', 'Tues', 'Wed', 'Thur', 'Fri'))
$WeekList.Size = New-Object System.Drawing.Size(400, 50)
$WeekList.Dock = 'Fill'

$TimeSubmission_Group = New-Object system.Windows.Forms.Groupbox
$TimeSubmission_Group.height = 100
$TimeSubmission_Group.width = 200
$TimeSubmission_Group.text = "Time Submission"
$TimeSubmission_Group.Location = New-Object System.Drawing.Point($($Form.Width * .05), $($($WeekList.Location.Y + $WeekList.Size.Height + 5)))
$TimeSubmission_Group.Dock = 'Fill'
#$TimeSubmission_Group.Enabled = $false

$TimeIn_TextBox = New-Object system.Windows.Forms.TextBox
$TimeIn_TextBox.multiline = $false
$TimeIn_TextBox.Size = New-Object System.Drawing.Size($($TimeSubmission_Group.Width * .33), 50)
$TimeIn_TextBox.Location = New-Object System.Drawing.Point($(($TimeSubmission_Group.Width / 16)), $(($TimeSubmission_Group.Height - $TimeIn_TextBox.Height) * .5))
$TimeIn_TextBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$TimeIn_Label = New-Object system.Windows.Forms.Label
$TimeIn_Label.text = "In"
$TimeIn_Label.AutoSize = $true
$TimeIn_Label.Size = New-Object System.Drawing.Size($(([System.Windows.Forms.TextRenderer]::MeasureText($TimeIn_Label.Text, $TimeIn_Label.Font)).Width), 20)
$TimeIn_Label.Location = New-Object System.Drawing.Point($($TimeIn_TextBox.Location.X + (($TimeIn_TextBox.Width / 2) - ($TimeIn_Label.Width / 2))), $($TimeIn_TextBox.Location.Y - 20))
$TimeIn_Label.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$Lunch_TextBox = New-Object system.Windows.Forms.TextBox
$Lunch_TextBox.multiline = $false
$Lunch_TextBox.Size = New-Object System.Drawing.Size($($TimeSubmission_Group.Width * .33), 50)
$Lunch_TextBox.Location = New-Object System.Drawing.Point($(($TimeIn_TextBox.Location.X + $TimeIn_TextBox.Width) + 10), $(($TimeSubmission_Group.Height - $Lunch_TextBox.Height) * .5))
$Lunch_TextBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$Lunch_Label = New-Object system.Windows.Forms.Label
$Lunch_Label.text = "Lunch"
$Lunch_Label.AutoSize = $true
$Lunch_Label.Size = New-Object System.Drawing.Size($(([System.Windows.Forms.TextRenderer]::MeasureText($Lunch_Label.Text, $Lunch_Label.Font)).Width), 20)
$Lunch_Label.Location = New-Object System.Drawing.Point($($Lunch_TextBox.Location.X + (($Lunch_TextBox.Width / 2) - ($Lunch_Label.Width / 2))), $($Lunch_TextBox.Location.Y - 20))
$Lunch_Label.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$LunchTimeAmount_CheckBox = New-Object system.Windows.Forms.CheckBox
$LunchTimeAmount_CheckBox.text = "Half Hour"
$LunchTimeAmount_CheckBox.AutoSize = $false
$LunchTimeAmount_CheckBox.width = 95
$LunchTimeAmount_CheckBox.height = 20
$LunchTimeAmount_CheckBox.location = New-Object System.Drawing.Point($($Lunch_TextBox.Location.X), $($Lunch_TextBox.Location.Y + $Lunch_TextBox.height))
$LunchTimeAmount_CheckBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$TimeOut_TextBox = New-Object system.Windows.Forms.TextBox
$TimeOut_TextBox.multiline = $false
$TimeOut_TextBox.Size = New-Object System.Drawing.Size($($TimeSubmission_Group.Width * .33), 50)
$TimeOut_TextBox.Location = New-Object System.Drawing.Point($(($Lunch_TextBox.Location.X + $Lunch_TextBox.Width) + 10), $(($TimeSubmission_Group.Height - $TimeOut_TextBox.Height) * .5))
$TimeOut_TextBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$TimeOut_Label = New-Object system.Windows.Forms.Label
$TimeOut_Label.text = "Out"
$TimeOut_Label.AutoSize = $true
$TimeOut_Label.Size = New-Object System.Drawing.Size($(([System.Windows.Forms.TextRenderer]::MeasureText($TimeOut_Label.Text, $TimeOut_Label.Font)).Width), 20)
$TimeOut_Label.Location = New-Object System.Drawing.Point($($TimeOut_TextBox.Location.X + (($TimeOut_TextBox.Width / 2) - ($TimeOut_Label.Width / 2))), $($TimeOut_TextBox.Location.Y - 20))
$TimeOut_Label.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$Options_Group = New-Object system.Windows.Forms.Groupbox
$Options_Group.height = 100
$Options_Group.width = 200
$Options_Group.text = "Options"
$Options_Group.Location = New-Object System.Drawing.Point($($TimeSubmission_Group.Location.X + $TimeSubmission_Group.Width), $($TimeSubmission_Group.Location.Y))
$Options_Group.Dock = 'Fill'
#$Options_Group.Enabled = $false

$QuickSelect_Button = New-Object system.Windows.Forms.Button
$QuickSelect_Button.text = "Select M-F"
$QuickSelect_Button.Size = New-Object System.Drawing.Size($(([System.Windows.Forms.TextRenderer]::MeasureText($QuickSelect_Button.Text, $QuickSelect_Button.Font)).Width * 1.2), 30)
$QuickSelect_Button.Dock = 'Top'

$ClearTimesheet_Button = New-Object system.Windows.Forms.Button
$ClearTimesheet_Button.text = "Clear Timesheet"
$ClearTimesheet_Button.Size = New-Object System.Drawing.Size($(([System.Windows.Forms.TextRenderer]::MeasureText($ClearTimesheet_Button.Text, $ClearTimesheet_Button.Font)).Width * 1.2), 30)
$ClearTimesheet_Button.Dock = 'Top'

$Submit_Button = New-Object system.Windows.Forms.Button
$Submit_Button.text = "Submit"
$Submit_Button.Dock = 'Bottom'
$Form.AcceptButton = $Submit_Button

$panel.Controls.Add($TimeSheets, 0, 0)
$panel.Controls.Add($WeekList, 0, 1)
$panel.Controls.Add($TimeSubmission_Group, 0, 2)
$panel.Controls.Add($Options_Group, 2, 1)
$panel.Controls.Add($Submit_Button, 1, 2)
$Form.controls.AddRange(@($panel))
$TimeSubmission_Group.controls.AddRange(@($TimeIn_TextBox, $TimeIn_Label, $Lunch_TextBox, $Lunch_Label, $TimeOut_TextBox, $TimeOut_Label, $LunchTimeAmount_CheckBox))
$Options_Group.controls.AddRange(@($QuickSelect_Button, $ClearTimesheet_Button))

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
        startDate = [DateTime]::ParseExact("$($startDay)$($startMonth)$($startYear)1200am", 'ddMMMyyyyhhmmtt', $null)
        endDate   = [DateTime]::ParseExact("$($endDay)$($endMonth)$($endYear)1159pm", 'ddMMMyyyyhhmmtt', $null)
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

$QuickSelect_Button.Add_Click( {
        foreach ($Day in $WeekList.items[2..7]) {
            $Day.Checked = $true
        }
    })

$Submit_Button.Add_Click( {
        foreach ($PayRange in $TimeSheets.CheckedItems) {
            Get-SeSelectionOption -Element $PayPeriodDropdown -ByPartialText "$($PayRange.Text)"
            
            # Click Time Sheet button
            Find-SeElement -XPath '/html/body/div[3]/form/table[2]/tbody/tr/td/input' -Driver $Driver | Invoke-SeClick -Driver $Driver

            $JobsSeqNo = Find-SeElement -Name 'JobsSeqNo' -Driver $Driver | Get-SeElementAttribute -Attribute 'Value'

            # This code is meant to pull the timehseet table and return which days already have time filled out. I am other thinking it at the moment, will get back to someday...
            <#
            $TimeSheetTable = Get-SeElement -Driver $Driver -XPath '/html/body/div[3]/table[1]/tbody/tr[5]/td/form/table[1]/tbody'
            /html/body/div[3]/table[1]/tbody/tr[5]/td/form/table[1]/tbody/tr[1]/td[6]

            $TimeSheetTableRows = $TimeSheetTable.FindElementsByTagName('tr')

            $EnteredHours = @()

            $headers = $TimeSheetTableRows[0].text -split '\n'

            for ($i = 4; $i -lt $headers.count; $i++) {
                $headers[$i] | Select-String -Pattern "(?<startMonth>\w{3}) (?<startDay>\d{2}), (?<startYear>\d{4})" |
                ForEach-Object {
                    $startMonth, $startDay, $startYear = $_.Matches[0].Groups['startMonth', 'startDay', 'startYear'].Value
                    $EnteredHours += [PSCustomObject] @{
                        header = [DateTime]::ParseExact("$($startDay)$($startMonth)$($startYear)", 'ddMMMyyyy', $null)
                        earnCode = $null
                        day = $null
                    }
                }
            }

            #Ignore Total Hours/Unit Rows
            foreach ($Row in $TimeSheetTableRows[1..($TimeSheetTableRows.Count - 3)]) {
                $dataSeperated = $Row.text -split '\n'
                
                for ($i = 0; $i -lt $EarnCodes.Count; $i++) {
                    if ($EarnCodes[$i][1] -match $dataSeperated[0]) {
                        $EnteredHours.earnCode = $EarnCodes[$i][1]
                        break
                    }
                }
                
                else {
                    
                }

                for ($i = 4; $i -lt $dataSeperated.Count; $i++) {
                    if ($day -notmatch 'Enter Hours') {
                        $EnteredHours[$i]
                    }
                }
                
                $EnteredHours += [PSCustomObject] @{
                    Earning = $EarnCode
                    days    = $null
    
                }
            }
            #>
            
            # Add DateTime data for each Checkbox day, used to simplfy TimeSheet URL creation
            for ($i = 0; $i -le [math]::Round(($PayRange.Tag.enddate - $PayRange.Tag.startdate).Totaldays, 1); $i++) {
                foreach ($day in $WeekList.Items) {
                    if ($PayRange.Tag.startDate.addDays($i).dayofweek -match $day.text) {
                        $day.Tag += @($PayRange.Tag.startdate.addDays($i))
                        break
                    }
                }
            }

            foreach ($checkedDay in $WeekList.CheckedItems.Tag) {
                if (($checkedDay -ge $PayRange.Tag.startdate) -and ($checkedDay -le $PayRange.Tag.enddate)) {
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
                if ($checkedDay -eq $WeekList.CheckedItems.Tag[-1]) {
                    # Click Time Sheet button
                    Find-SeElement -XPath '/html/body/div[3]/form/table[3]/tbody/tr[1]/td/input[1]' -Driver $Driver | Invoke-SeClick -Driver $Driver
                }
                
            }

            # Click Position Selection button
            Find-SeElement -XPath '/html/body/div[3]/table[1]/tbody/tr[5]/td/form/table[2]/tbody/tr/td[1]/input' -Driver $Driver | Invoke-SeClick -Driver $Driver

            #Reload dropdown, becasue the object becomes stale
            $PayPeriodDropdown = Find-SeElement -Driver $Driver -Id "period_1_id"
                       
        }
    })


$ClearTimesheet_Button.Add_Click( {

    })

[void]$Form.ShowDialog()

#Stop-SeDriver -Driver $Driver