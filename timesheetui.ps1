# https://www.powershellgallery.com/packages/Selenium/3.0.0

#$Credentials = Get-Credential

#Install-Module -Name Selenium -RequiredVersion 3.0.0
. .\Functions.ps1
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
$panel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 60)))
$panel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 40)))
$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33)))
$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33)))
$panel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25)))

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
$TimeSubmission_Group.Dock  = 'Fill'
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


$panel.Controls.Add($TimeSheets,0,0)
$panel.Controls.Add($WeekList,0,1)
$panel.Controls.Add($TimeSubmission_Group,0,2)
$panel.Controls.Add($Options_Group,2,1)
$panel.Controls.Add($Submit_Button,1,2)
$Form.controls.AddRange(@($panel))
$TimeSubmission_Group.controls.AddRange(@($TimeIn_TextBox, $TimeIn_Label, $Lunch_TextBox, $Lunch_Label, $TimeOut_TextBox, $TimeOut_Label, $LunchTimeAmount_CheckBox))
$Options_Group.controls.AddRange(@($QuickSelect_Button, $ClearTimesheet_Button))



[void]$Form.ShowDialog()