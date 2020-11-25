#region GUI
#https://colorhunt.co/
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$Global:ErrorProvider = New-Object System.Windows.Forms.ErrorProvider

$Form = New-Object system.Windows.Forms.Form
$Form.FormBorderStyle = "Sizable"
$Form.ClientSize = "400,400"
$Form.text = "Equipment Repair Form"
$Form.TopMost = $true
$Form.StartPosition = 'CenterScreen'

$LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$LayoutPanel.Dock = "Fill"
$LayoutPanel.ColumnCount = 7
$LayoutPanel.RowCount = 4
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 1)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))

$ID_Num_Text = New-Object system.Windows.Forms.TextBox
$ID_Num_Text.multiline = $false
$ID_Num_Text.Dock = 'Fill'
$ID_Num_Group = New-Object system.Windows.Forms.Groupbox
$ID_Num_Group.text = "ID Number"
$ID_Num_Group.Dock = 'Fill'
$ID_Num_Group.controls.Add($ID_Num_Text)

$Location_Dropdown = New-Object system.Windows.Forms.ComboBox
$Location_Dropdown.DropDownStyle = "DropDown"
$Location_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Location_Dropdown.AutoCompleteSource = 'ListItems'
$Location_Dropdown.DropDownWidth = '100'
$Location_Dropdown.Dock = 'Fill'
$Location_Group = New-Object system.Windows.Forms.Groupbox
$Location_Group.text = "Location"
$Location_Group.Dock = 'Fill'
$Location_Group.controls.Add($Location_Dropdown)

$Equipment_Dropdown = New-Object system.Windows.Forms.ComboBox
$Equipment_Dropdown.AutoCompleteMode = 'SuggestAppend'
$Equipment_Dropdown.AutoCompleteSource = 'ListItems'
$Equipment_Dropdown.Dock = 'Fill'
$Equipment_Group = New-Object system.Windows.Forms.Groupbox
$Equipment_Group.text = "Equipment"
$Equipment_Group.Dock = 'Fill'
$Equipment_Group.controls.Add($Equipment_Dropdown)

$Desc_Text = New-Object system.Windows.Forms.TextBox
$Desc_Text.multiline = $true
$Desc_Text.AutoSize = $true
$Desc_Text.Dock = 'Fill'
$Desc_Group = New-Object system.Windows.Forms.Groupbox
$Desc_Group.text = "Description of Issue"
$Desc_Group.Dock = 'Fill'
$Desc_Group.Controls.Add($Desc_Text)

$Current_Issues = New-Object system.Windows.Forms.DataGridView
$Current_Issues.ScrollBars = "Vertical"
$Current_Issues.AutoSizeColumnsMode = 'Fill'
$Current_Issues.RowHeadersVisible = $false
$Current_Issues.ColumnCount = 3
$Current_Issues.Columns[0].Name = 'Equipment'
$Current_Issues.Columns[1].Name = 'Issue'
$Current_Issues.Columns[2].Name = 'Submitted'
$Current_Issues.Columns[2].DefaultCellStyle.Format = "ddMMMyy hh:mm tt"
$Current_Issues.TabStop = $false
$Current_Issues.Dock = 'Fill'
$Current_Issues.Anchor = 'Left,Right'
$Current_Issues_Group = New-Object system.Windows.Forms.Groupbox
$Current_Issues_Group.text = "Current Issues"
$Current_Issues_Group.Dock = 'Fill'
$Current_Issues_Group.Controls.Add($Current_Issues)

$Submit_Button = New-Object system.Windows.Forms.Button
$Submit_Button.text = "Submit"
$Submit_Button.Dock = 'Fill'
$Submit_Button.Anchor = 'Left,Right'

$Form.AcceptButton = $Submit_Button

$Form.controls.Add($LayoutPanel)
#endregion

#region Main Layout
$LayoutPanel.Controls.Add($ID_Num_Group, 1, 0)
$LayoutPanel.Controls.Add($Location_Group, 3, 0)
$LayoutPanel.Controls.Add($Equipment_Group, 5, 0)

$LayoutPanel.Controls.Add($Desc_Group, 1, 1)
$LayoutPanel.SetColumnSpan($Desc_Group, 5)

$LayoutPanel.Controls.Add($Current_Issues_Group, 1, 2)
$LayoutPanel.SetColumnSpan($Current_Issues_Group, 5)

$LayoutPanel.Controls.Add($Submit_Button, 3, 3)
#endregion

#region UI Theme
$Theme = Get-Content -Path $PSScriptRoot\theme.json | ConvertFrom-Json
$Form.Font =  $Theme.Form.Font

if ($null -eq $Theme.LayoutPanel.BackColor) {
    $LayoutPanel.BackgroundImage = [System.Drawing.Image]::Fromfile((get-item $PSScriptRoot\$($Theme.LayoutPanel.BackgroundImage)))
    $LayoutPanel.BackgroundImageLayout = 'Stretch'
}
else{
    $LayoutPanel.BackColor = $Theme.LayoutPanel.BackColor
}

#$LayoutPanel.CellBorderStyle = 1

$ID_Num_Text.Font = $Theme.ID_Num_Text.Font
$ID_Num_Text.ForeColor = $Theme.ID_Num_Text.ForeColor
$ID_Num_Text.BorderStyle = 1
if ($null -eq $Theme.ID_Num_Text.BackColor) {
    $ID_Num_Text.BackgroundImage = [System.Drawing.Image]::Fromfile((get-item $($Theme.ID_Num_Text.BackgroundImage)))
    $ID_Num_Text.BackgroundImageLayout = 'Stretch'
}
else{
    $ID_Num_Text.BackColor = $Theme.ID_Num_Text.BackColor
}

$ID_Num_Group.Font =  $Theme.ID_Num_Group.Font
$ID_Num_Group.ForeColor =  $Theme.ID_Num_Group.ForeColor
if ($null -eq $Theme.ID_Num_Group.BackColor) {
    $ID_Num_Group.BackgroundImage = [System.Drawing.Image]::Fromfile((get-item $($Theme.ID_Num_Group.BackgroundImage)))
    $ID_Num_Group.BackgroundImageLayout = 'Stretch'
}
else{
    $ID_Num_Group.BackColor =  $Theme.ID_Num_Group.BackColor
}

$Location_Dropdown.Font =  $Theme.Location_Dropdown.Font
$Location_Dropdown.ForeColor = $Theme.Location_Dropdown.ForeColor
$Location_Dropdown.FlatStyle = 0
if ($null -eq $Theme.Location_Dropdown.BackColor) {
    $Location_Dropdown.BackgroundImage = [System.Drawing.Image]::Fromfile((get-item $($Theme.Location_Dropdown.BackgroundImage)))
    $Location_Dropdown.BackgroundImageLayout = 'Stretch'
}
else{
    $Location_Dropdown.BackColor =  $Theme.Location_Dropdown.BackColor
}

$Location_Group.Font =  $Theme.Location_Group.Font
$Location_Group.ForeColor = $Theme.Location_Group.ForeColor
if ($null -eq $Theme.Location_Group.BackColor) {
    $Location_Group.BackgroundImage = [System.Drawing.Image]::Fromfile((get-item $($Theme.Location_Group.BackgroundImage)))
    $Location_Group.BackgroundImageLayout = 'Stretch'
}
else{
    $Location_Group.BackColor =  $Theme.Location_Group.BackColor
}

$Equipment_Dropdown.Font =  $Theme.Equipment_Dropdown.Font
$Equipment_Dropdown.ForeColor = $Theme.Equipment_Dropdown.ForeColor
$Equipment_Dropdown.FlatStyle = 0
if ($null -eq $Theme.Equipment_Dropdown.BackColor) {
    $Equipment_Dropdown.BackgroundImage = [System.Drawing.Image]::Fromfile((get-item $($Theme.Equipment_Dropdown.BackgroundImage)))
    $Equipment_Dropdown.BackgroundImageLayout = 'Stretch'
}
else{
    $Equipment_Dropdown.BackColor =  $Theme.Equipment_Dropdown.BackColor
}

$Equipment_Group.Font =  $Theme.Equipment_Group.Font
$Equipment_Group.ForeColor = $Theme.Equipment_Group.ForeColor
if ($null -eq $Theme.Equipment_Group.BackColor) {
    $Equipment_Group.BackgroundImage = [System.Drawing.Image]::Fromfile((get-item $($Theme.Equipment_Group.BackgroundImage)))
    $Equipment_Group.BackgroundImageLayout = 'Stretch'
}
else{
    $Equipment_Group.BackColor =  $Theme.Equipment_Group.BackColor
}

$Desc_Text.Font =  $Theme.Desc_Text.Font
$Desc_Text.ForeColor = $Theme.Desc_Text.ForeColor
$Desc_Text.BorderStyle = 1
if ($null -eq $Theme.Desc_Text.BackColor) {
    $Desc_Text.BackgroundImage = [System.Drawing.Image]::Fromfile((get-item $($Theme.Desc_Text.BackgroundImage)))
    $Desc_Text.BackgroundImageLayout = 'Stretch'
}
else{
    $Desc_Text.BackColor =  $Theme.Desc_Text.BackColor
}

$Desc_Group.Font =  $Theme.Desc_Group.Font
$Desc_Group.ForeColor = $Theme.Desc_Group.ForeColor
if ($null -eq $Theme.Desc_Group.BackColor) {
    $Desc_Group.BackgroundImage = [System.Drawing.Image]::Fromfile((get-item $($Theme.Desc_Group.BackgroundImage)))
    $Desc_Group.BackgroundImageLayout = 'Stretch'
}
else{
    $Desc_Group.BackColor =  $Theme.Desc_Group.BackColor
}

$Current_Issues.Font =  $Theme.Current_Issues.Font
$Current_Issues.ForeColor = $Theme.Current_Issues.ForeColor
if ($null -eq $Theme.Current_Issues.BackColor) {
    $Current_Issues.BackgroundImage = [System.Drawing.Image]::Fromfile((get-item $($Theme.Current_Issues.BackgroundImage)))
    $Current_Issues.BackgroundImageLayout = 'Stretch'
}
else{
    $Current_Issues.BackColor =  $Theme.Current_Issues.BackColor
}
$Current_Issues.GridColor = $Theme.Current_Issues.GridColor
$Current_Issues.RowsDefaultCellStyle.Backcolor  = $Theme.Current_Issues.RowsDefaultCellStyleBackcolor
$Current_Issues.RowsDefaultCellStyle.Forecolor = $Theme.Current_Issues.RowsDefaultCellStyleForecolor

$Current_Issues.ColumnHeadersDefaultCellStyle.Backcolor  = $Theme.Current_Issues.ColumnHeadersDefaultCellStyleBackcolor
$Current_Issues.ColumnHeadersDefaultCellStyle.Forecolor = $Theme.Current_Issues.ColumnHeadersDefaultCellStyleForecolor
$Current_Issues.EnableHeadersVisualStyles = $false
#$Current_Issues.ColumnHeadersBorderStyle = "1"  

$Current_Issues_Group.Font =  $Theme.Current_Issues_Group.Font
$Current_Issues_Group.ForeColor = $Theme.Current_Issues_Group.ForeColor
if ($null -eq $Theme.Current_Issues_Group.BackColor) {
    $Current_Issues_Group.BackgroundImage = [System.Drawing.Image]::Fromfile((get-item $($Theme.Current_Issues_Group.BackgroundImage)))
    $Current_Issues_Group.BackgroundImageLayout = 'Stretch'
}
else{
    $Current_Issues_Group.BackColor =  $Theme.Current_Issues_Group.BackColor
}

$Submit_Button.Font =  $Theme.Submit_Button.Font
$Submit_Button.ForeColor = $Theme.Submit_Button.ForeColor
$Submit_Button.FlatStyle = 1
$Submit_Button.FlatAppearance.BorderSize = 0
if ($null -eq $Theme.Submit_Button.BackColor) {
    $Submit_Button.BackgroundImage = [System.Drawing.Image]::Fromfile((get-item $($Theme.Submit_Button.BackgroundImage)))
    $Submit_Button.BackgroundImageLayout = 'Stretch'
}
else{
    $Submit_Button.BackColor =  $Theme.Submit_Button.BackColor
}
#endregion