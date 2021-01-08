Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Region UI
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "AD Computer Search"
$Form.StartPosition = 'CenterScreen'
$Form.ClientSize = '225, 150'

$LayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$LayoutPanel.Dock = "Fill"
$LayoutPanel.ColumnCount = 4
$LayoutPanel.RowCount = 3
$LayoutPanel.CellBorderStyle = 1
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 15)))
[void]$LayoutPanel.ColumnStyles.Add((new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 15)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
[void]$LayoutPanel.RowStyles.Add((new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))

$SearchField = New-Object System.Windows.Forms.TextBox
$SearchField.Text = ""
$SearchField.Dock = 'Fill'
$SearchField.Anchor = 'Left, Right'

$SearchButton = New-Object System.Windows.Forms.Button
$SearchButton.Text = "Search"
$SearchButton.AutoSize = $True
$SearchButton.Dock = 'Fill'

$ExecuteButton = New-Object System.Windows.Forms.Button
$ExecuteButton.Text = "Execute"
$ExecuteButton.AutoSize = $True
$ExecuteButton.Dock = 'Fill'

$ComputerList = New-Object system.Windows.Forms.ListBox
$ComputerList.Dock = 'Fill'

$Output = New-Object system.Windows.Forms.TextBox
$Output.multiline = $true
$Output.Dock = 'Fill'

$LayoutPanel.Controls.Add($SearchField, 0, 0)
$LayoutPanel.SetColumnSpan($SearchField, 2)
$LayoutPanel.Controls.Add($SearchButton, 2, 0)
$LayoutPanel.Controls.Add($ExecuteButton, 3, 0)
$LayoutPanel.Controls.Add($ComputerList, 0, 1)
$LayoutPanel.SetColumnSpan($ComputerList, 2)
$LayoutPanel.SetRowSpan($ComputerList, 2)
$LayoutPanel.Controls.Add($Output, 1, 1)
$LayoutPanel.SetColumnSpan($Output, 2)
$LayoutPanel.SetRowSpan($Output, 2)


$Form.Controls.Add($LayoutPanel)
#EndRegion

#Region UI Events
$SearchButton.Add_Click( {
        $Computer = Get-ADComputer -Filter ('Name -Like "*' + $SearchField.Text + '*"')
        $Computer | ForEach-Object { $ComputerList.Items.Add($_.Name) }
    })

$ExecuteButton.Add_Click( {
        foreach ($Computer in $ComputerList.SelectedItems) {
            $Output.Text += $Computer
        }
    })

[void]$Form.ShowDialog()