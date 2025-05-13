#region Run Report Tab
# Status Group Box
$grpStatus = New-Object System.Windows.Forms.GroupBox
$grpStatus.Text = "Status"
$grpStatus.Location = New-Object System.Drawing.Point(10, 20)
$grpStatus.Size = New-Object System.Drawing.Size(740, 380)
$tabRun.Controls.Add($grpStatus)

# Status TextBox
$txtStatus = New-Object System.Windows.Forms.TextBox
$txtStatus.Location = New-Object System.Drawing.Point(20, 25)
$txtStatus.Size = New-Object System.Drawing.Size(700, 340)
$txtStatus.Multiline = $true
$txtStatus.ScrollBars = "Vertical"
$txtStatus.ReadOnly = $true
$txtStatus.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtStatus.Text = "Ready to run AD Snapshot report.`r`n`r`nClick 'Run Report' to begin."
$grpStatus.Controls.Add($txtStatus)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 410)
$progressBar.Size = New-Object System.Drawing.Size(740, 25)
$progressBar.Style = "Continuous"
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$tabRun.Controls.Add($progressBar)

# Run Report Button
$btnRunReport = New-Object System.Windows.Forms.Button
$btnRunReport.Text = "Run Report"
$btnRunReport.Location = New-Object System.Drawing.Point(10, 445)
$btnRunReport.Size = New-Object System.Drawing.Size(120, 30)
$tabRun.Controls.Add($btnRunReport)

# View Report Button
$btnViewReport = New-Object System.Windows.Forms.Button
$btnViewReport.Text = "View Report"
$btnViewReport.Location = New-Object System.Drawing.Point(140, 445)
$btnViewReport.Size = New-Object System.Drawing.Size(120, 30)
$btnViewReport.Enabled = $false
$tabRun.Controls.Add($btnViewReport)

# Open Report Folder Button
$btnOpenReportFolder = New-Object System.Windows.Forms.Button
$btnOpenReportFolder.Text = "Open Report Folder"
$btnOpenReportFolder.Location = New-Object System.Drawing.Point(270, 445)
$btnOpenReportFolder.Size = New-Object System.Drawing.Size(150, 30)
$btnOpenReportFolder.Enabled = $false
$tabRun.Controls.Add($btnOpenReportFolder)

# Cancel Button
$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(630, 445)
$btnCancel.Size = New-Object System.Drawing.Size(120, 30)
$btnCancel.Enabled = $false
$tabRun.Controls.Add($btnCancel)
