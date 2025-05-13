#region Import Modules and Add Windows Forms Assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

#region Form Creation
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Active Directory Snapshot Tool"
$mainForm.Size = New-Object System.Drawing.Size(800, 600)
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$mainForm.MaximizeBox = $false
$mainForm.MinimizeBox = $true
$mainForm.Icon = [System.Drawing.SystemIcons]::Shield

# Create Tab Control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(770, 500)
$tabControl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$mainForm.Controls.Add($tabControl)

# Create Tabs
$tabConfig = New-Object System.Windows.Forms.TabPage
$tabConfig.Text = "Configuration"
$tabControl.Controls.Add($tabConfig)

$tabOutput = New-Object System.Windows.Forms.TabPage
$tabOutput.Text = "Output Options"
$tabControl.Controls.Add($tabOutput)

$tabRun = New-Object System.Windows.Forms.TabPage
$tabRun.Text = "Run Report"
$tabControl.Controls.Add($tabRun)

#region Configuration Tab
# OU List
$lblOUList = New-Object System.Windows.Forms.Label
$lblOUList.Text = "OU List (comma separated):"
$lblOUList.Location = New-Object System.Drawing.Point(10, 20)
$lblOUList.Size = New-Object System.Drawing.Size(200, 20)
$tabConfig.Controls.Add($lblOUList)

$txtOUList = New-Object System.Windows.Forms.TextBox
$txtOUList.Location = New-Object System.Drawing.Point(220, 20)
$txtOUList.Size = New-Object System.Drawing.Size(530, 20)
$txtOUList.Text = "XXX"
$tabConfig.Controls.Add($txtOUList)

# Realm
$lblRealm = New-Object System.Windows.Forms.Label
$lblRealm.Text = "Realm Name:"
$lblRealm.Location = New-Object System.Drawing.Point(10, 50)
$lblRealm.Size = New-Object System.Drawing.Size(200, 20)
$tabConfig.Controls.Add($lblRealm)

$txtRealm = New-Object System.Windows.Forms.TextBox
$txtRealm.Location = New-Object System.Drawing.Point(220, 50)
$txtRealm.Size = New-Object System.Drawing.Size(530, 20)
$txtRealm.Text = "example"
$tabConfig.Controls.Add($txtRealm)

# Admin Group
$lblAdminGroup = New-Object System.Windows.Forms.Label
$lblAdminGroup.Text = "Admin Group:"
$lblAdminGroup.Location = New-Object System.Drawing.Point(10, 80)
$lblAdminGroup.Size = New-Object System.Drawing.Size(200, 20)
$tabConfig.Controls.Add($lblAdminGroup)

$txtAdminGroup = New-Object System.Windows.Forms.TextBox
$txtAdminGroup.Location = New-Object System.Drawing.Point(220, 80)
$txtAdminGroup.Size = New-Object System.Drawing.Size(530, 20)
$txtAdminGroup.Text = "Example Admin"
$tabConfig.Controls.Add($txtAdminGroup)

# Server Uptime Alarm
$lblServerUpTimeAlarm = New-Object System.Windows.Forms.Label
$lblServerUpTimeAlarm.Text = "Server Uptime Alarm (days):"
$lblServerUpTimeAlarm.Location = New-Object System.Drawing.Point(10, 110)
$lblServerUpTimeAlarm.Size = New-Object System.Drawing.Size(200, 20)
$tabConfig.Controls.Add($lblServerUpTimeAlarm)

$numServerUpTimeAlarm = New-Object System.Windows.Forms.NumericUpDown
$numServerUpTimeAlarm.Location = New-Object System.Drawing.Point(220, 110)
$numServerUpTimeAlarm.Size = New-Object System.Drawing.Size(80, 20)
$numServerUpTimeAlarm.Minimum = 1
$numServerUpTimeAlarm.Maximum = 365
$numServerUpTimeAlarm.Value = 30
$tabConfig.Controls.Add($numServerUpTimeAlarm)

# Computer Stale Days
$lblComputerStaleDays = New-Object System.Windows.Forms.Label
$lblComputerStaleDays.Text = "Computer Stale Days:"
$lblComputerStaleDays.Location = New-Object System.Drawing.Point(10, 140)
$lblComputerStaleDays.Size = New-Object System.Drawing.Size(200, 20)
$tabConfig.Controls.Add($lblComputerStaleDays)

$numComputerStaleDays = New-Object System.Windows.Forms.NumericUpDown
$numComputerStaleDays.Location = New-Object System.Drawing.Point(220, 140)
$numComputerStaleDays.Size = New-Object System.Drawing.Size(80, 20)
$numComputerStaleDays.Minimum = 1
$numComputerStaleDays.Maximum = 365
$numComputerStaleDays.Value = 80
$tabConfig.Controls.Add($numComputerStaleDays)

# User Stale Days
$lblUserStaleDays = New-Object System.Windows.Forms.Label
$lblUserStaleDays.Text = "User Stale Days:"
$lblUserStaleDays.Location = New-Object System.Drawing.Point(10, 170)
$lblUserStaleDays.Size = New-Object System.Drawing.Size(200, 20)
$tabConfig.Controls.Add($lblUserStaleDays)

$numUserStaleDays = New-Object System.Windows.Forms.NumericUpDown
$numUserStaleDays.Location = New-Object System.Drawing.Point(220, 170)
$numUserStaleDays.Size = New-Object System.Drawing.Size(80, 20)
$numUserStaleDays.Minimum = 1
$numUserStaleDays.Maximum = 365
$numUserStaleDays.Value = 80
$tabConfig.Controls.Add($numUserStaleDays)

# PC Sort
$lblPCSort = New-Object System.Windows.Forms.Label
$lblPCSort.Text = "PC Sort By:"
$lblPCSort.Location = New-Object System.Drawing.Point(10, 200)
$lblPCSort.Size = New-Object System.Drawing.Size(200, 20)
$tabConfig.Controls.Add($lblPCSort)

$cboPCSort = New-Object System.Windows.Forms.ComboBox
$cboPCSort.Location = New-Object System.Drawing.Point(220, 200)
$cboPCSort.Size = New-Object System.Drawing.Size(150, 20)
$cboPCSort.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboPCSort.Items.Add("Name")
$cboPCSort.Items.Add("Last Seen Date")
$cboPCSort.SelectedIndex = 0
$tabConfig.Controls.Add($cboPCSort)

# User Sort
$lblUserSort = New-Object System.Windows.Forms.Label
$lblUserSort.Text = "User Sort By:"
$lblUserSort.Location = New-Object System.Drawing.Point(10, 230)
$lblUserSort.Size = New-Object System.Drawing.Size(200, 20)
$tabConfig.Controls.Add($lblUserSort)

$cboUserSort = New-Object System.Windows.Forms.ComboBox
$cboUserSort.Location = New-Object System.Drawing.Point(220, 230)
$cboUserSort.Size = New-Object System.Drawing.Size(150, 20)
$cboUserSort.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboUserSort.Items.Add("Name")
$cboUserSort.Items.Add("Last Login Date")
$cboUserSort.SelectedIndex = 0
$tabConfig.Controls.Add($cboUserSort)

# Hide Disabled Users
$chkHideDisabledUsers = New-Object System.Windows.Forms.CheckBox
$chkHideDisabledUsers.Text = "Hide Disabled Users"
$chkHideDisabledUsers.Location = New-Object System.Drawing.Point(10, 260)
$chkHideDisabledUsers.Size = New-Object System.Drawing.Size(200, 20)
$chkHideDisabledUsers.Checked = $false
$tabConfig.Controls.Add($chkHideDisabledUsers)

# List ALL Computers
$chkListAllComputers = New-Object System.Windows.Forms.CheckBox
$chkListAllComputers.Text = "List ALL Computers (including servers)"
$chkListAllComputers.Location = New-Object System.Drawing.Point(10, 290)
$chkListAllComputers.Size = New-Object System.Drawing.Size(250, 20)
$chkListAllComputers.Checked = $true
$tabConfig.Controls.Add($chkListAllComputers)

# Skip Sections Group Box
$grpSkipSections = New-Object System.Windows.Forms.GroupBox
$grpSkipSections.Text = "Skip Sections"
$grpSkipSections.Location = New-Object System.Drawing.Point(10, 320)
$grpSkipSections.Size = New-Object System.Drawing.Size(740, 100)
$tabConfig.Controls.Add($grpSkipSections)

# Skip Admin
$chkSkipAdmin = New-Object System.Windows.Forms.CheckBox
$chkSkipAdmin.Text = "Skip Admin Section"
$chkSkipAdmin.Location = New-Object System.Drawing.Point(20, 25)
$chkSkipAdmin.Size = New-Object System.Drawing.Size(200, 20)
$chkSkipAdmin.Checked = $false
$grpSkipSections.Controls.Add($chkSkipAdmin)

# Skip Server
$chkSkipServer = New-Object System.Windows.Forms.CheckBox
$chkSkipServer.Text = "Skip Server Section"
$chkSkipServer.Location = New-Object System.Drawing.Point(20, 50)
$chkSkipServer.Size = New-Object System.Drawing.Size(200, 20)
$chkSkipServer.Checked = $false
$grpSkipSections.Controls.Add($chkSkipServer)

# Skip Users
$chkSkipUsers = New-Object System.Windows.Forms.CheckBox
$chkSkipUsers.Text = "Skip Users Section"
$chkSkipUsers.Location = New-Object System.Drawing.Point(20, 75)
$chkSkipUsers.Size = New-Object System.Drawing.Size(200, 20)
$chkSkipUsers.Checked = $false
$grpSkipSections.Controls.Add($chkSkipUsers)

# Skip Computers
$chkSkipComputers = New-Object System.Windows.Forms.CheckBox
$chkSkipComputers.Text = "Skip Computers Section"
$chkSkipComputers.Location = New-Object System.Drawing.Point(250, 25)
$chkSkipComputers.Size = New-Object System.Drawing.Size(200, 20)
$chkSkipComputers.Checked = $false
$grpSkipSections.Controls.Add($chkSkipComputers)

# Skip BitLocker Status
$chkSkipBitlockerStatus = New-Object System.Windows.Forms.CheckBox
$chkSkipBitlockerStatus.Text = "Skip BitLocker Status"
$chkSkipBitlockerStatus.Location = New-Object System.Drawing.Point(250, 50)
$chkSkipBitlockerStatus.Size = New-Object System.Drawing.Size(200, 20)
$chkSkipBitlockerStatus.Checked = $true
$grpSkipSections.Controls.Add($chkSkipBitlockerStatus)
