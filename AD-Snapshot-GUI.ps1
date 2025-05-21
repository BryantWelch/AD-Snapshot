#region Import Modules and Add Windows Forms Assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Create a single ToolTip object for all controls
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 5000
$toolTip.InitialDelay = 500
$toolTip.ReshowDelay = 500
$toolTip.ShowAlways = $true

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
$toolTip.SetToolTip($txtOUList, "Enter the OU(s) you want to run the script against. For multiple OUs, use comma-separated values like 'OU1,OU2,OU3'.")
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
$toolTip.SetToolTip($txtRealm, "Enter your organization's realm name (e.g., 'contoso'). This identifies your organization in the report.")
$tabConfig.Controls.Add($txtRealm)

# Domain Controller
$lblDomainController = New-Object System.Windows.Forms.Label
$lblDomainController.Text = "Domain Controller:"
$lblDomainController.Location = New-Object System.Drawing.Point(10, 80)
$lblDomainController.Size = New-Object System.Drawing.Size(200, 20)
$tabConfig.Controls.Add($lblDomainController)

$txtDomainController = New-Object System.Windows.Forms.TextBox
$txtDomainController.Location = New-Object System.Drawing.Point(220, 80)
$txtDomainController.Size = New-Object System.Drawing.Size(530, 20)
$txtDomainController.Text = "domaincontroller.example.org"
$toolTip.SetToolTip($txtDomainController, "Enter the fully qualified domain name of your domain controller (e.g., 'DC01.contoso.com').")
$tabConfig.Controls.Add($txtDomainController)

# Domain Common Name
$lblDomainCN = New-Object System.Windows.Forms.Label
$lblDomainCN.Text = "Domain Common Name:"
$lblDomainCN.Location = New-Object System.Drawing.Point(10, 110)
$lblDomainCN.Size = New-Object System.Drawing.Size(200, 20)
$tabConfig.Controls.Add($lblDomainCN)

$txtDomainCN = New-Object System.Windows.Forms.TextBox
$txtDomainCN.Location = New-Object System.Drawing.Point(220, 110)
$txtDomainCN.Size = New-Object System.Drawing.Size(530, 20)
$txtDomainCN.Text = "DC=example,DC=org"
$toolTip.SetToolTip($txtDomainCN, "Enter your domain's Common Name in LDAP format (e.g., 'DC=contoso,DC=com'). Use DC= prefix for each domain component.")
$tabConfig.Controls.Add($txtDomainCN)

# Admin Group
$lblAdminGroup = New-Object System.Windows.Forms.Label
$lblAdminGroup.Text = "Admin Group:"
$lblAdminGroup.Location = New-Object System.Drawing.Point(10, 140)
$lblAdminGroup.Size = New-Object System.Drawing.Size(200, 20)
$tabConfig.Controls.Add($lblAdminGroup)

$txtAdminGroup = New-Object System.Windows.Forms.TextBox
$txtAdminGroup.Location = New-Object System.Drawing.Point(220, 140)
$txtAdminGroup.Size = New-Object System.Drawing.Size(530, 20)
$txtAdminGroup.Text = "Example Admin"
$toolTip.SetToolTip($txtAdminGroup, "Enter the name of the admin group you want to report on (e.g., 'Domain Admins').")
$tabConfig.Controls.Add($txtAdminGroup)

# Server Uptime Alarm
$lblServerUpTimeAlarm = New-Object System.Windows.Forms.Label
$lblServerUpTimeAlarm.Text = "Server Uptime Alarm (days):"
$lblServerUpTimeAlarm.Location = New-Object System.Drawing.Point(10, 170)
$lblServerUpTimeAlarm.Size = New-Object System.Drawing.Size(200, 20)
$tabConfig.Controls.Add($lblServerUpTimeAlarm)

$numServerUpTimeAlarm = New-Object System.Windows.Forms.NumericUpDown
$numServerUpTimeAlarm.Location = New-Object System.Drawing.Point(220, 170)
$numServerUpTimeAlarm.Size = New-Object System.Drawing.Size(80, 20)
$numServerUpTimeAlarm.Minimum = 1
$numServerUpTimeAlarm.Maximum = 365
$numServerUpTimeAlarm.Value = 30
$toolTip.SetToolTip($numServerUpTimeAlarm, "Number of days a server can be up before its uptime is marked in red. Helps identify servers that need rebooting.")
$tabConfig.Controls.Add($numServerUpTimeAlarm)

# Computer Stale Days
$lblComputerStaleDays = New-Object System.Windows.Forms.Label
$lblComputerStaleDays.Text = "Computer Stale Days:"
$lblComputerStaleDays.Location = New-Object System.Drawing.Point(10, 200)
$lblComputerStaleDays.Size = New-Object System.Drawing.Size(200, 20)
$tabConfig.Controls.Add($lblComputerStaleDays)

$numComputerStaleDays = New-Object System.Windows.Forms.NumericUpDown
$numComputerStaleDays.Location = New-Object System.Drawing.Point(220, 200)
$numComputerStaleDays.Size = New-Object System.Drawing.Size(80, 20)
$numComputerStaleDays.Minimum = 1
$numComputerStaleDays.Maximum = 365
$numComputerStaleDays.Value = 80
$toolTip.SetToolTip($numComputerStaleDays, "Number of days a computer can be off the network before it's considered stale. Helps identify inactive computers.")
$tabConfig.Controls.Add($numComputerStaleDays)

# User Stale Days
$lblUserStaleDays = New-Object System.Windows.Forms.Label
$lblUserStaleDays.Text = "User Stale Days:"
$lblUserStaleDays.Location = New-Object System.Drawing.Point(10, 230)
$lblUserStaleDays.Size = New-Object System.Drawing.Size(200, 20)
$tabConfig.Controls.Add($lblUserStaleDays)

$numUserStaleDays = New-Object System.Windows.Forms.NumericUpDown
$numUserStaleDays.Location = New-Object System.Drawing.Point(220, 230)
$numUserStaleDays.Size = New-Object System.Drawing.Size(80, 20)
$numUserStaleDays.Minimum = 1
$numUserStaleDays.Maximum = 365
$numUserStaleDays.Value = 80
$toolTip.SetToolTip($numUserStaleDays, "Number of days a user can be inactive before they're considered stale. Helps identify unused accounts.")
$tabConfig.Controls.Add($numUserStaleDays)

# PC Sort
$lblPCSort = New-Object System.Windows.Forms.Label
$lblPCSort.Text = "PC Sort By:"
$lblPCSort.Location = New-Object System.Drawing.Point(10, 260)
$lblPCSort.Size = New-Object System.Drawing.Size(200, 20)
$tabConfig.Controls.Add($lblPCSort)

$cboPCSort = New-Object System.Windows.Forms.ComboBox
$cboPCSort.Location = New-Object System.Drawing.Point(220, 260)
$cboPCSort.Size = New-Object System.Drawing.Size(150, 20)
$cboPCSort.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboPCSort.Items.Add("Name")
$cboPCSort.Items.Add("Last Seen Date")
$cboPCSort.SelectedIndex = 0
$toolTip.SetToolTip($cboPCSort, "Choose how to sort computer entries in the report: by name or by the date last seen on the network.")
$tabConfig.Controls.Add($cboPCSort)

# User Sort
$lblUserSort = New-Object System.Windows.Forms.Label
$lblUserSort.Text = "User Sort By:"
$lblUserSort.Location = New-Object System.Drawing.Point(10, 290)
$lblUserSort.Size = New-Object System.Drawing.Size(200, 20)
$tabConfig.Controls.Add($lblUserSort)

$cboUserSort = New-Object System.Windows.Forms.ComboBox
$cboUserSort.Location = New-Object System.Drawing.Point(220, 290)
$cboUserSort.Size = New-Object System.Drawing.Size(150, 20)
$cboUserSort.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboUserSort.Items.Add("Name")
$cboUserSort.Items.Add("Last Login Date")
$cboUserSort.SelectedIndex = 0
$toolTip.SetToolTip($cboUserSort, "Choose how to sort user entries in the report: by name or by the date of last login.")
$tabConfig.Controls.Add($cboUserSort)

# Hide Disabled Users
$chkHideDisabledUsers = New-Object System.Windows.Forms.CheckBox
$chkHideDisabledUsers.Text = "Hide Disabled Users"
$chkHideDisabledUsers.Location = New-Object System.Drawing.Point(10, 320)
$chkHideDisabledUsers.Size = New-Object System.Drawing.Size(200, 20)
$chkHideDisabledUsers.Checked = $false
$toolTip.SetToolTip($chkHideDisabledUsers, "When checked, disabled user accounts will not be shown in the report.")
$tabConfig.Controls.Add($chkHideDisabledUsers)

# List ALL Computers
$chkListAllComputers = New-Object System.Windows.Forms.CheckBox
$chkListAllComputers.Text = "List ALL Computers (including servers)"
$chkListAllComputers.Location = New-Object System.Drawing.Point(10, 350)
$chkListAllComputers.Size = New-Object System.Drawing.Size(250, 20)
$chkListAllComputers.Checked = $true
$toolTip.SetToolTip($chkListAllComputers, "When checked, servers will also be displayed in the computers section of the report.")
$tabConfig.Controls.Add($chkListAllComputers)



# Skip Sections Group Box
$grpSkipSections = New-Object System.Windows.Forms.GroupBox
$grpSkipSections.Text = "Skip Sections"
$grpSkipSections.Location = New-Object System.Drawing.Point(10, 380)
$grpSkipSections.Size = New-Object System.Drawing.Size(740, 100)
$tabConfig.Controls.Add($grpSkipSections)

# Skip Admin
$chkSkipAdmin = New-Object System.Windows.Forms.CheckBox
$chkSkipAdmin.Text = "Skip Admin Section"
$chkSkipAdmin.Location = New-Object System.Drawing.Point(20, 25)
$chkSkipAdmin.Size = New-Object System.Drawing.Size(200, 20)
$chkSkipAdmin.Checked = $false
$toolTip.SetToolTip($chkSkipAdmin, "When checked, the admin membership section will be skipped in the report.")
$grpSkipSections.Controls.Add($chkSkipAdmin)

# Skip Server
$chkSkipServer = New-Object System.Windows.Forms.CheckBox
$chkSkipServer.Text = "Skip Server Section"
$chkSkipServer.Location = New-Object System.Drawing.Point(20, 50)
$chkSkipServer.Size = New-Object System.Drawing.Size(200, 20)
$chkSkipServer.Checked = $false
$toolTip.SetToolTip($chkSkipServer, "When checked, the server information section will be skipped in the report.")
$grpSkipSections.Controls.Add($chkSkipServer)

# Skip Users
$chkSkipUsers = New-Object System.Windows.Forms.CheckBox
$chkSkipUsers.Text = "Skip Users Section"
$chkSkipUsers.Location = New-Object System.Drawing.Point(20, 75)
$chkSkipUsers.Size = New-Object System.Drawing.Size(200, 20)
$chkSkipUsers.Checked = $false
$toolTip.SetToolTip($chkSkipUsers, "When checked, the user information section will be skipped in the report.")
$grpSkipSections.Controls.Add($chkSkipUsers)

# Skip Computers
$chkSkipComputers = New-Object System.Windows.Forms.CheckBox
$chkSkipComputers.Text = "Skip Computers Section"
$chkSkipComputers.Location = New-Object System.Drawing.Point(250, 25)
$chkSkipComputers.Size = New-Object System.Drawing.Size(200, 20)
$chkSkipComputers.Checked = $false
$toolTip.SetToolTip($chkSkipComputers, "When checked, the computer information section will be skipped in the report.")
$grpSkipSections.Controls.Add($chkSkipComputers)

# Skip BitLocker Status
$chkSkipBitlockerStatus = New-Object System.Windows.Forms.CheckBox
$chkSkipBitlockerStatus.Text = "Skip BitLocker Status"
$chkSkipBitlockerStatus.Location = New-Object System.Drawing.Point(250, 50)
$chkSkipBitlockerStatus.Size = New-Object System.Drawing.Size(200, 20)
$chkSkipBitlockerStatus.Checked = $true
$toolTip.SetToolTip($chkSkipBitlockerStatus, "When checked, the BitLocker status section will be skipped in the report.")
$grpSkipSections.Controls.Add($chkSkipBitlockerStatus)
