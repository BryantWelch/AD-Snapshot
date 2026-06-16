# Configuration tab content.
# The form, tab control, tabs and shared $toolTip are created in AD-Snapshot-GUI-Main.ps1,
# which sources this file. This file only ADDS controls to the existing $tabConfig.
# Controls are organized into labelled GroupBoxes for a cleaner, more professional layout.

#region Configuration Tab

# ============================ Domain & Scope ============================
$grpDomain = New-Object System.Windows.Forms.GroupBox
$grpDomain.Text = "Domain && Scope"
$grpDomain.Location = New-Object System.Drawing.Point(12, 8)
$grpDomain.Size = New-Object System.Drawing.Size(760, 180)
$tabConfig.Controls.Add($grpDomain)

# OU List
$lblOUList = New-Object System.Windows.Forms.Label
$lblOUList.Text = "OU(s):"
$lblOUList.Location = New-Object System.Drawing.Point(15, 30)
$lblOUList.Size = New-Object System.Drawing.Size(190, 20)
$grpDomain.Controls.Add($lblOUList)

$txtOUList = New-Object System.Windows.Forms.TextBox
$txtOUList.Location = New-Object System.Drawing.Point(210, 28)
$txtOUList.Size = New-Object System.Drawing.Size(533, 22)
$txtOUList.Text = ""
$toolTip.SetToolTip($txtOUList, "Leave blank to scan the entire domain, or enter OU name(s) / full OU distinguished name(s). For multiple OUs, separate entries with commas: 'Sales,IT' or 'OU=Sales,DC=contoso,DC=com'.")
$grpDomain.Controls.Add($txtOUList)

# Realm
$lblRealm = New-Object System.Windows.Forms.Label
$lblRealm.Text = "Realm Name:"
$lblRealm.Location = New-Object System.Drawing.Point(15, 60)
$lblRealm.Size = New-Object System.Drawing.Size(190, 20)
$grpDomain.Controls.Add($lblRealm)

$txtRealm = New-Object System.Windows.Forms.TextBox
$txtRealm.Location = New-Object System.Drawing.Point(210, 58)
$txtRealm.Size = New-Object System.Drawing.Size(533, 22)
$txtRealm.Text = ""
$toolTip.SetToolTip($txtRealm, "Report label / realm name. Leave blank to auto-detect the domain's NetBIOS name.")
$grpDomain.Controls.Add($txtRealm)

# Domain Controller
$lblDomainController = New-Object System.Windows.Forms.Label
$lblDomainController.Text = "Domain Controller:"
$lblDomainController.Location = New-Object System.Drawing.Point(15, 90)
$lblDomainController.Size = New-Object System.Drawing.Size(190, 20)
$grpDomain.Controls.Add($lblDomainController)

$txtDomainController = New-Object System.Windows.Forms.TextBox
$txtDomainController.Location = New-Object System.Drawing.Point(210, 88)
$txtDomainController.Size = New-Object System.Drawing.Size(533, 22)
$txtDomainController.Text = ""
$toolTip.SetToolTip($txtDomainController, "Leave blank to auto-detect a domain controller. Set it (e.g. 'DC01.contoso.com') to target a specific DC or another domain.")
$grpDomain.Controls.Add($txtDomainController)

# Domain Common Name
$lblDomainCN = New-Object System.Windows.Forms.Label
$lblDomainCN.Text = "Domain Common Name:"
$lblDomainCN.Location = New-Object System.Drawing.Point(15, 120)
$lblDomainCN.Size = New-Object System.Drawing.Size(190, 20)
$grpDomain.Controls.Add($lblDomainCN)

$txtDomainCN = New-Object System.Windows.Forms.TextBox
$txtDomainCN.Location = New-Object System.Drawing.Point(210, 118)
$txtDomainCN.Size = New-Object System.Drawing.Size(533, 22)
$txtDomainCN.Text = ""
$toolTip.SetToolTip($txtDomainCN, "Leave blank to auto-detect. Otherwise enter the domain DN in LDAP format (e.g. 'DC=contoso,DC=com').")
$grpDomain.Controls.Add($txtDomainCN)

# Admin Group
$lblAdminGroup = New-Object System.Windows.Forms.Label
$lblAdminGroup.Text = "Admin Group(s):"
$lblAdminGroup.Location = New-Object System.Drawing.Point(15, 150)
$lblAdminGroup.Size = New-Object System.Drawing.Size(190, 20)
$grpDomain.Controls.Add($lblAdminGroup)

$txtAdminGroup = New-Object System.Windows.Forms.TextBox
$txtAdminGroup.Location = New-Object System.Drawing.Point(210, 148)
$txtAdminGroup.Size = New-Object System.Drawing.Size(533, 22)
$txtAdminGroup.Text = "Domain Admins"
$toolTip.SetToolTip($txtAdminGroup, "Admin group(s) to report on, resolved anywhere in the domain. Multiple allowed, comma separated (e.g. 'Domain Admins,Enterprise Admins').")
$grpDomain.Controls.Add($txtAdminGroup)

# ============================ Thresholds & Sorting ============================
$grpThresholds = New-Object System.Windows.Forms.GroupBox
$grpThresholds.Text = "Thresholds && Sorting"
$grpThresholds.Location = New-Object System.Drawing.Point(12, 196)
$grpThresholds.Size = New-Object System.Drawing.Size(372, 170)
$tabConfig.Controls.Add($grpThresholds)

# Server Uptime Alarm
$lblServerUpTimeAlarm = New-Object System.Windows.Forms.Label
$lblServerUpTimeAlarm.Text = "Server Uptime Alarm (days):"
$lblServerUpTimeAlarm.Location = New-Object System.Drawing.Point(15, 28)
$lblServerUpTimeAlarm.Size = New-Object System.Drawing.Size(200, 20)
$grpThresholds.Controls.Add($lblServerUpTimeAlarm)

$numServerUpTimeAlarm = New-Object System.Windows.Forms.NumericUpDown
$numServerUpTimeAlarm.Location = New-Object System.Drawing.Point(230, 26)
$numServerUpTimeAlarm.Size = New-Object System.Drawing.Size(80, 22)
$numServerUpTimeAlarm.Minimum = 1
$numServerUpTimeAlarm.Maximum = 365
$numServerUpTimeAlarm.Value = 30
$toolTip.SetToolTip($numServerUpTimeAlarm, "Number of days a server can be up before its uptime is flagged. Helps identify servers that need rebooting.")
$grpThresholds.Controls.Add($numServerUpTimeAlarm)

# Computer Stale Days
$lblComputerStaleDays = New-Object System.Windows.Forms.Label
$lblComputerStaleDays.Text = "Computer Stale Days:"
$lblComputerStaleDays.Location = New-Object System.Drawing.Point(15, 58)
$lblComputerStaleDays.Size = New-Object System.Drawing.Size(200, 20)
$grpThresholds.Controls.Add($lblComputerStaleDays)

$numComputerStaleDays = New-Object System.Windows.Forms.NumericUpDown
$numComputerStaleDays.Location = New-Object System.Drawing.Point(230, 56)
$numComputerStaleDays.Size = New-Object System.Drawing.Size(80, 22)
$numComputerStaleDays.Minimum = 1
$numComputerStaleDays.Maximum = 3650
$numComputerStaleDays.Value = 80
$toolTip.SetToolTip($numComputerStaleDays, "Number of days a computer can be off the network before it's considered stale.")
$grpThresholds.Controls.Add($numComputerStaleDays)

# User Stale Days
$lblUserStaleDays = New-Object System.Windows.Forms.Label
$lblUserStaleDays.Text = "User Stale Days:"
$lblUserStaleDays.Location = New-Object System.Drawing.Point(15, 88)
$lblUserStaleDays.Size = New-Object System.Drawing.Size(200, 20)
$grpThresholds.Controls.Add($lblUserStaleDays)

$numUserStaleDays = New-Object System.Windows.Forms.NumericUpDown
$numUserStaleDays.Location = New-Object System.Drawing.Point(230, 86)
$numUserStaleDays.Size = New-Object System.Drawing.Size(80, 22)
$numUserStaleDays.Minimum = 1
$numUserStaleDays.Maximum = 3650
$numUserStaleDays.Value = 80
$toolTip.SetToolTip($numUserStaleDays, "Number of days a user can be inactive before they're considered stale.")
$grpThresholds.Controls.Add($numUserStaleDays)

# PC Sort
$lblPCSort = New-Object System.Windows.Forms.Label
$lblPCSort.Text = "PC Sort By:"
$lblPCSort.Location = New-Object System.Drawing.Point(15, 118)
$lblPCSort.Size = New-Object System.Drawing.Size(200, 20)
$grpThresholds.Controls.Add($lblPCSort)

$cboPCSort = New-Object System.Windows.Forms.ComboBox
$cboPCSort.Location = New-Object System.Drawing.Point(230, 116)
$cboPCSort.Size = New-Object System.Drawing.Size(130, 22)
$cboPCSort.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$cboPCSort.Items.Add("Name")
[void]$cboPCSort.Items.Add("Last Seen Date")
$cboPCSort.SelectedIndex = 0
$toolTip.SetToolTip($cboPCSort, "Choose how to sort computer entries in the report: by name or by the date last seen on the network.")
$grpThresholds.Controls.Add($cboPCSort)

# User Sort
$lblUserSort = New-Object System.Windows.Forms.Label
$lblUserSort.Text = "User Sort By:"
$lblUserSort.Location = New-Object System.Drawing.Point(15, 148)
$lblUserSort.Size = New-Object System.Drawing.Size(200, 20)
$grpThresholds.Controls.Add($lblUserSort)

$cboUserSort = New-Object System.Windows.Forms.ComboBox
$cboUserSort.Location = New-Object System.Drawing.Point(230, 146)
$cboUserSort.Size = New-Object System.Drawing.Size(130, 22)
$cboUserSort.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$cboUserSort.Items.Add("Name")
[void]$cboUserSort.Items.Add("Last Login Date")
$cboUserSort.SelectedIndex = 0
$toolTip.SetToolTip($cboUserSort, "Choose how to sort user entries in the report: by name or by the date of last login.")
$grpThresholds.Controls.Add($cboUserSort)

# ============================ Options ============================
$grpOptions = New-Object System.Windows.Forms.GroupBox
$grpOptions.Text = "Options"
$grpOptions.Location = New-Object System.Drawing.Point(396, 196)
$grpOptions.Size = New-Object System.Drawing.Size(376, 170)
$tabConfig.Controls.Add($grpOptions)

# Hide Disabled Users
$chkHideDisabledUsers = New-Object System.Windows.Forms.CheckBox
$chkHideDisabledUsers.Text = "Hide Disabled Users"
$chkHideDisabledUsers.Location = New-Object System.Drawing.Point(15, 30)
$chkHideDisabledUsers.Size = New-Object System.Drawing.Size(340, 20)
$chkHideDisabledUsers.Checked = $false
$toolTip.SetToolTip($chkHideDisabledUsers, "When checked, disabled user accounts will not be shown in the report.")
$grpOptions.Controls.Add($chkHideDisabledUsers)

# List ALL Computers
$chkListAllComputers = New-Object System.Windows.Forms.CheckBox
$chkListAllComputers.Text = "List ALL Computers (including servers)"
$chkListAllComputers.Location = New-Object System.Drawing.Point(15, 62)
$chkListAllComputers.Size = New-Object System.Drawing.Size(340, 20)
$chkListAllComputers.Checked = $true
$toolTip.SetToolTip($chkListAllComputers, "When checked, servers will also be displayed in the computers section of the report.")
$grpOptions.Controls.Add($chkListAllComputers)

# Use alternate credentials
$chkUseAltCreds = New-Object System.Windows.Forms.CheckBox
$chkUseAltCreds.Text = "Use alternate credentials (prompt at run)"
$chkUseAltCreds.Location = New-Object System.Drawing.Point(15, 94)
$chkUseAltCreds.Size = New-Object System.Drawing.Size(350, 20)
$chkUseAltCreds.Checked = $false
$toolTip.SetToolTip($chkUseAltCreds, "When checked, you'll be prompted for credentials when running. Use this to report on another domain or from a non-domain-joined machine (set the Domain Controller field too).")
$grpOptions.Controls.Add($chkUseAltCreds)

# ============================ Skip Sections ============================
$grpSkipSections = New-Object System.Windows.Forms.GroupBox
$grpSkipSections.Text = "Skip Sections"
$grpSkipSections.Location = New-Object System.Drawing.Point(12, 374)
$grpSkipSections.Size = New-Object System.Drawing.Size(760, 90)
$tabConfig.Controls.Add($grpSkipSections)

$chkSkipDomainOverview = New-Object System.Windows.Forms.CheckBox
$chkSkipDomainOverview.Text = "Skip Domain Overview"
$chkSkipDomainOverview.Location = New-Object System.Drawing.Point(15, 26)
$chkSkipDomainOverview.Size = New-Object System.Drawing.Size(170, 20)
$chkSkipDomainOverview.Checked = $false
$toolTip.SetToolTip($chkSkipDomainOverview, "When checked, the Domain Health Overview (functional levels, FSMO roles, DC inventory, password policy, trusts) is skipped.")
$grpSkipSections.Controls.Add($chkSkipDomainOverview)

$chkSkipAdmin = New-Object System.Windows.Forms.CheckBox
$chkSkipAdmin.Text = "Skip Admin"
$chkSkipAdmin.Location = New-Object System.Drawing.Point(205, 26)
$chkSkipAdmin.Size = New-Object System.Drawing.Size(140, 20)
$chkSkipAdmin.Checked = $false
$toolTip.SetToolTip($chkSkipAdmin, "When checked, the admin membership section will be skipped in the report.")
$grpSkipSections.Controls.Add($chkSkipAdmin)

$chkSkipServer = New-Object System.Windows.Forms.CheckBox
$chkSkipServer.Text = "Skip Servers"
$chkSkipServer.Location = New-Object System.Drawing.Point(395, 26)
$chkSkipServer.Size = New-Object System.Drawing.Size(140, 20)
$chkSkipServer.Checked = $false
$toolTip.SetToolTip($chkSkipServer, "When checked, the server information section will be skipped. (Server collection uses WMI and can be slow over a WAN.)")
$grpSkipSections.Controls.Add($chkSkipServer)

$chkSkipUsers = New-Object System.Windows.Forms.CheckBox
$chkSkipUsers.Text = "Skip Users"
$chkSkipUsers.Location = New-Object System.Drawing.Point(15, 56)
$chkSkipUsers.Size = New-Object System.Drawing.Size(170, 20)
$chkSkipUsers.Checked = $false
$toolTip.SetToolTip($chkSkipUsers, "When checked, the user information and Security Findings sections will be skipped in the report.")
$grpSkipSections.Controls.Add($chkSkipUsers)

$chkSkipComputers = New-Object System.Windows.Forms.CheckBox
$chkSkipComputers.Text = "Skip Computers"
$chkSkipComputers.Location = New-Object System.Drawing.Point(205, 56)
$chkSkipComputers.Size = New-Object System.Drawing.Size(150, 20)
$chkSkipComputers.Checked = $false
$toolTip.SetToolTip($chkSkipComputers, "When checked, the computer information section will be skipped in the report.")
$grpSkipSections.Controls.Add($chkSkipComputers)

$chkSkipBitlockerStatus = New-Object System.Windows.Forms.CheckBox
$chkSkipBitlockerStatus.Text = "Skip BitLocker"
$chkSkipBitlockerStatus.Location = New-Object System.Drawing.Point(395, 56)
$chkSkipBitlockerStatus.Size = New-Object System.Drawing.Size(150, 20)
$chkSkipBitlockerStatus.Checked = $true
$toolTip.SetToolTip($chkSkipBitlockerStatus, "When checked, the BitLocker recovery-key lookup is skipped (faster). Uncheck to include BitLocker status per computer.")
$grpSkipSections.Controls.Add($chkSkipBitlockerStatus)
