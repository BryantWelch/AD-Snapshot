#region Import Modules and Add Windows Forms Assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Global variables
$script:reportPath = $null
$script:reportFilename = $null
$script:reportPDFFilename = $null

#region Form Creation
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Active Directory Snapshot Tool"
$mainForm.Size = New-Object System.Drawing.Size(800, 600)
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$mainForm.MaximizeBox = $false
$mainForm.MinimizeBox = $true
$mainForm.Icon = [System.Drawing.SystemIcons]::Shield

# Create main menu
$mainMenu = New-Object System.Windows.Forms.MenuStrip
$mainForm.MainMenuStrip = $mainMenu
$mainForm.Controls.Add($mainMenu)

# File menu
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "File"
$mainMenu.Items.Add($fileMenu)

# Save Settings menu item
$saveSettingsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$saveSettingsMenuItem.Text = "Save Settings"
$saveSettingsMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::S
$fileMenu.DropDownItems.Add($saveSettingsMenuItem)

# Load Settings menu item
$loadSettingsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$loadSettingsMenuItem.Text = "Load Settings"
$loadSettingsMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::O
$fileMenu.DropDownItems.Add($loadSettingsMenuItem)

# Separator
$separator = New-Object System.Windows.Forms.ToolStripSeparator
$fileMenu.DropDownItems.Add($separator)

# Exit menu item
$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitMenuItem.Text = "Exit"
$exitMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::F4
$fileMenu.DropDownItems.Add($exitMenuItem)

# Help menu
$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$helpMenu.Text = "Help"
$mainMenu.Items.Add($helpMenu)

# About menu item
$aboutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$aboutMenuItem.Text = "About"
$helpMenu.DropDownItems.Add($aboutMenuItem)

# Create Tab Control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 30)
$tabControl.Size = New-Object System.Drawing.Size(770, 480)
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

# Bottom buttons
$btnSaveSettings = New-Object System.Windows.Forms.Button
$btnSaveSettings.Text = "Save Settings"
$btnSaveSettings.Location = New-Object System.Drawing.Point(10, 520)
$btnSaveSettings.Size = New-Object System.Drawing.Size(120, 30)
$mainForm.Controls.Add($btnSaveSettings)

$btnLoadSettings = New-Object System.Windows.Forms.Button
$btnLoadSettings.Text = "Load Settings"
$btnLoadSettings.Location = New-Object System.Drawing.Point(140, 520)
$btnLoadSettings.Size = New-Object System.Drawing.Size(120, 30)
$mainForm.Controls.Add($btnLoadSettings)

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Text = "Exit"
$btnExit.Location = New-Object System.Drawing.Point(660, 520)
$btnExit.Size = New-Object System.Drawing.Size(120, 30)
$mainForm.Controls.Add($btnExit)

# Source the tab content files
. "$PSScriptRoot\AD-Snapshot-GUI.ps1"
. "$PSScriptRoot\AD-Snapshot-GUI-Output.ps1"
. "$PSScriptRoot\AD-Snapshot-GUI-Run.ps1"
. "$PSScriptRoot\AD-Snapshot-GUI-Functions.ps1"

#region Event Handlers
# Browse Output Path button click
$btnBrowseOutputPath.Add_Click({
    $selectedPath = Get-FolderPath -Description "Select Output Folder"
    if ($selectedPath) {
        $txtOutputPath.Text = $selectedPath
    }
})

# Browse PDF Converter button click
$btnBrowsePDFConverter.Add_Click({
    $selectedFile = Get-FilePath -Filter "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*" -Title "Select PDF Converter"
    if ($selectedFile) {
        $txtPDFConverter.Text = $selectedFile
    }
})

# Run Report button click
$btnRunReport.Add_Click({
    Start-ADSnapshot
})

# View Report button click
$btnViewReport.Add_Click({
    Show-Report
})

# Open Report Folder button click
$btnOpenReportFolder.Add_Click({
    Open-ReportFolder
})

# Cancel button click
$btnCancel.Add_Click({
    Stop-RunningJob
})

# Save Settings button and menu item click
$btnSaveSettings.Add_Click({
    Save-Settings
})
$saveSettingsMenuItem.Add_Click({
    Save-Settings
})

# Load Settings button and menu item click
$btnLoadSettings.Add_Click({
    Import-AppSettings
})
$loadSettingsMenuItem.Add_Click({
    Import-AppSettings
})

# Exit button and menu item click
$btnExit.Add_Click({
    $mainForm.Close()
})
$exitMenuItem.Add_Click({
    $mainForm.Close()
})

# About menu item click
$aboutMenuItem.Add_Click({
    [System.Windows.Forms.MessageBox]::Show(
        "Active Directory Snapshot Tool`r`n`r`nVersion 1.0.0`r`n`r`nThis tool provides a GUI interface for the AD-Snapshot PowerShell script.`r`n`r`nOriginal script by Bryant Welch",
        "About AD Snapshot Tool",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
})

# Form Load event
$mainForm.Add_Load({
    # Try to load settings if they exist
    Import-AppSettings
})

# Show the form
[void]$mainForm.ShowDialog()
