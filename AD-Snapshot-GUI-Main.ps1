#region Assemblies and visual styles
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# Per-monitor friendly text rendering
try { [System.Windows.Forms.Application]::SetHighDpiMode([System.Windows.Forms.HighDpiMode]::SystemAware) | Out-Null } catch { }

# Native cue-banner (grey placeholder) support for text boxes
if (-not ("AdSnap.Native" -as [type])) {
    Add-Type -Namespace AdSnap -Name Native -MemberDefinition @"
[System.Runtime.InteropServices.DllImport("user32.dll", CharSet=System.Runtime.InteropServices.CharSet.Unicode)]
public static extern System.IntPtr SendMessage(System.IntPtr hWnd, int msg, System.IntPtr wParam, string lParam);
"@
}
function Set-CueBanner {
    param($TextBox, [string]$Text)
    try { [void][AdSnap.Native]::SendMessage($TextBox.Handle, 0x1501, [System.IntPtr]1, $Text) } catch { }
}

# Global state
$script:reportPath        = $null
$script:reportFilename    = $null
$script:reportPDFFilename = $null
$script:snapshotJob       = $null
$script:adCredential      = $null

# Shared palette
$clrBg     = [System.Drawing.Color]::FromArgb(244, 246, 249)
$clrPanel  = [System.Drawing.Color]::White
$clrBrand  = [System.Drawing.Color]::FromArgb(31, 111, 235)
$clrBrandD = [System.Drawing.Color]::FromArgb(22, 61, 115)
$clrInk    = [System.Drawing.Color]::FromArgb(31, 41, 51)
$fontUI    = New-Object System.Drawing.Font("Segoe UI", 9)
$fontHead  = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
$fontSub   = New-Object System.Drawing.Font("Segoe UI", 8.5)

# Shared ToolTip used by all tabs
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 8000
$toolTip.InitialDelay = 400
$toolTip.ReshowDelay  = 400
$toolTip.ShowAlways   = $true

#region Form
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Active Directory Snapshot Tool"
$mainForm.ClientSize = New-Object System.Drawing.Size(820, 660)
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$mainForm.MaximizeBox = $false
$mainForm.MinimizeBox = $true
$mainForm.BackColor = $clrBg
$mainForm.Font = $fontUI
$mainForm.Icon = [System.Drawing.SystemIcons]::Shield

#region Menu
$mainMenu = New-Object System.Windows.Forms.MenuStrip
$mainMenu.BackColor = $clrPanel
$mainForm.MainMenuStrip = $mainMenu
$mainForm.Controls.Add($mainMenu)

$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "&File"
[void]$mainMenu.Items.Add($fileMenu)

$saveSettingsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$saveSettingsMenuItem.Text = "Save Settings"
$saveSettingsMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::S
[void]$fileMenu.DropDownItems.Add($saveSettingsMenuItem)

$loadSettingsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$loadSettingsMenuItem.Text = "Load Settings"
$loadSettingsMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::O
[void]$fileMenu.DropDownItems.Add($loadSettingsMenuItem)

[void]$fileMenu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))

$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitMenuItem.Text = "Exit"
[void]$fileMenu.DropDownItems.Add($exitMenuItem)

$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$helpMenu.Text = "&Help"
[void]$mainMenu.Items.Add($helpMenu)

$aboutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$aboutMenuItem.Text = "About"
[void]$helpMenu.DropDownItems.Add($aboutMenuItem)

#region Header banner
$header = New-Object System.Windows.Forms.Panel
$header.Location = New-Object System.Drawing.Point(0, 24)
$header.Size = New-Object System.Drawing.Size(820, 66)
$header.BackColor = $clrBrandD
$mainForm.Controls.Add($header)

$header.Add_Paint({
    param($sender, $e)
    $rect = $sender.ClientRectangle
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, $clrBrandD, $clrBrand, [System.Drawing.Drawing2D.LinearGradientMode]::Horizontal)
    $e.Graphics.FillRectangle($brush, $rect)
    $brush.Dispose()
})

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Active Directory Snapshot"
$lblTitle.Font = $fontHead
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.BackColor = [System.Drawing.Color]::Transparent
$lblTitle.UseMnemonic = $false
$lblTitle.Location = New-Object System.Drawing.Point(18, 10)
$lblTitle.AutoSize = $true
$header.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = "Generate professional AD inventory & health reports"
$lblSubtitle.Font = $fontSub
$lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(210, 224, 245)
$lblSubtitle.BackColor = [System.Drawing.Color]::Transparent
$lblSubtitle.UseMnemonic = $false
$lblSubtitle.Location = New-Object System.Drawing.Point(20, 40)
$lblSubtitle.AutoSize = $true
$header.Controls.Add($lblSubtitle)

#region Tabs
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(12, 100)
$tabControl.Size = New-Object System.Drawing.Size(796, 510)
$tabControl.Font = $fontUI
$mainForm.Controls.Add($tabControl)

$tabConfig = New-Object System.Windows.Forms.TabPage
$tabConfig.Text = "  Configuration  "
$tabConfig.BackColor = $clrPanel
$tabControl.Controls.Add($tabConfig)

$tabOutput = New-Object System.Windows.Forms.TabPage
$tabOutput.Text = "  Output Options  "
$tabOutput.BackColor = $clrPanel
$tabControl.Controls.Add($tabOutput)

$tabRun = New-Object System.Windows.Forms.TabPage
$tabRun.Text = "  Run Report  "
$tabRun.BackColor = $clrPanel
$tabControl.Controls.Add($tabRun)

#region Helper to style flat buttons
function Set-FlatButton {
    param($Button, [bool]$Primary = $false)
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.Font = $fontUI
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    if ($Primary) {
        $Button.BackColor = $clrBrand
        $Button.ForeColor = [System.Drawing.Color]::White
    } else {
        $Button.BackColor = [System.Drawing.Color]::FromArgb(225, 230, 237)
        $Button.ForeColor = $clrInk
    }
}

#region Bottom buttons
$btnSaveSettings = New-Object System.Windows.Forms.Button
$btnSaveSettings.Text = "Save Settings"
$btnSaveSettings.Location = New-Object System.Drawing.Point(12, 622)
$btnSaveSettings.Size = New-Object System.Drawing.Size(120, 30)
Set-FlatButton $btnSaveSettings
$mainForm.Controls.Add($btnSaveSettings)

$btnLoadSettings = New-Object System.Windows.Forms.Button
$btnLoadSettings.Text = "Load Settings"
$btnLoadSettings.Location = New-Object System.Drawing.Point(140, 622)
$btnLoadSettings.Size = New-Object System.Drawing.Size(120, 30)
Set-FlatButton $btnLoadSettings
$mainForm.Controls.Add($btnLoadSettings)

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Text = "Exit"
$btnExit.Location = New-Object System.Drawing.Point(688, 622)
$btnExit.Size = New-Object System.Drawing.Size(120, 30)
Set-FlatButton $btnExit
$mainForm.Controls.Add($btnExit)

#region Source tab content and functions (these ADD controls to the tabs created above)
. "$PSScriptRoot\AD-Snapshot-GUI.ps1"
. "$PSScriptRoot\AD-Snapshot-GUI-Output.ps1"
. "$PSScriptRoot\AD-Snapshot-GUI-Run.ps1"
. "$PSScriptRoot\AD-Snapshot-GUI-Functions.ps1"

# Style the action buttons created in the Run tab
Set-FlatButton $btnRunReport $true
Set-FlatButton $btnViewReport
Set-FlatButton $btnOpenReportFolder
Set-FlatButton $btnCancel
if ($btnTestConnection) { Set-FlatButton $btnTestConnection }

#region Event Handlers
$btnBrowseOutputPath.Add_Click({
    $selectedPath = Get-FolderPath -Description "Select Output Folder"
    if ($selectedPath) { $txtOutputPath.Text = $selectedPath }
})

$btnRunReport.Add_Click({ Start-ADSnapshot })
$btnViewReport.Add_Click({ Show-Report })
$btnOpenReportFolder.Add_Click({ Open-ReportFolder })
$btnCancel.Add_Click({ Stop-RunningJob })
if ($btnTestConnection) { $btnTestConnection.Add_Click({ Test-ADConnection }) }

$btnSaveSettings.Add_Click({ Save-Settings })
$saveSettingsMenuItem.Add_Click({ Save-Settings })
$btnLoadSettings.Add_Click({ Import-AppSettings; Update-DependentControls })
$loadSettingsMenuItem.Add_Click({ Import-AppSettings; Update-DependentControls })

# Enable/disable dependent fields as their parent checkboxes change
$chkCreateFile.Add_CheckedChanged({ Update-DependentControls })
$chkExportCSV.Add_CheckedChanged({ Update-DependentControls })
$chkWantPDFFile.Add_CheckedChanged({ Update-DependentControls })
$chkSendEmail.Add_CheckedChanged({ Update-DependentControls })

$btnExit.Add_Click({ $mainForm.Close() })
$exitMenuItem.Add_Click({ $mainForm.Close() })

$aboutMenuItem.Add_Click({
    [System.Windows.Forms.MessageBox]::Show(
        "Active Directory Snapshot Tool`r`n`r`nVersion 2.0.0`r`n`r`nA GUI front end for the AD-Snapshot PowerShell report engine.`r`nCollects admin, server, user and computer inventory with a modern, self-contained HTML report.`r`n`r`nAuthor: Bryant Welch  -  MIT License",
        "About AD Snapshot Tool",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
})

$mainForm.Add_Load({
    Import-AppSettings
    # Grey placeholder hints for fields that can be auto-detected or scoped.
    Set-CueBanner $txtOUList           "Auto-detect entire domain, or enter OU names / DNs separated by commas"
    Set-CueBanner $txtRealm            "Auto-detect domain name"
    Set-CueBanner $txtDomainController "Auto-detect domain controller"
    Set-CueBanner $txtDomainCN         "Auto-detect domain DN (e.g. DC=contoso,DC=com)"
    Set-CueBanner $txtOutputPath       "Defaults to Desktop"
    Set-CueBanner $txtFromEmail        "Example: reports@contoso.com"
    Set-CueBanner $txtToEmail          "Example: admin@contoso.com"
    Set-CueBanner $txtCcList           "Example: security@contoso.com, helpdesk@contoso.com"
    Set-CueBanner $txtSmtpServer       "Example: smtp.contoso.com"
    # Reflect loaded settings in dependent control states
    Update-DependentControls
})

[void]$mainForm.ShowDialog()
