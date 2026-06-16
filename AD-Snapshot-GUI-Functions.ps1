#region Event Handlers and Functions

# Append a timestamped line to the status box (thread-safe-ish for the UI thread)
function Write-StatusLog {
    param ([string]$Message)
    if ([string]::IsNullOrWhiteSpace($Message)) { return }
    $txtStatus.AppendText("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message`r`n")
    $txtStatus.SelectionStart = $txtStatus.Text.Length
    $txtStatus.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# Browse for a folder
function Get-FolderPath {
    param (
        [string]$Description = "Select Folder",
        [string]$InitialDirectory = [Environment]::GetFolderPath("Desktop")
    )
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $Description
    $folderBrowser.SelectedPath = $InitialDirectory
    if ($folderBrowser.ShowDialog() -eq "OK") { return $folderBrowser.SelectedPath }
    return $null
}

# Browse for a file
function Get-FilePath {
    param (
        [string]$Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*",
        [string]$Title = "Select File",
        [string]$InitialDirectory = [Environment]::GetFolderPath("ProgramFiles")
    )
    $fileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $fileBrowser.Filter = $Filter
    $fileBrowser.Title = $Title
    $fileBrowser.InitialDirectory = $InitialDirectory
    if ($fileBrowser.ShowDialog() -eq "OK") { return $fileBrowser.FileName }
    return $null
}

# Resolve the configured output folder (blank = Desktop)
function Get-OutputFolder {
    if ([string]::IsNullOrWhiteSpace($txtOutputPath.Text)) {
        return [Environment]::GetFolderPath("Desktop")
    }
    return $txtOutputPath.Text
}

# Locate an installed Chromium browser (Edge/Chrome) for headless PDF rendering
function Get-PdfBrowser {
    $candidates = @(
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe')
        (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe')
        (Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe')
        (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe')
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\Application\msedge.exe')
    )
    foreach ($c in $candidates) { if ($c -and (Test-Path $c)) { return $c } }
    return $null
}

# Enable/disable controls based on which parent options are checked,
# so the form only offers fields that are actually in use.
function Update-DependentControls {
    # Output-path controls are relevant when an HTML file and/or CSV is being written
    $wantFiles = $chkCreateFile.Checked -or $chkExportCSV.Checked
    foreach ($c in @($lblOutputPath, $txtOutputPath, $btnBrowseOutputPath)) { $c.Enabled = $wantFiles }

    # Keep email fields enabled so cue-banner placeholders stay visible and
    # users can prefill settings before turning email delivery on.
    $pdfOn   = $chkWantPDFFile.Checked
    $emailOn = $chkSendEmail.Checked
    foreach ($c in @($lblFromEmail, $txtFromEmail, $lblToEmail, $txtToEmail, $lblCcList, $txtCcList, $lblSmtpServer, $txtSmtpServer)) { $c.Enabled = $true }

    # Attaching a PDF requires both a PDF and an email
    $chkAttachPDF.Enabled = ($pdfOn -and $emailOn)
}

# Validate required fields before running. Returns $false to abort.
function Test-RequiredFields {
    $missing  = @()   # hard stops
    $problems = @()   # soft warnings (user may continue)
    $emailRx  = '^[^@\s]+@[^@\s]+\.[^@\s]+$'

    if ($chkSendEmail.Checked) {
        if ([string]::IsNullOrWhiteSpace($txtSmtpServer.Text)) { $missing += "SMTP Server" }
        if ([string]::IsNullOrWhiteSpace($txtFromEmail.Text))  { $missing += "From Email" }
        if ([string]::IsNullOrWhiteSpace($txtToEmail.Text))    { $missing += "To Email" }

        if ($txtFromEmail.Text -and $txtFromEmail.Text -notmatch $emailRx) {
            $problems += "From Email '$($txtFromEmail.Text)' doesn't look like a valid address."
        }
        foreach ($addr in @($txtToEmail.Text -split '\s*,\s*' | Where-Object { $_ })) {
            if ($addr -notmatch $emailRx) { $problems += "To address '$addr' doesn't look valid." }
        }
        foreach ($addr in @($txtCcList.Text -split '\s*,\s*' | Where-Object { $_ })) {
            if ($addr -notmatch $emailRx) { $problems += "CC address '$addr' doesn't look valid." }
        }
    }

    if ($chkWantPDFFile.Checked -and -not (Get-PdfBrowser)) {
        $problems += "PDF is enabled but no Microsoft Edge / Google Chrome was found - the PDF will be skipped (the HTML report is still produced)."
    }

    if ($chkAttachPDF.Checked -and -not $chkWantPDFFile.Checked) {
        $problems += "'Attach PDF to Email' is on but 'Create PDF Report' is off - nothing will be attached."
    }

    if (($chkCreateFile.Checked -or $chkExportCSV.Checked) -and
        -not [string]::IsNullOrWhiteSpace($txtOutputPath.Text) -and
        -not (Test-Path -Path $txtOutputPath.Text)) {
        $problems += "Output path '$($txtOutputPath.Text)' does not exist."
    }

    if (-not $chkCreateFile.Checked -and -not $chkExportCSV.Checked -and -not $chkSendEmail.Checked) {
        $problems += "No output is selected (file, CSV, or email) - the report will run but nothing will be saved or sent."
    }

    if ($missing.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please complete the following required field(s):`r`n`r`n - " + ($missing -join "`r`n - "),
            "Missing Configuration",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return $false
    }

    if ($problems.Count -gt 0) {
        $res = [System.Windows.Forms.MessageBox]::Show(
            "Possible issues were found:`r`n`r`n - " + ($problems -join "`r`n - ") + "`r`n`r`nContinue anyway?",
            "Review Configuration",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) { return $false }
    }

    return $true
}

# Save settings to XML
function Save-Settings {
    param ([string]$FilePath = "$env:APPDATA\AD-Snapshot-GUI\settings.xml")
    $directory = Split-Path -Path $FilePath -Parent
    if (-not (Test-Path -Path $directory)) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }

    $settings = @{
        OUList              = $txtOUList.Text
        Realm               = $txtRealm.Text
        AdminGroup          = $txtAdminGroup.Text
        DomainController    = $txtDomainController.Text
        DomainCN            = $txtDomainCN.Text
        ServerUpTimeAlarm   = $numServerUpTimeAlarm.Value
        ComputerStaleDays   = $numComputerStaleDays.Value
        UserStaleDays       = $numUserStaleDays.Value
        PCSort              = $cboPCSort.SelectedIndex
        UserSort            = $cboUserSort.SelectedIndex
        HideDisabledUsers   = $chkHideDisabledUsers.Checked
        ListAllComputers    = $chkListAllComputers.Checked
        SkipDomainOverview  = $chkSkipDomainOverview.Checked
        SkipAdmin           = $chkSkipAdmin.Checked
        SkipServer          = $chkSkipServer.Checked
        SkipUsers           = $chkSkipUsers.Checked
        SkipComputers       = $chkSkipComputers.Checked
        SkipBitlockerStatus = $chkSkipBitlockerStatus.Checked
        UseAltCreds         = $chkUseAltCreds.Checked
        CreateFile          = $chkCreateFile.Checked
        ExportCSV           = $chkExportCSV.Checked
        OutputPath          = $txtOutputPath.Text
        WantPDFFile         = $chkWantPDFFile.Checked
        SendEmail           = $chkSendEmail.Checked
        FromEmail           = $txtFromEmail.Text
        ToEmail             = $txtToEmail.Text
        CcList              = $txtCcList.Text
        SmtpServer          = $txtSmtpServer.Text
        AttachPDF           = $chkAttachPDF.Checked
    }
    $settings | Export-Clixml -Path $FilePath
    Write-StatusLog "Settings saved to $FilePath"
}

# Load settings from XML
function Import-AppSettings {
    param ([string]$FilePath = "$env:APPDATA\AD-Snapshot-GUI\settings.xml")
    if (-not (Test-Path -Path $FilePath)) { Write-StatusLog "No saved settings found (using defaults)."; return }

    $settings = Import-Clixml -Path $FilePath
    if ($settings.OUList)           { $txtOUList.Text = $settings.OUList }
    if ($settings.Realm)            { $txtRealm.Text = $settings.Realm }
    if ($settings.AdminGroup)       { $txtAdminGroup.Text = $settings.AdminGroup }
    if ($settings.DomainController) { $txtDomainController.Text = $settings.DomainController }
    if ($settings.DomainCN)         { $txtDomainCN.Text = $settings.DomainCN }
    if ($null -ne $settings.ServerUpTimeAlarm) { $numServerUpTimeAlarm.Value = $settings.ServerUpTimeAlarm }
    if ($null -ne $settings.ComputerStaleDays) { $numComputerStaleDays.Value = $settings.ComputerStaleDays }
    if ($null -ne $settings.UserStaleDays)     { $numUserStaleDays.Value = $settings.UserStaleDays }
    if ($null -ne $settings.PCSort)            { $cboPCSort.SelectedIndex = $settings.PCSort }
    if ($null -ne $settings.UserSort)          { $cboUserSort.SelectedIndex = $settings.UserSort }
    if ($null -ne $settings.HideDisabledUsers) { $chkHideDisabledUsers.Checked = $settings.HideDisabledUsers }
    if ($null -ne $settings.ListAllComputers)  { $chkListAllComputers.Checked = $settings.ListAllComputers }
    if ($null -ne $settings.SkipDomainOverview){ $chkSkipDomainOverview.Checked = $settings.SkipDomainOverview }
    if ($null -ne $settings.SkipAdmin)         { $chkSkipAdmin.Checked = $settings.SkipAdmin }
    if ($null -ne $settings.SkipServer)        { $chkSkipServer.Checked = $settings.SkipServer }
    if ($null -ne $settings.SkipUsers)         { $chkSkipUsers.Checked = $settings.SkipUsers }
    if ($null -ne $settings.SkipComputers)     { $chkSkipComputers.Checked = $settings.SkipComputers }
    if ($null -ne $settings.SkipBitlockerStatus) { $chkSkipBitlockerStatus.Checked = $settings.SkipBitlockerStatus }
    if ($null -ne $settings.UseAltCreds)       { $chkUseAltCreds.Checked = $settings.UseAltCreds }
    if ($null -ne $settings.CreateFile)        { $chkCreateFile.Checked = $settings.CreateFile }
    if ($null -ne $settings.ExportCSV)         { $chkExportCSV.Checked = $settings.ExportCSV }
    if ($settings.OutputPath)                  { $txtOutputPath.Text = $settings.OutputPath }
    if ($null -ne $settings.WantPDFFile)       { $chkWantPDFFile.Checked = $settings.WantPDFFile }
    if ($null -ne $settings.SendEmail)         { $chkSendEmail.Checked = $settings.SendEmail }
    if ($settings.FromEmail)                   { $txtFromEmail.Text = $settings.FromEmail }
    if ($settings.ToEmail)                     { $txtToEmail.Text = $settings.ToEmail }
    if ($settings.CcList)                      { $txtCcList.Text = $settings.CcList }
    if ($settings.SmtpServer)                  { $txtSmtpServer.Text = $settings.SmtpServer }
    if ($null -ne $settings.AttachPDF)         { $chkAttachPDF.Checked = $settings.AttachPDF }

    Write-StatusLog "Settings loaded from $FilePath"
}

# Process one line of job output: detect the report-path marker, otherwise log it
function Write-JobLine {
    param([string]$Line)
    if ([string]::IsNullOrWhiteSpace($Line)) { return }
    if ($Line -like '*REPORTFILE::*') {
        $p = $Line.Substring($Line.IndexOf('REPORTFILE::') + 12).Trim()
        if ($p) {
            $script:reportPath        = Split-Path -Path $p -Parent
            $script:reportFilename    = Split-Path -Path $p -Leaf
            $script:reportPDFFilename = [System.IO.Path]::ChangeExtension($script:reportFilename, '.pdf')
            $btnViewReport.Enabled = $true
            $btnOpenReportFolder.Enabled = $true
            Write-StatusLog "Report saved: $p"
        }
        return
    }
    Write-StatusLog $Line
}

# Build the parameter hashtable passed to AD-Snapshot.ps1
function Get-SnapshotParameters {
    $ouArray = @($txtOUList.Text -split '\s*,\s*' | Where-Object { $_ -ne "" })
    $p = @{
        OUlist             = $ouArray
        realm              = $txtRealm.Text
        DomainCN           = $txtDomainCN.Text
        Server             = $txtDomainController.Text
        AdminGroup         = $txtAdminGroup.Text
        CreateFile         = $(if ($chkCreateFile.Checked) { "Y" } else { "N" })
        outputpath         = $txtOutputPath.Text
        WantPDFFile        = $(if ($chkWantPDFFile.Checked) { "true" } else { "false" })
        ExportCSV          = $(if ($chkExportCSV.Checked) { "true" } else { "false" })
        SendEmail          = $(if ($chkSendEmail.Checked) { "Y" } else { "N" })
        fromemail          = $txtFromEmail.Text
        defaultSentTo      = $txtToEmail.Text
        CcList             = $txtCcList.Text
        smtpserver         = $txtSmtpServer.Text
        ServerUpTimeAlarm  = [int]$numServerUpTimeAlarm.Value
        ComputerStaleDays  = [int]$numComputerStaleDays.Value
        UserStaleDays      = [int]$numUserStaleDays.Value
        PCsort             = $(if ($cboPCSort.SelectedIndex -eq 0) { "name" } else { "date" })
        Usersort           = $(if ($cboUserSort.SelectedIndex -eq 0) { "name" } else { "date" })
        HideDisabledUsers  = $(if ($chkHideDisabledUsers.Checked) { "true" } else { "false" })
        listALLComputers   = $(if ($chkListAllComputers.Checked) { "true" } else { "false" })
        skipDomainOverview = $(if ($chkSkipDomainOverview.Checked) { "true" } else { "false" })
        skipadmin          = $(if ($chkSkipAdmin.Checked) { "true" } else { "false" })
        skipServer         = $(if ($chkSkipServer.Checked) { "true" } else { "false" })
        skipUsers          = $(if ($chkSkipUsers.Checked) { "true" } else { "false" })
        skipComputers      = $(if ($chkSkipComputers.Checked) { "true" } else { "false" })
        skipBitlockerStatus = $(if ($chkSkipBitlockerStatus.Checked) { "true" } else { "false" })
    }
    if ($script:adCredential) { $p['Credential'] = $script:adCredential }
    return $p
}

# Run AD-Snapshot.ps1 as a background job using the GUI configuration
function Start-ADSnapshot {
    if (-not (Test-RequiredFields)) { return }

    $script:reportPath        = $null
    $script:reportFilename    = $null
    $script:reportPDFFilename = $null

    $btnRunReport.Enabled = $false
    $btnViewReport.Enabled = $false
    $btnOpenReportFolder.Enabled = $false
    $btnCancel.Enabled = $true
    # Marquee = honest "working" indicator (a background job can't report a real %)
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $progressBar.MarqueeAnimationSpeed = 30

    $txtStatus.Text = ""
    Write-StatusLog "Starting AD Snapshot Report..."

    try {
        $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "AD-Snapshot.ps1"
        if (-not (Test-Path $scriptPath)) { Write-StatusLog "ERROR: Cannot find AD-Snapshot.ps1 at $scriptPath"; return }

        # Prompt for alternate credentials when requested (for cross-domain / non-joined hosts)
        if ($chkUseAltCreds -and $chkUseAltCreds.Checked -and -not $script:adCredential) {
            $script:adCredential = Get-Credential -Message "Enter credentials for the target domain"
            if (-not $script:adCredential) { Write-StatusLog "Credential prompt cancelled - aborting."; return }
        }
        if (-not ($chkUseAltCreds -and $chkUseAltCreds.Checked)) { $script:adCredential = $null }

        $params = Get-SnapshotParameters
        $ouLabel = if ($params.OUlist.Count -gt 0) { $params.OUlist -join ', ' } else { 'ENTIRE DOMAIN (auto)' }
        $dcLabel = if ($params.Server) { $params.Server } else { 'auto-detect' }
        Write-StatusLog "Scope: $ouLabel  |  DC: $dcLabel"

        $script:snapshotJob = Start-Job -ScriptBlock {
            param($scriptPath, $params)
            & $scriptPath @params *>&1
        } -ArgumentList $scriptPath, $params

        $jobRunning = $true

        while ($jobRunning) {
            Start-Sleep -Milliseconds 400
            [System.Windows.Forms.Application]::DoEvents()

            # Stream incremental output as the job runs
            $incoming = Receive-Job -Id $script:snapshotJob.Id
            foreach ($line in $incoming) { Write-JobLine ([string]$line) }

            $jobState = (Get-Job -Id $script:snapshotJob.Id).State

            if ($jobState -eq "Completed") {
                $jobRunning = $false

                $tail = Receive-Job -Id $script:snapshotJob.Id
                foreach ($line in $tail) { Write-JobLine ([string]$line) }

                $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
                $progressBar.Value = 100

                if ($chkCreateFile.Checked -and -not $script:reportFilename) {
                    Write-StatusLog "Note: no report file path was reported (it may have been skipped or failed to write)."
                }
                Write-StatusLog "AD Snapshot Report completed successfully!"
            }
            elseif ($jobState -eq "Failed") {
                $jobRunning = $false
                $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
                $progressBar.Value = 0
                $jobError = (Get-Job -Id $script:snapshotJob.Id).ChildJobs[0].JobStateInfo.Reason
                Write-StatusLog "Error: $jobError"
                Write-StatusLog "AD Snapshot Report failed!"
            }
        }

        Remove-Job -Id $script:snapshotJob.Id -Force
        $script:snapshotJob = $null
    }
    catch {
        Write-StatusLog "Error: $_"
    }
    finally {
        $btnRunReport.Enabled = $true
        $btnCancel.Enabled = $false
    }
}

# Quick connectivity / prerequisites check
function Test-ADConnection {
    $tabControl.SelectedTab = $tabRun
    Write-StatusLog "----- Testing Active Directory connectivity -----"

    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-StatusLog "ActiveDirectory module: AVAILABLE"
    }
    catch {
        Write-StatusLog "ActiveDirectory module: NOT available. Install RSAT:"
        Write-StatusLog "  Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
        return
    }

    $dc = $txtDomainController.Text
    if ([string]::IsNullOrWhiteSpace($dc)) { Write-StatusLog "No domain controller configured."; return }

    if (Test-Connection -ComputerName $dc -Count 1 -Quiet -ErrorAction SilentlyContinue) {
        Write-StatusLog "Ping $dc : reachable"
    } else {
        Write-StatusLog "Ping $dc : no response (ICMP may be blocked)"
    }

    try {
        $root = Get-ADRootDSE -Server $dc -ErrorAction Stop
        Write-StatusLog "LDAP bind to $dc : OK"
        Write-StatusLog "  Default naming context: $($root.defaultNamingContext)"
    }
    catch {
        Write-StatusLog "LDAP bind to $dc FAILED: $($_.Exception.Message)"
    }
    Write-StatusLog "----- Connectivity test complete -----"
}

# Open the generated HTML report
function Show-Report {
    if ($script:reportPath -and $script:reportFilename) {
        $reportFullPath = Join-Path -Path $script:reportPath -ChildPath $script:reportFilename
        if (Test-Path -Path $reportFullPath) {
            Start-Process $reportFullPath
            Write-StatusLog "Opening report $reportFullPath"
        } else {
            Write-StatusLog "Report file not found at $reportFullPath"
        }
    } else {
        Write-StatusLog "No report file available to view"
    }
}

# Open the report output folder
function Open-ReportFolder {
    if ($script:reportPath -and (Test-Path -Path $script:reportPath)) {
        Start-Process $script:reportPath
        Write-StatusLog "Opening folder $script:reportPath"
    } else {
        Write-StatusLog "No report folder available to open"
    }
}

# Cancel a running job
function Stop-RunningJob {
    if ($script:snapshotJob) {
        Stop-Job -Id $script:snapshotJob.Id -ErrorAction SilentlyContinue
        Remove-Job -Id $script:snapshotJob.Id -Force -ErrorAction SilentlyContinue
        $script:snapshotJob = $null
        Write-StatusLog "Job cancelled"
    }
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $progressBar.Value = 0
    $btnRunReport.Enabled = $true
    $btnCancel.Enabled = $false
}

#endregion
