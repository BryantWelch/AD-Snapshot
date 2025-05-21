#region Event Handlers and Functions

# Function to log messages to the status textbox
function Write-StatusLog {
    param (
        [string]$Message
    )
    
    $txtStatus.AppendText("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message`r`n")
    $txtStatus.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# Function to browse for a folder
function Get-FolderPath {
    param (
        [string]$Description = "Select Folder",
        [string]$InitialDirectory = [Environment]::GetFolderPath("Desktop")
    )
    
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $Description
    $folderBrowser.SelectedPath = $InitialDirectory
    
    if ($folderBrowser.ShowDialog() -eq "OK") {
        return $folderBrowser.SelectedPath
    }
    return $null
}

# Function to browse for a file
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
    
    if ($fileBrowser.ShowDialog() -eq "OK") {
        return $fileBrowser.FileName
    }
    return $null
}

# Function to save settings to XML file
function Save-Settings {
    param (
        [string]$FilePath = "$env:APPDATA\AD-Snapshot-GUI\settings.xml"
    )
    
    # Create directory if it doesn't exist
    $directory = Split-Path -Path $FilePath -Parent
    if (-not (Test-Path -Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    
    $settings = @{
        OUList = $txtOUList.Text
        Realm = $txtRealm.Text
        AdminGroup = $txtAdminGroup.Text
        DomainController = $txtDomainController.Text
        DomainCN = $txtDomainCN.Text
        ServerUpTimeAlarm = $numServerUpTimeAlarm.Value
        ComputerStaleDays = $numComputerStaleDays.Value
        UserStaleDays = $numUserStaleDays.Value
        PCSort = $cboPCSort.SelectedIndex
        UserSort = $cboUserSort.SelectedIndex
        HideDisabledUsers = $chkHideDisabledUsers.Checked
        ListAllComputers = $chkListAllComputers.Checked
        SkipAdmin = $chkSkipAdmin.Checked
        SkipServer = $chkSkipServer.Checked
        SkipUsers = $chkSkipUsers.Checked
        SkipComputers = $chkSkipComputers.Checked
        SkipBitlockerStatus = $chkSkipBitlockerStatus.Checked
        CreateFile = $chkCreateFile.Checked
        OutputPath = $txtOutputPath.Text
        WantPDFFile = $chkWantPDFFile.Checked
        PDFConverter = $txtPDFConverter.Text
        SendEmail = $chkSendEmail.Checked
        FromEmail = $txtFromEmail.Text
        ToEmail = $txtToEmail.Text
        CcList = $txtCcList.Text
        SmtpServer = $txtSmtpServer.Text
        AttachPDF = $chkAttachPDF.Checked
    }
    
    $settings | Export-Clixml -Path $FilePath
    Write-StatusLog "Settings saved to $FilePath"
}

# Function to load settings from XML file
function Import-AppSettings {
    param (
        [string]$FilePath = "$env:APPDATA\AD-Snapshot-GUI\settings.xml"
    )
    
    if (Test-Path -Path $FilePath) {
        $settings = Import-Clixml -Path $FilePath
        
        $txtOUList.Text = $settings.OUList
        $txtRealm.Text = $settings.Realm
        $txtAdminGroup.Text = $settings.AdminGroup
        if ($settings.DomainController) { $txtDomainController.Text = $settings.DomainController }
        if ($settings.DomainCN) { $txtDomainCN.Text = $settings.DomainCN }
        $numServerUpTimeAlarm.Value = $settings.ServerUpTimeAlarm
        $numComputerStaleDays.Value = $settings.ComputerStaleDays
        $numUserStaleDays.Value = $settings.UserStaleDays
        $cboPCSort.SelectedIndex = $settings.PCSort
        $cboUserSort.SelectedIndex = $settings.UserSort
        $chkHideDisabledUsers.Checked = $settings.HideDisabledUsers
        $chkListAllComputers.Checked = $settings.ListAllComputers
        $chkSkipAdmin.Checked = $settings.SkipAdmin
        $chkSkipServer.Checked = $settings.SkipServer
        $chkSkipUsers.Checked = $settings.SkipUsers
        $chkSkipComputers.Checked = $settings.SkipComputers
        $chkSkipBitlockerStatus.Checked = $settings.SkipBitlockerStatus
        $chkCreateFile.Checked = $settings.CreateFile
        $txtOutputPath.Text = $settings.OutputPath
        $chkWantPDFFile.Checked = $settings.WantPDFFile
        $txtPDFConverter.Text = $settings.PDFConverter
        $chkSendEmail.Checked = $settings.SendEmail
        $txtFromEmail.Text = $settings.FromEmail
        $txtToEmail.Text = $settings.ToEmail
        $txtCcList.Text = $settings.CcList
        $txtSmtpServer.Text = $settings.SmtpServer
        $chkAttachPDF.Checked = $settings.AttachPDF
        
        Write-StatusLog "Settings loaded from $FilePath"
    }
    else {
        Write-StatusLog "No settings file found at $FilePath"
    }
}

# Function to run the AD Snapshot script with parameters from the GUI
function Start-ADSnapshot {
    $script:reportPath = $null
    $script:reportFilename = $null
    $script:reportPDFFilename = $null
    
    # Disable buttons during execution
    $btnRunReport.Enabled = $false
    $btnViewReport.Enabled = $false
    $btnOpenReportFolder.Enabled = $false
    $btnCancel.Enabled = $true
    $progressBar.Value = 0
    
    # Clear status
    $txtStatus.Text = ""
    Write-StatusLog "Starting AD Snapshot Report..."
    
    try {
        # Create a temporary script file with parameters from the GUI
        $tempScriptPath = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.ps1'
        
        # Build the script content with parameters from the GUI
        $scriptContent = @"
# AD Snapshot Script with parameters from GUI
`$StartupVars = @()
`$StartupVars = Get-Variable | Select-Object -ExpandProperty Name
`$StartupVars += "PSItem"

# Parameters from GUI
`$OUlist = "$($txtOUList.Text)"
`$realm = "$($txtRealm.Text)"
`$CreateFile = "$(if ($chkCreateFile.Checked) { "Y" } else { "N" })"
`$outputpath = "$($txtOutputPath.Text)"
`$WantPDFFile = "$(if ($chkWantPDFFile.Checked) { "true" } else { "false" })"
`$PDFConverter = "$($txtPDFConverter.Text)"
`$SendEmail = "$(if ($chkSendEmail.Checked) { "Y" } else { "N" })"
`$CcList = "$($txtCcList.Text)"
`$smtpserver = "$($txtSmtpServer.Text)"
`$ServerUpTimeAlarm = $($numServerUpTimeAlarm.Value)
`$ComputerStaleDays = $($numComputerStaleDays.Value)
`$UserStaleDays = $($numUserStaleDays.Value)
`$PCsort = "$(if ($cboPCSort.SelectedIndex -eq 0) { "name" } else { "date" })"
`$Usersort = "$(if ($cboUserSort.SelectedIndex -eq 0) { "name" } else { "date" })"
`$HideDisabledUsers = "$(if ($chkHideDisabledUsers.Checked) { "true" } else { "false" })"
`$listALLComputers = "$(if ($chkListAllComputers.Checked) { "true" } else { "false" })"
`$skipadmin = "$(if ($chkSkipAdmin.Checked) { "true" } else { "false" })"
`$skipServer = "$(if ($chkSkipServer.Checked) { "true" } else { "false" })"
`$skipUsers = "$(if ($chkSkipUsers.Checked) { "true" } else { "false" })"
`$skipComputers = "$(if ($chkSkipComputers.Checked) { "true" } else { "false" })"
`$skipBitlockerStatus = "$(if ($chkSkipBitlockerStatus.Checked) { "true" } else { "false" })"
`$fromemail = "$($txtFromEmail.Text)"
`$defaultSentTo = "$($txtToEmail.Text)"
`$AdminGroup = "$($txtAdminGroup.Text)"

# Custom Domain Settings
if ("$($txtRealm.Text)" -eq "example") {
    `$DomainCN = "DC=example,DC=org"
    `$Server = "domaincontroller.example.org"
} else {
    # Use custom domain settings from GUI
    `$DomainCN = "$($txtDomainCN.Text)"
    `$Server = "$($txtDomainController.Text)"
}

# Include the original script content
"@
        
        # Get the content of the original script file
        $originalScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "AD-Snapshot.ps1"
        $originalScriptContent = Get-Content -Path $originalScriptPath -Raw
        
        # Extract the main script body (excluding the parameter declarations)
        $scriptBody = $originalScriptContent -replace '(?sm)^.*?#########################################################################################################\s*###\s*DO NOT CHANGE ANYTHING BELOW THIS LINE\s*###\s*#########################################################################################################', '#########################################################################################################
### DO NOT CHANGE ANYTHING BELOW THIS LINE ###
#########################################################################################################'
        
        # Remove the hardcoded domain settings line from the original script
        $scriptBody = $scriptBody -replace 'if \( \$realm -eq "example"\)\{\$DomainCN = "DC=example,DC=org" ; \$Server = "domaincontroller\.example\.org"\}.*?\n', ''
        
        # Combine the parameter declarations with the script body
        $scriptContent += $scriptBody
        
        # Write the combined script to the temporary file
        Set-Content -Path $tempScriptPath -Value $scriptContent
        
        Write-StatusLog "Temporary script created at $tempScriptPath"
        
        # Create a PowerShell job to run the script
        $job = Start-Job -ScriptBlock {
            param($scriptPath)
            & $scriptPath
        } -ArgumentList $tempScriptPath
        
        # Update progress while job is running
        $progressBar.Value = 10
        $jobRunning = $true
        
        while ($jobRunning) {
            Start-Sleep -Milliseconds 500
            [System.Windows.Forms.Application]::DoEvents()
            
            $jobState = (Get-Job -Id $job.Id).State
            
            if ($jobState -eq "Running") {
                if ($progressBar.Value -lt 90) {
                    $progressBar.Value += 1
                }
            }
            elseif ($jobState -eq "Completed") {
                $progressBar.Value = 100
                $jobRunning = $false
                
                # Get the job output
                $jobOutput = Receive-Job -Id $job.Id
                foreach ($line in $jobOutput) {
                    Write-StatusLog $line
                }
                
                # Determine the report path and filename
                if ($chkCreateFile.Checked) {
                    if ([string]::IsNullOrEmpty($txtOutputPath.Text)) {
                        $script:reportPath = [Environment]::GetFolderPath("Desktop")
                    }
                    else {
                        $script:reportPath = $txtOutputPath.Text
                    }
                    
                    $subOUonly = $txtOUList.Text
                    if ($subOUonly.Contains(",")) {
                        $subOUonly = $subOUonly.Substring(0, $subOUonly.IndexOf(","))
                    }
                    
                    $script:reportFilename = "$subOUonly AD Snapshot Report.html"
                    $script:reportPDFFilename = "$subOUonly AD Snapshot Report.pdf"
                    
                    Write-StatusLog "Report generated at $script:reportPath\$script:reportFilename"
                    
                    if ($chkWantPDFFile.Checked) {
                        Write-StatusLog "PDF report generated at $script:reportPath\$script:reportPDFFilename"
                    }
                    
                    $btnViewReport.Enabled = $true
                    $btnOpenReportFolder.Enabled = $true
                }
                
                Write-StatusLog "AD Snapshot Report completed successfully!"
            }
            elseif ($jobState -eq "Failed") {
                $progressBar.Value = 0
                $jobRunning = $false
                
                # Get the job error
                $jobError = (Get-Job -Id $job.Id).Error
                Write-StatusLog "Error: $jobError"
                
                Write-StatusLog "AD Snapshot Report failed!"
            }
        }
        
        # Clean up
        Remove-Job -Id $job.Id -Force
        Remove-Item -Path $tempScriptPath -Force
    }
    catch {
        Write-StatusLog "Error: $_"
    }
    finally {
        # Re-enable buttons
        $btnRunReport.Enabled = $true
        $btnCancel.Enabled = $false
    }
}

# Function to view the generated report
function Show-Report {
    if ($script:reportPath -and $script:reportFilename) {
        $reportFullPath = Join-Path -Path $script:reportPath -ChildPath $script:reportFilename
        if (Test-Path -Path $reportFullPath) {
            Start-Process $reportFullPath
            Write-StatusLog "Opening report $reportFullPath"
        }
        else {
            Write-StatusLog "Report file not found at $reportFullPath"
        }
    }
    else {
        Write-StatusLog "No report file available to view"
    }
}

# Function to open the report folder
function Open-ReportFolder {
    if ($script:reportPath) {
        if (Test-Path -Path $script:reportPath) {
            Start-Process $script:reportPath
            Write-StatusLog "Opening folder $script:reportPath"
        }
        else {
            Write-StatusLog "Folder not found at $script:reportPath"
        }
    }
    else {
        Write-StatusLog "No report folder available to open"
    }
}

# Function to cancel the running job
function Stop-RunningJob {
    $jobs = Get-Job
    foreach ($job in $jobs) {
        if ($job.State -eq "Running") {
            Stop-Job -Id $job.Id
            Remove-Job -Id $job.Id -Force
            Write-StatusLog "Job cancelled"
        }
    }
    
    $progressBar.Value = 0
    $btnRunReport.Enabled = $true
    $btnCancel.Enabled = $false
}

#endregion
