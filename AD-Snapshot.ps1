<#
####################################################################################################
## Active Directory Snapshot Report  -  Version 2.0.0
## Author: Bryant Welch   |   License: MIT
####################################################################################################

 This script uses the Remote Server Administration Tools (RSAT) for Windows.
 Install on a workstation with:
     Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0

 Optional PDF output uses your installed Microsoft Edge or Google Chrome (Chromium) in
 headless mode - no extra software to download or configure. Edge ships with Windows 10/11.

 Run directly (edit the default values in the param() block below), via the GUI, or from
 Task Scheduler, e.g.:
     powershell.exe -ExecutionPolicy Bypass -File "C:\path\AD-Snapshot.ps1" -OUlist "MyOU" -realm "contoso"

 If a firewall blocks WMI, from an elevated prompt:
     netsh advfirewall firewall set rule group="Windows Management Instrumentation (WMI)" new enable=yes
####################################################################################################
#>

[CmdletBinding()]
param(
    # ===== Active Directory =====
    # Everything here can be left blank to auto-detect the current domain, so the
    # script is portable to any environment. Supply values only to override.
    [string[]]$OUlist          = @(),                                    # Blank = entire domain. Or OU name(s) / full OU DN(s). Multiple: "OU1","OU2"
    [string]  $realm           = "",                                     # Blank = auto-detect NetBIOS domain name
    [string]  $DomainCN        = "",                                     # Blank = auto-detect (e.g. DC=contoso,DC=com)
    [string]  $Server          = "",                                     # Blank = auto-detect a domain controller
    [string]  $AdminGroup      = "Domain Admins",                        # Admin group(s) to report on (comma separated)
    [System.Management.Automation.PSCredential]$Credential,             # Optional alternate credentials (for cross-domain / non-joined hosts)

    # ===== Output / File =====
    [string]  $CreateFile      = "Y",                                    # "Y" saves the HTML report to disk
    [string]  $outputpath      = "",                                     # Output folder (blank = Desktop)
    [string]  $WantPDFFile     = "false",                               # "true" to also produce a PDF (via headless Edge/Chrome)
    [string]  $ExportCSV       = "false",                               # "true" to export Users/Computers as CSV

    # ===== Email =====
    [string]  $SendEmail       = "N",                                    # "Y" to email the report
    [string]  $fromemail       = "user1@example.com",
    [string]  $defaultSentTo   = "user2@example.com",
    [string]  $CcList          = "",
    [string]  $smtpserver      = "smtp.example.org",

    # ===== Thresholds / Display =====
    [int]     $ServerUpTimeAlarm = 30,                                   # Days of uptime before flagged
    [int]     $ComputerStaleDays = 80,                                   # Days before a computer is stale
    [int]     $UserStaleDays     = 80,                                   # Days before a user is stale
    [string]  $PCsort            = "name",                               # "name" or "date"
    [string]  $Usersort          = "name",                               # "name" or "date"
    [string]  $HideDisabledUsers = "false",                             # "true" hides disabled users
    [string]  $listALLComputers  = "true",                             # "true" includes servers in computer list

    # ===== Section toggles =====
    [string]  $skipDomainOverview   = "false",
    [string]  $skipadmin            = "false",
    [string]  $skipServer           = "false",
    [string]  $skipUsers            = "false",
    [string]  $skipComputers        = "false",
    [string]  $skipBitlockerStatus  = "true"
)

$StartupVars  = @()                                              ###  DO NOT MODIFY
$StartupVars  = Get-Variable | Select-Object -ExpandProperty Name ###  DO NOT MODIFY
$StartupVars += "PSItem"                                         ###  DO NOT MODIFY

#########################################################################################################
###                            DO NOT CHANGE ANYTHING BELOW THIS LINE                                 ###
#########################################################################################################

# ----- Load the Active Directory module -----
$adAvailable = $false
try { Import-Module ActiveDirectory -ErrorAction Stop; $adAvailable = $true }
catch {
    Write-Host "ERROR: The ActiveDirectory module is not available."
    Write-Host "Install RSAT, e.g.: Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
    return
}

$Version = "AD Snapshot Report Version 2.0.0"
$DebugPreference   = "SilentlyContinue"
$VerbosePreference = "SilentlyContinue"

# ----- PDF support via headless Chromium (Edge/Chrome) - no external tooling required -----
function Get-ChromiumBrowser {
    $candidates = @(
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe')
        (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe')
        (Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe')
        (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe')
        (Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\Application\msedge.exe')
    )
    foreach ($c in $candidates) { if ($c -and (Test-Path $c)) { return $c } }

    # Registry App Paths fallback
    foreach ($exe in 'msedge.exe', 'chrome.exe') {
        foreach ($hive in 'HKLM:', 'HKCU:') {
            $key = "$hive\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$exe"
            try {
                $p = (Get-ItemProperty -Path $key -ErrorAction Stop).'(default)'
                if ($p -and (Test-Path $p)) { return $p }
            } catch { }
        }
    }
    return $null
}

function Convert-HtmlToPdf {
    param(
        [string]$HtmlPath,
        [string]$PdfPath,
        [string]$BrowserExe
    )
    if (-not $BrowserExe) { return $false }
    if (-not (Test-Path $HtmlPath)) { return $false }

    # Use a file:// URI so spaces/special chars are encoded safely
    $uri = ([uri](Resolve-Path -LiteralPath $HtmlPath).Path).AbsoluteUri
    # Isolated profile so it works even if the browser is already running interactively
    $profileDir = Join-Path $env:TEMP ("adsnap_pdf_" + [guid]::NewGuid().ToString('N'))

    if (Test-Path $PdfPath) { Remove-Item -LiteralPath $PdfPath -Force -ErrorAction SilentlyContinue }

    $chromeArgs = @(
        '--headless'
        '--disable-gpu'
        '--no-first-run'
        '--no-pdf-header-footer'
        "--user-data-dir=$profileDir"
        "--print-to-pdf=$PdfPath"
        $uri
    )

    try {
        $null = & $BrowserExe @chromeArgs 2>&1
        # Give the writer a moment to flush in case the process returned early
        $deadline = (Get-Date).AddSeconds(20)
        while (-not (Test-Path $PdfPath) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 200 }
        return (Test-Path $PdfPath)
    } catch {
        return $false
    } finally {
        Remove-Item -Path $profileDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ----- Severity color palette (soft, print-friendly) -----
$cNormal  = "#eef2f6"
$cDanger  = "#f8d7da"
$cWarning = "#fff3cd"
$cInfo    = "#d1ecf1"
$cSuccess = "#d4edda"
$cLock    = "#f5c6cb"
$cUnknown = "#e9ecef"

# Strips HTML tags for CSV / plain-text output
function ConvertTo-PlainText { param([string]$Text) return ($Text -replace '<[^>]+>', '' -replace '\s+', ' ').Trim() }

# Minimal HTML-encode for values that come from AD (names, descriptions)
function ConvertTo-HtmlText { param([string]$Text) if ($null -eq $Text) { return "" } return ([string]$Text -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;') }

# Build a key/value row for the Domain Health Overview, with optional severity background
function New-KvRow {
    param([string]$Key, [string]$Value, [string]$Severity = "")
    $bg = switch ($Severity) { "danger" { $cDanger } "warn" { $cWarning } "ok" { $cSuccess } default { "" } }
    $style = if ($bg) { " style='background:$bg;'" } else { "" }
    return "<tr><td>$(ConvertTo-HtmlText $Key)</td><td$style>$Value</td></tr>"
}

# Build a Security Findings row (label, affected accounts, severity). Empty list -> no row.
function New-FindingRow {
    param([string]$Label, $Accounts, [string]$Severity = "warn")
    $list = @($Accounts | Sort-Object -Unique)
    if ($list.Count -eq 0) { return "" }
    $bg = switch ($Severity) { "danger" { $cDanger } "warn" { $cWarning } default { $cNormal } }
    $names = ($list -join ', ')
    return "<tr><td style='background:$bg;'>$(ConvertTo-HtmlText $Label)</td><td style='background:$bg;'><center>$($list.Count)</center></td><td>$(ConvertTo-HtmlText $names)</td></tr>"
}

# Flag operating systems that are out of (mainstream/extended) support
function Test-OSEndOfLife {
    param([string]$OS)
    if (-not $OS) { return $false }
    return ($OS -match 'Windows (XP|Vista|7|8|2000)|Windows Server (2000|2003|2008|2012)|Windows NT')
}

# ----- Resolve domain context (auto-detect so the script works in any domain) -----
# Common parameters threaded through every AD cmdlet (Server / Credential when supplied)
$adParams = @{}
if ($Server)     { $adParams['Server']     = $Server }
if ($Credential) { $adParams['Credential'] = $Credential }

try {
    $domainObj = Get-ADDomain @adParams -ErrorAction Stop
    if (-not $Server)   { $Server = $domainObj.PDCEmulator; $adParams['Server'] = $Server }
    if (-not $DomainCN) { $DomainCN = $domainObj.DistinguishedName }
    if (-not $realm)    { $realm = $domainObj.NetBIOSName }
    Write-Host "Connected to domain $($domainObj.DNSRoot) (DC: $Server)"
}
catch {
    Write-Warning "Could not auto-detect the domain: $($_.Exception.Message)"
    if (-not $DomainCN) {
        Write-Host "ERROR: No domain could be reached and no -DomainCN was supplied. Provide -Server (and -Credential if needed)."
        return
    }
    Write-Warning "Continuing with the supplied -DomainCN / -Server values."
}

# Normalize the OU list: blank/placeholder = a single pass over the entire domain
$ouEntries = @($OUlist | Where-Object { $_ -and $_.Trim() -ne "" -and $_.Trim().ToUpper() -ne "XXX" })
if ($ouEntries.Count -eq 0) { $ouEntries = @("") }

########################################################################################################
## BEGIN DOMAIN HEALTH OVERVIEW (collected once, domain-wide) ##
$script:domainOverviewHTML = ""
if ($skipDomainOverview -ne "true" -and $domainObj) {
    $timenow = Get-Date -UFormat %r
    write-host "$timenow Collecting Domain Health Overview"

    try { $forestObj = Get-ADForest @adParams -ErrorAction Stop } catch { $forestObj = $null; Write-Warning "Could not read forest: $($_.Exception.Message)" }
    try { $pwPol     = Get-ADDefaultDomainPasswordPolicy @adParams -ErrorAction Stop } catch { $pwPol = $null }

    $recycleBin = "Unknown"
    try {
        $rb = Get-ADOptionalFeature -Filter "Name -eq 'Recycle Bin Feature'" @adParams -ErrorAction Stop
        if ($rb -and @($rb.EnabledScopes).Count -gt 0) { $recycleBin = "Enabled" } else { $recycleBin = "Disabled" }
    } catch { }

    $dcs = @()
    try { $dcs = @(Get-ADDomainController -Filter * @adParams -ErrorAction Stop) } catch { Write-Warning "Could not enumerate domain controllers: $($_.Exception.Message)" }

    $trusts = @()
    try { $trusts = @(Get-ADTrust -Filter * @adParams -ErrorAction Stop) } catch { }

    $krbAge = $null
    try { $krb = Get-ADUser -Identity krbtgt -Properties PasswordLastSet @adParams -ErrorAction Stop; if ($krb.PasswordLastSet) { $krbAge = ((Get-Date) - $krb.PasswordLastSet).Days } } catch { }

    $builtinAdmin = $null
    try { $builtinAdmin = Get-ADUser -Identity "$($domainObj.DomainSID.Value)-500" -Properties Enabled, PasswordLastSet, LastLogonDate @adParams -ErrorAction Stop } catch { }

    # ---- Summary table ----
    $ov = "<h3>Summary</h3><table class='kv'><tbody>"
    if ($forestObj) { $ov += New-KvRow "Forest" (ConvertTo-HtmlText $forestObj.Name) }
    $ov += New-KvRow "Domain (DNS)"     (ConvertTo-HtmlText $domainObj.DNSRoot)
    $ov += New-KvRow "Domain (NetBIOS)" (ConvertTo-HtmlText $domainObj.NetBIOSName)
    $domModeOld = ([string]$domainObj.DomainMode -match '2000|2003|2008|2012')
    $ov += New-KvRow "Domain functional level" (ConvertTo-HtmlText ([string]$domainObj.DomainMode)) ($(if ($domModeOld) { "warn" } else { "ok" }))
    if ($forestObj) {
        $forModeOld = ([string]$forestObj.ForestMode -match '2000|2003|2008|2012')
        $ov += New-KvRow "Forest functional level" (ConvertTo-HtmlText ([string]$forestObj.ForestMode)) ($(if ($forModeOld) { "warn" } else { "ok" }))
    }
    $ov += New-KvRow "AD Recycle Bin" $recycleBin ($(if ($recycleBin -eq "Enabled") { "ok" } elseif ($recycleBin -eq "Disabled") { "danger" } else { "" }))
    $ov += New-KvRow "Domain Controllers" ("{0}" -f @($dcs).Count)
    if ($null -ne $krbAge) {
        $krbSev = if ($krbAge -gt 365) { "danger" } elseif ($krbAge -gt 180) { "warn" } else { "ok" }
        $ov += New-KvRow "krbtgt password age" ("$krbAge days") $krbSev
    }
    if ($builtinAdmin) {
        $baEnabled = [bool]$builtinAdmin.Enabled
        $baPwAge = if ($builtinAdmin.PasswordLastSet) { ((Get-Date) - $builtinAdmin.PasswordLastSet).Days } else { $null }
        $baText = $(if ($baEnabled) { "Enabled" } else { "Disabled" })
        if ($null -ne $baPwAge) { $baText += " &bull; password $baPwAge days old" }
        $ov += New-KvRow "Built-in Administrator (RID 500)" $baText ($(if ($baEnabled) { "warn" } else { "ok" }))
    }
    $ov += "</tbody></table>"

    # ---- FSMO roles ----
    $ov += "<h3>FSMO Role Holders</h3><table class='kv'><tbody>"
    $ov += New-KvRow "PDC Emulator"          (ConvertTo-HtmlText ([string]$domainObj.PDCEmulator))
    $ov += New-KvRow "RID Master"            (ConvertTo-HtmlText ([string]$domainObj.RIDMaster))
    $ov += New-KvRow "Infrastructure Master" (ConvertTo-HtmlText ([string]$domainObj.InfrastructureMaster))
    if ($forestObj) {
        $ov += New-KvRow "Schema Master"        (ConvertTo-HtmlText ([string]$forestObj.SchemaMaster))
        $ov += New-KvRow "Domain Naming Master" (ConvertTo-HtmlText ([string]$forestObj.DomainNamingMaster))
    }
    $ov += "</tbody></table>"

    # ---- Password & lockout policy ----
    if ($pwPol) {
        $ov += "<h3>Default Domain Password Policy</h3><table class='kv'><tbody>"
        $minLen = [int]$pwPol.MinPasswordLength
        $ov += New-KvRow "Minimum password length" $minLen ($(if ($minLen -lt 8) { "danger" } elseif ($minLen -lt 14) { "warn" } else { "ok" }))
        $ov += New-KvRow "Complexity enabled" ([string]$pwPol.ComplexityEnabled) ($(if ($pwPol.ComplexityEnabled) { "ok" } else { "danger" }))
        $maxAge = [int]$pwPol.MaxPasswordAge.TotalDays
        $ov += New-KvRow "Maximum password age (days)" ($(if ($maxAge -le 0) { "Never expires" } else { "$maxAge" })) ($(if ($maxAge -le 0) { "warn" } else { "ok" }))
        $ov += New-KvRow "Minimum password age (days)" ([int]$pwPol.MinPasswordAge.TotalDays)
        $ov += New-KvRow "Password history count" ([int]$pwPol.PasswordHistoryCount)
        $lockThresh = [int]$pwPol.LockoutThreshold
        $ov += New-KvRow "Account lockout threshold" ($(if ($lockThresh -eq 0) { "0 (no lockout)" } else { "$lockThresh attempts" })) ($(if ($lockThresh -eq 0) { "warn" } else { "ok" }))
        $ov += New-KvRow "Lockout duration (min)" ([int]$pwPol.LockoutDuration.TotalMinutes)
        $ov += New-KvRow "Reversible encryption" ([string]$pwPol.ReversibleEncryptionEnabled) ($(if ($pwPol.ReversibleEncryptionEnabled) { "danger" } else { "ok" }))
        $ov += "</tbody></table>"
    }

    # ---- Domain Controller inventory ----
    if (@($dcs).Count -gt 0) {
        $ov += "<h3>Domain Controllers</h3><table class='sortable' id='tblDCs'><thead><tr><th>Name</th><th>Operating System</th><th>IPv4</th><th>Site</th><th>Global Catalog</th><th>Read-Only</th></tr></thead><tbody>"
        foreach ($dc in ($dcs | Sort-Object HostName)) {
            $dcOS = [string]$dc.OperatingSystem
            $osBg = if (Test-OSEndOfLife $dcOS) { " style='background:$cDanger;'" } else { "" }
            $ov += "<tr><td>$(ConvertTo-HtmlText ([string]$dc.HostName))</td><td$osBg>$(ConvertTo-HtmlText $dcOS)</td><td>$(ConvertTo-HtmlText ([string]$dc.IPv4Address))</td><td>$(ConvertTo-HtmlText ([string]$dc.Site))</td><td><center>$(if ($dc.IsGlobalCatalog) { 'Yes' } else { 'No' })</center></td><td><center>$(if ($dc.IsReadOnly) { 'Yes' } else { 'No' })</center></td></tr>"
        }
        $ov += "</tbody></table>"
    }

    # ---- Trusts ----
    if (@($trusts).Count -gt 0) {
        $ov += "<h3>Domain Trusts</h3><table class='sortable' id='tblTrusts'><thead><tr><th>Name</th><th>Direction</th><th>Type</th></tr></thead><tbody>"
        foreach ($t in $trusts) {
            $ov += "<tr><td>$(ConvertTo-HtmlText ([string]$t.Name))</td><td>$(ConvertTo-HtmlText ([string]$t.Direction))</td><td>$(ConvertTo-HtmlText ([string]$t.TrustType))</td></tr>"
        }
        $ov += "</tbody></table>"
    }

    $script:domainOverviewHTML = $ov
    $timenow = Get-Date -UFormat %r
    write-host "$timenow Completed Domain Health Overview"
}
########################################################################################################
## END DOMAIN HEALTH OVERVIEW ##

foreach ($subOU in $ouEntries) {      ###  DO NOT MODIFY
$starttime = Get-Date

# Resolve the effective search base and a friendly label for this pass
if ([string]::IsNullOrWhiteSpace($subOU)) {
    $effectiveBase = $DomainCN
    if ($realm) { $subOUonly = $realm } else { $subOUonly = (($DomainCN -split ',' | ForEach-Object { $_ -replace '^DC=','' }) -join '.') }
}
elseif ($subOU -match '(?i)DC=') {
    # A full distinguished name was supplied
    $effectiveBase = $subOU
    $subOUonly = (($subOU -split ',')[0] -replace '^[A-Za-z]+=','')
}
else {
    # Treat as an OU name and resolve it to a DN anywhere in the domain
    $subOUonly = $subOU
    $effectiveBase = $DomainCN
    try {
        $ouMatch = Get-ADOrganizationalUnit -Filter "Name -eq '$subOU'" @adParams -ErrorAction Stop | Select-Object -First 1
        if ($ouMatch) { $effectiveBase = $ouMatch.DistinguishedName }
        else { Write-Warning "OU '$subOU' not found; searching the entire domain instead." }
    }
    catch { Write-Warning "Could not look up OU '$subOU' ($($_.Exception.Message)); searching the entire domain instead." }
}
Write-Host "Search base for this pass: $effectiveBase"

# Email recipients and subject
[string[]] $SendToList = "$defaultSentTo"
$subjectline = "$subOUonly AD Snapshot Report"

# Output filenames (label sanitized so it is filesystem-safe)
$safeName = ($subOUonly -replace '[\\/:*?"<>|]', '_').Trim()
if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = "AD" }
$filename     = $safeName + " AD Snapshot Report.html"
$filenamePDF  = $safeName + " AD Snapshot Report.pdf"
$filenameUsersCSV     = $safeName + " AD Snapshot Users.csv"
$filenameComputersCSV = $safeName + " AD Snapshot Computers.csv"

# Dashboard counters
$script:adminCount     = 0
$script:serverReported = 0
$script:serverFound    = 0
$script:userCount      = 0
$script:userDisabled   = 0
$script:userStale      = 0
$script:computerCount  = 0
$script:computerStale  = 0
$script:securityFindingsHTML = ""

########################################################################################################
## BEGIN ADMIN MEMBERSHIP INFORMATION ##
if ($skipAdmin -eq "true") { } else {
    $timenow = Get-Date -UFormat %r
    write-host $timenow Starting Admin Membership collection

    # Support one or more comma-separated group names; resolve each anywhere in the domain
    $adminGroupNames = @($AdminGroup -split '\s*,\s*' | Where-Object { $_ -and $_.Trim() -ne "" })
    if ($adminGroupNames.Count -eq 0) { $adminGroupNames = @("Domain Admins") }

    $AdminLG = ""
    $script:adminCount = 0
    foreach ($gName in $adminGroupNames) {
        $AdminLG += "<table class='internal'><thead><tr><th>$(ConvertTo-HtmlText $gName) Members</th></tr></thead><tbody><tr><td>"
        try {
            $grp = Get-ADGroup -Filter "Name -eq '$gName'" @adParams -ErrorAction Stop | Select-Object -First 1
            if (-not $grp) { $grp = Get-ADGroup -Filter "SamAccountName -eq '$gName'" @adParams -ErrorAction SilentlyContinue | Select-Object -First 1 }
            if ($grp) {
                $members = @(Get-ADGroupMember -Identity $grp -Recursive @adParams -ErrorAction Stop | Sort-Object Name)
                foreach ($member in $members) { $AdminLG += (ConvertTo-HtmlText $member.Name) + "<br />" }
                $cnt = $members.Count
                $script:adminCount += $cnt
                $AdminLG += "</td></tr></tbody></table><center>$cnt Members.</center><br>"
            }
            else {
                $AdminLG += "<i>Group not found in this domain.</i></td></tr></tbody></table><br>"
                Write-Warning "Admin group '$gName' was not found."
            }
        }
        catch {
            $AdminLG += "<i>Error reading group: $(ConvertTo-HtmlText $_.Exception.Message)</i></td></tr></tbody></table><br>"
            Write-Warning "Could not read admin group '$gName': $($_.Exception.Message)"
        }
    }

    $timenow = Get-Date -UFormat %r
    write-host $timenow Completed Admin Membership collection
}
########################################################################################################
## END ADMIN MEMBERSHIP INFORMATION ##


########################################################################################################
## BEGIN SERVER INFORMATION ##
if ($skipServer -eq "true") { } else {
    $timenow = Get-Date -UFormat %r
    write-host $timenow Starting Server Collection
       # Discover servers by operating system anywhere under the search base (portable)
       $servers = @()
       try {
           $servers = @(Get-ADComputer -Filter 'OperatingSystem -like "*Server*"' -SearchBase $effectiveBase -SearchScope Subtree -Properties OperatingSystem, Enabled @adParams -ErrorAction Stop |
                        Where-Object { $_.Enabled } | Select-Object -ExpandProperty Name | Sort-Object)
       }
       catch { Write-Warning "Could not enumerate servers under $effectiveBase : $($_.Exception.Message)" }
       write-host "$timenow Found $($servers.Count) server object(s) in AD"
    # filter out inaccessible computers and create error log
    $errlog = ""
    $filteredservers = @()
    foreach($system in $servers) {
        $timenow = Get-Date -UFormat %r
        write-host $timenow Testing communication with $system
        if (Get-WMIObject -ea 0 -Errorvariable err -ComputerName $system Win32_LogicalDisk)
           { $filteredservers += $system } else {
                $errlog+="Could Not Connect to Server: $system -- $err[0]<br>"
                $timenow = Get-Date -UFormat %r
                Write-warning "$timenow Communication with $system failed"
    }}
    $numfound = $servers.Count
    $num = $filteredservers.Count
    $script:serverReported = $num
    $script:serverFound    = $numfound
    $serverinfo = "<table id='tblServers' class='sortable'><thead><tr><th>Server</th><th><table class='drive'><th style='width: 20px;'></th><th>Size</th><th>Free</th><th style='width: 60px;'>%</th><th style='width: 10px;'></th></table></th><th>Uptime</th><th>Administrators<br/>Members</th><th>RDP<br/>Members</th><th>Local<br/>Accounts</th></tr></thead><tbody>"
    foreach ($server in $filteredservers) {
    $timenow = Get-Date -UFormat %r
    write-host $timenow Getting information for $server
        try {
                $UptimeAlarm = $cNormal
                $i = Get-WMIObject -class Win32_OperatingSystem -ComputerName $server
                $Bootup = $i.LastBootUpTime
                $serverIP = Get-NetIPAddress
                $LastBootUpTime = [System.Management.ManagementDateTimeconverter]::ToDateTime($Bootup)
                $now = Get-Date
                $Uptime = $now - $LastBootUpTime
                $d = $Uptime.Days
                $h = $Uptime.Hours
                $m = $uptime.Minutes
                $ms= $uptime.Milliseconds
                # Display uptime
                $ServerUptime = "$d days<br/>$h hours<br />$m minutes"
                if ( $d -gt $ServerUpTimeAlarm ) { $UptimeAlarm = $cDanger }
                # Server Administrators
                $serverAdminCount = 0
                $admin = ""
                $group = "Group Error"
                $GMembers = "GMembers Error"
                $group = [ADSI]("WinNT://$server/Administrators,group")
                $GMembers = $group.psbase.invoke("Members")
                $GMembers | ForEach-Object { $admin = $admin + $_.GetType().InvokeMember("Name",'GetProperty', $null, $_, $null) + "<br>"  }
                Write-Verbose "$server Admins $admin"
                # Server Remote Desktop Users
                $serverRDPCount = 0
                $RDP = ""
                $group = "Group Error"
                $GMembers = "GMembers Error"
                $group = [ADSI]("WinNT://$server/Remote Desktop Users,group")
                $GMembers = $group.psbase.invoke("Members")
                $GMembers | ForEach-Object { $RDP = $RDP + $_.GetType().InvokeMember("Name",'GetProperty', $null, $_, $null) + "<br>" }
                Write-Verbose "$server RDP users $RDP"
                # Server IP addresses
                $ServerIPlist = ""
                $serverIPs = Get-WmiObject win32_networkadapterconfiguration -filter "ipenabled = 'True'" -ComputerName $server | Select IPAddress
                $ServerIPs | ForEach-Object {$ServerIPlist += "<br/>" + $_.IPAddress }
                # Server Local User Accounts
                $LocalAccts = ""
                $AllLocalAccounts = Get-WmiObject -Class Win32_UserAccount -Namespace "root\cimv2" -Filter "LocalAccount='$True'" -ComputerName $server
                Write-Verbose "$server"
                Foreach($LocalAccount in $AllLocalAccounts)
                {
                    $LocalUserName = $LocalAccount.Name
                    $LocalUserDisabled = $LocalAccount.Disabled
                    if ( $LocalUserDisabled -eq "True")
                    {
                        $LocalAccts += "$LocalUserName <span class='tag tag-off'>Disabled</span><br/>"
                    } else {
                        $LocalAccts += "$LocalUserName<br/>"
                    }
                    Write-Verbose "    $LocalUserName,$LocalUserDisabled"
                }
                # Server Disk Information
                $serverDriveinfo = "<table class='drive'>"
            foreach ( $i in (Get-WmiObject -Class Win32_LogicalDisk -ComputerName $server))
            {
                if ($i.DriveType -eq 3) {
                    $SystemName = $i.SystemName
                    $Drive = $i.Name
                    $timenow = Get-Date -UFormat %r
                    Write-Verbose "$timenow Collecting Drive $Drive information for $server"
                    $VolName = $i.VolumeName
                    $Size = (($i.Size/1gb))
                    $Free = (($i.freespace/1gb))
                    $alarm = 0
                    if ($i.Size*100 -lt 1) { $PercentFree = "----" } else { $PercentFree = ("{0:N4}" -f ($i.freespace/$i.Size)) }
                    if ($Free -lt 10) { $alarm = 1 }
                    if ($PercentFree -lt 0.080) { $alarm = 1 }
                    if ($PercentFree -eq "----") { $alarm = 0 }
                    if ($PercentFree -gt 0.160) { $alarm = 0 }
                    $Size = ("{0:N2}" -f ($i.Size/1gb))
                    $Free = ("{0:N2}" -f ($i.freespace/1gb))
                    if ($i.Size*100 -lt 1) { $PercentFree = "----" } else { $PercentFree = ("{0:N2}" -f ($i.freespace/$i.Size*100))}


                    if ($alarm -eq 1)  {
                        $serverDriveinfo = $serverDriveinfo + "<tr style='background:$cDanger;'><td class='drive'>$Drive</td><td class='drive'>$Size</td><td class='drive'>$Free</td><td class='drive'>$PercentFree%</td></tr>"
                        } else {
                        $serverDriveinfo = $serverDriveinfo + "<tr><td class='drive'>$Drive</td><td class='drive'>$Size</td><td class='drive'>$Free</td><td class='drive'>$PercentFree%</td></tr>"
                    }
                }
            }
            $serverDriveinfo = $serverDriveinfo + "</tr></table>"
            $serverinfo = $serverinfo + "<tr><td class='server'>$SystemName $ServerIPlist</td><td class='server'>$serverDriveinfo</td><td class='server' style='background-color:$UptimeAlarm;' >$ServerUptime</td><td class='server'>$admin</td><td class='server'>$RDP</td><td class='server'>$LocalAccts</td></tr>"
        }
        catch {
            $errlog += "Error collecting from server: $server -- $($_.Exception.Message)<br>"
            Write-Warning "Error collecting from server $server : $($_.Exception.Message)"
        }

        }

    $serverinfo = $serverinfo + "</tbody></table>$num Servers reported out of $numfound Servers found in AD.<br>"

    $timenow = Get-Date -UFormat %r
    write-host $timenow Completed Server collection
}
########################################################################################################
## END SERVER INFORMATION  ##

########################################################################################################
## BEGIN AD USER INFORMATION ##
if ($skipUsers -eq "true") { } else {
    $timenow = Get-Date -UFormat %r
    $adlogons = "<table id='tblUsers' class='sortable'><thead><tr><th>Name</th><th>Login</th><th>Description</th><th>Last <br/>Login</th><th class ='norm'>Password <br/>Changed</th><th class ='norm'>When<br/>Created</th><th class ='norm'>Never <br/>Expire</th><th class ='norm'>Active</th></tr></thead><tbody>"
    $UserRecinfo = @()
    $UserRecinfo_sort = @()
   write-host $timenow Starting User collection
    # Single domain/OU-wide query (works in any environment, no hardcoded OU=Users path)
    $userObjects = @()
    try {
        $userObjects = @(Get-ADUser -Filter * -SearchBase $effectiveBase -SearchScope Subtree -ResultPageSize 500 -Properties WhenCreated, LastLogon, LastLogonTimeStamp, userAccountControl, logonCount, displayName, description, LockoutTime, pwdLastSet, givenName, sn, Enabled, ServicePrincipalName, adminCount, AccountExpirationDate @adParams -ErrorAction Stop)
    }
    catch { Write-Warning "Could not query users under $effectiveBase : $($_.Exception.Message)" }
    $num = $userObjects.Count
    $script:userCount = $num
    # Security hygiene collectors (B)
    $secNoPreAuth = @(); $secSPN = @(); $secDelegation = @(); $secPwNotReq = @(); $secAdminCount = @(); $secExpired = @(); $secNeverExpire = @()
   if ($num -eq 0) { Write-Warning "No users found under $effectiveBase" } else {
    foreach ($u in $userObjects) {
    $sam = $u.SamAccountName
    if ($u.displayName) { $name = [string]$u.displayName } else { $name = [string]$u.Name }
    $timenow = Get-Date -UFormat %r
    write-host $timenow Collecting information for $name
     $name = ConvertTo-HtmlText $name
     $description = $u.description
     if($description){$description = ConvertTo-HtmlText ([string]$description)} else {$description =""}
     $userAC = $u.userAccountControl
     $lockout = $u.LockoutTime
     $logonCount = $u.logonCount
     $pwdLastSet = [DateTime]::FromFileTime([Int64] $u.pwdLastSet)
     $nameBG = $cNormal
     $logonBg = $cNormal
     $ActiveUser = "A"
### WhenCreated
     $ADWhenCreated = [DateTime]$u.WhenCreated
     $WhenCreatedYear = $ADWhenCreated.Year
     $WhenCreatedMonth = $ADWhenCreated.Month.ToString("00")
     $WhenCreatedDay = $ADWhenCreated.Day.ToString("00")
     $WhenCreated1 = "$WhenCreatedMonth/$WhenCreatedDay/$WhenCreatedYear"
     $WhenCreatedBG = $cNormal
     $ADlastLogon = [DateTime]::FromFileTime([Int64] $u.LastLogon)
     $ADlastLogonTimeStamp = [DateTime]::FromFileTime([Int64] $u.LastLogonTimeStamp)
     if ($ADlastLogon -gt $ADlastLogonTimeStamp) {$lastLogon = $ADlastLogon} else {$lastLogon = $ADlastLogonTimeStamp}
     $lastLogonYear = $lastLogon.Year ; $lastLogonMonth = $lastLogon.Month.ToString("00") ; $lastLogonDay = $lastLogon.Day.ToString("00")
     $lastLogon1 = "$lastLogonMonth/$lastLogonDay/$lastLogonYear"
     # Robust flag detection via bit masks / the Enabled property (portable across all account types)
     if (($userAC -band 0x10000) -ne 0) { $neverexpire = "X" ; $expireBG = $cDanger } else { $expireBG = $cNormal ; $neverexpire = " " }
     $isDisabled = (-not $u.Enabled) -or (($userAC -band 0x2) -ne 0)
     if ($isDisabled) { $nameBG = $cInfo; $name = $name + " <span class='tag tag-off'>Disabled</span>" ; $ActiveUser = "D" }
     if ($lockout -gt 2) {$nameBG = $cLock; $name = $name + " <span class='tag tag-lock'>Locked</span>" }
     $now = Get-Date
     if ($lastLogon.Year -lt 2000)  {
         $logonBg = $cDanger  ## Users who never logged in will flag as red
         $lastLogon1 = "Never"
         $Uptime = $now - $ADWhenCreated
         if($Uptime.Days -gt $UserStaleDays) { $WhenCreatedBG = $cDanger } ## Users created over $UserStaleDays ago
     }
     $Uptime = $now - $lastLogon
     $d = $Uptime.Days
     if ($d -gt $UserStaleDays){ $logonBg = $cDanger}	## Users over $UserStaleDays flag red
     if ($logonBg -eq $cDanger) { $script:userStale++ }

     if ($pwdLastSet.Year -gt 2000) {
        $pwdLastSetYear= $pwdLastSet.Year ; $pwdLastSetMonth = $pwdLastSet.Month.ToString("00") ; $pwdLastSetDay = $pwdLastSet.Day.ToString("00")
        $pwdLastSet1 = "$pwdLastSetMonth/$pwdLastSetDay/$pwdLastSetYear"
        $pwdage = $now - $pwdLastSet
        $d = $pwdage.Days
        if ($neverexpire -eq "X") {
           if ($d -gt "89"){$pwdBG = $cWarning} else {$pwdBG = $cNormal}
           if ($d -gt "360"){$pwdBG = $cDanger}
        } else { if ($d -gt "82"){$pwdBG = $cWarning} else {$pwdBG = $cNormal}
           if ($d -gt "82") {              ## One week until expiration
              $pwdBG = $cWarning }         ## Highlight yellow
              else {
              $pwdBG = $cNormal }
           if ($d -gt "89") {              ## Expired
           $pwdBG = $cDanger }         ## Highlight red
        }  }
     else {
        $pwdLastSet1 = "Never" ;$pwdBG = $cDanger }


     write-Verbose  "Password changed on $pwdLastSet - $d days ago"
     if ($ActiveUser -eq "D") { $script:userDisabled++ }

     # ----- Security hygiene flags (B) -----
     if (($userAC -band 0x20) -ne 0)     { $secPwNotReq   += $sam; $name += " <span class='tag tag-risk'>PW NotReq</span>" }
     if (($userAC -band 0x400000) -ne 0) { $secNoPreAuth  += $sam; $name += " <span class='tag tag-risk'>No PreAuth</span>" }
     if (($userAC -band 0x80000) -ne 0)  { $secDelegation += $sam; $name += " <span class='tag tag-risk'>Deleg</span>" }
     if (@($u.ServicePrincipalName).Count -gt 0) { $secSPN += $sam; $name += " <span class='tag tag-svc'>SPN</span>" }
     if ($u.adminCount -eq 1) { $secAdminCount += $sam; $name += " <span class='tag tag-priv'>Priv</span>" }
     if ($neverexpire -eq "X") { $secNeverExpire += $sam }
     if ($u.AccountExpirationDate -and $u.AccountExpirationDate -lt $now -and -not $isDisabled) { $secExpired += $sam; $name += " <span class='tag tag-off'>Expired</span>" }

     if ($HideDisabledUsers -eq "true" -and $ActiveUser -eq "D") {
        write-Verbose "Disabled account for $name not displayed on report."
        } Else {
        $UserRecinfo += ,@( $name, $sam, $description, $lastLogon1, $neverexpire, $logonBg, $expireBG,$nameBG,$WhenCreated1,$pwdLastSet1,$pwdBG,$lastLogon,$ActiveUser,$WhenCreatedBG)
        ##                   0      1     2             3            4             5         6         7       8             9            10     11         12          13
        }
     }
     if ($Usersort -eq "name" ) {
     $UserRecinfo_sort = $UserRecinfo | sort-object @{Expression={$_[0]}; Ascending=$true} } else{
     $UserRecinfo_sort = $UserRecinfo | sort-object @{Expression={$_[11]}; Ascending=$true} }
     foreach($UserRec in $UserRecinfo_sort) {
            $adlogons = $adlogons +  "<tr><td style='background:" + $UserRec[7] + ";'>" + $UserRec[0] + "</td><td>" + $UserRec[1] + "</td><td>" + $UserRec[2] + "</td><td style='background:" + $UserRec[5] + ";'><center>" + $UserRec[3] + "</center></td><td style='background:" + $UserRec[10] + ";'><center>" + $UserRec[9] + "</center></td><td style='background:" + $UserRec[13] + ";'><center>" + $UserRec[8] + "</center></td><td style='background:" + $UserRec[6] + ";'><center>" + $UserRec[4] + "</center></td><td><center>" + $UserRec[12] + "</center></td></tr>"
     }
     $adlogons = $adlogons + "</tbody></table><center>$num Users Found.</center><br>" }

    # ----- Security Findings rollup (B) -----
    $findingRows  = ""
    $findingRows += New-FindingRow "Kerberos pre-auth not required (AS-REP roastable)" $secNoPreAuth   "danger"
    $findingRows += New-FindingRow "Trusted for unconstrained delegation"             $secDelegation  "danger"
    $findingRows += New-FindingRow "Password not required"                            $secPwNotReq    "danger"
    $findingRows += New-FindingRow "Service accounts with SPN (Kerberoastable)"        $secSPN         "warn"
    $findingRows += New-FindingRow "Protected / privileged (adminCount = 1)"           $secAdminCount  "warn"
    $findingRows += New-FindingRow "Expired but still enabled"                         $secExpired     "warn"
    $findingRows += New-FindingRow "Password never expires"                            $secNeverExpire "warn"
    if ($findingRows) {
        $script:securityFindingsHTML = "<table class='sortable' id='tblSec'><thead><tr><th>Finding</th><th>Count</th><th>Accounts</th></tr></thead><tbody>$findingRows</tbody></table>"
    } else {
        $script:securityFindingsHTML = "<center>No notable user security findings.</center>"
    }

    $timenow = Get-Date -UFormat %r
    write-host $timenow Completed AD User collection
  }
########################################################################################################
## END AD USER INFORMATION ##

########################################################################################################
## BEGIN AD COMPUTER INFORMATION ##
if ($skipComputers -eq "true") { } else {
    $timenow = Get-Date -UFormat %r
    write-host $timenow Starting Computer collection
    $now = Get-Date
    $num = 0
    $script:computerinfo = @()
    $adcomputers = "<table id='tblComputers' class='sortable'><thead><tr><th>Name</th><th>Description</th><th>Last<br/>Seen</th><th>Operating System</th><th>OS<br/>Version</th></tr></thead><tbody>"
        function AD_CompInfo {
        Param($c)

            $OSBG = $cNormal ; $alarm = 0
            $TPM = ""
            $samCN = $c.Name
            $timenow = Get-Date -UFormat %r
            write-host $timenow Collecting information for $samCN
            $BitlockerKey = ""
            $lastLogon = ""
            $description = ConvertTo-HtmlText ([string]$c.description)
            $sam = [string]$c.SamAccountName
            $OS = $c.operatingsystem
            $OSV = [string]$c.operatingsystemversion
            if ($OS) {$OS = ConvertTo-HtmlText ([string]$OS)} else { $OS = "" ; $OSBG = $cUnknown}
            $seenBG = $cNormal

        ########## Start Bitlocker ################################
            if ($skipBitlockerStatus -eq "true") { Write-Verbose "$samCN - $description - $OS "  } else {
                $BitLockerDataHash = @{}
                $BitLockerDataArray = @("BitCN", "BitKey")
                #Check if the computer object has had a BitLocker Recovery Password
                try { $BitlockerObject = Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase $c.DistinguishedName -Properties 'msFVE-RecoveryPassword' @adParams -ErrorAction Stop | Select-Object -Last 1
                } catch { $BitlockerObject = $null ; Write-Verbose "No BitLocker recovery object for $samCN" }
                if($BitlockerObject.'msFVE-RecoveryPassword'){
                    $BitLockerKey = $BitLockerObject.'msFVE-RecoveryPassword'
                    $BitLockerDataArray += @( $samCN, $BitLockerKey)
                    $BitLockerDataHash += @{ $samCN = $BitLockerKey}
                    Write-Verbose "$samCN - $description - $OS - Bitlocker Key= $BitlockerKey"
                    $TPM = "true"
                }else{ Write-Host -ForegroundColor Red "$samCN - WARNING!!! - The Bitlocker key is missing. - WARNING!!!" }
            }
            if ($TPM -eq "true") { $OSBG = $cSuccess; $OS += " <span class='tag tag-bl'>BitLocker</span>"}
        ########## End Bitlocker   ################################
            $ADlastLogon = [DateTime]::FromFileTime([Int64] $c.LastLogon)
            $ADlastLogonTimeStamp = [DateTime]::FromFileTime([Int64] $c.LastLogonTimeStamp)
            if ($ADlastLogon -gt $ADlastLogonTimeStamp) {$lastLogon = $ADlastLogon} else {$lastLogon = $ADlastLogonTimeStamp}
            $lastLogonYear = $lastLogon.Year ; $lastLogonMonth = $lastLogon.Month.ToString("00") ; $lastLogonDay = $lastLogon.Day.ToString("00")
            $lastLogon1 = "$lastLogonMonth/$lastLogonDay/$lastLogonYear"
                $age = $now - $lastLogon
                if($age.Days -gt $ComputerStaleDays) { $seenBG = $cDanger } ## Stale computers flag red


            if ($lastLogon.Year -lt 2000)  {
                $logonBg = $cDanger  ## Computers that never logged in flag red
                $lastLogon1 = "Never"
                $age = $now - $c.whenCreated
                if($age.Days -gt $ComputerStaleDays) { $seenBG = $cDanger } ## Stale computers flag red
                if ($c.logonCount -lt 10 ) {
                    $logonBg = $cWarning  ## New computers flag yellow
                    $lastLogonYear = $c.whenCreated.Year ; $lastLogonMonth = $c.whenCreated.Month.ToString("00") ; $lastLogonDay = $c.whenCreated.Day.ToString("00")
                    $lastLogon1 = "Created<br />$lastLogonMonth/$lastLogonDay/$lastLogonYear"
                }
            }
        if ($seenBG -eq $cDanger) { $script:computerStale++ }
        $script:computerinfo += ,@( $samCN, $description, $lastLogon1, $seenBG, $OS, $OSV, $OSBG)
        #                     0       1             2            3        4    5      6

     }

     # Single domain/OU-wide computer query (portable). Optionally exclude servers.
     $allComputers = @()
     try {
        $allComputers = @(Get-ADComputer -Filter * -SearchBase $effectiveBase -SearchScope Subtree -ResultPageSize 500 -Properties SamAccountName, lastLogon, lastLogonTimeStamp, displayName, description, operatingsystem, operatingsystemversion, whenCreated, logonCount, DistinguishedName, Enabled @adParams -ErrorAction Stop)
     }
     catch { Write-Warning "Could not query computers under $effectiveBase : $($_.Exception.Message)" }
     if ($listALLComputers -ne "true") { $allComputers = @($allComputers | Where-Object { $_.operatingsystem -notlike '*Server*' }) }
     foreach ($comp in $allComputers) {
        $num ++
        try { AD_CompInfo $comp }
        catch { Write-Warning "Error processing computer $($comp.Name): $($_.Exception.Message)" }
     }

    $script:computerCount = $num

         if ($PCsort -eq "name" ){
         $computerinfo_sort = $script:computerinfo | sort-object @{Expression={$_[0]}; Ascending=$true} } else {
         $computerinfo_sort = $script:computerinfo | sort-object @{Expression={$_[2]}; Ascending=$true} }
     foreach($computer in $computerinfo_sort) { $adcomputers += "
     <tr><td>" + $computer[0] + "</td><td>&nbsp; &nbsp;" + $computer[1] + "</td><td style='background:" + $computer[3] + "; width:90px;'>" + $computer[2] + "</td><td style='background:" + $computer[6] + "; width:170px;'>" + $computer[4] + "</td><td style='width:90px;'>" + $computer[5] + "</td></tr>" }

     $adcomputers+= "</tbody></table><center>$num Computers Found.</center><br>"

    $timenow = Get-Date -UFormat %r
    write-host $timenow Completed Computer collection
}
########################################################################################################
## END AD COMPUTER INFORMATION ##


########################################################################################################
## BEGIN CSV EXPORT ##
if ($ExportCSV -eq "true") {
    $timenow = Get-Date -UFormat %r
    write-host $timenow Exporting CSV files
    if ($outputpath -eq "") { $csvTarget = ([Environment]::GetFolderPath("Desktop")) } else { $csvTarget = $outputpath }

    if ($skipUsers -ne "true" -and $UserRecinfo_sort) {
        $UserRecinfo_sort | ForEach-Object {
            [PSCustomObject]@{
                Name            = (ConvertTo-PlainText $_[0])
                SamAccountName  = $_[1]
                Description     = (ConvertTo-PlainText $_[2])
                LastLogin       = $_[3]
                PasswordChanged = $_[9]
                NeverExpire     = ($_[4]).Trim()
                Active          = $_[12]
            }
        } | Export-Csv -Path (Join-Path $csvTarget $filenameUsersCSV) -NoTypeInformation -Encoding UTF8
        write-host "$timenow Users CSV written to $csvTarget\$filenameUsersCSV"
    }

    if ($skipComputers -ne "true" -and $computerinfo_sort) {
        $computerinfo_sort | ForEach-Object {
            [PSCustomObject]@{
                Name            = $_[0]
                Description     = (ConvertTo-PlainText $_[1])
                LastSeen        = (ConvertTo-PlainText $_[2])
                OperatingSystem = (ConvertTo-PlainText $_[4])
                OSVersion       = $_[5]
            }
        } | Export-Csv -Path (Join-Path $csvTarget $filenameComputersCSV) -NoTypeInformation -Encoding UTF8
        write-host "$timenow Computers CSV written to $csvTarget\$filenameComputersCSV"
    }
}
########################################################################################################
## END CSV EXPORT ##


########################################################################################################
## BEGIN HTML INFORMATION ##
    $timenow = Get-Date -UFormat %r
    write-host $timenow Starting Message Formatting
$endtime = Get-Date -UFormat %T

# ----- Build dashboard cards -----
$dashboardCards = ""
if ($skipAdmin -ne "true")     { $dashboardCards += "<div class='card'><div class='card-value'>$($script:adminCount)</div><div class='card-label'>$AdminGroup Members</div></div>" }
if ($skipServer -ne "true")    { $dashboardCards += "<div class='card'><div class='card-value'>$($script:serverReported)<span class='card-sub'>/ $($script:serverFound)</span></div><div class='card-label'>Servers Reported</div></div>" }
if ($skipUsers -ne "true")     {
    $dashboardCards += "<div class='card'><div class='card-value'>$($script:userCount)</div><div class='card-label'>Total Users</div></div>"
    $dashboardCards += "<div class='card card-warn'><div class='card-value'>$($script:userStale)</div><div class='card-label'>Stale Users (&gt;$UserStaleDays d)</div></div>"
    $dashboardCards += "<div class='card card-muted'><div class='card-value'>$($script:userDisabled)</div><div class='card-label'>Disabled Users</div></div>"
    if (@($secNoPreAuth).Count -gt 0) { $dashboardCards += "<div class='card card-danger'><div class='card-value'>$(@($secNoPreAuth).Count)</div><div class='card-label'>AS-REP Roastable</div></div>" }
    if (@($secSPN).Count -gt 0)       { $dashboardCards += "<div class='card card-warn'><div class='card-value'>$(@($secSPN).Count)</div><div class='card-label'>Kerberoastable (SPN)</div></div>" }
    if (@($secPwNotReq).Count -gt 0)  { $dashboardCards += "<div class='card card-danger'><div class='card-value'>$(@($secPwNotReq).Count)</div><div class='card-label'>Password Not Required</div></div>" }
}
if ($skipComputers -ne "true") {
    $dashboardCards += "<div class='card'><div class='card-value'>$($script:computerCount)</div><div class='card-label'>Total Computers</div></div>"
    $dashboardCards += "<div class='card card-warn'><div class='card-value'>$($script:computerStale)</div><div class='card-label'>Stale Computers (&gt;$ComputerStaleDays d)</div></div>"
}

$HTMLmessage = @"
<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8" /><meta name="viewport" content="width=device-width, initial-scale=1" />
<title>$subOUonly AD Snapshot Report</title>
<meta name="author" content="AD Snapshot Tool" /><meta name="description" content="$subOUonly AD Snapshot Report - $Version" />
<style>
:root{
  --bg:#f4f6f9; --panel:#ffffff; --ink:#1f2933; --muted:#6b7280; --line:#e3e8ee;
  --brand:#1f6feb; --brand-dark:#163d73; --head:#243b53; --row:#ffffff; --row-alt:#f7f9fc;
}
*{box-sizing:border-box;}
body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.5 "Segoe UI",Roboto,Helvetica,Arial,sans-serif;}
.wrap{max-width:1120px;margin:0 auto;padding:24px;}
.appbar{background:linear-gradient(135deg,var(--brand-dark),var(--brand));color:#fff;border-radius:14px;padding:22px 26px;box-shadow:0 6px 18px rgba(22,61,115,.25);}
.appbar h1{margin:0;font-size:22px;font-weight:600;letter-spacing:.2px;}
.appbar .meta{margin-top:6px;font-size:13px;opacity:.9;}
.appbar .badge{display:inline-block;background:rgba(255,255,255,.18);padding:3px 10px;border-radius:999px;font-size:12px;margin-right:8px;}
.cards{display:flex;flex-wrap:wrap;gap:14px;margin:22px 0;}
.card{flex:1 1 150px;background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:16px 18px;box-shadow:0 1px 3px rgba(16,24,40,.04);}
.card-value{font-size:28px;font-weight:700;color:var(--head);}
.card-value .card-sub{font-size:15px;font-weight:500;color:var(--muted);margin-left:4px;}
.card-label{font-size:12px;color:var(--muted);margin-top:4px;text-transform:uppercase;letter-spacing:.4px;}
.card-warn{border-top:3px solid #e0a800;}
.card-muted{border-top:3px solid #9aa5b1;}
.card-danger{border-top:3px solid #c0392b;}
.section{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:18px 20px;margin-bottom:22px;box-shadow:0 1px 3px rgba(16,24,40,.04);}
.section h2{margin:0 0 14px;font-size:16px;color:var(--head);border-left:4px solid var(--brand);padding-left:10px;}
.section h3{margin:18px 0 8px;font-size:13px;color:var(--head);text-transform:uppercase;letter-spacing:.4px;}
table.kv{width:auto;min-width:440px;margin-bottom:6px;}
table.kv td:first-child{font-weight:600;color:var(--head);width:280px;}
.toolbar{margin-bottom:12px;}
.toolbar input{width:280px;max-width:60%;padding:8px 12px;border:1px solid var(--line);border-radius:8px;font-size:13px;outline:none;}
.toolbar input:focus{border-color:var(--brand);box-shadow:0 0 0 3px rgba(31,111,235,.12);}
.toolbar .count{color:var(--muted);font-size:12px;margin-left:8px;}
table{width:100%;border-collapse:collapse;background:var(--panel);border:1px solid var(--line);border-radius:8px;overflow:hidden;}
thead th{background:var(--head);color:#fff;font-weight:600;font-size:12px;text-transform:uppercase;letter-spacing:.3px;padding:9px 10px;text-align:left;vertical-align:middle;white-space:nowrap;}
table.sortable thead th{cursor:pointer;user-select:none;}
table.sortable thead th:hover{background:#2f4d6e;}
table.sortable thead th.sorted-asc::after{content:" \25B2";font-size:10px;}
table.sortable thead th.sorted-desc::after{content:" \25BC";font-size:10px;}
tbody td{padding:7px 10px;border-bottom:1px solid var(--line);vertical-align:middle;}
tbody tr:nth-child(even){background:var(--row-alt);}
tbody tr:hover{background:#eef4ff;}
td.server{border-top:2px solid #cdd6e0;}
table.drive{width:100%;border:none;background:transparent;}
table.drive th{background:transparent;color:var(--muted);font-size:11px;padding:1px 4px;text-transform:none;}
table.drive td.drive{text-align:right;border:none;padding:1px 4px;}
table.internal{width:340px;}
.tag{display:inline-block;font-size:11px;font-weight:600;padding:1px 7px;border-radius:999px;line-height:1.6;vertical-align:middle;}
.tag-off{background:#f8d7da;color:#842029;}
.tag-lock{background:#e2d6f3;color:#4a2c82;}
.tag-bl{background:#d4edda;color:#155724;}
.tag-risk{background:#fde2cf;color:#8a3b00;}
.tag-svc{background:#d7e6fb;color:#1b4f8a;}
.tag-priv{background:#ede0c8;color:#7a4f00;}
.errlog{color:#b02a37;font-weight:600;margin:8px 0;}
center{font-size:12px;color:var(--muted);}
.footer{text-align:center;color:var(--muted);font-size:12px;padding:10px 0 30px;}
@media print{body{background:#fff;}.appbar{box-shadow:none;}.section,.card{box-shadow:none;}.toolbar{display:none;}}
</style>
<script>
function adSort(table, col, th){
  var tb=table.tBodies[0]; if(!tb) return;
  var rows=Array.prototype.slice.call(tb.rows);
  var dir=th.classList.contains('sorted-asc')?'desc':'asc';
  var hs=table.tHead.rows[0].cells;
  for(var i=0;i<hs.length;i++){hs[i].classList.remove('sorted-asc','sorted-desc');}
  th.classList.add(dir==='asc'?'sorted-asc':'sorted-desc');
  rows.sort(function(a,b){
    var x=(a.cells[col]?a.cells[col].innerText:'').trim().toLowerCase();
    var y=(b.cells[col]?b.cells[col].innerText:'').trim().toLowerCase();
    var nx=parseFloat(x), ny=parseFloat(y); var nums=!isNaN(nx)&&!isNaN(ny);
    var dx=Date.parse(x), dy=Date.parse(y); var dts=!isNaN(dx)&&!isNaN(dy);
    var cmp; if(nums){cmp=nx-ny;} else if(dts){cmp=dx-dy;} else {cmp=x<y?-1:(x>y?1:0);}
    return dir==='asc'?cmp:-cmp;
  });
  for(var r=0;r<rows.length;r++){tb.appendChild(rows[r]);}
}
function adInitSort(){
  var tables=document.querySelectorAll('table.sortable');
  for(var t=0;t<tables.length;t++){(function(table){
    if(!table.tHead) return;
    var hs=table.tHead.rows[0].cells;
    for(var i=0;i<hs.length;i++){(function(idx){
      var th=hs[idx]; th.addEventListener('click',function(){adSort(table,idx,th);});
    })(i);}
  })(tables[t]);}
}
function adFilter(inputId, tableId){
  var box=document.getElementById(inputId); var table=document.getElementById(tableId);
  if(!box||!table||!table.tBodies[0]) return;
  var q=box.value.toLowerCase(); var rows=table.tBodies[0].rows; var shown=0;
  for(var i=0;i<rows.length;i++){
    var hit=rows[i].innerText.toLowerCase().indexOf(q)>-1;
    rows[i].style.display=hit?'':'none'; if(hit)shown++;
  }
  var c=document.getElementById(inputId+'Count'); if(c)c.innerText=shown+' shown';
}
document.addEventListener('DOMContentLoaded',adInitSort);
</script>
</head>
<body>
<div class="wrap">
<div class="appbar">
  <h1>$subOUonly &mdash; Active Directory Snapshot</h1>
  <div class="meta"><span class="badge">$realm</span><span class="badge">$Version</span>Generated: $starttime &ndash; $endtime</div>
</div>
<div class="cards">$dashboardCards</div>
"@

    if ($skipDomainOverview -ne "true" -and $script:domainOverviewHTML) {
        $HTMLmessage += @"
        <div class="section"><h2>Domain Health Overview</h2>
        $($script:domainOverviewHTML)
        </div>
"@  }

    if ($skipAdmin -eq "true") { } else {
        $HTMLmessage +=  @"
        <div class="section"><h2>Active Directory Group Memberships</h2>
        $AdminLG
        </div>
"@  }
    if ($skipServer -eq "true") { } else {
        $errBlock = ""
        if ($errlog) { $errBlock = "<div class='errlog'>$errlog</div>" }
        $HTMLmessage = $HTMLmessage + @"
        <div class="section"><h2>$subOUonly Server Report</h2>
        $serverinfo
        $errBlock
        </div>
"@  }
    if ($skipUsers -eq "true") { } else {
        if ($script:securityFindingsHTML) {
            $HTMLmessage = $HTMLmessage + @"
        <div class="section"><h2>$subOUonly Security Findings</h2>
        $($script:securityFindingsHTML)
        </div>
"@      }
        $HTMLmessage = $HTMLmessage + @"
        <div class="section"><h2>$subOUonly Active Directory User List</h2>
        <div class="toolbar"><input type="text" id="userSearch" placeholder="Filter users..." onkeyup="adFilter('userSearch','tblUsers')" /><span class="count" id="userSearchCount"></span></div>
        $adlogons
        </div>
"@  }
    if ($skipComputers -eq "true") { } else {
        $HTMLmessage = $HTMLmessage + @"
        <div class="section"><h2>$subOUonly Active Directory Computer List</h2>
        <div class="toolbar"><input type="text" id="compSearch" placeholder="Filter computers..." onkeyup="adFilter('compSearch','tblComputers')" /><span class="count" id="compSearchCount"></span></div>
        $adcomputers
        </div>
"@  }
    $HTMLmessage = $HTMLmessage + @"
<div class="footer">Generated: $starttime &ndash; $endtime &bull; $Version</div>
</div>
</body></html>
"@
    $timenow = Get-Date -UFormat %r
    write-host $timenow Completed Message Formatting
########################################################################################################
## END HTML INFORMATION ##

########################################################################################################
## START PDF Engine Detect ##
    if ($WantPDFFile -eq "true") {
        $script:browserExe = Get-ChromiumBrowser
        if ($script:browserExe) {
            Write-Host "PDF engine: $script:browserExe"
        } else {
            Write-Host "PDF requested, but no Microsoft Edge / Google Chrome was found - skipping PDF."
            Write-Host "  Tip: the self-contained HTML report can be saved to PDF from any browser (Ctrl+P -> Save as PDF)."
            $WantPDFFile = "false"
        }
    }
########################################################################################################
## END PDF Engine Detect ##


if ($CreateFile -eq "Y") {
    $timenow = Get-Date -UFormat %r
    write-host $timenow Creating File
    if ([string]::IsNullOrWhiteSpace($outputpath)) { $targetFolder = ([Environment]::GetFolderPath("Desktop")) } else { $targetFolder = $outputpath }
    if (-not (Test-Path -Path $targetFolder)) { New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null }

    $reportFull = Join-Path $targetFolder $filename
    $htmlmessage | Out-File -FilePath $reportFull -Encoding UTF8
    # Marker consumed by the GUI to locate the report regardless of auto-detected naming
    write-host "REPORTFILE::$reportFull"

    if ($WantPDFFile -eq "true") {
        write-host $timenow Creating PDF
        $pdfFull = Join-Path $targetFolder $filenamePDF
        if (Convert-HtmlToPdf -HtmlPath $reportFull -PdfPath $pdfFull -BrowserExe $script:browserExe) {
            write-host "PDF saved: $pdfFull"
        } else {
            write-host "WARNING: PDF generation failed (the HTML report was still saved)."
        }
    }
}
if ($SendEmail -eq "Y") {
    $timenow = Get-Date -UFormat %r
    write-host Sending Email to $SendToList
    $tempHtml = Join-Path $env:temp $filename
    $tempPdf  = Join-Path $env:temp $filenamePDF
    $pdfReady = $false
    if ($WantPDFFile -eq "true") {
        $htmlmessage | Out-File -FilePath $tempHtml -Encoding UTF8
        $pdfReady = Convert-HtmlToPdf -HtmlPath $tempHtml -PdfPath $tempPdf -BrowserExe $script:browserExe
        if (-not $pdfReady) { write-host "WARNING: PDF generation failed - sending email without the PDF attachment." }
    }

    if ($pdfReady) {
        if ($CcList -eq "") { send-mailmessage -From $fromemail -To $SendToList -Subject $subjectline -BodyAsHTML -Body $HTMLmessage -SmtpServer $smtpserver -Attachments $tempPdf }
        else                { send-mailmessage -From $fromemail -To $SendToList -Cc $CcList -Subject $subjectline -BodyAsHTML -Body $HTMLmessage -SmtpServer $smtpserver -Attachments $tempPdf }
    } else {
        if ($CcList -eq "") { send-mailmessage -From $fromemail -To $SendToList -Subject $subjectline -BodyAsHTML -Body $HTMLmessage -SmtpServer $smtpserver }
        else                { send-mailmessage -From $fromemail -To $SendToList -Cc $CcList -Subject $subjectline -BodyAsHTML -Body $HTMLmessage -SmtpServer $smtpserver }
    }

    Remove-Item -LiteralPath $tempHtml -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $tempPdf  -Force -ErrorAction SilentlyContinue
}

}     ###  DO NOT MODIFY
if ($StartupVars) {
        $UserVars = Get-Variable -Exclude $StartupVars -Scope Global
}
