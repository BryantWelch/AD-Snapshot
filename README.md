# AD Snapshot Report v2.0.0
----

This PowerShell tool collects information about Active Directory (AD) objects, including admin membership, server health, user accounts, and computer inventory. It ships with both a command-line script and a modern graphical user interface (GUI) to make it easy to generate comprehensive AD reports. The tool uses standard AD/WMI cmdlets to retrieve data from your domain and produces a clean, self-contained HTML report.

![image](https://github.com/user-attachments/assets/1fc448cc-b4d1-4314-bdcc-8edef2707f74)


## Works in Any Domain (Portable)

This tool auto-detects its environment, so you can drop it onto a machine in **any** Active Directory domain and run it without editing OU paths:

  * **Auto-detects the domain** — domain DN, a domain controller, and the NetBIOS name are discovered automatically (`Get-ADDomain`). Leave the Realm / Domain Controller / Domain Common Name fields blank and it just works.
  * **No hardcoded OU structure** — users, computers, and servers are discovered with domain-wide queries (`Get-ADUser` / `Get-ADComputer`) rather than fixed paths like `OU=Users` or `OU=Servers`. Leave the OU field blank to scan the whole domain, or enter an OU name or full OU distinguished name to scope it.
  * **Servers identified by OS**, not by OU — any computer whose operating system contains "Server" is treated as a server.
  * **Admin groups resolved by name anywhere** in the domain (supports nested membership and multiple comma-separated groups, e.g. `Domain Admins,Enterprise Admins`).
  * **Cross-domain / non-joined support** — tick "Use alternate credentials" in the GUI (or pass `-Credential` and `-Server` on the command line) to report against a different domain or run from a workstation that isn't domain-joined.
  * **Resilient** — unreachable servers, missing OUs, or query errors are logged and skipped instead of aborting the whole report.

## What's New in v2.0.0

  * **Portable across any domain** (see above) with auto-detection and alternate-credential support.
  * **Modern, self-contained HTML report** — redesigned UI with a summary dashboard, soft severity colors, and status badges. No external internet resources are required, so reports render correctly offline and in email.
  * **At-a-glance dashboard** — key metrics (admin members, servers reported, total/stale/disabled users, total/stale computers) shown as cards at the top of every report.
  * **Interactive tables** — click any column header to sort, and use the per-table search box to instantly filter users and computers.
  * **CSV export** — optionally export the user and computer lists as CSV files for spreadsheets and auditing.
  * **Parameterized engine** — `AD-Snapshot.ps1` now accepts parameters, so the GUI runs it directly and it integrates cleanly with Task Scheduler and automation (no fragile script rewriting).
  * **Test AD Connection** — verify the AD module, domain controller reachability, and LDAP bind before running a full report.
  * **Reworked GUI** — header banner, flat buttons, consistent Segoe UI styling, required-field validation, and live streaming progress output.
  * **Bug fixes** — fixed a duplicate-form issue in the GUI, the "Save Report as File" option that never saved, and the broken "Last Login" highlight in the user table.

## Features 

  * **Domain Health Overview** — domain & forest functional levels, the five FSMO role holders, a Domain Controller inventory (OS/IP/site/GC/RODC, with end-of-life OS flagged), the default password & lockout policy, domain trusts, AD Recycle Bin status, and krbtgt / built-in Administrator password age. Risky values are color-coded.
  * **Security Findings** — surfaces common AD hygiene risks among user accounts: Kerberos pre-auth disabled (AS-REP roastable), accounts with SPNs (Kerberoastable), unconstrained delegation, "password not required", protected/privileged (`adminCount=1`), expired-but-enabled, and password-never-expires. Affected accounts are listed and tagged inline in the user table.
  * Retrieves group members for the configured Admin group(s).
  * Retrieves server information for the servers in scope.
    * Collects Uptime, Server Administrators, Remote Desktop Users, IP Address, Local User Accounts, Disk Information (Size, Free, Percent Free), Total Server Count.
  * Retrieves user information for the users in scope.
    * Collects Full Name, SAM Account Name, Description, Last Logon, Last Password Change, Account Creation Date, Account Active or Disabled...and more.
  * Retrieves computer information for the computers in scope.
    * Collects Computer Name, Description, Last Seen on Network, Operating System & Version, BitLocker status...and more. 
     
## Usage 

### GUI Application (Recommended)

  1. Simply double-click the `Launch-AD-Snapshot-GUI.vbs` file to start the application without showing a PowerShell console window
  2. Configure your settings in the GUI:
     * **Configuration Tab**: Set OUs, realm, domain controller, domain common name, admin group, and other core settings
     * **Output Options Tab**: Configure file saving options, PDF generation, and email settings
     * **Run Report Tab**: Execute the report and view progress
  3. Click the "Run Report" button to generate your AD snapshot
  4. Use the "View Report" button to open the generated report
  5. Settings can be saved and loaded for future use
  6. Hover over any field to see helpful tooltips explaining each option

### Command-Line Script

The script now uses a `param()` block, so you can either edit the defaults at the top of `AD-Snapshot.ps1` or pass values on the command line.

  1. Edit the default values in the `param()` block at the top of `AD-Snapshot.ps1` (or supply them as parameters). Key settings:
     * **Active Directory**:
       * `-OUlist` : The OU(s) to run against. Multiple OUs: `"OU1","OU2"`
       * `-realm` : Your organization's realm/name
       * `-DomainCN` : Domain Common Name (e.g. `"DC=contoso,DC=com"`)
       * `-Server` : The domain controller FQDN
       * `-AdminGroup` : The admin group to report on
     * **Output**:
       * `-CreateFile "Y"` saves the HTML report (default Desktop, or set `-outputpath`)
       * `-ExportCSV "true"` also writes Users/Computers CSV files
       * `-WantPDFFile "true"` creates a PDF using your installed Edge/Chrome (no extra software)
     * **Email** (optional): `-SendEmail "Y"`, `-smtpserver`, `-fromemail`, `-defaultSentTo`, `-CcList`
  2. Run it. The simplest case auto-detects everything for the current domain:

```powershell
# Scan the entire current domain (everything auto-detected)
powershell.exe -ExecutionPolicy Bypass -File "C:\path\AD-Snapshot.ps1"
```

```powershell
# Scope to one OU (by name or full DN)
powershell.exe -ExecutionPolicy Bypass -File "C:\path\AD-Snapshot.ps1" -OUlist "Sales"
```

```powershell
# Report against a different / non-joined domain using alternate credentials
$cred = Get-Credential
powershell.exe -ExecutionPolicy Bypass -File "C:\path\AD-Snapshot.ps1" -Server "DC01.contoso.com" -Credential $cred
```

  3. The report is written to your Desktop (or `-outputpath`) and can optionally be emailed and/or saved as PDF/CSV.
     

## Notes 

  * This tool is designed for use on your local machine or a trusted environment.
  * Make sure you have the necessary permissions and credentials to access your AD domain.
  * The tool uses various AD cmdlets, which may require additional installation or configuration.
  * PDF generation uses headless Microsoft Edge or Google Chrome (Chromium), auto-detected at runtime - no extra software to install or configure. Edge ships with Windows 10/11. The self-contained HTML report can also be saved to PDF from any browser with Ctrl+P.

## Requirements

  * Windows PowerShell 5.0 or later
  * Remote Server Administration Tools (RSAT) for Active Directory
    * Can be installed using: `Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0`
  * Microsoft Edge or Google Chrome (optional, for PDF generation - Edge is built into Windows 10/11)

## File Structure

  * **AD-Snapshot.ps1**: The report engine / command-line script (parameterized)
  * **Launch-AD-Snapshot-GUI.vbs**: VBScript launcher for the GUI (recommended entry point)
  * **AD-Snapshot-GUI-Main.ps1**: Main GUI script
  * **AD-Snapshot-GUI.ps1**: Configuration tab components
  * **AD-Snapshot-GUI-Output.ps1**: Output options tab components
  * **AD-Snapshot-GUI-Run.ps1**: Run report tab components
  * **AD-Snapshot-GUI-Functions.ps1**: Helper functions for the GUI
  * **AD-Snapshot-GUI-Launcher.ps1**: PowerShell launcher with prerequisite checks
     

### Credits 
This script was developed by Bryant Welch and is licensed under the [MIT License](https://opensource.org/license/MIT). 
