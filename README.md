# AD Snapshot Report 
----

This PowerShell tool is designed to collect information about Active Directory (AD) objects, including admin membership, user information, and computer information. It provides both a command-line script and a graphical user interface (GUI) to make it easy to generate comprehensive AD reports. The tool uses various AD cmdlets to retrieve data from your domain.

![image](https://github.com/user-attachments/assets/4527492e-8dd5-4cb5-b4ee-2cd63ade70e0)


## Features 

  * Retrieves group memebers for configured Admin group.
  * Retrieves server information for all servers within the configured OU.
    * Collects Uptime, Server Administrators, Remote Desktop Users, IP Address, Local Users Accounts, Disk Information (Size, Free, Percent Free), Total Server Count.
  * Retrieves User information for all user within the configured OU.
    * Collects Full Name, SAM Account Name, Description, Last Logon, Last Password Change, Account Creation Date, Account Active or Disabled...and more.
  * Retrieves computer information for all computers within the configured OU.
    * Collects Computer Name, Description, Last Seen on Network, Operating System & Version...and more. 
     
## Usage 

### GUI Application (Recommended)

  1. Simply double-click the `Launch-AD-Snapshot-GUI.vbs` file to start the application without showing a PowerShell console window
  2. Configure your settings in the GUI:
     * **Configuration Tab**: Set OUs, realm, admin group, and other core settings
     * **Output Options Tab**: Configure file saving options, PDF generation, and email settings
     * **Run Report Tab**: Execute the report and view progress
  3. Click the "Run Report" button to generate your AD snapshot
  4. Use the "View Report" button to open the generated report
  5. Settings can be saved and loaded for future use

### Command-Line Script

  1. Open and edit the AD-Snapshot.ps1 file with a text editor like Notepad++, specifying the necessary parameters:
     * $OUlist: The OU you want to run this script against
     * $realm: The Realm name of your organization
     * $AdminGroup: The OU Admin group you want the Admin section to use
     * $DomainCN: Domain Common Name (Example "DC=example,DC=org")
     * $Server: The Domain Controller of your organization
  2. By default the script will save a file in the location where the .ps1 is ran. If you want to email this file or create a .pdf, specify the optional parameters:
     * $SendEmail: "Y" will send a report email. "N" will not send a report email
     * $smtpserver: The SMTP server that will be used to email the resulting report
     * $fromemail: The from address that will be used to send the report email
     * $defaultSentTo: The to address that will receive the report email - Sample: "email.one@example.com", "email.two@example.com"
     * $WantPDFFile: true will create a PDF file saved to a folder (and/or attached to email)   
  3. Run the script using PowerShell
  4. The script will output the collected data in a human-readable format
     

## Notes 

  * This tool is designed for use on your local machine or a trusted environment.
  * Make sure you have the necessary permissions and credentials to access your AD domain.
  * The tool uses various AD cmdlets, which may require additional installation or configuration.
  * For PDF generation, you need to download and install wkhtmltopdf from http://wkhtmltopdf.org/downloads.html

## Requirements

  * Windows PowerShell 5.0 or later
  * Remote Server Administration Tools (RSAT) for Active Directory
    * Can be installed using: `Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0`
  * wkhtmltopdf (optional, for PDF generation)

## File Structure

  * **AD-Snapshot.ps1**: The original command-line script
  * **Launch-AD-Snapshot-GUI.vbs**: VBScript launcher for the GUI (recommended entry point)
  * **AD-Snapshot-GUI-Main.ps1**: Main GUI script
  * **AD-Snapshot-GUI.ps1**: Configuration tab components
  * **AD-Snapshot-GUI-Output.ps1**: Output options tab components
  * **AD-Snapshot-GUI-Run.ps1**: Run report tab components
  * **AD-Snapshot-GUI-Functions.ps1**: Helper functions for the GUI
  * **AD-Snapshot-GUI-Launcher.ps1**: PowerShell launcher with prerequisite checks
     

### Credits 
This script was developed by Bryant Welch and is licensed under the [MIT License](https://opensource.org/license/MIT). 
