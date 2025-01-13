# AD Snapshot Report 
----

This PowerShell script is designed to collect information about Active Directory (AD) objects, including admin membership, user information, and computer information. The script uses various AD cmdlets to retrieve data from your domain. 

## Features 

  * Retrieves group memebers for configured Admin group.
  * Retrieves server information for all servers within the configured OU.
    * Collects Uptime, Server Administrators, Remote Desktop Users, IP Address, Local Users Accounts, Disk Information (Size, Free, Percent Free), Total Server Count.
  * Retrieves User information for all user within the configured OU.
    * Collects Full Name, SAM Account Name, Description, Last Logon, Last Password Change, Account Creation Date, Account Active or Disabled...and more.
  * Retrieves computer information for all computers within the configured OU.
    * Collects Computer Name, Description, Last Seen on Network, Operating System & Version...and more. 
     
## Usage 

  1. Save this script as a PowerShell file (e.g., AD-Snapshot.ps1)
  2. Open and edit the .ps1 file with a text editor like Notepad++, specifying the necessary parameters:
     * $OUlist: The OU you want to run this script against
     * $realm: The Realm name of your organizaion
     * $AdminGroup: The OU Admin group you want the Admin section to use
     * $DomainCN: Domain Common Name (Example "DC=example,DC=org")
     * $Server: The Domain Controller of your organization
  3. By Defalut the script will save a file in the location where the .ps1 is ran. If you want to email this file or create a .pdf, specify the below optional parameters:
     * $SendEmail: "Y" will send a report email. "N" will not send a report email
     * $smtpserver: The SMTP (Simple Mail Transfer Protocol) that will be used to email the resulting report
     * $fromemail: The from address that will be used to send the report email
     * $defaultSentTo: The to address that will receive the report email - Sample: "email.one@example.com", "email.two@example.com"
     * $WantPDFFile: true will create a PDF file saved to a folder. (and/or attached to email)   
  4. The script will output the collected data in a human-readable format
     

## Notes 

  * This script is designed for use on your local machine or a trusted environment.
  * Make sure you have the necessary permissions and credentials to access your AD domain.
  * The script uses various AD cmdlets, which may require additional installation or configuration.
     

### Credits 
This script was developed by Bryant Welch and is licensed under the [MIT License](https://opensource.org/license/MIT). 
