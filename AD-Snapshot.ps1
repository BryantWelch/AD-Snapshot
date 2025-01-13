<######################################################################################################
## Active Directory Snapshot Report Version 1.0.0 last updated on 1/13/2025
## updated by Bryant Welch 
#######################################################################################################

  This script uses the Remote Server Administration Tools for Windows ( You will need to install if not already done )
      https://4sysops.com/archives/how-to-install-the-powershell-active-directory-module/
  Here is a link to the installers for non server OSs:
      https://www.microsoft.com/en-us/search/result.aspx?q=Remote%20Server%20Administration%20Tools%20for%20Windows%20Server%202008&form=DLC


  There is the option to save the Report to PDF and/or email a PDF
  You will need to download and install wkhtmltopdf.exe  for that funcion to work ...http://wkhtmltopdf.org/downloads.html
      64bit = https://downloads.wkhtmltopdf.org/0.12/0.12.5/wkhtmltox-0.12.5-1.msvc2015-win64.exe
      32bit = https://downloads.wkhtmltopdf.org/0.12/0.12.5/wkhtmltox-0.12.5-1.msvc2015-win32.exe

  To execute using Task Scheduler use "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" "C:\path to location\AD-Snapshot.ps1"
  
## Set-ExecutionPolicy RemoteSigned <<<>>> You will want to run this from an elevated PowerShell window if you have never ran PowerShell Scripts before.

  If the firewall is blocking it use an Elevated command prompt and paste the below string:
  netsh advfirewall firewall set rule group="Windows Management Instrumentation (WMI)" new enable=yes

  If you are missing only the Group information, run this in an Elevated command prompt:
  netsh advfirewall firewall set rule name="File and Printer Sharing (SMB-In)" new enable=yes
########################################################################################################>
$StartupVars = @()        ###  DO NOT MODIFY
$StartupVars = Get-Variable | Select-Object -ExpandProperty Name  ###  DO NOT MODIFY
$StartupVars += "PSItem"  ###  DO NOT MODIFY
########################################################################################################

$OUlist = "XXX"            # Change to the OU you want to run this script against 
      ###    If you have severial OUs, you can list them like: "MARSHA1","MARSHA2","MARSHA3"  ...Do NOT run the Server Section across the WAN
$realm = "example"                # Change to your organization's Realm name

$CreateFile = "Y"            # Y will save the report as file.
      # Saving this report to file is the recomended action and will allow for sortable user and computer lists.
$outputpath = ""             # If this is blank, the report will be saved on your desktop.
$WantPDFFile = "false"        # true will create a PDF file saved to a folder. (and/or attached to email)
      # If you want PDFs created by this script you need to download and install wkhtmltopdf.exe  (see above)
$PDFConverter = "c:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe" # Path to wkhtmltopdf.exe 

$SendEmail = "N"                                # "Y" will send a report email. "N" will not send a report email
$CcList = ""                                    # additional recipients for the email - Sample: "email.one@example.com", "email.two@example.com"
$smtpserver = "smtp.example.org"				# change to the smtp server you want used to email the report

$ServerUpTimeAlarm = 30                         # number of days on before marking red 
$ComputerStaleDays = 80                         # number of days a computer is allowed off of the network
$UserStaleDays = 80                             # number of days a User is allowed off of the network
$PCsort = "name"                                # use "name" will sort by computer name, otherwise date Last Seen is used
$Usersort = "name"                              # use "name" will sort by user name, otherwise date of last Login is used
$HideDisabledUsers = "false"                    # when "true" will not show any user that is disabled 
$listALLComputers = "true"                      # when "true" servers will also display with computers

##  Used to skip sections
$skipadmin = "false"                            # when "true" admin section will be skipped
$skipServer = "false" 							# when "true" server section will be skipped
$skipUsers = "false"    						# when "true" users section will be skipped
$skipComputers = "false" 						# when "true" computers section will be skipped
$skipBitlockerStatus = "true"					# when "true" bitlocker section will be skipped
########################################################################################################
##   DO NOT CHANGE THE BELOW LINES
########################################################################################################
Import-Module ActiveDirectory 
foreach ($subOU in $OUlist) {         ###  DO NOT MODIFY
try {$subOUonly = $subOU.substring(0,$subOU.IndexOf(",",1))           ###  DO NOT MODIFY
} catch { $subOUonly = $subOU }       ###  DO NOT MODIFY
$starttime = Get-Date
########################################################################################################
##   DO NOT CHANGE THE ABOVE LINES
########################################################################################################


$fromemail = "user1@example.com"     # Change to the from address that will be used to send the report email
$defaultSentTo = "user2@example.com" # Change to the to address that will receive the report email - Sample: "email.one@example.com", "email.two@example.com"
[string[]] $SendToList =  "$defaultSentTo"      # The email will be sent to
$subjectline = "$subOUonly AD Snapshot Report"  # Subject line
$AdminGroup = "Example Admin"             # Change to the OU Admin group you want the Admin section to use
$filename = $subOUonly + " AD Snapshot Report.html"		# html filename
$filenamePDF = $subOUonly + " AD Snapshot Report.pdf"	# pdf filename


#########################################################################################################
### # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # ###
###                                                                                                   ###
###            DO NOT CHANGE ANYTHING BELOW THIS LINE                                                 ###
###                                                                                                   ###
### # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # ###
#########################################################################################################

if ( $realm -eq "example"){$DomainCN = "DC=example,DC=org" ; $Server = "domaincontroller.example.org"}
	# Change the above options to match your organization

$Version = "AD Snapshot Report Version 1.0.0"
#$DebugPreference = "Continue"
$DebugPreference = "SilentlyContinue"
#$VerbosePreference = "Continue"
$VerbosePreference = "SilentlyContinue"
$needsServer = ""
########################################################################################################
## BEGIN ADMIN MEMBERSHIP INFORMATION ##
if ($skipAdmin -eq "true") { } else {
    $timenow = Get-Date -UFormat %r
    write-host $timenow Starting Admin Membership collection
    $LDAPAdminLG = "LDAP://cn=$AdminGroup,OU=Groups,OU=$subOU,$DomainCN"
        $AdminMembersGroupList = @()
        $AdminLG = "<table class='internal'><th>$AdminGroup Members</th><tr><td>" 
        ([adsi]"$LdapAdminLG").member | 
        ForEach-Object { $AdminMembersGroupList += ,@( ([adsi]"LDAP://$_" | sort displayName | select -expand CN  ) ) }
        $AdminMembersGroupList = $AdminMembersGroupList | sort
        $num = $AdminMembersGroupList.Count
        foreach($member in $AdminMembersGroupList) { $AdminLG += $member + "<br />" }
        $AdminLG += "</td></tr></table><center>$num Users Found.</center><br>"

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
       # create computer array from OU container
       $servers = @()
         $ouquery1= [ADSI]"LDAP://OU=Servers,OU=$subOU,$DomainCN"  
            foreach ($child in $ouquery1.psbase.Children) { if ($child.ObjectCategory -like '*computer*') { $servers += $child.Name } }
         $ouquery2= [ADSI]"LDAP://OU=Windows Task Scheduler Enabled,OU=Servers,OU=$subOU,$DomainCN"  
            foreach ($child in $ouquery2.psbase.Children) { if ($child.ObjectCategory -like '*computer*') { $servers += $child.Name } }
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
    $serverinfo = "<table><th>Server</th><th><table class='drive'><th style='width: 20px;'></th><th>Size</th><th>Free</th><th style='width: 60px;'>%</th><th style='width: 10px;'></th></table></th><th>Uptime</th><th>Administrors<br/>Members</th><th>RDP<br/>Members</th><th>Local<br/>Accounts</th>"
    foreach ($server in $filteredservers) {
    $timenow = Get-Date -UFormat %r
    write-host $timenow Getting information for $server
                $UptimeAlarm = "#ebebeb"
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
                if ( $d -gt $ServerUpTimeAlarm ) { $UptimeAlarm = "#ff0000" }
                # Server Administrators       
                $serverAdminCount = 0
                $admin = ""
                $group = "Group Error"
                $GMembers = "GMembers Error"
                $group = [ADSI]("WinNT://$server/Administrators,group")               
                $GMembers = $group.psbase.invoke("Members")
                $GMembers | ForEach-Object { $admin = $admin + $_.GetType().InvokeMember("Name",'GetProperty', $null, $_, $null) + "<br>"  }
                Write-Verbose "$server Admins $admin"
                # Server Server Remote Desktop Users
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
                        $LocalAccts += "$LocalUserName<IMG SRC='http://upload.wikimedia.org/wikipedia/commons/1/14/Red_x_small.PNG' width='11' height='11' /><small>Disabled</small><br/>"
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
                        $serverDriveinfo = $serverDriveinfo + "<tr bgcolor='#FF0000'><td class='drive'>$Drive</td><td class='drive'>$Size</td><td class='drive'>$Free</td><td class='drive'>$PercentFree%</td></tr>"
                        } else {
                        $serverDriveinfo = $serverDriveinfo + "<tr><td class='drive'>$Drive</td><td class='drive'>$Size</td><td class='drive'>$Free</td><td class='drive'>$PercentFree%</td></tr>"
                    }
                } 
            }                        
            $serverDriveinfo = $serverDriveinfo + "</tr></table>"
            $serverinfo = $serverinfo + "<tr><td class='server'>$SystemName $ServerIPlist</td><td class='server'>$serverDriveinfo</td><td class='server' style='background-color:$UptimeAlarm;' >$ServerUptime</td><td class='server'>$admin</td><td class='server'>$RDP</td><td class='server'>$LocalAccts</td></tr>"

        } 
    
    $serverinfo = $serverinfo + "</table>$num Servers reported out of $numfound Servers found in AD.<br>"

    $timenow = Get-Date -UFormat %r
    write-host $timenow Completed Server collection
}
########################################################################################################
## END SERVER INFORMATION  ## 

########################################################################################################
## BEGIN AD USER INFORMATION ##
if ($skipUsers -eq "true") { } else {
    $timenow = Get-Date -UFormat %r
    $adlogons = "<table class='sortable'><th><u>Name</u></th><th><u>Login</u></th><th><u>Description</u></th><th><u>Last <br/>Login</u></th><th class ='norm'><u>Password <br/>Changed</u></th><th class ='norm'><u>When<br/>Created</u></th><th class ='norm'><u>Never <br/>Expire</u></th><th class ='norm'><u>Active</u></th>"
    $UserRecinfo = @()
    $UserRecinfo_sort = @()
   write-host $timenow Starting User collection
    $ADSearch = New-Object System.DirectoryServices.DirectorySearcher
    $ADSearch.PageSize = 200
    $ADSearch.SearchScope = "subtree"
    $ADSearch.SearchRoot = "LDAP://OU=Users,OU=$subOU,$DomainCN"
    write-Verbose "LDAP://OU=Users,OU=$subOU,$DomainCN"
    $ADSearch.Filter = "(&(objectCategory=person)(objectClass=user))"
    $ADSearch.PropertiesToLoad.Add("sAMAccountName")
    $ADSearch.PropertiesToLoad.Add("displayName")
    $userObjects = $ADSearch.FindAll()
    $num =  $userObjects.Count
   if ($userObjects -eq "") { $ErrorMSG = "The Search for users in OU=Users,OU=" + $subOU + "," + $DomainCN + " came up blank!" ; Write-Error $ErrorMSG  } else {
    foreach ($user in $userObjects) {
    $sam = $user.Properties.Item("sAMAccountName")[0].ToString()
    $name = $user.Properties.Item("displayName")[0].ToString()
    $timenow = Get-Date -UFormat %r
    write-host $timenow Collecting information for $name
     try { $u = Get-ADUser $sam -Properties WhenCreated, LastLogon, LastLogonTimeStamp, userAccountControl, logonCount, displayName, description, LockoutTime, pwdLastSet, givenName, sn 
     } catch {
     $u = Get-ADUser $sam -Server $Server -Properties WhenCreated, LastLogon, LastLogonTimeStamp, userAccountControl, logonCount, displayName, description, LockoutTime, pwdLastSet, givenName, sn 
     $needsServer = "true"}
     $description = $u.description
     if($description){$description = $description.ToString()} else {$description =""}
     $userAC = $u.userAccountControl
     $lockout = $u.LockoutTime
     $logonCount = $u.logonCount
     $pwdLastSet = [DateTime]::FromFileTime([Int64] $u.pwdLastSet)   
     $nameBG = "#efefef"
     $logonBG = "#efefef"
     $ActiveUser = "A"
### WhenCreated
     $ADWhenCreated = [DateTime]$u.WhenCreated
     $WhenCreatedYear = $ADWhenCreated.Year
     $WhenCreatedMonth = $ADWhenCreated.Month.ToString("00")
     $WhenCreatedDay = $ADWhenCreated.Day.ToString("00")
     $WhenCreated1 = "$WhenCreatedMonth/$WhenCreatedDay/$WhenCreatedYear"
     $WhenCreatedBG = "#efefef"
     $ADlastLogon = [DateTime]::FromFileTime([Int64] $u.LastLogon)
     $ADlastLogonTimeStamp = [DateTime]::FromFileTime([Int64] $u.LastLogonTimeStamp)
     if ($ADlastLogon -gt $ADlastLogonTimeStamp) {$lastLogon = $ADlastLogon} else {$lastLogon = $ADlastLogonTimeStamp}
     $lastLogonYear = $lastLogon.Year ; $lastLogonMonth = $lastLogon.Month.ToString("00") ; $lastLogonDay = $lastLogon.Day.ToString("00")
     $lastLogon1 = "$lastLogonMonth/$lastLogonDay/$lastLogonYear"
     if ($userAC -eq 66048) { $neverexpire = "X" ; $expireBG = "#FF0000" } else { $expireBG = "#efefef" ; $neverexpire = " "}
     if ($userAC -eq 66050) { $neverexpire = "X" ; $expireBG = "#FF0000" ; $nameBG = "#00C4F4"; $name = $name + "<IMG SRC='http://upload.wikimedia.org/wikipedia/commons/1/14/Red_x_small.PNG' width='11' height='11' /><small>Disabled</small>" ; $ActiveUser = "D" }
     if ($userAC -eq 514) {$nameBG = "#00C4F4"; $name = $name + "<IMG SRC='http://upload.wikimedia.org/wikipedia/commons/1/14/Red_x_small.PNG' width='11' height='11' /><small>Disabled</small>" ; $ActiveUser = "D" }
     if ($lockout -gt 2) {$nameBG = "#880000"; $name = $name + "<IMG SRC='http://upload.wikimedia.org/wikipedia/commons/d/d2/Padlock-closed.png' width='20' height='20' />" }
     $now = Get-Date
     if ($lastLogon.Year -lt 2000)  {
         $logonBG = "#FF0000"  ## Users who never logged in will flag as red 
         $lastLogon1 = "Never" 
         $Uptime = $now - $ADWhenCreated
         if($Uptime.Days -gt $UserStaleDays) { $WhenCreatedBG = "#FF0000" } ## Users will flag as red if they were created over $UserStaleDays
     }   
     $Uptime = $now - $lastLogon
     $d = $Uptime.Days
     if ($d -gt $UserStaleDays){ $logonBG = "#FF0000"}	## Users will flag as red if they are over $UserStaleDays

     if ($pwdLastSet.Year -gt 2000) {
        $pwdLastSetYear= $pwdLastSet.Year ; $pwdLastSetMonth = $pwdLastSet.Month.ToString("00") ; $pwdLastSetDay = $pwdLastSet.Day.ToString("00")
        $pwdLastSet1 = "$pwdLastSetMonth/$pwdLastSetDay/$pwdLastSetYear" 
        $pwdage = $now - $pwdLastSet
        $d = $pwdage.Days
        if ($neverexpire -eq "X") {
           if ($d -gt "89"){$pwdBG = "#FFFF00"} else {$pwdBG = "#efefef"}
           if ($d -gt "360"){$pwdBG = "#FF0000"}
        } else { if ($d -gt "82"){$pwdBG = "#FFFF00"} else {$pwdBG = "#efefef"} 
           if ($d -gt "82") {              ## One week until expiration
              $pwdBG = "#FFFF00" }         ## Highlight yellow
              else {
              $pwdBG = "#efefef" }
           if ($d -gt "89") {              ## Expired
           $pwdBG = "#FF0000" }         ## Highlight red
        }  }
     else {
        $pwdLastSet1 = "Never" ;$pwdBG = "#FF0000" } 
    


     write-Verbose  "Password changed on $pwdLastSet - $d days ago"
     if ($HideDisabledUsers -eq "true" -and $ActiveUser -eq "D") {
        write-Verbose "Disabled account for $name not displayed on report."
        } Else {
        $UserRecinfo += ,@( $name, $sam, $description, $lastLogon1, $neverexpire, $UserRecBG, $expireBG,$nameBG,$WhenCreated1,$pwdLastSet1,$pwdBG,$lastLogon,$ActiveUser,$WhenCreatedBG)  
        ##                   0      1     2             3            4             5           6         7       8             9            10     11         12          13
        }
     }
     if ($Usersort -eq "name" ) {
     $UserRecinfo_sort = $UserRecinfo | sort-object @{Expression={$_[0]}; Ascending=$true} } else{
     $UserRecinfo_sort = $UserRecinfo | sort-object @{Expression={$_[11]}; Ascending=$true} }
     foreach($UserRec in $UserRecinfo_sort) {            
            $adlogons = $adlogons +  "<tr><td style='background:" + $UserRec[7] + ";'>" + $UserRec[0] + "</td><td>" + $UserRec[1] + "</td><td>" + $UserRec[2] + "</td><td style='background:" + $UserRec[5] + ";'><center>" + $UserRec[3] + "|" + "</center></td><td style='background:" + $UserRec[10] + ";'><center>" + $UserRec[9] + "|" + "</center></td><td style='background:" + $UserRec[13] + ";'><center>" + $UserRec[8] + "</center></td><td style='background:" + $UserRec[6] + ";'><center>" + $UserRec[4] + "</center></td><td><center>" + $UserRec[12] + "</center></td></tr>"
     }
     $adlogons = $adlogons + "</table><center>$num Users Found.</center><br>" }
  
    $timenow = Get-Date -UFormat %r
    write-host $timenow Completed AD User collectionn
  }
########################################################################################################
## END AD USER INFORMATION ##

########################################################################################################
## BEGIN AD COMPUTER INFORMATION ##
if ($skipComputers -eq "true") { } else {
    $timenow = Get-Date -UFormat %r
    write-host $timenow Starting Computer collectionn
    $now = Get-Date
    $num = 0
    $needsServer = "false"
    Try { Get-ADComputer -Filter 'ObjectClass -eq "computer"' -SearchBase "OU=Computers,OU=$subOU,$DomainCN"
    } catch { Get-ADComputer -Filter 'ObjectClass -eq "computer"' -Server $Server -SearchBase "OU=Computers,OU=$subOU,$DomainCN" 
     $needsServer = "true" }
    $script:computerinfo = @()
    $adcomputers = "<table class='sortable'><th><u>Name</u></th><th><u>Description</u></th><th><u>Last<br/>Seen</u></th><th><u>Operating System</u></th><th><u>OS<br/>Version</u></th>"
        function AD_CompInfo {
        Param($Computer)
        
        #Check if the Computer Object exists
        if ($needsServer -eq "false") {$c = Get-ADComputer -Filter {cn -eq $Computer} -Property SamAccountName, cn, lastLogon, lastLogonTimeStamp, displayName, description, operatingsystem, operatingsystemversion, whenCreated, logonCount, msTPM-OwnerInformation, msTPM-TpmInformationForComputer}
        if ($needsServer -eq "true") {$c = Get-ADComputer -Filter {cn -eq $Computer} -Server $Server -Property SamAccountName, cn, lastLogon, lastLogonTimeStamp, displayName, description, operatingsystem, operatingsystemversion, whenCreated, logonCount, msTPM-OwnerInformation, msTPM-TpmInformationForComputer}
        #if($c -eq $null){ Write-Host "Error..." } else {
            $OSBG = "#ebebeb" ; $alarm = 0
            $TPM = ""
            $samCN = $c.cn
            $timenow = Get-Date -UFormat %r
            write-host $timenow Collecting information for $samCN
            $BitlockerKey = ""
            $lastLogon = ""
            $description = $c.description
            $sam = $c.SamAccountName.ToString() 
            $OS = $c.operatingsystem
            $OSV = $c.operatingsystemversion
            if ($OS) {$OS = $OS.ToString()} else { $OS = "" ; $OSBG = "#888888"}
            $seenBG = "#efefef"

        ########## Start Bitlocker ################################
            if ($skipBitlockerStatus -eq "true") { Write-Verbose "$samCN - $description - $OS "  } else {
                $BitLockerDataHash = @{}
                $BitLockerDataArray = @("BitCN", "BitKey")
                #Check if the computer object has had a BitLocker Recovery Password
                try { $BitlockerObject = Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase $c.DistinguishedName -Properties 'msFVE-RecoveryPassword' | Select-Object -Last 1
                } catch { $BitlockerObject = Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -Server $Server -SearchBase $c.DistinguishedName -Properties 'msFVE-RecoveryPassword' | Select-Object -Last 1 }
                if($BitlockerObject.'msFVE-RecoveryPassword'){
                    $BitLockerKey = $BitLockerObject.'msFVE-RecoveryPassword'
                    $BitLockerDataArray += @( $Computer, $BitLockerKey)
                    $BitLockerDataHash += @{ $Computer = $BitLockerKey}
                    Write-Verbose "$samCN - $description - $OS - Bitlocker Key= $BitlockerKey"
                    $TPM = "true"
                }else{ Write-Host -ForegroundColor Red "$Computer - WARNING!!! - The Bitlocker key is missing. - WARNING!!!" }
            }
            if ($TPM -eq "true") { $OSBG = "#00BB22"; $OS += "  <IMG SRC='http://upload.wikimedia.org/wikipedia/commons/7/7a/Encrypted.png'  width='16' height='16' >"}
        ########## End Bitlocker   ################################
            $ADlastLogon = [DateTime]::FromFileTime([Int64] $c.LastLogon)
            $ADlastLogonTimeStamp = [DateTime]::FromFileTime([Int64] $c.LastLogonTimeStamp)
            if ($ADlastLogon -gt $ADlastLogonTimeStamp) {$lastLogon = $ADlastLogon} else {$lastLogon = $ADlastLogonTimeStamp}
            $lastLogonYear = $lastLogon.Year ; $lastLogonMonth = $lastLogon.Month.ToString("00") ; $lastLogonDay = $lastLogon.Day.ToString("00")
            $lastLogon1 = "$lastLogonMonth/$lastLogonDay/$lastLogonYear"
                $age = $now - $lastLogon
                if($age.Days -gt $ComputerStaleDays) { $seenBG = "#FF0000" } ## Users will flag as red if they were created over $UserStaleDays


            if ($lastLogon.Year -lt 2000)  {
                $logonBG = "#FF0000"  ## Users who never logged in will flag as red 
                $lastLogon1 = "Never" 
                $age = $now - $c.whenCreated
                if($age.Days -gt $ComputerStaleDays) { $seenBG = "#FF0000" } ## Users will flag as red if they were created over $UserStaleDays
                if ($c.logonCount -lt 10 ) {
                    $logonBG = "#FFFF00"  ## New Computers will flag as Yellow 
                    $lastLogonYear = $c.whenCreated.Year ; $lastLogonMonth = $c.whenCreated.Month.ToString("00") ; $lastLogonDay = $c.whenCreated.Day.ToString("00")
                    $lastLogon1 = "Created<br />$lastLogonMonth/$lastLogonDay/$lastLogonYear" 
                } 
            }  
        #}
        $script:computerinfo += ,@( $samCN, $description, $lastLogon1, $seenBG, $OS, $OSV, $OSBG)
        #                     0       1             2            3        4    5      6
        
     }

     if ($needsServer -ne "true"){
        Get-ADComputer -Filter 'ObjectClass -eq "computer"' -SearchBase "OU=Computers,OU=$subOU,$DomainCN" | foreach-object {
            $num ++
            $Computer = $_.name
            AD_CompInfo $Computer $script:computerinfo
        }
        if ($listALLComputers -eq "true") {
            Get-ADComputer -Filter 'ObjectClass -eq "computer"' -SearchBase "OU=Member Servers,OU=$subOU,$DomainCN" | foreach-object {
            $num ++
            $Computer = $_.name
            AD_CompInfo $Computer $script:computerinfo
            }
            Get-ADComputer -Filter 'ObjectClass -eq "computer"' -SearchBase "OU=Unmanaged Servers,OU=$subOU,$DomainCN" | foreach-object {
            $num ++
            $Computer = $_.name
            AD_CompInfo $Computer $script:computerinfo
            }
        }
     }else{
        Get-ADComputer -Filter 'ObjectClass -eq "computer"' -Server $Server -SearchBase "OU=Computers,OU=$subOU,$DomainCN" | foreach-object {
            $num ++
            $Computer = $_.name
            AD_CompInfo $Computer $script:computerinfo
        }
        if ($listALLComputers -eq "true") {
            Get-ADComputer -Filter 'ObjectClass -eq "computer"' -Server $Server -SearchBase "OU=Member Servers,OU=$subOU,$DomainCN" | foreach-object {
            $num ++
            $Computer = $_.name
            AD_CompInfo $Computer $script:computerinfo
            }
            Get-ADComputer -Filter 'ObjectClass -eq "computer"' -Server $Server -SearchBase "OU=Unmanaged Servers,OU=$subOU,$DomainCN" | foreach-object {
            $num ++
            $Computer = $_.name
            AD_CompInfo $Computer $script:computerinfo
            }
        }
      }



         if ($PCsort -eq "name" ){
         $computerinfo_sort = $script:computerinfo | sort-object @{Expression={$_[0]}; Ascending=$true} } else {
         $computerinfo_sort = $script:computerinfo | sort-object @{Expression={$_[2]}; Ascending=$true} }
     foreach($computer in $computerinfo_sort) { $adcomputers += "
     <tr><td>" + $computer[0] + "</td><td>&nbsp; &nbsp;" + $computer[1] + "</td><td style='background:" + $computer[3] + "; width:90px;'>" + $computer[2] + "</td><td style='background:" + $computer[6] + "; width:170px;'>" + $computer[4] + "</td><td style='width:90px;'>" + $computer[5] + "</td></tr>" }  
     
     $adcomputers+= "</table><center>$num Computers Found.</center><br>"

    $timenow = Get-Date -UFormat %r
    write-host $timenow Completed Computer collectionn
}
########################################################################################################
## END AD COMPUTER INFORMATION ##    


########################################################################################################
## BEGIN HTML INFORMATION ##
    $timenow = Get-Date -UFormat %r
    write-host $timenow Starting Message Formatting
$endtime = Get-Date -UFormat %T
$HTMLmessage = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en-us"><head><title>$subOUonly AD Snapshot Report</title>
<meta name="author" content="Daniel Kisrow" /><meta name="description" content="$subOUonly AD Snapshot Report - $Version " />
    <script type="text/javascript" src="http://www.kryogenix.org/code/browser/sorttable/sorttable.js"></script> 
    <meta name="Source for above file" content="http://freepages.genealogy.rootsweb.ancestry.com/~gearyfamily/expression-web/sort-table.html" />
<style type="text/css">
body{font: .8em "Lucida Grande", Tahoma, Arial, Helvetica, sans-serif;}
ol{margin:0;padding: 0 1.5em;}
table{color:#000000;background:#ebebeb;border-collapse:collapse;width:960px;padding: 0px;border:1px solid #000000;vertical-align:top;}
table.drive{color:#000000;background:#ebebeb;border-collapse:collapse;width:170px;padding: 0px;border:none;vertical-align:top;text-align:left;}
table.external{color:#000000;background:#ffffff;border-collapse:collapse;width:960px;padding: 0px;border:0px solid #FFF;vertical-align:top;}
table.external2{color:#000000;background:#ffffff;border-collapse:collapse;width:318px;padding: 0px;border:0px solid #FFF;vertical-align:top;}
table.internal{color:#000000;background:#ebebeb;border-collapse:collapse;width:318px;padding: 0px;border:1px solid #000000; vertical-align:top;}
th{padding:0px;font-size:120%;text-align:center;background:#999999;}
th.norm {font-size: 100%;}
th.small {font-size: 80%;}
th.rotate {height: 140px;white-space: nowrap;}
th.rotate > div {transform: translate(25px, 51px) rotate(315deg);width: 30px;}
th.rotate > div > span {border-bottom: 1px solid #ccc;padding: 5px 10px;}
td{vertical-align:middle; text-align:left}
td.external{vertical-align:top;}
td.external2{vertical-align:top;}
td.internal{vertical-align:top;}
td.internal2{vertical-align:top;}
td.server{border-top:2px solid black;}
td.drive{text-align:right;cellpadding:0;cellspacing:0;}

img{vertical-align: middle;}
tfoot{}
tfoot td{padding-bottom:1.5em;}
tfoot tr{}
a href{color:#000000;}
u{color:blue;text-decoration:underline;}
#middle{background-color:#900;}
.MsoChpDefault
	{mso-style-type:export-only;
	mso-default-props:yes;
	font-size:10.0pt;
	mso-ansi-font-size:10.0pt;
	mso-bidi-font-size:10.0pt;}
@page WordSection1
	{size:8.5in 11.0in;
	margin:.3in .3in .3in .3in;
	mso-header-margin:.3in;
	mso-footer-margin:.3in;
	mso-paper-source:0;}
div.WordSection1
	{page:WordSection1;}
-->
</style> <!border-bottom:1px dotted #FFF;>
</head>
<body bgcolor="white"><center><div style="font-family:Arial;font-size:22px;font-weight:bold;">$subOUonly Active Directory Snapshot Report</div>
<b>Report Generated: $starttime - $endtime</b><br /><a name="grp" /><br />
"@


    if ($skipAdmin -eq "true") { } else {     
        $HTMLmessage +=  @"
        <font color="black" face="Arial, Verdana" size="3"><b>Active Direcory Group Memberships</b></font><br />
        <table class='external' bgcolor="#FFFFFF" ><tr>
        <td class='external2' valign='top'>$AdminLG</td>
        <td class='external2' valign='top'></td></tr></table>
"@  }  
    if ($skipServer -eq "true") { } else {
        $HTMLmessage = $HTMLmessage + @"
        <font color="black" face="Arial, Verdana" size="3"><b>$subOUonly Server Report</b></font><br />
        $serverinfo
        <font color="#FF0000"><b>$errlog</b></font><br />
"@  }
    if ($skipUsers -eq "true") { } else {
        $HTMLmessage = $HTMLmessage + @"
        <font color="black" face="Arial, Verdana" size="3"><b>$subOUonly Active Directory User List</b></font><br />
        $adlogons<br/>
"@  }
    if ($skipComputers -eq "true") { } else {
        $HTMLmessage = $HTMLmessage + @"
        <font color="black" face="Arial, Verdana" size="3"><b>$subOUonly Active Directory Computer List</b></font><br />
        $adcomputers<br/>
"@  }
    $HTMLmessage = $HTMLmessage + @"
    Generated: $starttime - $endtime<br />$Version <br />
    </body></html> 
"@
    $timenow = Get-Date -UFormat %r
    write-host $timenow Completeted Message Formatting 
########################################################################################################
## END HTML INFORMATION ##

########################################################################################################
## START PDF Converter Verify ##
    if ($WantPDFFile -eq "true") { 
        $PDFConverterFound = Test-Path $PDFConverter
        if ($PDFConverterFound){Write-Verbose "PDF Converter Found"} else {Write-Verbose "PDF Converter is Missing"
            $wshell = New-Object -ComObject Wscript.Shell
            $PDFReturn = $wshell.Popup("Unable to Find PDF converter `n   $PDFConverter `n`nClick [Cancle] to turn off PDF file output.         `n`nIf wkhtmltopdf is installed click [Try Again] to enter the correct path.        `n`nClick [Continue] to download wkHTMLtoPDF. `n   You will be prompted to enter the correct path.",90,"Oops - PDF Converter is Missing",0x6) 
            Write-Verbose "Return is $PDFReturn"
            
            if($PDFReturn -eq -1){$WantPDFFile = "false"} # -1 = Timed Out
            if($PDFReturn -eq 2 ){$WantPDFFile = "false"} # 2  = Cancel
            if($PDFReturn -eq 10){ # 10 = Try Again
                [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $PDFConverter = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the path to wkhtmltopdf.exe", "Path", "$PDFConverter")
                $PDFConverterFound = Test-Path $PDFConverter
                if ($PDFConverterFound -eq "True" ){Write-Verbose "Try Again - PDF Converter Found"} else {Write-Verbose "Try Again - PDF Converter is Missing" ; $WantPDFFile = "false"}
            }
            if($PDFReturn -eq 11){ # 11 = Continue
                (New-Object -Com Shell.Application).Open("http://wkhtmltopdf.org/downloads.html") 
                [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $PDFConverter = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the path to wkhtmltopdf.exe", "Path", "$PDFConverter")
                $PDFConverterFound = Test-Path $PDFConverter
                if ($PDFConverterFound -eq "True" ){Write-Verbose "Continue - PDF Converter Found"} else {Write-Verbose "Continue - PDF Converter is Missing" ; $WantPDFFile = "false"}
            }
        }
    }
########################################################################################################
## END PDF Converter Verify ##


if ($CreateFile -eq "N") {
    $timenow = Get-Date -UFormat %r
    write-host $timenow Creating File
    if ($outputpath -eq "") { $desktop = ([Environment]::GetFolderPath("Desktop")) ; $htmlmessage | out-File "$desktop\$filename"
    if ($WantPDFFile -eq "true") {  write-host $timenow Creating PDF ; &$PDFConverter -s Letter -q "$desktop\$filename" "$desktop\$filenamePDF"}}
    else { $htmlmessage | out-File $outputpath\$filename 
    if ($WantPDFFile -eq "true") { &$PDFConverter -s Letter -q "$outputpath\$filename" "$outputpath\$filenamePDF"} }
}
if ($SendEmail -eq "Y") {
    $timenow = Get-Date -UFormat %r
    write-host Sending Email to $SendToList
    if ($WantPDFFile -eq "true") { 
         $htmlmessage | out-File "$env:temp\$filename" 
       &"$PDFConverter" -s Letter -q "$env:temp\$filename" "$env:temp\$filenamePDF"
            if ($CcList -eq "") {send-mailmessage -From $fromemail -To $SendToList -Subject $subjectline -BodyAsHTML -Body $HTMLmessage -SmtpServer $smtpserver -Attachments "$env:temp\$filenamePDF" }
            else {send-mailmessage -From $fromemail -To $SendToList -Cc $CcList -Subject $subjectline -BodyAsHTML -Body $HTMLmessage -SmtpServer $smtpserver -Attachments "$env:temp\$filenamePDF" }
       del "$env:temp\$filename"
       del "$env:temp\$filenamePDF"
       } else {
    if ($CcList -eq "") {send-mailmessage -From $fromemail -To $SendToList -Subject $subjectline -BodyAsHTML -Body $HTMLmessage -SmtpServer $smtpserver }
    else { send-mailmessage -From $fromemail -To $SendToList -Cc $CcList -Subject $subjectline -BodyAsHTML -Body $HTMLmessage -SmtpServer $smtpserver }
} }

}     ###  DO NOT MODIFY 
if ($StartupVars) {
        $UserVars = Get-Variable -Exclude $StartupVars -Scope Global
        
#foreach ($var in $UserVars){ try {
#        Remove-Variable -Name $var.Name -Force -Scope Global -ErrorAction Stop
#        Write-Verbose -Message "Variable '$($var.Name)' has been successfully removed."
#        } catch { Write-Warning -Message "An error has occured. Error Details: $($_.Exception.Message)" }            
#    } }else { Write-Verbose -Message "`$StartupVars already exists in '$($profile.$Location)'" 
}