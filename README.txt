<#==============================================================================
         File Name : RicohAudit.ps1
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr DOT com)
                   :
       Description : This script will open an Excel spreadsheet stored in SharePoint, 
                   : extract printer IP addresses, and then poll each using SNMP for
                   : current toner levels.  It then emails the results to the designated
                   : recipient(s) using an HTML form.  Current OID specifications are 
                   : for Ricoh printers.  If desired, the script may be run with the 
                   : "remote" option set to "true" and it will create the IP input file
                   : on a remote system and exit.  This allows compensating for issues
                   : accessing the Excel file from the system running the full script.
                   : The remote otion on the remote system would be left as "false" which
                   : causes the script to only look in it's local folder for the IP file
                   : and not attempt to load the spreadsheet. Stored, encrypted credentials
                   : can be used for remote operations (see below).  If not you will be 
                   : prompted to enter them during each run, but only if remote ops are enabled.
                   :
         Arguments : N/A
                   :
   External Config : An external XML configuration file is required to set customized settings.
                   : In this way the script becomes generic and can be posted online
                   : with no potentially leaked private data.  Create this in the script folder
                   : with the same name as the script and .XML
                   : <?xml version="1.0" encoding="utf-8"?>
                   :    <Settings>
                   :        <General>
	                 :            <TriggerLevel>20</TriggerLevel>       <!-- NOTE: Level at which toner triggers alert -->
	                 :		    <DNS>8.8.8.8</DNS>
	                 :            <SmtpServer>mymail.myorg.com</SmtpServer>
	                 :            <SmtpPort>25</SmtpPort>
	                 :		    <EmailRecipient>groupemail@myorg.com</EmailRecipient>
	                 :		    <EmailAltRecipient>me@myorg.com</EmailAltRecipient>
	                 :		    <EmailSender>RicohPrinters@myorg.com</EmailSender>
	                 : 		    <ExcelFileName>Ricoh_Master_Inventory.xlsx</ExcelFileName>
	                 :		    <ExcelFilePath>C:\Users\me\Documents\Ricoh</ExcelFilePath>
	                 :            <Remote>false</Remote>
	                 :	        <RemoteHost>server01</RemoteHost>
	                 :	        <RemotePath>c$\scripts\ricohaudit</RemotePath>
	                 :        </General>
                   :        <HTML>
                   :            <HTML1>If you don't have a MyRicoh account, please submit a request via &lt;a href='https://my.ricoh-usa.com'&gt;Ricohusa.com&lt;/a&gt;.</HTML1>
                   :            <HTML2>My Company IT Department</HTML2>
                   :        </HTML>
                   :    	<Credentials>
		               :            <CredDrive>x:</CredDrive>
		               :            <PasswordFile>Pass.txt</PasswordFile>
		               :            <KeyFile>Key.txt</KeyFile>
	                 :        </Credentials>
	                 :    </Settings> 
                   :                   
      Requirements : PowerShell v5 or newer.  
                   : Net-SNMP available in your path.  http://www.net-snmp.org/
                   : PowerShell SNMP module.  https://github.com/lahell/SNMPv3
                   : To create the encrypted credetials file see this: https://github.com/kcmazie/CredentialsWithKey
                   :
             Notes : Adjust SNMP settings as needed.  To bypass Excel place a flat text 
                   : file in the script folder name IPLIST.TXT with one IP per line.
                   : The following info is to allow extraction directly from a Teams or SharePoint share:
                   :    PnP PowerShell   https://github.com/pnp/powershell
                   :    PnP PowerShell is a .NET 6 based PowerShell Module providing over 
                   :    650 cmdlets that work with Microsoft 365 environments such as 
                   :    SharePoint Online, Microsoft Teams, Microsoft Project, Security 
                   :    & Compliance, Azure Active Directory, and more. Requires PS v7.2 or newer.
                   :
          Warnings : None.  Excel is opened read-only.  Run once manually to assure that modules get installed.
                   :
             Legal : Public Domain. Modify and redistribute freely. No rights reserved.
                   : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF
                   : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.
                   :
           Credits : Code snippets and/or ideas came from many sources around the web.
                   :
    Last Update by : Kenneth C. Mazie
   Version History : v1.00 - 07-26-23 - Original
    Change History : v2.00 - 08-08-23 - Added error checking for Excel column changes.  Added
                   :                  - Added script editor detection for testing.
                   : v2.10 - 12-26-23 - Fixed missing offline detection
                   : v3.00 - 02-21-24 - Added new data from spreadsheet to report. 
                   : v3.10 - 04-15-24 - Added check to not send email if device is offline.
                   : v3.20 - 05-31-24 - Switched columns so that toner levels show first
                   : v3.30 - 06-06-24 - Added weekend abort
                   : v4.00 - 08-30-24 - Added high priority email option if toner is out and color
                   :                  - adjustment in report for DNS fail or toner out.
                   : v4.10 - 09-03-24 - Fixed DNS resolution bug.
                   : v4.20 - 09-06-24 - Added line to report to denote console/debug mode.
                   : v4.30 - 09-09-24 - Added option to specify a DNS server for use in nslookup.
                   : v5.00 - 01-29-25 - Added option to run from an external server.  Adjusted error 
                   :                  - handling, adjusted email send options, moved triggerlevel to
                   :                  - XML file, switched to per-run log file and 10 copy retention,
                   :                  - added log file to email as an attachment when remote is true. 
                   : v5.10 - 02-03-25 - Fixed email send when remote = true
                   : v5.20 - 02-07-25 - Added more info to remote operation email.
                   : v6.00 - 05-08-25 - Added credentials to connect to the remote system when remote 
                   :                    functioning is enabled.                     
                   : #>
                   $ScriptVer = "6.00"    <#--[ Current version # used in script ]--
#===============================================================================#>
