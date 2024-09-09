<#==============================================================================
         File Name : RicohAudit.ps1
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr DOT com)
                   :
       Description : This script will open an Excel spreadsheet stored in SharePoint, 
                   : extract printer IP addresses, and then poll each using SNMP for
                   : current toner levels.  It then emails the results to the designated
                   : recipient(s) using an HTML form.  Current OID specifications are 
                   : for Ricoh printers.
                   :
         Arguments : N/A
                   :
      Requirements : PowerShell v5 or newer.  
                   : Net-SNMP available in your path.  http://www.net-snmp.org/
                   : PowerShell SNMP module.  https://github.com/lahell/SNMPv3
                   :
             Notes : Adjust SNMP settings as needed.  To bypass Excel place a flat text 
                   : file in the  script folder name IPLIST.TXT with one IP per line.
                   : The following info is to allow extraction directly from a Teams or SharePoint share:
                   :    PnP PowerShell   https://github.com/pnp/powershell
                   :    PnP PowerShell is a .NET 6 based PowerShell Module providing over 
                   :    650 cmdlets that work with Microsoft 365 environments such as 
                   :    SharePoint Online, Microsoft Teams, Microsoft Project, Security 
                   :    & Compliance, Azure Active Directory, and more. Requires PS v7.2 or newer.
                   :
          Warnings : None.  Excel is opened read-only.
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
                   :
#===============================================================================#>
