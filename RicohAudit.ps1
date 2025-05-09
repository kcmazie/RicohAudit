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
	               :		    <ExcelFileName>Ricoh_Master_Inventory.xlsx</ExcelFileName>
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
#requires -version 5.0
clear-host
 
[string]$DateTime = Get-Date -Format MM-dd-yyyy_HHmmss 
$CloseAnyOpenXL = $False
$ExtOption = New-Object -TypeName psobject #--[ Object to hold runtime options ]--
$ScriptName = ($MyInvocation.MyCommand.Name).Replace(".ps1","" ) 
$ExtOption | Add-Member -Force -MemberType NoteProperty -Name "LogFile" -Value ($PSScriptRoot+'\'+$ScriptName+"_"+$DateTime+'.log')

#--[ Only retain 10 of the most recent log files ]--
Get-ChildItem -Path $PSScriptRoot | Where-Object {(-not $_.PsIsContainer) -and ($_.Name -like "*log*")} | Sort-Object -Descending -Property LastTimeWrite | Select-Object -Skip 10 | Remove-Item | Out-Null

#--[ Runtime Adjustments ]--
$Console = $false
$Debug = $false
#--------------------------
If ($Debug){
    $Console = $True
}
$erroractionpreference = "stop"

$Today =(get-date).DayOfWeek  #--[ Dont run on weekends ]-----------
If (($Today -eq "Saturday") -or ($Today -eq "Sunday")){
   break
}

#==[ Functions, Arrays, & Lists ]===========================================
Function GetCred ($ExtOption){
	#--[ Prepare Credentials ]--
	$UN = $Env:USERNAME
	$DN = $Env:USERDOMAIN
	$UID = $DN+"\"+$UN

	#--[ Test location of encrypted files, remote or local ]--
	If (Test-Path -path ($ExtOption.CredDrive+'\'+$ExtOption.PasswordFile)){
	    $PF = ($ExtOption.CredDrive+'\'+$ExtOption.PasswordFile)
	    $KF = ($ExtOption.CredDrive+'\'+$ExtOption.KeyFile)
	}Else{
	    $PF = ($PSScriptRoot+'\'+$ExtOption.PasswordFile)
	    $KF = ($PSScriptRoot+'\'+$ExtOption.KeyFile)
	}

	If (Test-Path -Path $PF){
	    $Base64String = (Get-Content $KF)
	    $ByteArray = [System.Convert]::FromBase64String($Base64String)
	    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UID, (Get-Content $PF | ConvertTo-SecureString -Key $ByteArray)
	    $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Credential" -Value $Credential
	}Else{
	    $Credential = $Host.ui.PromptForCredential("Enter your credentials","Please enter your Domain\UserID and Password.","","")
	    $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Credential" -Value $Credential
	}

    Return $ExtOption    
}

Function InstallModules{
    Try{
        if (!(Get-Module -Name PnP.PowerShell)) {       
            Get-Module -ListAvailable PnP.PowerShell | Import-Module | Out-Null    
            Install-Module -Name PnP.PowerShell -RequiredVersion 2.2.156
            Install-Module -Name PnP.PowerShell -RequiredVersion 1.12.0
        }
    }Catch{
        Write-Host "Error installing PNP module" -ForegroundColor "Red"
        Add-Content -path $ExtOption.Logfile -value $_.Error.Message
        Add-Content -path $ExtOption.Logfile -value $_.Exception.Message 
    }

    Try{
        if (!(Get-Module -Name SNMPv3)) {
            StatusMsg "Installing PowerShell SNMP module" "Magenta" $ExtOption
            Get-Module -ListAvailable SNMPv3 | Import-Module | Out-Null
            Install-Module -Name SNMPv3
        }
    }Catch{
        Write-Host "Error installing SNMP module" -ForegroundColor "Red"
        Add-Content -path $ExtOption.Logfile -value $_.Error.Message
        Add-Content -path $ExtOption.Logfile -value $_.Exception.Message        
    }
}

Function LoadConfig ($ExtOption,$ConfigFile){  #--[ Read and load configuration file ]-------------------------------------
    StatusMsg "Loading external config file..." "Magenta" $ExtOption
    if (Test-Path -Path $ConfigFile -PathType Leaf){                       #--[ Error out if configuration file doesn't exist ]--
        [xml]$Config = Get-Content $ConfigFile  #--[ Read & Load XML ]--    
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "TriggerLevel" -Value $Config.Settings.General.TriggerLevel
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailRecipient" -Value $Config.Settings.General.EmailRecipient
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailAltRecipient" -Value $Config.Settings.General.EmailAltRecipient
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailSender" -Value $Config.Settings.General.EmailSender
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "ExcelFileName" -Value $Config.Settings.General.ExcelFileName
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "ExcelFilePath" -Value $Config.Settings.General.ExcelFilePath
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SmtpServer" -Value $Config.Settings.General.SmtpServer
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SmtpPort" -Value $Config.Settings.General.SmtpPort
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "DNS" -Value $Config.Settings.General.DNS
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Remote" -Value $Config.Settings.General.Remote
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "RemoteHost" -Value $Config.Settings.General.RemoteHost
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "RemotePath" -Value $Config.Settings.General.RemotePath
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "HTML1" -Value $Config.Settings.HTML.HTML1
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "HTML2" -Value $Config.Settings.HTML.HTML2
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "CredDrive" -Value $Config.Settings.Credentials.CredFile
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "PasswordFile" -Value $Config.Settings.Credentials.PasswordFile
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "KeyFile" -Value $Config.Settings.Credentials.KeyFile
	    
        #--[ Prepare Credentials if remote ]--
        If ($ExtOption.Remote -eq "True"){
    	    $UN = "a"+$Env:USERNAME   #--[ NOTE the addition of "a" to denote an admin user.  You may need to remove this. ]--
	        $DN = $Env:USERDOMAIN
        	$UID = $DN+"\"+$UN

        	#--[ Test location of encrypted files, remote or local ]--
    	    If (Test-Path -path ($ExtOption.CredDrive+'\'+$ExtOption.PasswordFile)){
	            $PF = ($ExtOption.CredDrive+'\'+$ExtOption.PasswordFile)
        	    $KF = ($ExtOption.CredDrive+'\'+$ExtOption.KeyFile)
    	    }Else{
	            $PF = ($PSScriptRoot+'\'+$ExtOption.PasswordFile)
        	    $KF = ($PSScriptRoot+'\'+$ExtOption.KeyFile)
    	    }

        	If (Test-Path -Path $PF){
    	        $Base64String = (Get-Content $KF)
	            $ByteArray = [System.Convert]::FromBase64String($Base64String)
	            $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UID, (Get-Content $PF | ConvertTo-SecureString -Key $ByteArray)
	            $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Credential" -Value $Credential
        	}Else{
	            $Credential = $Host.ui.PromptForCredential("Enter your credentials","Please enter your Domain\UserID and Password.","","")
	            $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Credential" -Value $Credential
        	}
        }
    }Else{
        StatusMsg "MISSING XML CONFIG FILE.  File is required.  Script aborted..." " Red" $ExtOption
        break;break;break
    }
    Return $ExtOption
}

Function GetConsoleHost ($ConfigFile){  #--[ Detect if we are using a script editor or the console ]--
    Switch ($Host.Name){
        'consolehost'{
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $False -force
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "ConsoleMessage" -Value "PowerShell Console detected." -Force
        }
        'Windows PowerShell ISE Host'{
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $True -force
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "ConsoleMessage" -Value "PowerShell ISE editor detected.  Console mode enabled." -Force
        }
        'PrimalScriptHostImplementation'{
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $True -force
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "COnsoleMessage" -Value "PrimalScript or PowerShell Studio editor detected.  Console mode enabled." -Force
        }
        "Visual Studio Code Host" {
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "ConsoleState" -Value $True -force
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "ConsoleMessage" -Value "Visual Studio Code editor detected.  Console mode enabled. " -Force
        }
    }
    If ($ConfigFile.ConsoleState){
        StatusMsg "Detected session running from an editor..." "Cyan" $ConfigFile
    }
    Return $ConfigFile
}

Function StatusMsg ($Msg, $Color, $ExtOption){
    If ($Null -eq $Color){
        $Color = "Magenta"
    }
    Add-content -path $ExtOption.LogFile -value $msg
    If ($ExtOption.ConsoleState){
        Write-Host "-- Script Status: $Msg" -ForegroundColor $Color
    }
    $Msg = ""
}

Function SendEmail ($MessageBody,$ExtOption) {      
    $ErrorActionPreference = "Stop"
    $Smtp = New-Object Net.Mail.SmtpClient($ExtOption.SmtpServer,$ExtOption.SmtpPort) 
    $Email = New-Object System.Net.Mail.MailMessage  
    $Email.IsBodyHTML = $true
    $Email.From = $ExtOption.EmailSender

    If ($ExtOption.Priority){
        $Email.Priority = "High"
    }

    If ($ExtOption.Remote -eq "true"){
        $Email.To.Add($ExtOption.EmailAltRecipient) 
        $Email.Subject = "Ricoh Printer Inventory Update Report"
    }Else{
        $Email.Subject = "Ricoh Printer Status Report"
        If ($ExtOption.ConsoleState){  #--[ If running out of an IDE console, send only to the user for testing ]-- 
            $Email.To.Add($ExtOption.EmailAltRecipient)         
            $Email.Attachments.Add($ExtOption.LogFile)
        }Else{
            If (($ExtOption.Alert) -or ($ExtOption.Priority)){  #--[ If a device failed self-test or trigger day is matched send to main recipient ]--
                $Email.To.Add($ExtOption.EmailRecipient)  
                #$Email.To.Add($ExtOption.EmailAltRecipient)    #--[ In case this user isn't part of the group email ]--  
                ForEach ($Address in ($ExtOption.AddEmail.split(";"))){
                    If ($Address -ne ""){
                        $Email.To.Add($Address)
                    }
                }
            }
        }
    }

    $Email.Body = $MessageBody

    Try {
        If ($ExtOption.Alert){
            $ErrorActionPreference = "silentlycontinue"
            $Smtp.Send($Email)
            If ($ExtOption.ConsoleState){Write-Host `n"--- Email Sent ---" -ForegroundColor red }
            $ErrorMsg = "None"
            $ExceptionMsg = "N/A"
            Start-Sleep -millisec 500
        }
    }Catch{
        Write-host "-- Error sending email --" -ForegroundColor Red
    }

    If ($ExtOption.Debug){
        $Msg="-- Debug Parameters --" 
        StatusMsg $Msg "yellow" $ExtOption
        $Msg="Priority       = "+$Email.Priority
        StatusMsg $Msg "yellow" $ExtOption 
        $Msg="Send Error     = "+$ErrorMsg
        StatusMsg $Msg "yellow" $ExtOption
        $Msg="Send Exception = "+$ExceptionMsg
        StatusMsg $Msg "yellow" $ExtOption
        $ExtOption 
    }
}

Function OctetString2String ($Result){
    $Bytes = [System.Text.Encoding]::Unicode.GetBytes($Result)
    $SaveVal = "" 
    ForEach ($Value in $Bytes){
        If ($Value -ne " "){
            $SaveVal += ([System.Text.Encoding]::ASCII.GetString($Value)).trim()                
        }
    }  
    Return $SaveVal
}

Function TCPportTest ($Target, $Port, $Debug){
    Try{
        #$Result = Test-NetConnection -ComputerName $Target -Port $Port #-ErrorAction SilentlyContinue -WarningAction SilentlyContinue #-InformationLevel Quiet
        $Result = New-Object System.Net.Sockets.TcpClient($Target, $Port) -ErrorAction:Stop
    }Catch{
        Return $_.Exception.Message
    }
    If ($Debug){
        Write-Host "`nFTP Debug :" $Result.connected -foregroundcolor red
    }
    return $Result
}

Function SMNPv3Walk ($Target,$OID,$Debug){
    $WalkRequest = @{
        UserName   = $Script:SMNPv3User
        Target     = $Target
        OID        = $OID
        AuthType   = 'MD5'
        AuthSecret = $Script:SNMPv3Secret
        PrivType   = 'DES'
        PrivSecret = $Script:SNMPv3Secret
        #Context    = ''
    }
    $Result = Invoke-SNMPv3Walk @WalkRequest | Format-Table -AutoSize
    If ($Debug){
        Write-Host "SNMpv3 Debug :" $Result 
    }
    Return $Result
}

Function GetSNMPv1 ($Target,$OID,$Debug) {
    $SNMP = New-Object -ComObject olePrn.OleSNMP
    $OID = $OID.Split(",")[1]
    $ErrorActionPreference = "Stop"
    Try{
        $SNMP.open($Target,"public",2,1000)
        $Result = $SNMP.get($OID)
    }Catch{       
        $Result = "N/A"
    }
    If ($Debug){
        Write-Host "SNMpv1 Debug :" $Result 
    }
    Return $Result   
}

Function GetSMNPv3 ($Target,$OID,$Debug,$Test){
    If ($Test){  #--[ If 1st user tests positive on 1st use, use it by setting the global variable below ]--
        $GetRequest1 = @{
            UserName   = $Script:SMNPv3User
            Target     = $Target
            OID        = $OID.Split(",")[1]
            AuthType   = 'MD5'
            AuthSecret = $Script:SNMPv3Secret
            PrivType   = 'DES'
            PrivSecret = $Script:SNMPv3Secret
        }
        Try{
            $Result = Invoke-SNMPv3Get @GetRequest1 -ErrorAction:Stop
            If ($Result -like "*Exception*"){
                $Script:v3UserTest = $False  
            }Else{
                $Script:v3UserTest = $True  #--[ Global v3 user variable ]--
            }
        }Catch{
            If ($Debug){
                Write-Host $_.Exception.Message -ForegroundColor Cyan
                Write-Host " -- SNMPv3 User 1 failed..." -ForegroundColor red
            }
        }
    }Else{  #--[ User 1 has failed so use user 2 instead ]--
        $GetRequest2 = @{
            UserName   = $Script:SMNPv3AltUser
            Target     = $Target
            OID        = $OID.Split(",")[1]
            #AuthType   = 'MD5'
            #AuthSecret = $Script:SNMPv3Secret
            #PrivType   = 'DES'
            #PrivSecret = $Script:SNMPv3Secret
        }
        Try{
            $Result = Invoke-SNMPv3Get @GetRequest2 -ErrorAction:Stop
        }Catch{
            If ($Result -like "*Exception*"){
                Write-Host " -- SNMPv3 User 2 failed... No SNMPv3 access..." -ForegroundColor red
                Write-Host $_.Exception.Message -ForegroundColor Blue
            }
        }
    }
    If ($Debug){
        Write-Host "  -- SNMPv3 Debug: " -ForegroundColor Yellow -NoNewline
        If ($Test){
            Write-Host "SNMP User 2  " -ForegroundColor Green -NoNewline
        }Else{
            Write-Host "SNMP user 1  " -ForegroundColor Green -NoNewline
        }
        Write-Host $OID.Split(",")[0]"  " -ForegroundColor Cyan -NoNewline
        Write-Host $Result    
    }
    Return $Result
}

#--[ SNMP OID Array ]--
$OIDArray = @()
#$OIDArray += ,@('Black','.1.3.6.1.2.1.43.11.1.1.9.1.4')
#$OIDArray += ,@('Cyan','.1.3.6.1.2.1.43.11.1.1.9.1.1')
#$OIDArray += ,@('Magenta','.1.3.6.1.2.1.43.11.1.1.9.1.2')
#$OIDArray += ,@('Yellow','.1.3.6.1.2.1.43.11.1.1.9.1.3')
#$OIDArray += ,@('Black','.1.3.6.1.2.1.43.11.1.1.9.1.1')  # Ricoh
#$OIDArray += ,@('Cyan','.1.3.6.1.2.1.43.11.1.1.9.1.2')
#$OIDArray += ,@('Magenta','.1.3.6.1.2.1.43.11.1.1.9.1.3')
#$OIDArray += ,@('Yellow','.1.3.6.1.2.1.43.11.1.1.9.1.4')
$OIDArray += ,@('MFG','.1.3.6.1.4.1.367.3.2.1.1.1.7.0')  # Ricoh
$OIDArray += ,@('Model','.1.3.6.1.4.1.367.3.2.1.1.1.1.0')
$OIDArray += ,@('MachineID','.1.3.6.1.4.1.367.3.2.1.2.1.4.0') # Serial #
$OIDArray += ,@('Firmware','.1.3.6.1.4.1.367.3.2.1.1.1.2.0')
$OIDArray += ,@('Black','.1.3.6.1.4.1.367.3.2.1.2.24.1.1.5.1')
$OIDArray += ,@('Cyan','.1.3.6.1.4.1.367.3.2.1.2.24.1.1.5.2')
$OIDArray += ,@('Magenta','.1.3.6.1.4.1.367.3.2.1.2.24.1.1.5.3')
$OIDArray += ,@('Yellow','.1.3.6.1.4.1.367.3.2.1.2.24.1.1.5.4')

#==[ End of Functions, Arrays, & Lists ]=============================================    
$ErrorActionPreference = "stop"
$Msg = "--[ Begin ]------------------------------------" 
StatusMsg $Msg "Yellow" $ExtOption

#--[ Load required PowerShell modules ]--
InstallModules

#--[ Load external XML options file ]--
$ConfigFile = $PSScriptRoot+"\"+$MyInvocation.MyCommand.Name.Replace(".ps1", ".xml")
$ExtOption = LoadConfig $ExtOption $ConfigFile

If ($NUll -ne $ExtOption.SmtpServer){
    StatusMsg "External config file loaded successfully." "Magenta" $ExtOption
}
Add-content -path $ExtOption.LogFile -value (Get-Date -f MM-dd-yyyy_hh:mm:ss)

#--[ Detect Runspace ]--
$ExtOption = GetConsoleHost $ExtOption 
If ($ExtOption.ConsoleState){ 
    StatusMsg $ExtOption.ConsoleMessage "Cyan" $ExtOption
}
#--[ Detect debug mode ]--
If ($Debug){
    $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Debug" -Value $True
}

If (Test-Path -Path "$PSScriptroot\IPList.txt"){
    $TargetList = Get-Content -Path "$PSScriptRoot\IPList.txt"
}Else{
    $IPList = @() 
    $PreviousIP = ""
    $Row = 5
    # $Col = 13  #--[ SET THIS BELOW.  This is no longer used ]--
    $Counter = 0
    $AddEmail = "" #--[ Used to add contact to email report if needed ]--

    #Read more: https://www.sharepointdiary.com/2016/09/sharepoint-online-download-file-from-library-using-powershell.html#ixzz8QKFappQS
    #<ExcelFileName> AHLM_Ricoh_Master_Inventory.xlsx
    #<ExcelFilePath> https://adventisthealthwest.sharepoint.com/:x:/r/sites/AHLMITTechnologyTeam/Shared%20Documents/General/Inventory/Ricoh
    #onedrivepath  = "C:\Users\maziekc\OneDrive - ADVENTIST HEALTH SYSTEM WEST\General - AHLM IT Technology Team\Inventory\Ricoh"

    $SourceFile = $ExtOption.ExcelFilePath+"\"+$ExtOption.ExcelFileName
    $LocalFile = $PSScriptRoot+"\"+$ExtOption.ExcelFileName

    If ($ExtOption.Debug){
        StatusMsg $SourceFile "Yellow" $ExtOption
    }

    If (Test-Path -Path $SourceFile -PathType Leaf){   #--[ Is there a valid source file? ]--
        StatusMsg "Source file located, copying to local folder..." "Magenta" $ExtOption  
        Try{ 
            Copy-Item -Path $SourceFile -destination $LocalFile -Force -ErrorAction "stop"
            $Source = $LocalFile
        }Catch{
            StatusMsg "Source file copy failed..." "Red" $ExtOption
        }
    }Else{ 
        StatusMsg "Source file NOT located, checking local folder..." "Red" $ExtOption 
        If (Test-Path -Path $LocalFile -PathType Leaf){
            StatusMsg "Local file detected, loading it..." "Magenta" $ExtOption
            $Source = $LocalFile
        }Else{
            StatusMsg "No target file can be located.  Exiting..." "Red" $ExtOption
            break;break;break
        }
    }

    If ($CloseAnyOpenXL){  #--[ Close copies of Excel that PowerShell has open ]--
        $ProcID = Get-CimInstance Win32_Process | Where-Object {$_.name -like "*excel*"}
        ForEach ($ID in $ProcID){  #--[ Kill any open instances to avoid issues ]--
            Foreach ($Proc in (get-process -id $id.ProcessId)){
                if (($ID.CommandLine -like "*/automation -Embedding") -Or ($proc.MainWindowTitle -like "$ExcelWorkingCopy*")){
                    Stop-Process -ID $ID.ProcessId -Force
                    StatusMsg "Killing any lingering PowerShell spawned instances of Excel..." "Red" $ExtOption
                    Start-Sleep -Milliseconds 100
                }
            }
        }
    }

    Try{
        $Excel = New-Object -ComObject Excel.Application -ErrorAction Stop
        StatusMsg "Creating new Excel COM object..." "Magenta" $ExtOption
    }Catch{
        StatusMsg "Creation of Excel COM object has failed..." "Red" $ExtOption
        StatusMsg $_.Error.Message "red" $ExtOption
        StatusMsg $_.Exception.Message "red" $ExtOption
    }

    Try{
        $WorkBook = $Excel.Workbooks.Open($Source)
        $WorkSheet = $Workbook.Sheets.Item("Sheet1")
        $WorkSheet.activate()
        $Excel.Visible = $False
        $Excel.displayalerts = $False
        StatusMsg "Spreadsheet opened..." "Magenta" $ExtOption
    }Catch{
        StatusMsg $_.Error.Message "red" $ExtOption
        StatusMsg $_.Exception.Message "red" $ExtOption
        $Excel.Quit()
        break;break
    }

    StatusMsg "Reading Spreadsheet data..." "Magenta" $ExtOption
    If ($ExtOption.ConsoleState){Write-Host "   ." -NoNewline -ForegroundColor Cyan}
    Do {
        $CurrentIP = $WorkSheet.Cells.Item($Row,14).Text 
        $Hostname = $WorkSheet.Cells.Item($Row,13).Text 
        $Building = $WorkSheet.Cells.Item($Row,2).Text             
        $Department = $WorkSheet.Cells.Item($Row,3).Text 
        $Address = $WorkSheet.Cells.Item($Row,5).Text 
        $Contact = $WorkSheet.Cells.Item($Row,6).Text 
        $Email = $WorkSheet.Cells.Item($Row,7).Text 
        If ($CurrentIP -ne $PreviousIP){  #--[ Make sure IPs are added only once ]--
            If ($ExtOption.ConsoleState){Write-Host "." -NoNewline -ForegroundColor Cyan}
            $IPList += ,@($CurrentIP+";"+$Hostname+";"+$Department+";"+$Building+";"+$Address+";"+$Contact+";"+$Email) 
            $PreviousIP = $CurrentIP
        }
        $Row++
        $Counter++
    } Until (
        $WorkSheet.Cells.Item($Row,2).Text -eq ""   #--[ Condition that stops the loop if it returns true ]--
    )
    If ($Counter -eq 0){
        $IPlist = "No Entries"
    }
    If ($ExtOption.ConsoleState){Write-Host ""}
    $Msg = "$Counter Total IP addresses detected."
    StatusMsg $Msg "Magenta" $ExtOption
    $Excel.DisplayAlerts = $false
    Try{
        $WorkBook.Close($true)
        $Excel.Quit()
    }catch{
        StatusMsg $_.Error.Message "red" $ExtOption
        StatusMsg $_.Exception.Message "red" $ExtOption
    }
    $TargetList = $IPlist
}

#==[ Remote Execution:  Create IP list and copy it to the remote execution server ]==
If ($ExtOption.Remote -eq "true"){     
    StatusMsg "Remote mode enabled."  "Cyan" $ExtOption
    $RemoteFile = "\\"+$ExtOption.RemoteHost+"\"+$ExtOption.RemotePath+"\IPList.txt"
    $Msg = "Current remote IP filename        = $RemoteFile"
    $RmtMsg = $Msg+"<br>"
    StatusMsg $Msg  "Cyan" $ExtOption
    StatusMsg "Attempting remote connection..." "Cyan" $ExtOption
    #--[ Validate connectivity to remote file location & attempt to connect ]--
    $Password = $ExtOption.Credential.GetNetworkCredential().Password
    $Domain = $ExtOption.Credential.GetNetworkCredential().Domain
    $Username = $Domain+"\"+$ExtOption.Credential.GetNetworkCredential().UserName
    $Command = "\\"+$ExtOption.RemoteHost+"\ipc$"
    Try{
        $Msg = @(net use $Command /user:$Username $Password *>&1) 
        $Color = "Green"
        StatusMsg $Msg $Color $ExtOption
        Try{
            If (Test-Path -Path $RemoteFile -PathType Leaf){
                $RemoteDate = (get-item $RemoteFile).LastWriteTime
                $Msg = "Current remote IP file date stamp = $RemoteDate"
                $RmtMsg += $Msg+"<br>"
                StatusMsg $Msg  "Cyan" $ExtOption
                Remove-Item $RemoteFile -Force
                Start-Sleep -sec 2
                $Msg = "Current remote IP file deleted successfully..."
                $RmtMsg += $Msg+"<br>"
                StatusMsg $Msg  "Cyan" $ExtOption
            }
            Add-Content -Path $RemoteFile -Value $TargetList -force 
            $RemoteDate = (get-item $RemoteFile).LastWriteTime
            $Msg = "New remote IP file date stamp     = $RemoteDate"
            $RmtMsg += $Msg+"<br>"
            StatusMsg $Msg  "Cyan" $ExtOption
            $Msg = "Ricoh Printer Audit remote IP file update successful..."
            $RmtMsg += $Msg
            StatusMsg $Msg  "Green" $ExtOption
        }Catch{
            $Msg = "Ricoh Printer Audit remote IP file delete/create has failed..."
            $RmtMsg += $Msg
            StatusMsg $Msg "Red" $ExtOption
            Add-Content -Path "$PSScriptRoot/RicohIPList.txt" -Value $TargetList -force 
            $Msg = "Remote IP file has been written locally to script folder.  Please copy then rename it manually..."
            $RmtMsg += $Msg
            StatusMsg $Msg "Green" $ExtOption
            #Break;Break;Break
        }
    }Catch{
        $Msg = "Remote conenction has failed..."
        $Color = "Red"
        StatusMsg $Msg $Color $ExtOption
        $Msg = "Ricoh Printer Audit remote IP file delete/create has failed..."
        $RmtMsg += $Msg
        StatusMsg $Msg "Red" $ExtOption
        Add-Content -Path "$PSScriptRoot/RicohIPList.txt" -Value $TargetList -force 
        $Msg = "Remote IP file has been written locally to script folder.  Please copy then rename it manually..."
        $RmtMsg += $Msg
        StatusMsg $Msg "Green" $ExtOption
        #Break;Break;Break
    }
    $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "Alert" -Value $True
    $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailRecipient" -Value $ExtOption.EmailAltRecipient
    SendEmail $RmtMsg $ExtOption
    #--[ Remove the connection ]--
    StatusMsg "Attempting remote disconnect..." "Cyan" $ExtOption
    $Command = "\\"+$ExtOption.RemoteHost+"\ipc$"
    Try{
        $Msg = @(net use $Command /delete *>&1) 
        StatusMsg $Msg "Cyan" $ExtOption
    }Catch{
        StatusMsg "Remote connection disconnect failed or was not found..." "Red" $ExtOption
    }
    If ($ExtOption.ConsoleState){
        Write-Host "`n--- Completed ---" -foregroundcolor red
    }
    Break;Break;Break
}

$ColSpan = 15
StatusMsg "Creating report header... " "Magenta" $ExtOption
$HtmlData = @() 
$HtmlHeader = @() 
$HtmlHeader += '
    <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html>
    <head>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style type="text/css">
        table.myTable { border:10px solid black;border-collapse:collapse; }
        table.myTable td { border:2px solid black;padding:5px;background: #E6E6E6 } 
        table.myTable th { border:2px solid black;padding:5px;background: #B4B4AB }
        table.bottomBorder { border-collapse:collapse; }
        table.bottomBorder td, table.bottomBorder th { border-bottom:1px dotted black;padding:5px; }
        .content {
            background: white;
        }
    html {
        width: 100%;
        height: 100%;
        }
    body {
        display: flex;
        justify-content: center;
        align-items: center;
    }
    </style>
</head>
<body><div class="content">
'


$HtmlData += '<table class="myTable">'
$HtmlData += '<tr><td colspan='+$ColSpan+'><center><h2 style="color: DarkCyan"><strong>Ricoh Printer Status Report</strong></h2></center></td></tr>'
If ($ExtOption.ConsoleState){
    $HtmlData += '<tr><td colspan='+$ColSpan+'><center><font color=Magenta>--- Script running in Console/Debug mode ---</center></td></tr>'
}
$HtmlData += '<tr><td colspan='+$ColSpan+'><center><font color=black>Levels displayed below indicate toner percent remaining.&nbsp;&nbsp;&nbsp;Alert triggered at levels below '+$ExtOption.TriggerLevel+'%.</center></td></tr>'

If ($Null -eq $TargetList){
        StatusMsg "There was an error with the target list." "red" $ExtOption
        StatusMsg "No IP data has been found." "red" $ExtOption
        StatusMsg "Have the columns been moved, added, or deleted?" "red" $ExtOption
        $HtmlData += '<tr><td colspan=10><Strong><Center><font color=DarkRed>There has been an error.  No IP address data has been found.</strong></center></td></tr>'
        $HtmlData += '<tr><td colspan=10><Strong><Center><font color=DarkRed>Have the columns been changed on the spreadsheet?</strong></center></td></tr>'
}Else{
    ForEach ($Target in $TargetList){
        If ($Target -eq "#Local"){
            $HtmlData += '<tr><td colspan='+$ColSpan+'><strong><center><font color=red>NOTICE: Spreadsheet not detected.  A local copy of the spreadsheet has been used.</center></strong></td></tr>'
        }
    }
    $HtmlData += '<tr><strong><td><center>Host Name</center></td><td><center>Status</td>'
    $HtmlData += '<td><center>Black Level</td>'
    $HtmlData += '<td><font color=darkcyan><center>Cyan Level</Font></td>'
    $HtmlData += '<td><font color=darkmagenta><center>Magenta Level</font></td>'
    $HtmlData += '<td><font color=orange><center>Yellow Level</font></strong></center></td>'    
    $HtmlData += '<td><center>IP Address</td><td><center>Contact</td><td><center>Department</td>'
    $HtmlData += '<td><center>Building</td><td><center>Address</td>'
    $HtmlData += '<td><center>Mfg</td><td>Model</td><td><center>Serial</td><td><center>Firmware</td></tr>'

    ForEach ($Target in $TargetList | Where-Object {$_.Substring(0,1) -ne "#"}){
        $Low = $False
        $InkOut = $False
        $Obj = New-Object -TypeName psobject   #--[ Collection for Results ]--
        Try{
            If ($Null -eq $ExtOption.DNS){
                $HostLookup = (nslookup $Target.Split(";")[0] ($Env:LogonServer.Split("\")[2]) 2>&1)          
            }Else{
                $HostLookup = (nslookup $Target.Split(";")[0] ($ExtOption.DNS) 2>&1)          
            }
            $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value (($HostLookup[3].split(":")[1].TrimStart()).Split(".")[0]).ToUpper() -force
            $Obj | Add-Member -MemberType NoteProperty -Name "HostnameLookup" -Value $True
        }Catch{
            Try{
                $HostLookup = (nslookup $Target.Split(";")[0] 2>&1)          
                $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value (($HostLookup[3].split(":")[1].TrimStart()).Split(".")[0]).ToUpper() -force
                $Obj | Add-Member -MemberType NoteProperty -Name "HostnameLookup" -Value $True
            }Catch{
                #$Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value "DNS Not Found" -force
                $Obj | Add-Member -MemberType NoteProperty -Name "Hostname" -Value $Target.Split(";")[1] -force
                $Obj | Add-Member -MemberType NoteProperty -Name "HostnameLookup" -Value $False
                If ($ExtOption.Debug){
                    Add-Content -path $ExtOption.Logfile -value $_.Error.Message
                    Add-Content -path $ExtOption.Logfile -value $_.Exception.Message
                }
            }
        }

        $Obj | Add-Member -MemberType NoteProperty -Name "IPAddress" -Value $Target.Split(";")[0] -force        
        $Obj | Add-Member -MemberType NoteProperty -Name "Contact" -Value $Target.Split(";")[5] -force
        $Obj | Add-Member -MemberType NoteProperty -Name "Email" -Value $Target.Split(";")[6] -force
        $Obj | Add-Member -MemberType NoteProperty -Name "Department" -Value $Target.Split(";")[2] -force  
        $Obj | Add-Member -MemberType NoteProperty -Name "Building" -Value $Target.Split(";")[3] -force
        $Obj | Add-Member -MemberType NoteProperty -Name "Address" -Value $Target.Split(";")[4] -force    

        If (Test-Connection -ComputerName $Target.Split(";")[0] -count 1 -BufferSize 16 -Quiet){  #--[ Ping Test ]--
            $Obj | Add-Member -MemberType NoteProperty -Name "Connection" -Value "Online" -force
            $Msg = "Querying SNMP on "+$Target.Split(";")[0]+"..."
            StatusMsg $Msg "Yellow" $ExtOption
            If ($Console){Write-Host "     Working." -NoNewline}
            ForEach ($OID in $OIDArray){     
                If ($Obj.Connection -eq "Online"){  #--[ Only process OIDs if online  ]--------------------------
                    $Result = GetSNMPv1 $Target.Split(";")[0] $OID $False #$Script:v3UserTest
                    If ($Console){Write-Host "." -NoNewline -ForegroundColor cyan}
                }Else{
                    $Result = "N/A"
                }             
                Switch ($OID.Split(",")[0]) {  #--[ Returned values are quantity remaining in % ]--
                    "MFG"{
                        $SaveVal = $Result 
                        $Obj | Add-Member -MemberType NoteProperty -Name "MFG" -Value $SaveVal -force
                    } #
                    "Model"{
                        $SaveVal = $Result 
                        $Obj | Add-Member -MemberType NoteProperty -Name "Model" -Value $SaveVal -force
                    } #
                    "MachineID"{
                        $SaveVal = $Result 
                        $Obj | Add-Member -MemberType NoteProperty -Name "Serial" -Value $SaveVal -force
                    } #
                    "Firmware"{
                        $SaveVal = $Result 
                        $Obj | Add-Member -MemberType NoteProperty -Name "Firmware" -Value $SaveVal -force
                    } #
                    "Black"{
                        $SaveVal = $Result 
                        $Obj | Add-Member -MemberType NoteProperty -Name "BlackLevel" -Value $SaveVal -force
                    } #
                    "Cyan"{
                        $SaveVal = $Result
                        $Obj | Add-Member -MemberType NoteProperty -Name "CyanLevel" -Value $SaveVal -force
                    } #
                    "Magenta"{
                        $SaveVal = $Result
                        $Obj | Add-Member -MemberType NoteProperty -Name "MagentaLevel" -Value $SaveVal -force
                    } #
                    "Yellow"{
                        $SaveVal = $Result 
                        $Obj | Add-Member -MemberType NoteProperty -Name "YellowLevel" -Value $SaveVal -force
                    } #           
                }
            }
            If ($Console){Write-Host " "}
        }Else{
            $Obj | Add-Member -MemberType NoteProperty -Name "Connection" -Value "Offline" -force
        }

        If ($ExtOption.ConsoleState){
            Write-Host " "
            If ($obj.HostNameLookup){
                Write-Host "  Hostname :"$obj.Hostname -ForegroundColor Magenta
            }Else{
                Write-Host "  Hostname :"$obj.Hostname -ForegroundColor Red
            }
            Write-Host "        IP :"$obj.IPAddress -ForegroundColor Magenta
            If ($obj.Connection -eq "OffLine"){
                Write-Host "    Status :"$obj.Connection -ForegroundColor Red
            }Else{
                Write-Host "    Status :"$obj.Connection -ForegroundColor Magenta
            }   
            Write-Host "   Contact :"$obj.Contact -ForegroundColor Magenta
            Write-Host "     Email :"$obj.Email -ForegroundColor Magenta
            Write-Host "Department :"$obj.Department -ForegroundColor Magenta
            Write-Host "  Building :"$obj.Building -ForegroundColor Magenta
            Write-Host "   Address :"$obj.Address -ForegroundColor Magenta
            Write-Host "       MFG :"$obj.MFG -ForegroundColor Magenta
            Write-Host "     Model :"$obj.Model -ForegroundColor Magenta
            Write-Host "    Serial :"$obj.Serial -ForegroundColor Magenta
            Write-Host "  Firmware :"$obj.Firmware -ForegroundColor Magenta
            Write-Host "   Black % :"$obj.BlackLevel -ForegroundColor Magenta
            Write-Host "    Cyan % :"$obj.CyanLevel -ForegroundColor Magenta
            Write-Host " Magenta % :"$obj.MagentaLevel -ForegroundColor Magenta
            Write-Host "  Yellow % :"$obj.YellowLevel -ForegroundColor Magenta
            Write-Host " "
        }

        $HtmlData += '<tr>'
        If ($obj.Connection -eq "Offline"){
            $StatusCol = '<td><strong><font color=red>'+$Obj.Connection+'</strong></font></td>'
        }Else{
            $StatusCol = '<td><font color=green>'+$Obj.Connection+'</font></td>'
        }
        If ($Obj.BlackLevel -le $ExtOption.TriggerLevel){
            $BlackCol = '<td><strong><font color=red>'+$Obj.BlackLevel+'</strong></font></td>'
            $Low = $True
            If ([Int]$Obj.BlackLevel -le 0){
                $InkOut = $True
                $ExtOption | Add-Member -MemberType NoteProperty -Name "Priority" -Value $InkOut -force
            }
        }Else{
            $BlackCol = '<td>'+$Obj.BlackLevel+'</td>'
        }   
        If ($Obj.CyanLevel -le $ExtOption.TriggerLevel){
            $CyanCol = '<td><strong><font color=red>'+$Obj.CyanLevel+'</strong></font></td>'
            $Low = $True
            If ([Int]$Obj.CyanLevel -le 0){
                $InkOut = $True
                $ExtOption | Add-Member -MemberType NoteProperty -Name "Priority" -Value $InkOut -force
            }
        }Else{
            $CyanCol = '<td>'+$Obj.CyanLevel+'</td>'
        }
        If ($Obj.MagentaLevel -le $ExtOption.TriggerLevel){
            $MagCol = '<td><strong><font color=red>'+$Obj.MagentaLevel+'</strong></font></td>'
            $Low = $True        
            If ([Int]$Obj.MagentaLevel -le 0){
                $InkOut = $True
                $ExtOption | Add-Member -MemberType NoteProperty -Name "Priority" -Value $InkOut -force
            }
        }Else{
            $MagCol = '<td>'+$Obj.MagentaLevel+'</td>'
        }
        If ($Obj.YellowLevel -le $ExtOption.TriggerLevel){
            $YellowCol = '<td><strong><font color=red>'+$Obj.YellowLevel+'</strong></font></td>'
            $Low = $True
            If ([Int]$Obj.YellowLevel -le 0){
                $InkOut = $True
                $ExtOption | Add-Member -MemberType NoteProperty -Name "Priority" -Value $InkOut -force
            }
        }Else{
            $YellowCol = '<td>'+$Obj.YellowLevel+'</td>'
        }

        If ($obj.HostNameLookup){
            If ($InkOut){
                $HostCol = '<td><font color=green>'+$Obj.HostName+'<br><strong><font color=red><center>Toner Out!</center></strong></font></td>'
            }Else{
                $HostCOl = '<td><font color=green>'+$Obj.HostName+'</font></td>'
            }
        }Else{
            If ($InkOut){
                $HostCol = '<td><font color=red>'+$Obj.HostName+'<br><strong><center>DNS & Toner Issues !</center></strong></font></td>'
            }Else{
                $HostCol = '<td><font color=red>'+$Obj.HostName+'<br><strong><center>DNS Issue !</center></strong></font></td>'
            }
        }

        $HtmlData += $HostCol+$StatusCol+$BlackCol+$CyanCol+$MagCol+$YellowCol
        $HtmlData += '<td>'+$Obj.IPAddress+'</td>'
        $HtmlData += '<td>'+$Obj.Contact+'</td>'
        $HtmlData += '<td>'+$Obj.Department+'</td>'
        $HtmlData += '<td>'+$Obj.Building+'</td>'
        $HtmlData += '<td>'+$Obj.Address+'</td>'
        $HtmlData += '<td>'+$Obj.MFG+'</td>'
        $HtmlData += '<td>'+$Obj.Model+'</td>'
        $HtmlData += '<td>'+$Obj.Serial+'</td>'
        $HtmlData += '<td>'+$Obj.Firmware+'</td>'
        $HtmlData += '</tr>'
        If (($Low) -and ($Obj.Connection -ne "Offline")){
            $AddEmail = $AddEmail+$obj.Email+";"            
        }
    }
    $ExtOption | Add-Member -MemberType NoteProperty -Name "AddEmail" -Value $AddEmail -force
}

$DateTime = Get-Date -Format MM-dd-yyyy_hh:mm:ss 
$HtmlData += '<tr><td colspan='+$ColSpan+'15><center><font color=darkcyan><strong>Audit completed at: '+$DateTime+'</strong></center></td></tr>'   
$HtmlData += '</table>'
$HtmlData += 'Script Version: '+$ScriptVer
$HtmlData += '</div></body></html>'

$HtmlNotice = @()
$HtmlNotice += 'Good morning,'
$HtmlNotice += '<br><br>This email is a courtesy notice because you are listed as the contact person for one or more Ricoh printers.'
$HtmlNotice += '&nbsp;&nbsp;One or more of your printers has at least one toner cartridge that is low.'
$HtmlNotice += '&nbsp;&nbsp;Please review the included report and locate any printers you have responsibility for.'
$HtmlNotice += '&nbsp;&nbsp;Printers with low toner levels have those levels denoted in red.'
$HtmlNotice += '&nbsp;&nbsp;Please note that ordering of consumable items such as toner are the responsibility of the end user or department.'
$HtmlNotice += '<br><br>We recommend ordering toner now so that the printer does not completely run out.&nbsp;&nbsp;To order toner, please visit <a href="https://MyRicoh.com">MyRicoh.com</a>. '
$HtmlNotice +=  '<br><br>'+$ExtOption.HTML1
If ($extOption.Priority){
    $HtmlNotice += '<br><br><font color=red><strong>NOTICE!!!  At lease one printer has run out of toner.  Please check below for a negative number.</strong>'
}
$HtmlNotice += '<br><br><font color=black>Thanks,<br>'+$ExtOption.HTML2+'<br><br>'

#--[ Only add the notice and/or send email if something triggered the 20% limit ]--
If ($Null -ne $ExtOption.AddEmail){
    $HtmlData = $HtmlHeader+$HtmlNotice+$HtmlData
    $ExtOption | Add-Member -MemberType NoteProperty -Name "AddEmail" -Value $AddEmail -force
    $ExtOption | Add-Member -MemberType NoteProperty -Name "Alert" -Value $True -force
}Else{  
    $HtmlData = $HtmlHeader+$HtmlData
}

If ($ExtOption.ConsoleState){Write-Host ""}
$Msg = "Ready to send report..."
StatusMsg $Msg "Yellow" $ExtOption
    
#--[ Set the alternate email recipient if running out of an IDE console for testing ]-- 
If ($Env:Username.SubString(0,1) -eq "a"){       #--[ Filter out admin accounts ]--
    $ThisUser = ($Env:Username.SubString(1))+"@"+$Env:USERDNSDOMAIN 
    $ExtOption | Add-Member -MemberType NoteProperty -Name "EmailAltRecipient" -Value $ThisUser -force
}Else{
    $ThisUser = $Env:USERNAME+"@"+$Env:USERDNSDOMAIN 
    $ExtOption | Add-Member -MemberType NoteProperty -Name "EmailAltRecipient" -Value $ThisUser -force
}

SendEmail $HtmlData $ExtOption

If ($ExtOption.ConsoleState){Write-Host "`n--- Completed ---" -foregroundcolor red}
