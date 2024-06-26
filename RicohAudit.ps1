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
                   :
#===============================================================================#>
#requires -version 5.0
clear-host
 
[string]$DateTime = Get-Date -Format MM-dd-yyyy_HHmmss 
$CloseAnyOpenXL = $False
$ExtOption = New-Object -TypeName psobject #--[ Object to hold runtime options ]--
$ScriptName = ($MyInvocation.MyCommand.Name).split(".")[0] 
$ExtOption | Add-Member -Force -MemberType NoteProperty -Name "LogFile" -Value ($PSScriptRoot+'\'+$ScriptName+'.log')
$TriggerLevel = 20

#--[ Runtime Adjustments ]--
$Console = $false
$Debug = $false
#--------------------------
If ($Debug){
    $Console = $True
}
$erroractionpreference = "stop"

#==[ Functions, Arrays, & Lists ]===========================================
Function InstallModules{
    Try{
        if (!(Get-Module -Name PnP.PowerShell)) {       
            Get-Module -ListAvailable PnP.PowerShell | Import-Module | Out-Null    
            Install-Module -Name PnP.PowerShell -RequiredVersion 2.2.156
            Install-Module -Name PnP.PowerShell -RequiredVersion 1.12.0
        }
    }Catch{
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
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailRecipient" -Value $Config.Settings.General.EmailRecipient
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "EmailSender" -Value $Config.Settings.General.EmailSender
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "ExcelFileName" -Value $Config.Settings.General.ExcelFileName
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "ExcelFilePath" -Value $Config.Settings.General.ExcelFilePath
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SmtpServer" -Value $Config.Settings.General.SmtpServer
        $ExtOption | Add-Member -Force -MemberType NoteProperty -Name "SmtpPort" -Value $Config.Settings.General.SmtpPort
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
    Write-Host "-- Script Status: $Msg" -ForegroundColor $Color
    $Msg = ""
}

Function SendEmail ($MessageBody,$ExtOption) {    
    $ErrorActionPreference = "Stop"
    $Smtp = New-Object Net.Mail.SmtpClient($ExtOption.SmtpServer,$ExtOption.SmtpPort) 
    $Email = New-Object System.Net.Mail.MailMessage  
    $Email.IsBodyHTML = $true
    $Email.From = $ExtOption.EmailSender
    If ($ExtOption.ConsoleState){  #--[ If running out of an IDE console, send only to the user for testing ]-- 
        $Email.To.Add($ExtOption.EmailAltRecipient)  
    }Else{
        If ($ExtOption.Alert){  #--[ If a device failed self-test or trigger day is matched send to main recipient ]--
            $Email.To.Add($ExtOption.EmailRecipient)  
            #$Email.To.Add($ExtOption.EmailAltRecipient)   #--[ In case this user isn't part of the group email ]--  
            ForEach ($Address in ($ExtOption.AddEmail.split(";"))){
                If ($Address -ne ""){
                    $Email.To.Add($Address)
                }
            }
        }
    }

    $Email.Subject = "Ricoh Printer Status Report"
    $Email.Body = $MessageBody
    If ($ExtOption.Debug){
        $Msg="-- Email Parameters --" 
        StatusMsg $Msg "yellow" $ExtOption
        $Msg="Error Msg     = "+$_.Error.Message
        StatusMsg $Msg "yellow" $ExtOption
        $Msg="Exception Msg = "+$_.Exception.Message
        StatusMsg $Msg "yellow" $ExtOption
        $Msg="Local Sender  = "+$ThisUser
        StatusMsg $Msg "yellow" $ExtOption
        $Msg="Recipient     = "+$ExtOption.EmailRecipient
        StatusMsg $Msg "yellow" $ExtOption
        $Msg="SMTP Server   = "+$ExtOption.SmtpServer
        StatusMsg $Msg "yellow" $ExtOption
    }
    $ErrorActionPreference = "stop"
    Try {
        If ($ExtOption.Alert){
            $Smtp.Send($Email)
            If ($ExtOption.ConsoleState){Write-Host `n"--- Email Sent ---" -ForegroundColor red }
        }
    }Catch{
        Write-host "-- Error sending email --" -ForegroundColor Red
        Write-host "Error Msg     = "$_.Error.Message
        StatusMsg  $_.Error.Message "red" $ExtOption
        Write-host "Exception Msg = "$_.Exception.Message
        StatusMsg  $_.Exception.Message "red" $ExtOption
        Write-host "Local Sender  = "$ThisUser
        Write-host "Recipient     = "$ExtOption.EmailRecipient
        Write-host "SMTP Server   = "$ExtOption.SmtpServer
        add-content -path $psscriptroot -value  $_.Error.Message
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

InstallModules

#--[ Load external XML options file ]--
$ConfigFile = $PSScriptRoot+"\"+($MyInvocation.MyCommand.Name.Split("_")[0]).Split(".")[0]+".xml"
$ExtOption = LoadConfig $ExtOption $ConfigFile
If ($NUll -ne $ExtOption.SmtpServer){
    StatusMsg "External config file loaded successfully." "Magenta" $ExtOption
}

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

    # $NetworkFile = $ExtOption.ExcelFilePath+"/"+$ExtOption.ExcelFileName
    $SourceFile = $ExtOption.ExcelFilePath+"\"+$ExtOption.ExcelFileName
    $LocalFile = $PSScriptRoot+"\"+$ExtOption.ExcelFileName
    # $TempFile = $PsScriptroot+"\"+$ExtOption.ExcelFileName+".bak"

    If (Test-Path -Path $SourceFile -PathType Leaf){   #--[ Is there a valid source file? ]--
        StatusMsg "Source file located, copying to local folder..." "Magenta" $ExtOption  
        Try{ 
            Copy-Item -Path $SourceFile -destination $LocalFile -Force -ErrorAction "stop"
            $Source = $LocalFile
        }Catch{
            StatusMsg "Source file copy failed..." "Red" $ExtOption
        }
    }Else{ 
        If (Test-Path -Path $LocalFile -PathType Leaf){
            StatusMsg "Local file detected, loading it..." "Magenta" $ExtOption
            #Rename-Item $LocalFile -NewName $TempFile
            $Source = $LocalFile
            #Start-Sleep -sec 1
   #         }ElseIf (Test-Path -Path $TempFile -PathType Leaf){
   #             StatusMsg "Local backup file detected, renaming and loading it..." "Magenta" $ExtOption
   #             Rename-Item $TempFile -NewName $LocalFile
   #             $Source = $LocalFile
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
    Write-Host "   ." -NoNewline -ForegroundColor Cyan
    Do {
        $CurrentIP = $WorkSheet.Cells.Item($Row,14).Text 
        $Hostname = $WorkSheet.Cells.Item($Row,13).Text 
        $Building = $WorkSheet.Cells.Item($Row,2).Text             
        $Department = $WorkSheet.Cells.Item($Row,3).Text 
        $Address = $WorkSheet.Cells.Item($Row,5).Text 
        $Contact = $WorkSheet.Cells.Item($Row,6).Text 
        $Email = $WorkSheet.Cells.Item($Row,7).Text 
        If ($CurrentIP -ne $PreviousIP){  #--[ Make sure IPs are added only once ]--
            Write-Host "." -NoNewline -ForegroundColor Cyan
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

    Write-Host ""
    $Msg = "$Counter Total IP addresses detected."
    StatusMsg $Msg "Magenta" $ExtOption
    Write-Host ""
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
$HtmlData += '<tr><td colspan='+$ColSpan+'><center><font color=black>Levels displayed below indicate toner percent remaining.&nbsp;&nbsp;&nbsp;Alert triggered at levels below '+$TriggerLevel+'%.</center></td></tr>'

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
        $Low = 0
        If ($Console){Write-Host "" } #`nCurrent Target  :"$Target -ForegroundColor Yellow }
        $Obj = New-Object -TypeName psobject   #--[ Collection for Results ]--
        Try{
            $HostLookup = (nslookup $Target.Split(";")[0] $Env:LogonServer 2>&1) 
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
            If ($Console){Write-Host "  Working." -NoNewline}
            ForEach ($OID in $OIDArray){     
                If ($Obj.Connection -eq "Online"){  #--[ Only process OIDs if online  ]--------------------------
                    $Result = GetSNMPv1 $Target.Split(";")[0] $OID $False #$Script:v3UserTest
                    Write-Host "." -NoNewline
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
            Write-Host "."
        }Else{
            $Obj | Add-Member -MemberType NoteProperty -Name "Connection" -Value "Offline" -force
        }

        If ($Console){
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
        }

        $HtmlData += '<tr>'
        If ($obj.HostNameLookup){
            $HtmlData += '<td><font color=green>'+$Obj.HostName+'</font></td>'
        }Else{
            $HtmlData += '<td><font color=red>'+$Obj.HostName+'</font></td>'
        }
        If ($obj.Connection -eq "Offline"){
            $HtmlData += '<td><strong><font color=red>'+$Obj.Connection+'</strong></font></td>'
        }Else{
            $HtmlData += '<td><font color=green>'+$Obj.Connection+'</font></td>'
        }
        If ($Obj.BlackLevel -le $TriggerLevel){
            $HtmlData += '<td><strong><font color=red>'+$Obj.BlackLevel+'</strong></font></td>'
            $Low ++
        }Else{
            $HtmlData += '<td>'+$Obj.BlackLevel+'</td>'
        }   
        If ($Obj.CyanLevel -le $TriggerLevel){
            $HtmlData += '<td><strong><font color=red>'+$Obj.CyanLevel+'</strong></font></td>'
            $Low ++
        }Else{
            $HtmlData += '<td>'+$Obj.CyanLevel+'</td>'
        }
        If ($Obj.MagentaLevel -le $TriggerLevel){
            $HtmlData += '<td><strong><font color=red>'+$Obj.MagentaLevel+'</strong></font></td>'
            $Low ++
        }Else{
            $HtmlData += '<td>'+$Obj.MagentaLevel+'</td>'
        }
        If ($Obj.YellowLevel -le $TriggerLevel){
            $HtmlData += '<td><strong><font color=red>'+$Obj.YellowLevel+'</strong></font></td>'
            $Low ++
        }Else{
            $HtmlData += '<td>'+$Obj.YellowLevel+'</td>'
        }
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
        If (($Low -gt 0) -and ($Obj.Connection -ne "Offline")){
            $AddEmail = $AddEmail+$obj.Email+";"            
        }
    }
    $ExtOption | Add-Member -MemberType NoteProperty -Name "AddEmail" -Value $AddEmail -force
}

$DateTime = Get-Date -Format MM-dd-yyyy_hh:mm:ss 
$HtmlData += '<tr><td colspan='+$ColSpan+'15><center><font color=darkcyan><strong>Audit completed at: '+$DateTime+'</strong></center></td></tr>'   
$HtmlData += '</table></div></body></html>'

$HtmlNotice = @()
$HtmlNotice += 'Good morning,'
$HtmlNotice += '<br><br>This email is a courtesy notice because you are listed as the contact person for one or more Ricoh printers.'
$HtmlNotice += '&nbsp;&nbsp;One or more of your printers has at least one toner cartridge that is low.'
$HtmlNotice += '&nbsp;&nbsp;Please review the included report and locate any printers you have responsibility for.'
$HtmlNotice += '&nbsp;&nbsp;Printers with low toner levels have those levels denoted in red.'
$HtmlNotice += '&nbsp;&nbsp;Please note that ordering of consumable items such as toner are the responsibility of the end user or department.'
$HtmlNotice += '<br><br>We recommend ordering toner now so that the printer does not completely run out.&nbsp;&nbsp;To order toner, please visit <a href="https://MyRicoh.com">MyRicoh.com</a>. '
$HtmlNotice += "<br><br>If you don't have a MyRicoh account, please submit a request via <a href='https://ithelp.ah.org'>ithelp.ah.org</a> and someone will provide you with instructions."
$HtmlNotice += '<br><br>Thanks,<br>Adventist Health Lodi Memorial IT<br><br>'

#--[ Only add the notice and/or send email if something triggered the 20% limit ]--
If ($Null -ne $ExtOption.AddEmail){
    $HtmlData = $HtmlHeader+$HtmlNotice+$HtmlData
    $ExtOption | Add-Member -MemberType NoteProperty -Name "AddEmail" -Value $AddEmail -force
    $ExtOption | Add-Member -MemberType NoteProperty -Name "Alert" -Value $True -force
}Else{  
    $HtmlData = $HtmlHeader+$HtmlData
}

Write-Host ""
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

If ($Console){Write-Host "`n--- Completed ---" -foregroundcolor red}
