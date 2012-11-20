Param($ReportFor)
$curDir = $MyInvocation.MyCommand.Definition | split-path -parent
$xml = [xml](get-content "$curDir\cmconf.xml")
$ConnectionUri = $xml.config.ConnectionUri
$SearchOU = $xml.config.SearchOU
$mailserver = $xml.config.mailserver# Сервер с которого будет отпраляться почта. 
$mailrecepient = $xml.config.mailrecepient # ящик получателя 
$Mailsender = $xml.config.Mailsender
$mailUser = $xml.config.mailUser
$mailUserPassword = $xml.config.mailUserPassword
$archiveDatabase = $xml.config.archivedatabase
$defaultsendsize = $xml.config.defaultsendsize
$defaultReceivesize = $xml.config.defaultReceivesize

$LogFileName = ($env:TEMP + "\Mail change config " + (Get-Date -Format d) + ".log")
("Запуск скрипта -- " + (Get-Date)) >> $LogFileName

# Подключаем оснастку ActiveDirectory . Данная команда работает на WS2008 и W7
import-module activedirectory -Cmdlet get-adgroup, get-aduser

# Подключаем сеанс Exchange 
$session = New-PSSession -Configurationname Microsoft.Exchange –ConnectionUri $ConnectionUri
Import-PSSession $session -AllowClobber -commandname Get-User, Get-Group, Enable-Mailbox, Get-CASMailbox, Set-CASMailbox, Get-MailBox, Set-Mailbox

# Создаем сеанс для отправки почты.
$SmtpClient = New-object system.net.mail.SmtpClient 
$SmtpClient.Host = $mailserver 
$SmtpClient.Port = 25
$SmtpClient.Credentials = New-Object System.Net.NetworkCredential($mailUser, $mailUserPassword)
$SmtpClient.EnableSsl = $true

# Функция отправки письма 
Function SendMailMessage
     {
      Param($MailSubject,$logFile,$Body)
        $MailMessage = New-Object system.net.mail.MailMessage
	    $MailMessage.from = $Mailsender
	    $Mailmessage.Subject = $MailSubject
		If ($logFile -ne $null)
			{
			$Attachment = New-Object Net.mail.Attachment($logFile)
        	$MailMessage.Attachments.Add($Attachment)
			}
		$Mailmessage.To.Add($mailrecepient)
		$Mailmessage.IsBodyHtml = $True
        $mailmessage.Body = ($Body)
		$SmtpClient.Send($mailmessage)
        If ($logFile -ne $null){$Attachment.Dispose()}     
      
      }

If ($ReportFor -ne $null)
	{
	$ReportFileName = ($env:TEMP + "\Mail change config report " + (Get-Date -uFormat "%Y _ %m _ %d") + ".html")
	# Начанаем формиравание HTML файла отчета.
	"<!DOCTYPE ШТМЛ PUBLIC `"-//W3C//DTD ШТМЛ 4.01//EN`" `"http://www.w3.org/TR/ШТМЛ4/strict.dtd`">" > $ReportFileName
	"<html>" >> $ReportFileName
	"<head>" >> $ReportFileName
	"<meta http-equiv=`"Content-Type`" content=`"text/html; charset=windows-1251`">" >> $ReportFileName
	"<style type=`"text/css`">" >> $ReportFileName
	"table {width: 800px; border: 1px solid black;}" >> $ReportFileName
	"TD {padding: 3px; border: 1px solid black;}" >> $ReportFileName
	"TD.tl {border-bottom: 2px solid black;}" >> $ReportFileName
	"TH {text-align: center; padding: 3px; border: 1px solid black;}" >> $ReportFileName
	"TH.vl {border-bottom: 2px solid black;}" >> $ReportFileName
	"TH.v2 {font-size: 80%; font-weight: normal;}" >> $ReportFileName
	"TH.v3 {font-size: 80%; font-weight:normal; background-color:green}" >> $ReportFileName
	"TH.v4 {font-size: 80%; font-weight:normal; background-color:yellow}" >> $ReportFileName
	"</style>" >> $ReportFileName
	"</head>" >> $ReportFileName
	"<body>" >> $ReportFileName
	"<table>"  >> $ReportFileName
	"<tr class=`"t1`"><th>Пользователь</th><th class=`"v1`">OWAEnabled</th><th class=`"v1`">PopEnabled</th><th class=`"v1`">ImapEnabled</th><th class=`"v1`">ActiveSyncEnabled</th><th class=`"v1`">Send Limit</th><th class=`"v1`">Receive Limit</th></tr>" >> $ReportFileName	
		
	If ($ReportFor -eq "all")
		{
		$AllCasMailFutures = Get-CASMailbox -OrganizationalUnit $SearchOU
		$allmailboxes = Get-Mailbox -OrganizationalUnit $SearchOU
		
		Foreach ($usermailbox in $allmailboxes)
			{
			"<tr><th class=`"v2`">" + $usermailbox.DisplayName + "</th> `
			<th class=`"v2`">" + (($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).OWAEnabled) + "</th> `
			<th class=`"v2`">" + (($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).PopEnabled) + "</th> `
			<th class=`"v2`">" + (($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).ImapEnabled) + "</th> `
			<th class=`"v2`">" + (($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).ActiveSyncEnabled) + "</th> `
			<th class=`"v2`">" + $usermailbox.MaxSendSize + "</th> `
			<th class=`"v2`">" + $usermailbox.MaxReceiveSize + "</th> `
			</tr>" >> $ReportFileName
			}
		"</table></body></html>" >> $ReportFileName
		}
	else
		{
		$usermailbox = Get-Mailbox -Identity $ReportFor
		$userCasFuture = Get-CASMailbox -Identity $ReportFor
		"<tr><th class=`"v2`">" + $usermailbox.DisplayName + "</th> `
		<th class=`"v2`">" + $userCasFuture.OWAEnabled + "</th> `
		<th class=`"v2`">" + $userCasFuture.PopEnabled + "</th> `
		<th class=`"v2`">" + $userCasFuture.ImapEnabled + "</th> `
		<th class=`"v2`">" + $userCasFuture.ActiveSyncEnabled + "</th> `
		<th class=`"v2`">" + $usermailbox.MaxSendSize + "</th> `
		<th class=`"v2`">" + $usermailbox.MaxReceiveSize + "</th> `
		</tr>" >> $ReportFileName
		}
	.$ReportFileName
	Exit
	}

# Полчаем участников всех групп разрешения почтовых функций.
$AllCasMailFutures = Get-CASMailbox -OrganizationalUnit $SearchOU 
$OWAEnabledGroupMembers = (Get-Group -Identity "НП_OWAEnabled").Members
$PopEnabledGroupMembers = (Get-Group -Identity "НП_PopEnabled").Members
$ActiveSyncEnabledGroupMembers = (Get-Group -Identity "НП_ActiveSyncEnabled").Members
$ImapEnabledGroupMembers = (Get-Group -Identity "НП_ImapEnabled").Members
$SendMailMessageFlag = $false


#Основное тело скрипта.
$allmailboxes = Get-Mailbox -OrganizationalUnit $SearchOU 
Foreach ($usermailbox in $allmailboxes) 
	{
	$error.clear()
	If ($usermailbox.ArchiveDatabase -eq $null)
		{
		Enable-Mailbox -Identity $usermailbox.Identity -Archive -ArchiveDatabase $archiveDatabase
		if ($error.Count -ge 0)
			{
			$SendMailMessageFlag = $true
			$Body += "`n Произошла ошибка при выполнении активации архивного почтового ящика для пользователя $usermailbox.name. Подробности в log файле."
			"Произошла ошибка при выполнении активации архивного почтового ящика для пользователя $usermailbox.name." >> $LogFileName
			$Error[0] >> $LogFileName
			$error.clear()
			}
		}
		
# OWA 		
	If ($OWAEnabledGroupMembers -contains $usermailbox.Identity)										# Блок проверки разрешений на указанную функцию
		{																								# Проверяем, учавствует ли пользователь в группе разрешения функции
		If (!($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).OWAEnabled)		# Если учавствует, запрашиваем текущее состояние функции
			{																							#
			Set-CASMailbox -Identity $usermailbox.Identity -OWAEnabled $true							# Если Функция отключена, включаем её.
			if ($error.Count -ge 0)
				{																							# Блок ловли ошибок.
				$SendMailMessageFlag = $true
				$Body += "`n Произошла ошибка при выполнении активации функции OWA для пользователя $usermailbox.name. Подробности в log файле."
				"Произошла ошибка при выполнении активации функции OWA для пользователя $usermailbox.name." >> $LogFileName
				$Error[0] >> $LogFileName
				$error.clear()
				}
			}																							#
		}																								#
	else																								# Если пользователь не учавствует в группе разрешения функции
		{																								#
		If (($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).OWAEnabled)		# Проверяем её текущее состояние
			{
			if ($error.Count -ge 0)
				{																						# Блок ловли ошибок.
				$SendMailMessageFlag = $true
				$Body += "`n Произошла ошибка при выполнении деактивации функции OWA для пользователя $usermailbox.name. Подробности в log файле."
				"Произошла ошибка при выполнении деактивации функции OWA для пользователя $usermailbox.name." >> $LogFileName
				$Error[0] >> $LogFileName
				$error.clear()
				}#
			Set-CASMailbox -Identity $usermailbox.Identity -OWAEnabled $false							# Если функция включена, отключаем её.
			}																							#
		}
		
# POP3
	If ($PopEnabledGroupMembers -contains $usermailbox.Identity)										
		{																								
		If (!($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).PopEnabled)		
			{																							
			Set-CASMailbox -Identity $usermailbox.Identity -PopEnabled $true
			if ($error.Count -ge 0)
				{																							# Блок ловли ошибок.
				$SendMailMessageFlag = $true
				$Body += "`n Произошла ошибка при выполнении активации функции POP3 для пользователя $usermailbox.name. Подробности в log файле."
				"Произошла ошибка при выполнении активации функции POP3 для пользователя $usermailbox.name." >> $LogFileName
				$Error[0] >> $LogFileName
				$error.clear()
				}
			}																							
		}																								
	else																								
		{																								
		If (($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).PopEnabled)		
			{																							
			Set-CASMailbox -Identity $usermailbox.Identity -PopEnabled $false
			if ($error.Count -ge 0)
				{																							# Блок ловли ошибок.
				$SendMailMessageFlag = $true
				$Body += "`n Произошла ошибка при выполнении деактивации функции POP3 для пользователя $usermailbox.name. Подробности в log файле."
				"Произошла ошибка при выполнении деактивации функции POP3 для пользователя $usermailbox.name." >> $LogFileName
				$Error[0] >> $LogFileName
				$error.clear()
				}
			}																							
		}
		
# ActiveSync
	If ($ActiveSyncEnabledGroupMembers -contains $usermailbox.Identity)										
		{																								
		If (!($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).ActiveSyncEnabled)		
			{																							
			Set-CASMailbox -Identity $usermailbox.Identity -ActiveSyncEnabled $true
			if ($error.Count -ge 0)
				{																							# Блок ловли ошибок.
				$SendMailMessageFlag = $true
				$Body += "`n Произошла ошибка при выполнении активации функции ActiveSync для пользователя $usermailbox.name. Подробности в log файле."
				"Произошла ошибка при выполнении активации функции ActiveSync для пользователя $usermailbox.name." >> $LogFileName
				$Error[0] >> $LogFileName
				$error.clear()
				}
			}																							
		}																								
	else																								
		{																								
		If (($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).ActiveSyncEnabled)		
			{																							
			Set-CASMailbox -Identity $usermailbox.Identity -ActiveSyncEnabled $false
			if ($error.Count -ge 0)
				{																							# Блок ловли ошибок.
				$SendMailMessageFlag = $true
				$Body += "`n Произошла ошибка при выполнении деактивации функции ActiveSync для пользователя $usermailbox.name. Подробности в log файле."
				"Произошла ошибка при выполнении деактивации функции ActiveSync для пользователя $usermailbox.name." >> $LogFileName
				$Error[0] >> $LogFileName
				$error.clear()
				}
			}																							
		}
		
# IMAP
	If ($ImapEnabledGroupMembers -contains $usermailbox.Identity)										
		{																								
		If (!($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).ImapEnabled)		
			{																							
			Set-CASMailbox -Identity $usermailbox.Identity -ImapEnabled $true
			if ($error.Count -ge 0)
				{																							# Блок ловли ошибок.
				$SendMailMessageFlag = $true
				$Body += "`n Произошла ошибка при выполнении активации функции IMAP для пользователя $usermailbox.name. Подробности в log файле."
				"Произошла ошибка при выполнении активации функции IMAP для пользователя $usermailbox.name." >> $LogFileName
				$Error[0] >> $LogFileName
				$error.clear()
				}
			}																							
		}																								
	else																								
		{																								
		If (($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).ImapEnabled)		
			{																							
			Set-CASMailbox -Identity $usermailbox.Identity -ImapEnabled $false
			if ($error.Count -ge 0)
				{																							# Блок ловли ошибок.
				$SendMailMessageFlag = $true
				$Body += "`n Произошла ошибка при выполнении деактивации функции IMAP для пользователя $usermailbox.name. Подробности в log файле."
				"Произошла ошибка при выполнении деактивации функции IMAP для пользователя $usermailbox.name." >> $LogFileName
				$Error[0] >> $LogFileName
				$error.clear()
				}
			}																							
		}
		
#Получаем информацию о группах в готорые включен пользователь.
	
	$AlluserGroups = (Get-ADUser $usermailbox.Guid -Properties memberof).memberof							# Получаем все группы в которых учавствует пользователь
	$AllSendLimitGroups = $AlluserGroups -match "НП_ReceiveLimit_" 											# Выделяем только группы в названии которых есть "ReceiveLimit_"
	If ($AllSendLimitGroups.Count -eq 0)																	# Если ни одной такой группы нет то  
		{																									# проверяем отличие лимита от настроек по умолчанию
		If ($usermailbox.MaxReceiveSize -notmatch ("$defaultReceivesize" + " MB"))							#
			{																								# 
			Set-Mailbox -Identity $usermailbox.Identity -MaxReceiveSize ("$defaultReceivesize" + " MB")		# если настройки не верные - задаем лимит по умолчанию
			if ($error.Count -ge 0)
				{																							# Блок ловли ошибок.
				$SendMailMessageFlag = $true
				$Body += "`n Произошла ошибка при выполнении установки персонального лимита на размер получаемого письма для пользователя $usermailbox.name. Подробности в log файле."
				"Произошла ошибка при выполнении установки персонального лимита на размер получаемого письма для пользователя $usermailbox.name." >> $LogFileName
				$Error[0] >> $LogFileName
				$error.clear()
				}
			}																								#
		}																									# 
	if ($AllSendLimitGroups.Count -eq 1)																	# Если есть 1-на группа
		{																									# 
		$maxReceveLimit = (Get-ADGroup $AllSendLimitGroups[0]).name -replace "НП_ReceiveLimit_",""				# выделяем из её названия число - размер ограничения
		If ($usermailbox.MaxReceiveSize -notmatch ("$maxReceveLimit" + " MB"))								# проверяем соответствие установленного лимита и указанного ограничения
			{																								#
			Set-Mailbox -Identity $usermailbox.Identity -MaxReceiveSize ("$maxReceveLimit" + " MB")			# если значения не соответствуют - перезаписываем его
			if ($error.Count -ge 0)
				{																							# Блок ловли ошибок.
				$SendMailMessageFlag = $true
				$Body += "`n Произошла ошибка при выполнении установки персонального лимита на размер получаемого письма для пользователя $usermailbox.name. Подробности в log файле."
				"Произошла ошибка при выполнении установки персонального лимита на размер получаемого письма для пользователя $usermailbox.name." >> $LogFileName
				$Error[0] >> $LogFileName
				$error.clear()
				}
			}																								#
		}																									#
 	if ($AllSendLimitGroups.Count -gt 1)																	
		{
		$SendMailMessageFlag = $true
		$Body += "`n У пользователя $usermailbox.name более одной группы НП_ReceiveLimit_. Необходимо установить только 1-н лимит на размер получаемого письма"
		}
		
		
	$AllSendLimitGroups = $AlluserGroups -match "НП_SendLimit_" 												# Выделяем только группы в названии которых есть "SendLimit_"
	If ($AllSendLimitGroups.Count -eq 0)																	# Если ни одной такой группы нет то  
		{																									# проверяем отличие лимита от настроек по умолчанию
		If ($usermailbox.MaxSendSize -notmatch ("$defaultsendsize" + " MB"))								#
			{																								# 
			Set-Mailbox -Identity $usermailbox.Identity -MaxSendSize ("$defaultsendsize" + " MB")			# если настройки не верные - задаем лимит по умолчанию
			if ($error.Count -ge 0)
				{																							# Блок ловли ошибок.
				$SendMailMessageFlag = $true
				$Body += "`n Произошла ошибка при выполнении установки персонального лимита на размер отправляемого письма для пользователя $usermailbox.name. Подробности в log файле."
				"Произошла ошибка при выполнении установки персонального лимита на размер отправляемого письма для пользователя $usermailbox.name." >> $LogFileName
				$Error[0] >> $LogFileName
				$error.clear()
				}
			}																								#
		}																									# 
	if ($AllSendLimitGroups.Count -eq 1)																	# Если есть 1-на группа
		{																									# 
		$maxSendLimit = (Get-ADGroup $AllSendLimitGroups[0]).name -replace "НП_SendLimit_",""					# выделяем из её названия число - размер ограничения
		If ($usermailbox.MaxSendSize -notmatch ("$maxSendLimit" + " MB"))									# проверяем соответствие установленного лимита и указанного ограничения
			{																								#
			Set-Mailbox -Identity $usermailbox.Identity -MaxSendSize ("$maxSendLimit" + " MB")				# если значения не соответствуют - перезаписываем его
			if ($error.Count -ge 0)
				{																							# Блок ловли ошибок.
				$SendMailMessageFlag = $true
				$Body += "`n Произошла ошибка при выполнении установки персонального лимита на размер отправляемого письма для пользователя $usermailbox.name. Подробности в log файле."
				"Произошла ошибка при выполнении установки персонального лимита на размер отправляемого письма для пользователя $usermailbox.name." >> $LogFileName
				$Error[0] >> $LogFileName
				$error.clear()
				}
			}																								#
		}																									#
 	if ($AllSendLimitGroups.Count -gt 1)																	
		{
		$SendMailMessageFlag = $true
		$Body += "`n У пользователя $usermailbox.name более одной группы НП_SendLimit_. Необходимо установить только 1-н лимит на размер отправляемого письма"
		}	
	
	}
	If ($SendMailMessageFlag)
		{
		$MailSubject = "Отчет о работе скрипта изменения настроек для почтовых ящиков пользователей"
		SendMailMessage $MailSubject $LogFileName $Body
		}
	Remove-Item $LogFileName -Force














	  
	  