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
$defaultresiavesize = $xml.config.defaultresiavesize

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
	    $Mailmessage.Subject = ($MailSubject)
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

# Функция установки индивидуальных лимитов на размер отправляемого и получаемого письма. Вход $usermailbox - почтовый ящик пользователя
# $SendLimit - лимит на размер отправляемого сообщения (число в Мб), $Resivelimit - лимит на размер получаемого сообщения (число в Мб)
#Function ChangeSendReciveLimit
#	{
#	Praram($usermailbox = $null,$SendLimit = $defaultsendsize,$Resivelimit = $defaultresiavesize)
#	If ($usermailbox.MaxSendSize -notmatch ("$SendLimit" + " MB"))
#		{
#		Set-Mailbox -Identity $usermailbox -MaxSendSize $SendLimit
#		}
#	If ($usermailbox.MaxReceiveSize -notmatch ("$Resivelimit" + " MB"))
#		{
#		Set-Mailbox -Identity $usermailbox -MaxReceiveSize $Resivelimit
#		}
#	}

# Полчаем участников всех групп разрешения почтовых функций.
$AllCasMailFutures = Get-CASMailbox -OrganizationalUnit $SearchOU 
$OWAEnabledGroupMembers = (Get-Group -Identity "OWAEnabled").Members
$PopEnabledGroupMembers = (Get-Group -Identity "PopEnabled").Members
$ActiveSyncEnabledGroupMembers = (Get-Group -Identity "ActiveSyncEnabled").Members
$ImapEnabledGroupMembers = (Get-Group -Identity "ImapEnabled").Members


#Основное тело скрипта.
$allmailboxes = Get-Mailbox -OrganizationalUnit $SearchOU 
Foreach ($usermailbox in $allmailboxes) 
	{
	If ($usermailbox.ArchiveDatabase -eq $null)
		{
		Enable-Mailbox -Identity $usermailbox.Identity -Archive -ArchiveDatabase $archiveDatabase
		}
		
# OWA 		
	If ($OWAEnabledGroupMembers -contains $usermailbox.Identity)										# Блок проверки разрешений на указанную функцию
		{																								# Проверяем, учавствует ли пользователь в группе разрешения функции
		If (!($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).OWAEnabled)		# Если учавствует, запрашиваем текущее состояние функции
			{																							#
			Set-CASMailbox -Identity $usermailbox.Identity -OWAEnabled $true							# Если Функция отключена, включаем её.
			}																							#
		}																								#
	else																								# Если пользователь не учавствует в группе разрешения функции
		{																								#
		If (($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).OWAEnabled)		# Проверяем её текущее состояние
			{																							#
			Set-CASMailbox -Identity $usermailbox.Identity -OWAEnabled $false							# Если функция включена, отключаем её.
			}																							#
		}
		
# POP3
	If ($PopEnabledGroupMembers -contains $usermailbox.Identity)										
		{																								
		If (!($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).PopEnabled)		
			{																							
			Set-CASMailbox -Identity $usermailbox.Identity -PopEnabled $true							
			}																							
		}																								
	else																								
		{																								
		If (($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).PopEnabled)		
			{																							
			Set-CASMailbox -Identity $usermailbox.Identity -PopEnabled $false							
			}																							
		}
		
# ActiveSync
	If ($ActiveSyncEnabledGroupMembers -contains $usermailbox.Identity)										
		{																								
		If (!($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).ActiveSyncEnabled)		
			{																							
			Set-CASMailbox -Identity $usermailbox.Identity -ActiveSyncEnabled $true							
			}																							
		}																								
	else																								
		{																								
		If (($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).ActiveSyncEnabled)		
			{																							
			Set-CASMailbox -Identity $usermailbox.Identity -ActiveSyncEnabled $false							
			}																							
		}
		
# IMAP
	If ($ImapEnabledGroupMembers -contains $usermailbox.Identity)										
		{																								
		If (!($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).ImapEnabled)		
			{																							
			Set-CASMailbox -Identity $usermailbox.Identity -ImapEnabled $true							
			}																							
		}																								
	else																								
		{																								
		If (($AllCasMailFutures | Where-Object {$_.name -eq $usermailbox.DisplayName}).ImapEnabled)		
			{																							
			Set-CASMailbox -Identity $usermailbox.Identity -ImapEnabled $false							
			}																							
		}
		
#Получаем информацию о группах в готорые включен пользователь.
	
	$AlluserGroups = (Get-ADUser $usermailbox.Guid -Properties memberof).memberof							# Получаем все группы в которых учавствует пользователь
	$AllSendLimitGroups = $AlluserGroups -match "ReceiveLimit_" 											# Выделяем только группы в названии которых есть "ReceiveLimit_"
	If ($AllSendLimitGroups.Count -eq 0)																	# Если ни одной такой группы нет то  
		{																									# проверяем отличие лимита от настроек по умолчанию
		If ($usermailbox.MaxReceiveSize -notmatch ("$defaultresiavesize" + " MB"))							#
			{																								# 
			Set-Mailbox -Identity $usermailbox.Identity -MaxReceiveSize ("$defaultresiavesize" + " MB")		# если настройки не верные - задаем лимит по умолчанию
			}																								#
		}																									# 
	if ($AllSendLimitGroups.Count -eq 1)																	# Если есть 1-на группа
		{																									# 
		$maxReceveLimit = (Get-ADGroup $AllSendLimitGroups[0]).name -replace "ReceiveLimit_",""				# выделяем из её названия число - размер ограничения
		If ($usermailbox.MaxReceiveSize -notmatch ("$maxReceveLimit" + " MB"))								# проверяем соответствие установленного лимита и указанного ограничения
			{																								#
			Set-Mailbox -Identity $usermailbox.Identity -MaxReceiveSize ("$maxReceveLimit" + " MB")			# если значения не соответствуют - перезаписываем его
			}																								#
		}																									#
 	if ($AllSendLimitGroups.Count -gt 1)																	# Если групп больше одной - пока не знаю что делать.
		{
		# тут надо написать обработчик наличия у пользователя нескольких групп изменения размера принимаемого сообщения
		}
		
		
	$AllSendLimitGroups = $AlluserGroups -match "SendLimit_" 												# Выделяем только группы в названии которых есть "SendLimit_"
	If ($AllSendLimitGroups.Count -eq 0)																	# Если ни одной такой группы нет то  
		{																									# проверяем отличие лимита от настроек по умолчанию
		If ($usermailbox.MaxReceiveSize -notmatch ("$defaultsendsize" + " MB"))								#
			{																								# 
			Set-Mailbox -Identity $usermailbox.Identity -MaxSendSize ("$defaultsendsize" + " MB")			# если настройки не верные - задаем лимит по умолчанию
			}																								#
		}																									# 
	if ($AllSendLimitGroups.Count -eq 1)																	# Если есть 1-на группа
		{																									# 
		$maxSendLimit = (Get-ADGroup $AllSendLimitGroups[0]).name -replace "SendLimit_",""					# выделяем из её названия число - размер ограничения
		If ($usermailbox.MaxReceiveSize -notmatch ("$maxSendLimit" + " MB"))								# проверяем соответствие установленного лимита и указанного ограничения
			{																								#
			Set-Mailbox -Identity $usermailbox.Identity -MaxSendSize ("$maxSendLimit" + " MB")			# если значения не соответствуют - перезаписываем его
			}																								#
		}																									#
 	if ($AllSendLimitGroups.Count -gt 1)																	# Если групп больше одной - пока не знаю что делать.
		{
		# тут надо написать обработчик наличия у пользователя нескольких групп изменения размера принимаемого сообщения
		}	
		
		
	
	}















	  
	  