<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim wms
set wms = server.createobject("easymail.WebMailSet")
wms.Load

dim ei
set ei = server.createobject("easymail.sysinfo")
ei.Load

ei.ErrSender = trim(request("ErrSender"))

ei.MaxSendNum = trim(request("MaxSendNum"))
ei.ErrMaxSendNum = trim(request("ErrMaxSendNum"))
ei.smtpMailMaxSize = trim(request("smtpMailMaxSize"))
ei.pop3MailMaxSize = trim(request("pop3MailMaxSize"))
ei.defaultMailBoxSize = trim(request("defaultMailBoxSize"))

if trim(request("restrictSmtpMailSize")) <> "" then
	ei.restrictSmtpMailSize = true
else
	ei.restrictSmtpMailSize = false
end if

if trim(request("restrictPop3MailSize")) <> "" then
	ei.restrictPop3MailSize = true
else
	ei.restrictPop3MailSize = false
end if

if trim(request("onlyMailFromSystem")) <> "" then
	ei.onlyMailFromSystem = true
else
	ei.onlyMailFromSystem = false
end if

if trim(request("onlyRcptToSystem")) <> "" then
	ei.onlyRcptToSystem = true
else
	ei.onlyRcptToSystem = false
end if

ei.smtpport = trim(request("smtpport"))
ei.pop3port = trim(request("pop3port"))
ei.daytimeport = trim(request("daytimeport"))

if trim(request("canUseIsp")) <> "" then
	ei.canUseIsp = true
else
	ei.canUseIsp = false
end if

ei.ispServerName = trim(request("ispServerName"))

if trim(request("isppop3Port")) <> "" then
	ei.isppop3Port = CInt(trim(request("isppop3Port")))
end if

ei.ispUserName = trim(request("ispUserName"))
ei.ispPassword = trim(request("ispPassword"))


if trim(request("recvByName")) <> "" then
	ei.recvByName = true
else
	ei.recvByName = false
end if

if trim(request("useList")) <> "" then
	ei.useList = true
else
	ei.useList = false
end if

ei.listSender = trim(request("listSender"))


if trim(request("useSmtp")) <> "" then
	ei.useSmtp = true
else
	ei.useSmtp = false
end if

if trim(request("usePOP3")) <> "" then
	ei.usePOP3 = true
else
	ei.usePOP3 = false
end if

if trim(request("useDayTime")) <> "" then
	ei.useDayTime = true
else
	ei.useDayTime = false
end if

ei.POP3DownSleepTime = trim(request("POP3DownSleepTime"))
ei.IspDownSleepTime = trim(request("IspDownSleepTime"))

if trim(request("startKill")) <> "" then
	ei.startKill = true
else
	ei.startKill = false
end if

if trim(request("LogSave")) <> "" then
	ei.LogSave = true
else
	ei.LogSave = false
end if


ei.DNS = trim(request("DNS"))

if trim(request("useLogonPass")) <> "" then
	ei.useLogonPass = true
else
	ei.useLogonPass = false
end if


if trim(request("manageisp")) <> "" then
	ei.manageisp = CInt(trim(request("manageisp")))
end if

if trim(request("backupmode")) <> "" then
	ei.backupmode = CInt(trim(request("backupmode")))
end if

if trim(request("backupdate")) <> "" then
	ei.backupdate = CInt(trim(request("backupdate")))
end if


'----- 3.2.0.1
if trim(request("useStakeOut")) <> "" then
	ei.useStakeOut = true
else
	ei.useStakeOut = false
end if

ei.stakeOutTo = trim(request("stakeOutTo"))


if trim(request("webMailMaxLen")) <> "" then
	ei.webMailMaxLen = CLng(trim(request("webMailMaxLen")))
end if

if trim(request("useAutoMailClean")) <> "" then
	ei.useAutoMailClean = true
else
	ei.useAutoMailClean = false
end if

if trim(request("mailMoveDays")) <> "" then
	ei.mailMoveDays = CInt(trim(request("mailMoveDays")))
end if

ei.mailMoveTo = trim(request("mailMoveTo"))

if trim(request("mailDeleteDays")) <> "" then
	ei.mailDeleteDays = CInt(trim(request("mailDeleteDays")))
end if

if trim(request("useAutoUserClean")) <> "" then
	ei.useAutoUserClean = true
else
	ei.useAutoUserClean = false
end if

if trim(request("forbidUserDays")) <> "" then
	ei.forbidUserDays = CInt(trim(request("forbidUserDays")))
end if

if trim(request("deleteUserDays")) <> "" then
	ei.deleteUserDays = CInt(trim(request("deleteUserDays")))
end if


if trim(request("useMailList")) <> "" then
	ei.useMailList = true
else
	ei.useMailList = false
end if

if trim(request("defaultMailsNumber")) <> "" then
	ei.defaultMailsNumber = CInt(trim(request("defaultMailsNumber")))
end if

if trim(request("MaxRecipients")) <> "" then
	ei.MaxRecipients = CInt(trim(request("MaxRecipients")))
end if


if trim(request("useAuth")) <> "" then
	ei.useAuth = true
else
	ei.useAuth = false
end if

if trim(request("useDistributeISPMail")) <> "" then
	ei.useDistributeISPMail = true
else
	ei.useDistributeISPMail = false
end if


ei.distributeISPMailinglist = trim(request("distributeISPMailinglist"))
ei.timeZone = trim(request("timeZone"))


if trim(request("EnablePreHacker")) <> "" then
	ei.EnablePreHacker = true
else
	ei.EnablePreHacker = false
end if


if IsNumeric(trim(request("bWaitMinute"))) = true then
	ei.bWaitMinute = CLng(trim(request("bWaitMinute")))
end if


if trim(request("EnableLimitSmtp")) <> "" then
	ei.EnableLimitSmtp = true
else
	ei.EnableLimitSmtp = false
end if

if trim(request("EnableLimitPop3")) <> "" then
	ei.EnableLimitPop3 = true
else
	ei.EnableLimitPop3 = false
end if

if trim(request("EnableSMS")) <> "" then
	ei.EnableSMS = true
else
	ei.EnableSMS = false
end if

if trim(request("EnableKeywordFilter")) <> "" then
	ei.EnableKeywordFilter = true
else
	ei.EnableKeywordFilter = false
end if


if trim(request("enableOpenRelay")) <> "" then
	ei.enableOpenRelay = true
else
	ei.enableOpenRelay = false
end if

if trim(request("enableCatchAll")) <> "" then
	ei.enableCatchAll = true
else
	ei.enableCatchAll = false
end if

if trim(request("enableCatchToOut")) <> "" then
	ei.enableCatchToOut = true
else
	ei.enableCatchToOut = false
end if

if trim(request("cleanMailIncludeExpiresUser")) <> "" then
	ei.cleanMailIncludeExpiresUser = true
else
	ei.cleanMailIncludeExpiresUser = false
end if

if trim(request("cleanAccoutIncludeExpiresUser")) <> "" then
	ei.cleanAccoutIncludeExpiresUser = true
else
	ei.cleanAccoutIncludeExpiresUser = false
end if

if trim(request("enableDomainMonitor")) <> "" then
	ei.enableDomainMonitor = true
else
	ei.enableDomainMonitor = false
end if

if trim(request("enableKillAttacker")) <> "" then
	ei.enableKillAttacker = true
else
	ei.enableKillAttacker = false
end if

if trim(request("enableAutoKillAttacker")) <> "" then
	ei.enableAutoKillAttacker = true
else
	ei.enableAutoKillAttacker = false
end if

if IsNumeric(trim(request("autoKillAttackerConnectMaxNumber"))) = true then
	ei.autoKillAttackerConnectMaxNumber = CLng(trim(request("autoKillAttackerConnectMaxNumber")))
end if

if IsNumeric(trim(request("autoKillAttackerConnectRate"))) = true then
	ei.autoKillAttackerConnectRate = CLng(trim(request("autoKillAttackerConnectRate")))
end if

if IsNumeric(trim(request("autoKillAttackerExpiresMinute"))) = true then
	ei.autoKillAttackerExpiresMinute = CLng(trim(request("autoKillAttackerExpiresMinute")))
end if

if IsNumeric(trim(request("sysBackupTime"))) = true then
	ei.sysBackupTime = CLng(trim(request("sysBackupTime")))
end if

if trim(request("enableLogAutoRemove")) <> "" then
	ei.enableLogAutoRemove = true
else
	ei.enableLogAutoRemove = false
end if

if IsNumeric(trim(request("logAutoRemoveDay"))) = true then
	ei.logAutoRemoveDay = CLng(trim(request("logAutoRemoveDay")))
end if

if trim(request("enableWebAdminIPLimit")) <> "" then
	ei.enableWebAdminIPLimit = true
else
	ei.enableWebAdminIPLimit = false
end if

ei.webAdminIP = trim(request("webAdminIP"))



if trim(request("enableRec_Cortrol")) <> "" then
	ei.enableRec_Cortrol = true
else
	ei.enableRec_Cortrol = false
end if

if trim(request("enableRec_BelieveDomains")) <> "" then
	ei.enableRec_BelieveDomains = true
else
	ei.enableRec_BelieveDomains = false
end if

if trim(request("enableRec_SendDomains")) <> "" then
	ei.enableRec_SendDomains = true
else
	ei.enableRec_SendDomains = false
end if

if trim(request("enable_IndictSpam")) <> "" then
	ei.enable_IndictSpam = true
else
	ei.enable_IndictSpam = false
end if

if trim(request("enable_AttachmentExName_Filter")) <> "" then
	ei.enable_AttachmentExName_Filter = true
else
	ei.enable_AttachmentExName_Filter = false
end if


if IsNumeric(trim(request("MaxConnNum"))) = true then
	ei.MaxConnNum = CLng(trim(request("MaxConnNum")))
end if

if trim(request("enableScanVirus")) <> "" then
	ei.enableScanVirus = true
else
	ei.enableScanVirus = false
end if

if trim(request("Enable_NoSpam_Affirm")) <> "" then
	ei.Enable_NoSpam_Affirm = true
else
	ei.Enable_NoSpam_Affirm = false
end if



if trim(request("useSSL_POP3")) <> "" then
	ei.useSSL_POP3 = true
else
	ei.useSSL_POP3 = false
end if

if IsNumeric(trim(request("sslpop3port"))) = true then
	ei.sslpop3port = CLng(trim(request("sslpop3port")))
end if

if IsNumeric(trim(request("sslsmtpport"))) = true then
	ei.sslsmtpport = CLng(trim(request("sslsmtpport")))
end if

if trim(request("useSSL_SMTP")) <> "" then
	ei.useSSL_SMTP = true
else
	ei.useSSL_SMTP = false
end if

if IsNumeric(trim(request("SSL_Mode"))) = true then
	ei.SSL_Mode = CLng(trim(request("SSL_Mode")))
end if



if trim(request("use_IMAP4")) <> "" then
	ei.use_IMAP4 = true
else
	ei.use_IMAP4 = false
end if

if IsNumeric(trim(request("imap4port"))) = true then
	ei.imap4port = CLng(trim(request("imap4port")))
end if

if trim(request("useSSL_IMAP4")) <> "" then
	ei.useSSL_IMAP4 = true
else
	ei.useSSL_IMAP4 = false
end if

if IsNumeric(trim(request("sslimap4port"))) = true then
	ei.sslimap4port = CLng(trim(request("sslimap4port")))
end if


if trim(request("enableKillAttacker_With_IMAP4")) <> "" then
	ei.enableKillAttacker_With_IMAP4 = true
else
	ei.enableKillAttacker_With_IMAP4 = false
end if


if trim(request("enableSystemFilter")) <> "" then
	ei.enableSystemFilter = true
else
	ei.enableSystemFilter = false
end if

if trim(request("enableHandPoint2")) <> "" then
	ei.enableHandPoint2 = true
else
	ei.enableHandPoint2 = false
end if

ei.HELO_STRING = trim(request("HELO_STRING"))

if IsNumeric(trim(request("TempRegInfoKeepDays"))) = true then
	ei.TempRegInfoKeepDays = CLng(trim(request("TempRegInfoKeepDays")))
end if

if trim(request("enableSpeedScanVirus")) <> "" then
	ei.enableSpeedScanVirus = true
else
	ei.enableSpeedScanVirus = false
end if


if IsNumeric(trim(request("Web_Max_Recipients"))) = true then
	ei.Web_Max_Recipients = CLng(trim(request("Web_Max_Recipients")))
end if


ei.System_Mail_CharSet = trim(request("System_Mail_CharSet"))


if trim(request("EnableSystemCollectionMail")) <> "" then
	ei.EnableSystemCollectionMail = true
else
	ei.EnableSystemCollectionMail = false
end if

if trim(request("enableAutoStopUserOutGoing")) <> "" then
	ei.enableAutoStopUserOutGoing = true
else
	ei.enableAutoStopUserOutGoing = false
end if

if IsNumeric(trim(request("autoStopUserOutGoingMaxNumber"))) = true then
	ei.autoStopUserOutGoingMaxNumber = CLng(trim(request("autoStopUserOutGoingMaxNumber")))
end if

if IsNumeric(trim(request("autoStopUserOutGoingExpiresMinute"))) = true then
	ei.autoStopUserOutGoingExpiresMinute = CLng(trim(request("autoStopUserOutGoingExpiresMinute")))
end if

if trim(request("enableUserExpAffirm")) <> "" then
	ei.enableUserExpAffirm = true
else
	ei.enableUserExpAffirm = false
end if

if IsNumeric(trim(request("daysUserExpAffirm"))) = true then
	ei.daysUserExpAffirm = CLng(trim(request("daysUserExpAffirm")))
end if

if trim(request("catchAllNeedBack")) <> "" then
	ei.catchAllNeedBack = true
else
	ei.catchAllNeedBack = false
end if

if trim(request("EnableIngoingAuth")) <> "" then
	ei.EnableIngoingAuth = true
else
	ei.EnableIngoingAuth = false
end if

if trim(request("EnableErrBackToOut")) <> "" then
	ei.EnableErrBackToOut = true
else
	ei.EnableErrBackToOut = false
end if

if trim(request("EnableErrBackToOutForLocalNoSuchUser")) <> "" then
	ei.EnableErrBackToOutForLocalNoSuchUser = true
else
	ei.EnableErrBackToOutForLocalNoSuchUser = false
end if

if trim(request("EnableRelayServerSend")) <> "" then
	ei.EnableRelayServerSend = true
else
	ei.EnableRelayServerSend = false
end if

if trim(request("DNS_can_by_ROOT")) <> "" then
	ei.DNS_can_by_ROOT = true
else
	ei.DNS_can_by_ROOT = false
end if

if trim(request("EnableLocalNetwork")) <> "" then
	ei.EnableLocalNetwork = true
else
	ei.EnableLocalNetwork = false
end if

if trim(request("EnableKillHeloDomain")) <> "" then
	ei.EnableKillHeloDomain = true
else
	ei.EnableKillHeloDomain = false
end if

if IsNumeric(trim(request("DNSExpiresDays"))) = true then
	ei.DNSExpiresDays = CLng(trim(request("DNSExpiresDays")))
end if

if trim(request("EnableSmtpDomainCheck")) <> "" then
	ei.EnableSmtpDomainCheck = true
else
	ei.EnableSmtpDomainCheck = false
end if

if trim(request("EnableCheckMailFromDomainIsGood")) <> "" then
	ei.EnableCheckMailFromDomainIsGood = true
else
	ei.EnableCheckMailFromDomainIsGood = false
end if

if trim(request("EnableCheckMailFromIP")) <> "" then
	ei.EnableCheckMailFromIP = true
else
	ei.EnableCheckMailFromIP = false
end if

if trim(request("EnableCheckHeloIP")) <> "" then
	ei.EnableCheckHeloIP = true
else
	ei.EnableCheckHeloIP = false
end if

if trim(request("EnableCheckMailFromIP_WhenCheckHeloIP_Error")) <> "" then
	ei.EnableCheckMailFromIP_WhenCheckHeloIP_Error = true
else
	ei.EnableCheckMailFromIP_WhenCheckHeloIP_Error = false
end if

if IsNumeric(trim(request("CheckIPClass"))) = true then
	ei.CheckIPClass = CLng(trim(request("CheckIPClass")))
end if

if IsNumeric(trim(request("EnableSmtpCheckError2Trash"))) = true then
	if trim(request("EnableSmtpCheckError2Trash")) = "1" then
		ei.EnableSmtpCheckError2Trash = false
	else
		ei.EnableSmtpCheckError2Trash = true
	end if
end if

if trim(request("EnableAlwaryCanAccessIP")) <> "" then
	ei.EnableAlwaryCanAccessIP = true
else
	ei.EnableAlwaryCanAccessIP = false
end if

if trim(request("EnableTrustUser")) <> "" then
	ei.EnableTrustUser = true
else
	ei.EnableTrustUser = false
end if

if IsNumeric(trim(request("Delivery_Retry_Interval"))) = true then
	ei.Delivery_Retry_Interval = CLng(trim(request("Delivery_Retry_Interval")))
end if

if IsNumeric(trim(request("DeliveryTimeout"))) = true then
	ei.DeliveryTimeout = CLng(trim(request("DeliveryTimeout")))
end if

if trim(request("EnableMailFromNULL")) <> "" then
	ei.EnableMailFromNULL = true
else
	ei.EnableMailFromNULL = false
end if

if trim(request("DisableRelayEmail")) <> "" then
	ei.DisableRelayEmail = true
else
	ei.DisableRelayEmail = false
end if

if IsNumeric(trim(request("DisableRelayEmail_Mode"))) = true then
	ei.DisableRelayEmail_Mode = CLng(trim(request("DisableRelayEmail_Mode")))
end if

if trim(request("EnableGreylisting")) <> "" then
	ei.EnableGreylisting = true
else
	ei.EnableGreylisting = false
end if

if trim(request("EnableAuthTrustIP")) <> "" then
	ei.EnableAuthTrustIP = true
else
	ei.EnableAuthTrustIP = false
end if

if trim(request("EnableMailFromTrustEmail")) <> "" then
	ei.EnableMailFromTrustEmail = true
else
	ei.EnableMailFromTrustEmail = false
end if

if trim(request("EnableCheckMailHeader")) <> "" then
	ei.EnableCheckMailHeader = true
else
	ei.EnableCheckMailHeader = false
end if


if trim(request("EnableResponse_SMTP_NoUser")) <> "" then
	ei.EnableResponse_SMTP_NoUser = true
else
	ei.EnableResponse_SMTP_NoUser = false
end if

if trim(request("EnableReceiveOutMail")) <> "" then
	ei.EnableReceiveOutMail = true
else
	ei.EnableReceiveOutMail = false
end if

if trim(request("EnableAutoReply")) <> "" then
	ei.EnableAutoReply = true
else
	ei.EnableAutoReply = false
end if

if trim(request("Enable_MailRecall")) <> "" then
	ei.Enable_MailRecall = true
else
	ei.Enable_MailRecall = false
end if

if trim(request("Enable_System_TrashMsg")) <> "" then
	ei.Enable_System_TrashMsg = true
else
	ei.Enable_System_TrashMsg = false
end if

Application("em_Enable_MailRecall") = ei.Enable_MailRecall

if IsNumeric(trim(request("MailRecall_SaveDays"))) = true then
	ei.MailRecall_SaveDays = CLng(trim(request("MailRecall_SaveDays")))
end if

if trim(request("Enable_SendOutMonitor")) <> "" then
	ei.Enable_SendOutMonitor = true
else
	ei.Enable_SendOutMonitor = false
end if

if IsNumeric(trim(request("SendOutMonitor_SaveDays"))) = true then
	ei.SendOutMonitor_SaveDays = CLng(trim(request("SendOutMonitor_SaveDays")))
end if

if trim(request("Enable_SendOut_Auto_Monitor")) <> "" then
	ei.Enable_SendOut_Auto_Monitor = true
else
	ei.Enable_SendOut_Auto_Monitor = false
end if

if IsNumeric(trim(request("Auto_Monitor_Start_Max_SendNum"))) = true then
	ei.Auto_Monitor_Start_Max_SendNum = CLng(trim(request("Auto_Monitor_Start_Max_SendNum")))
end if

if IsNumeric(trim(request("Auto_Monitor_Keep_Days"))) = true then
	ei.Auto_Monitor_Keep_Days = CLng(trim(request("Auto_Monitor_Keep_Days")))
end if

if trim(request("ZATT_Is_Enable")) <> "" then
	ei.ZATT_Is_Enable = true
else
	ei.ZATT_Is_Enable = false
end if

ei.ZATT_URL = trim(request("ZATT_URL"))

if IsNumeric(trim(request("ZATT_Validity_Days"))) = true then
	ei.ZATT_Validity_Days = CLng(trim(request("ZATT_Validity_Days")))
end if

if trim(request("EnableVerification")) <> "" then
	ei.EnableVerification = true
else
	ei.EnableVerification = false
end if
Application("em_EnableVerification") = ei.EnableVerification

if trim(request("Enable_ErrBackToTrustMail")) <> "" then
	ei.Enable_ErrBackToTrustMail = true
else
	ei.Enable_ErrBackToTrustMail = false
end if

if trim(request("Enable_OnlyAutoForwardToLocal")) <> "" then
	ei.Enable_OnlyAutoForwardToLocal = true
else
	ei.Enable_OnlyAutoForwardToLocal = false
end if

if IsNumeric(trim(request("NotifyModeForVirus"))) = true then
	ei.NotifyModeForVirus = CLng(trim(request("NotifyModeForVirus")))
end if

if trim(request("Enable_AttackerSaveLog")) <> "" then
	ei.Enable_AttackerSaveLog = true
else
	ei.Enable_AttackerSaveLog = false
end if

if trim(request("Enable_KillSaveLog")) <> "" then
	ei.Enable_KillSaveLog = true
else
	ei.Enable_KillSaveLog = false
end if

if IsNumeric(trim(request("Collection_SysError_Mail_Mode"))) = true then
	ei.Collection_SysError_Mail_Mode = CLng(trim(request("Collection_SysError_Mail_Mode")))
end if

ei.Collection_SysError_Mail_To = trim(request("Collection_SysError_Mail_To"))

if trim(request("EnableArchive")) <> "" then
	wms.EnableArchive = true
else
	wms.EnableArchive = false
end if

if IsNumeric(trim(request("Archive_MaxNumber"))) = true then
	wms.Archive_MaxNumber = CLng(trim(request("Archive_MaxNumber")))
end if

ei.Save
wms.Save
set ei = nothing
set wms = nothing

Response.Redirect "ok.asp?" & getGRSN() & "&gourl=showsysinfo.asp"
%>
