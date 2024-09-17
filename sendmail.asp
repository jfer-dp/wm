<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" --> 

<%
dim mailsend
set mailsend = server.createobject("easymail.MailSend")

dim sendisok
sendisok = false

dim createads
createads = false

dim ads_i
ads_i = 0

dim statestr
statestr = b_lang_075

dim decTextArea
decTextArea = ""

sname = trim(request("sname"))
sfname = trim(request("sfname"))
YH_sendmode = trim(request("YH_sendmode"))
YH_filename = trim(request("YH_filename"))
YH_backurl = trim(request("YH_backurl"))
YH_prev = trim(request("YH_prev"))
YH_next = trim(request("YH_next"))
prev_gourl = replace(YH_backurl, YH_filename, YH_prev)
next_gourl = replace(YH_backurl, YH_filename, YH_next)

if sname <> "" and sfname <> "" then
	openresult = mailsend.OpenFriendFolder(Session("wem"), sname, sfname, false)

	if openresult = -1 then
		set mailsend = nothing
		Response.Redirect "err.asp?errstr=" & b_lang_076
	elseif  openresult = 1 then
		set mailsend = nothing
		Response.Redirect "err.asp?errstr=" & b_lang_077
	elseif  openresult = 2 then
		set mailsend = nothing
		Response.Redirect "err.asp?errstr=" & b_lang_078
	end if
end if

mailsend.createnew Session("wem"), Session("tid")

if trim(request("EasyMail_CertServer")) = "1" then
	if Session("scpw") = "" then
		mailsend.SecCertPassword = trim(request("EasyMail_CertPW"))

		if trim(request("EasyMail_CertPW")) <> "" then
			Session("scpw") = trim(request("EasyMail_CertPW"))
		end if
	else
		mailsend.SecCertPassword = Session("scpw")
	end if

	mailsend.needSignAndEncrypt = true
elseif trim(request("EasyMail_CertServer")) = "2" then
	if Session("scpw") = "" then
		mailsend.SecCertPassword = trim(request("EasyMail_CertPW"))

		if trim(request("EasyMail_CertPW")) <> "" then
			Session("scpw") = trim(request("EasyMail_CertPW"))
		end if
	else
		mailsend.SecCertPassword = Session("scpw")
	end if

	mailsend.needSign = true
end if

mailsend.CharSet = trim(request("EasyMail_CharSet"))

mailsend.MailName = trim(request("MailName"))
mailsend.EM_BackAddress = trim(request("EasyMail_BackAddress"))
mailsend.EM_Bcc = trim(request("EasyMail_Bcc"))
mailsend.EM_Cc = trim(request("EasyMail_Cc"))
mailsend.EM_OrMailName = trim(request("EasyMail_OrMailName"))
mailsend.EM_Priority = trim(request("EasyMail_Priority"))
mailsend.EM_Zatt = trim(request("zAttFileString"))

if trim(request("EasyMail_ReadBack")) <> "" then
	mailsend.EM_ReadBack = true
else
	mailsend.EM_ReadBack = false
end if


mailsend.needAddInDebarList = false

if trim(request("needAddInDebarList")) <> "" then
	if isadmin() = true then
		mailsend.needAddInDebarList = true
	else
		dim dmadmin
		set dmadmin = server.createobject("easymail.Domain")
		dmadmin.Load

		if dmadmin.GetUserManagerDomainCount(Session("wem")) > 0 then
			mailsend.needAddInDebarList = true
		end if

		set dmadmin = nothing
	end if
end if

if IsNumeric(trim(request("EasyMail_SignNo"))) = true then
	mailsend.EM_SignNo = CLng(trim(request("EasyMail_SignNo")))
end if

mailsend.EM_Subject = trim(request("EasyMail_Subject"))

if trim(request("useRichEditer")) = "true" then
	i = 10
	do while i < 20
		mailsend.EM_Text = Request.Form("Mdec_RichEdit_Text" & i)
		i = i + 1
	loop

	i = 20
	do while i < 30
		mailsend.EM_HTML_Text = Request.Form("Mdec_RichEdit_Html" & i)
		i = i + 1
	loop

	mailsend.useRichEditer = true
else
	i = 0
	do while i < 10
		mailsend.EM_Text = Request.Form("Mdec_EasyMail_Text" & i)
		i = i + 1
	loop

	mailsend.useRichEditer = false
end if


mailsend.EM_TimerSend = trim(request("EasyMail_TimerSend"))
mailsend.EM_To = trim(request("EasyMail_To"))
mailsend.ForwardAttString = trim(request("EasyMail_OrAtt"))

mailsend.AddFromAttFileString = trim(request("AddFromAttFileString"))


if trim(request("EasyMail_SystemMail")) <> "" and isadmin() = true then
	mailsend.SystemMessage = true
else
	mailsend.SystemMessage = false
end if


if trim(request("EasyMail_SendBackup")) <> "" then
	mailsend.SendBackup = true
else
	mailsend.SendBackup = false
end if


if trim(request("SendMode")) = "send" then
	if mailsend.Send() = false then
		Set mailsend = nothing
		sendisok = false
	else
		if trim(request("EasyMail_OrMailName")) <> "" then
			mailsend.SetForward trim(request("EasyMail_R_F_MailName"))
		else
			mailsend.SetReply trim(request("EasyMail_R_F_MailName"))
		end if

		Set mailsend = nothing

		sendisok = true

		createads = true
	end if
elseif trim(request("SendMode")) = "save" then
	statestr = b_lang_079

	if mailsend.Save() = false then
		Set mailsend = nothing
		sendisok = false
	else
		if trim(request("EasyMail_OrMailName")) <> "" then
			mailsend.SetForward trim(request("EasyMail_R_F_MailName"))
		else
			mailsend.SetReply trim(request("EasyMail_R_F_MailName"))
		end if

		Set mailsend = nothing
		sendisok = true
	end if
elseif trim(request("SendMode")) = "timersend" then
	statestr = b_lang_080

	if mailsend.TimerSend() = false then
		Set mailsend = nothing
		sendisok = false
	else
		Set mailsend = nothing
		sendisok = true
	end if
elseif trim(request("SendMode")) = "post" then
	statestr = b_lang_081

	if IsNumeric(trim(request("face"))) = true then
		mailsend.face = CInt(trim(request("face")))
	end if

	canView = false

	if isadmin() = false then
		set pfvl = server.createobject("easymail.PubFolderViewLimit")
		pfvl.Load trim(request("iniid"))

		if pfvl.IsShow(Session("mail")) = true then
			canView = true
		end if

		set pfvl = nothing
	end if

	if isadmin() = true or canView = true then
		if mailsend.Post(trim(request("iniid")), trim(request("pid")), trim(request("searchkey"))) = false then
			Set mailsend = nothing
			sendisok = false
		else
			Set mailsend = nothing
			sendisok = true
		end if
	else
		Set mailsend = nothing
		sendisok = false
	end if
elseif trim(request("SendMode")) = "editpost" then
	statestr = b_lang_081

	if IsNumeric(trim(request("face"))) = true then
		mailsend.face = CInt(trim(request("face")))
	end if

	canView = false
	isok = false

	if isadmin() = false then
		set pfvl = server.createobject("easymail.PubFolderViewLimit")
		pfvl.Load trim(request("iniid"))

		if pfvl.IsShow(Session("mail")) = true then
			canView = true
		end if

		set pfvl = nothing
	end if

	if canView = true then
		dim pf
		set pf = server.createobject("easymail.PubFolderManager")
		pf.load trim(request("iniid"))

		permission = pf.Permission

		if LCase(pf.admin) <> LCase(Session("wem")) then
			if (permission = 0 or permission = 1) and Session("mail") = pf.GetPostName(trim(request("oname"))) then
				isok = true
			end if
		else
			isok = true
		end if

		set pf = nothing
	end if

	if isadmin() = true or isok = true then
		if mailsend.EditPost(trim(request("iniid")), trim(request("oname")), trim(request("searchkey"))) = false then
			Set mailsend = nothing
			sendisok = false
		else
			Set mailsend = nothing
			sendisok = true
		end if
	else
		Set mailsend = nothing
		sendisok = false
	end if
elseif trim(request("SendMode")) = "domainslistmail" then
	statestr = b_lang_075

	dim dm
	set dm = server.createobject("easymail.Domain")
	dm.Load

	if dm.GetUserManagerDomainCount(Session("wem")) < 1 then
		if isadmin() = false then
			Set mailsend = nothing
			set dm = nothing
			response.redirect "noadmin.asp"
		end if
	end if


	dim ss
	dim se
	ss = 2
	se = 2
	allnum = dm.GetUserManagerDomainCount(Session("wem"))
	allisok = true
	msg = trim(request("EasyMail_DomainList"))

	Do While 1
		se = InStr(ss, msg, Chr(9))

        If se <> 0 Then
			tempdomain = Mid(msg, ss, se - ss)

			isok = false
			i = 0
			do while i < allnum
				if tempdomain = dm.GetUserManagerDomain(Session("wem"), i) then
					isok = true
    		        exit do
				end if

				i = i + 1
			loop

			tempdomain = NULL

			if isok = false then
				allisok = false
   		        Exit Do
			end if
		Else
			Exit Do
		End If

		ss = se + 1
	Loop

	set dm = nothing


	if allisok = true or isadmin() = true then
		if mailsend.SendDomains(trim(request("EasyMail_DomainList"))) = false then
			Set mailsend = nothing
			sendisok = false
		else
			Set mailsend = nothing
			sendisok = true
		end if
	else
		Set mailsend = nothing
		response.redirect "noadmin.asp"
	end if
end if
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
</HEAD>

<BODY>
<form method=post name="form1" action="sendend.asp">
<%
if sendisok = true then
%>
<table width="90%" align="center" border="0" cellspacing="0" cellpadding="0" style="margin-top:40px;">
	<tr style="background:#EFF7FF; color:#104A7B;">
	<td colspan="2" height="30" style="border:1px solid #8CA5B5;">&nbsp;&nbsp;<%=b_lang_082 %><%=statestr %><%=b_lang_083 %></td>
	</tr>
<%
	if createads = true then
		dim ads
		set ads = server.createobject("easymail.Addresses")
		ads.Load Session("wem")

		dim allsendstr
		allsendstr = trim(request("EasyMail_To"))

		if trim(request("EasyMail_Cc")) <> "" then
			allsendstr = allsendstr & "," & trim(request("EasyMail_Cc"))
		end if

		if trim(request("EasyMail_Bcc")) <> "" then
			allsendstr = allsendstr & "," & trim(request("EasyMail_Bcc"))
		end if

		ads.SetSendMailToInfo allsendstr
		allnum = ads.SendMail_Addresses_NoInPrivateAddresses_Count

		ads_i = 0
		if allnum > 0 then
%>
	<tr>
	<td style="color:#444;">&nbsp;&nbsp;<%=b_lang_084 %></td>
	<td nowrap align="center" style="height:22px; width:140px; background:#104A7B; color:white; border:1px solid #104A7B;">
	<input type="checkbox" onclick="checkall(this);">&nbsp;<%=b_lang_085 %>
	</td>
	</tr>
<%
			do while ads_i < allnum
				nemail = ads.GetEmail_NoInPriAddresses_ByIndex(ads_i)
%>
	<tr><td align="right" style="height:20px; border-bottom:1px solid #8CA5B5; padding-right:12px;"><%=GetAdsName(nemail) %></td>
	<td align="center" style="background:#EFF7FF; border-left:1px solid #8CA5B5; border-right:1px solid #8CA5B5; border-bottom:1px solid #8CA5B5;">
	<input type="checkbox" name="newadd<%=ads_i %>" value="<%=nemail %>">
	</td></tr>
<%
				ads_i = ads_i + 1
			loop
		else
%>
	<tr><td colspan="2" style="border-bottom:1px solid #8CA5B5; height:24px;">&nbsp;</td></tr>
<%
		end if

		set ads = nothing
	end if
else
%>
<table width="90%" align="center" border="0" cellspacing="0" cellpadding="0" style="margin-top:40px;">
	<tr style="background:#EFF7FF; color:#104A7B;">
	<td colspan="2" height="30" style="border:1px solid #8CA5B5;">&nbsp;<img src="images/error.gif" align="absmiddle" border="0">&nbsp;<%=b_lang_086 %></td>
	</tr>
	<tr><td colspan="2" style="border-bottom:1px solid #8CA5B5; height:24px;">&nbsp;</td></tr>
<%
end if
%>
	<tr><td colspan="2" style="height:24px;">&nbsp;</td></tr>
	<tr><td align="right" colspan="2" style="padding-right:30px;">
	<a class="wwm_btnDownload btn_blue" style="width:40px;" href="javascript:goback();"><%=s_lang_ok %></a>
	</td></tr>
</table>
<%
if createads = true then
%>
	<input type="hidden" name="newaddnum" value="<%=allnum %>">
<%
end if
%>
<INPUT NAME="gourl" TYPE="hidden" Value="<%=trim(request("gourl")) %>">
</form>

<script type="text/javascript">
<!--
function goback() {
<%
if YH_backurl <> "" then
%>
	document.form1.gourl.value = "<%=YH_backurl %>";
<%
end if
%>
	document.form1.submit();
}

function checkall(tgobj) {
	var i = 0;
	var theObj;

	for(; i<<%=ads_i %>; i++)
	{
		theObj = eval("document.form1.newadd" + i);

		if (theObj != null)
			theObj.checked = tgobj.checked;
	}
}
// -->
</script>

</BODY>
</HTML>

<%
function GetAdsName(ostr)
	gan_s = InStr(ostr, "<")
	if gan_s > 0 then
		gan_e = InStr(gan_s, ostr, ">")

		if gan_e > 0 then
			GetAdsName = server.htmlencode(Mid(ostr, 1, gan_s - 1))
		else
			GetAdsName = ostr
		end if
	else
		GetAdsName = ostr
	end if
end function
%>
