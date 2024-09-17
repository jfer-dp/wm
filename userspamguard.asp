<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" -->

<%
isamg = false
showmail = ""

amg = trim(request("amg"))
gourl = trim(request("gourl"))
userid = trim(request("id"))
name = trim(request("name"))
domain = trim(request("domain"))

if amg = "1" and Len(gourl) > 0 and Len(userid) > 0 and Len(name) > 0 and Len(domain) > 0 then
	isamg = true
end if

if isamg = true then
	if isadmin() = false and isAccountsAdmin() = false then
		dim ed
		set ed = server.createobject("easymail.domain")
		ed.Load

		if ed.GetUserManagerDomainCount(Session("wem")) < 1 then
			set ed = nothing
			Response.Redirect "noadmin.asp"
		end if

		i = 0
		allnum = ed.GetUserManagerDomainCount(Session("wem"))

		dim isok
		isok = false

		do while i < allnum
			cdomainstr = ed.GetUserManagerDomain(Session("wem"), i)

			if LCase(cdomainstr) = LCase(domain) then
				isok = true
			end if

			cdomainstr = NULL

			i = i + 1
		loop

		set ed = nothing

		if isok = false then
			Response.Redirect "noadmin.asp"
		end if
	end if


	sp = InStr(1, name, "@")
	if sp > 0 then
		showmail = Mid(name, 1, sp) & domain
	else
		showmail = name & "@" & domain
	end if
end if


dim ei
set ei = server.createobject("easymail.UserSpamGuard")

dim mam
set mam = server.createobject("easymail.AdminManager")
mam.Load

dim userweb
set userweb = server.createobject("easymail.UserWeb")
if isamg = false then
	userweb.Load Session("wem")
else
	userweb.Load name
end if

dim tm
set tm = server.createobject("easymail.TrashMsg")
if isamg = false then
	tm.Load Session("wem")
else
	tm.Load name
end if


if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if isamg = false then
		ei.Load Session("wem")
	else
		ei.Load name
	end if

	if trim(request("setdefault")) = "true" then
		ei.SetDefault
		ei.EnableReceiveAllMail = false
		ei.EnableTrashMsg = true
		ei.EnableSizeLimitNoSpam = false
		ei.NoSpamSizeLimit = 200
		ei.Save

		userweb.EnableClearWhenFull = false
		userweb.EnableClearSendBox = false
		userweb.useAutoClearTrashBox = false
		userweb.autoClearTrashBoxDays = 15
		userweb.Save

		tm.IntervalMin = "1440"
		tm.Save()

		set ei = nothing
		set mam = nothing
		set userweb = nothing
		set tm = nothing

		if isamg = false then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=userspamguard.asp"
		else
			Response.Redirect "changepw.asp?" & getGRSN() & "&fo=1&id=" & userid & "&gourl=" & Server.URLEncode(gourl)
		end if
	end if


	if trim(request("EnableClearWhenFull")) = "" then
		userweb.EnableClearWhenFull = false
	else
		userweb.EnableClearWhenFull = true
	end if

	if trim(request("EnableClearSendBox")) = "" then
		userweb.EnableClearSendBox = false
	else
		userweb.EnableClearSendBox = true
	end if

	if trim(request("enableAutoClear")) = "" then
		userweb.useAutoClearTrashBox = false
	else
		userweb.useAutoClearTrashBox = true
	end if

	if trim(request("autoClearDays")) <> "" and IsNumeric(trim(request("autoClearDays"))) = true then
		userweb.autoClearTrashBoxDays = CInt(trim(request("autoClearDays")))
	else
		userweb.autoClearTrashBoxDays = 15
	end if

	userweb.Save


	if trim(request("EnableReceiveAllMail")) = "" then
		ei.EnableReceiveAllMail = false
	else
		ei.EnableReceiveAllMail = true
	end if

	if trim(request("EnableTrashMsg")) = "" then
		ei.EnableTrashMsg = false
	else
		ei.EnableTrashMsg = true
	end if

	if trim(request("EnableUser_Receive_MailToCc_MyEmail")) <> "" then
		ei.EnableUser_Receive_MailToCc_MyEmail = true
	else
		ei.EnableUser_Receive_MailToCc_MyEmail = false
	end if

	if trim(request("EnableUser_SpamGuard")) <> "" then
		ei.EnableUser_SpamGuard = true
	else
		ei.EnableUser_SpamGuard = false
	end if

	if trim(request("EnableUser_ReceiveLocal")) <> "" then
		ei.EnableUser_ReceiveLocal = true
	else
		ei.EnableUser_ReceiveLocal = false
	end if

	if trim(request("EnableUser_ReceiveAddressBook")) <> "" then
		ei.EnableUser_ReceiveAddressBook = true
	else
		ei.EnableUser_ReceiveAddressBook = false
	end if

	if trim(request("EnableUser_ReceiveFromOutEmails")) <> "" then
		ei.EnableUser_ReceiveFromOutEmails = true
	else
		ei.EnableUser_ReceiveFromOutEmails = false
	end if

	if trim(request("EnableUser_ReceiveDomain")) <> "" then
		ei.EnableUser_ReceiveDomain = true
	else
		ei.EnableUser_ReceiveDomain = false
	end if

	if mam.Enable_NoSpam_Affirm = true then
		if trim(request("Enable_User_NoSpam_Affirm")) <> "" then
			ei.Enable_User_NoSpam_Affirm = true
		else
			ei.Enable_User_NoSpam_Affirm = false
		end if
	end if

	if IsNumeric(trim(request("SpamProcessMode"))) = true then
		ei.SpamProcessMode = CLng(trim(request("SpamProcessMode")))
	end if

	if trim(request("EnableUser_OneDayMulRepeatReceive_Guard")) <> "" then
		ei.EnableUser_OneDayMulRepeatReceive_Guard = true
	else
		ei.EnableUser_OneDayMulRepeatReceive_Guard = false
	end if

	if IsNumeric(trim(request("OneDayMulRepeatReceive_Guard_ProcessMode"))) = true then
		ei.OneDayMulRepeatReceive_Guard_ProcessMode = CLng(trim(request("OneDayMulRepeatReceive_Guard_ProcessMode")))
	end if

	if trim(request("EnableSizeLimitNoSpam")) = "" then
		ei.EnableSizeLimitNoSpam = false
	else
		ei.EnableSizeLimitNoSpam = true
	end if

	if IsNumeric(trim(request("NoSpamSizeLimit"))) = true then
		ei.NoSpamSizeLimit = CLng(trim(request("NoSpamSizeLimit")))
	end if

	ei.Save


	if IsNumeric(trim(request("IntervalMin"))) = true then
		tm.IntervalMin = CLng(trim(request("IntervalMin")))
		tm.Save()
	end if


	set ei = nothing
	set mam = nothing
	set userweb = nothing
	set tm = nothing

	if isamg = false then
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=userspamguard.asp"
	else
		Response.Redirect "changepw.asp?" & getGRSN() & "&fo=1&id=" & userid & "&gourl=" & Server.URLEncode(gourl)
	end if
end if

if isamg = false then
	ei.LightLoad Session("wem")
else
	ei.LightLoad name
end if
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<STYLE type=text/css>
<!--
html {overflow: scroll; overflow-x: hidden; overflow-y: auto !important;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.st_l {height:24px; text-align:left; white-space:nowrap; border-left:1px solid #A5B6C8; border-top:1px solid #A5B6C8; border-right:1px solid #A5B6C8; padding-left:4px;}
.st_r {height:24px; text-align:right; white-space:nowrap; border-top:1px solid #A5B6C8; border-right:1px solid #A5B6C8; padding-right:4px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function SpamProcessMode_onchange() {
<%
if mam.Enable_NoSpam_Affirm = true then
%>
	if (document.f1.EnableUser_SpamGuard.checked == true)
	{
		if (document.f1.SpamProcessMode.value == "0")
			document.f1.Enable_User_NoSpam_Affirm.disabled = true;
		else
			document.f1.Enable_User_NoSpam_Affirm.disabled = false;
	}
	else
	{
		document.f1.Enable_User_NoSpam_Affirm.disabled = true;
	}
<%
end if
%>
}

function godef() {
	if (confirm("<%=b_lang_276 %>") == false)
		return ;

	document.f1.setdefault.value = "true"
	document.f1.submit();
}

function gosub(){
	if (document.f1.EnableReceiveAllMail.checked == true)
	{
		document.f1.EnableReceiveAllMail.checked = false;
		modify_receiveall();
		document.f1.EnableReceiveAllMail.checked = true;
	}

	document.f1.EnableUser_ReceiveLocal.disabled = false;
	document.f1.EnableUser_ReceiveAddressBook.disabled = false;
	document.f1.EnableUser_ReceiveDomain.disabled = false;
<%
if mam.Enable_NoSpam_Affirm = true then
%>
	document.f1.Enable_User_NoSpam_Affirm.disabled = false;
<%
end if
%>
	document.f1.SpamProcessMode.disabled = false;
	document.f1.OneDayMulRepeatReceive_Guard_ProcessMode.disabled = false;

	document.f1.submit();
}

function modifyit(){
	if (document.f1.EnableUser_SpamGuard.checked == true)
	{
		document.f1.EnableUser_ReceiveLocal.disabled = false;
		document.f1.EnableUser_ReceiveAddressBook.disabled = false;
		document.f1.EnableUser_ReceiveFromOutEmails.disabled = false;
		document.f1.EnableUser_ReceiveDomain.disabled = false;
	}
	else
	{
		document.f1.EnableUser_ReceiveLocal.disabled = true;
		document.f1.EnableUser_ReceiveAddressBook.disabled = true;
		document.f1.EnableUser_ReceiveFromOutEmails.disabled = true;
		document.f1.EnableUser_ReceiveDomain.disabled = true;
	}

<%
if mam.Enable_NoSpam_Affirm = true then
%>
	if (document.f1.Enable_User_NoSpam_Affirm.disabled == true)
	{
		if (document.f1.EnableUser_SpamGuard.checked == true)
			if (document.f1.SpamProcessMode.value == "1")
				document.f1.Enable_User_NoSpam_Affirm.disabled = false;
	}
	else
	{
		if (document.f1.EnableUser_SpamGuard.checked == false)
			document.f1.Enable_User_NoSpam_Affirm.disabled = true;
	}
<%
end if
%>
}

function level_onchange()
{
	if (document.f1.levelsel.value == "1")
	{
		document.f1.SpamProcessMode.value = "1";

		document.f1.EnableUser_SpamGuard.checked = true;

		modifyit();

		document.f1.EnableUser_ReceiveLocal.checked = true;
		document.f1.EnableUser_ReceiveAddressBook.checked = true;
		document.f1.EnableUser_ReceiveFromOutEmails.checked = true;
		document.f1.EnableUser_ReceiveDomain.checked = true;
<%
if mam.Enable_NoSpam_Affirm = true then
%>
		document.f1.Enable_User_NoSpam_Affirm.checked = true;
<%
end if
%>

		document.f1.EnableUser_Receive_MailToCc_MyEmail.checked = false;

		document.f1.enableAutoClear.checked = true;
		document.f1.EnableClearWhenFull.checked = false;
		document.f1.EnableClearSendBox.checked = false;

		document.f1.EnableUser_OneDayMulRepeatReceive_Guard.checked = false;
		document.f1.OneDayMulRepeatReceive_Guard_ProcessMode.disabled = true;
	}
	else if (document.f1.levelsel.value == "2")
	{
		document.f1.SpamProcessMode.value = "1";

		document.f1.EnableUser_SpamGuard.checked = true;

		modifyit();

		document.f1.EnableUser_ReceiveLocal.checked = true;
		document.f1.EnableUser_ReceiveAddressBook.checked = true;
		document.f1.EnableUser_ReceiveFromOutEmails.checked = true;
		document.f1.EnableUser_ReceiveDomain.checked = false;
<%
if mam.Enable_NoSpam_Affirm = true then
%>
		document.f1.Enable_User_NoSpam_Affirm.checked = true;
<%
end if
%>

		document.f1.EnableUser_Receive_MailToCc_MyEmail.checked = false;

		document.f1.enableAutoClear.checked = true;
		document.f1.EnableClearWhenFull.checked = false;
		document.f1.EnableClearSendBox.checked = false;

		document.f1.EnableUser_OneDayMulRepeatReceive_Guard.checked = true;
		document.f1.OneDayMulRepeatReceive_Guard_ProcessMode.disabled = false;
	}
	else if (document.f1.levelsel.value == "3")
	{
		document.f1.SpamProcessMode.value = "1";

		document.f1.EnableUser_SpamGuard.checked = true;

		modifyit();

		document.f1.EnableUser_ReceiveLocal.checked = true;
		document.f1.EnableUser_ReceiveAddressBook.checked = true;
		document.f1.EnableUser_ReceiveFromOutEmails.checked = false;
		document.f1.EnableUser_ReceiveDomain.checked = false;
<%
if mam.Enable_NoSpam_Affirm = true then
%>
		document.f1.Enable_User_NoSpam_Affirm.checked = true;
<%
end if
%>

		document.f1.EnableUser_Receive_MailToCc_MyEmail.checked = true;

		document.f1.enableAutoClear.checked = true;
		document.f1.EnableClearWhenFull.checked = false;
		document.f1.EnableClearSendBox.checked = false;

		document.f1.EnableUser_OneDayMulRepeatReceive_Guard.checked = true;
		document.f1.OneDayMulRepeatReceive_Guard_ProcessMode.disabled = false;
	}
}

function modifyMulRepeatReceive()
{
	if (document.f1.EnableUser_OneDayMulRepeatReceive_Guard.checked == true)
		document.f1.OneDayMulRepeatReceive_Guard_ProcessMode.disabled = false;
	else
		document.f1.OneDayMulRepeatReceive_Guard_ProcessMode.disabled = true;
}

function window_onload() {
	document.f1.IntervalMin.value = "<%=tm.IntervalMin %>";
	document.f1.NoSpamSizeLimit.value = "<%=ei.NoSpamSizeLimit %>";

	if (document.f1.EnableReceiveAllMail.checked == true)
		modify_receiveall();

	modifyEnableTrashMsg();
	modifyEnableSizeLimitNoSpam();
	modifyEnableClearWhenFull();
}

function goback()
{
<%
if isamg = false then
%>
	location.href = "user_right.asp?<%=getGRSN() %>";
<%
else
%>
	location.href = "changepw.asp?<%=getGRSN() %>&fo=1&id=<%=userid %>&gourl=<%=Server.URLEncode(gourl) %>";
<%
end if
%>
}

function modify_receiveall()
{
	if (document.f1.EnableReceiveAllMail.checked == true)
	{
		document.f1.SpamProcessMode.disabled = true;
		document.f1.levelsel.disabled = true;
		document.f1.EnableUser_SpamGuard.disabled = true;
<%
if mam.Enable_NoSpam_Affirm = true then
%>
		document.f1.Enable_User_NoSpam_Affirm.disabled = true;
<%
end if
%>
		document.f1.EnableUser_Receive_MailToCc_MyEmail.disabled = true;
		document.f1.EnableUser_OneDayMulRepeatReceive_Guard.disabled = true;
		document.f1.OneDayMulRepeatReceive_Guard_ProcessMode.disabled = true;

		document.f1.EnableUser_ReceiveLocal.disabled = true;
		document.f1.EnableUser_ReceiveAddressBook.disabled = true;
		document.f1.EnableUser_ReceiveFromOutEmails.disabled = true;
		document.f1.EnableUser_ReceiveDomain.disabled = true;
	}
	else
	{
		document.f1.SpamProcessMode.disabled = false;
		document.f1.levelsel.disabled = false;
		document.f1.EnableUser_SpamGuard.disabled = false;
		document.f1.EnableUser_Receive_MailToCc_MyEmail.disabled = false;
		document.f1.EnableUser_OneDayMulRepeatReceive_Guard.disabled = false;

		modifyit();
		modifyMulRepeatReceive();
	}
}

function modifyEnableTrashMsg()
{
	if (document.f1.EnableTrashMsg.checked == true)
		document.f1.IntervalMin.disabled = false;
	else
		document.f1.IntervalMin.disabled = true;
}

function modifyEnableSizeLimitNoSpam()
{
	if (document.f1.EnableSizeLimitNoSpam.checked == true)
		document.f1.NoSpamSizeLimit.disabled = false;
	else
		document.f1.NoSpamSizeLimit.disabled = true;
}

function modifyEnableClearWhenFull() {
	if (document.f1.EnableClearWhenFull.checked == true)
		document.f1.EnableClearSendBox.disabled = false;
	else
		document.f1.EnableClearSendBox.disabled = true;
}
//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload()">
<FORM ACTION="userspamguard.asp" METHOD="POST" NAME="f1">
<input name="setdefault" type="hidden">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_277 %>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white" style="border-bottom:1px solid #A5B6C8;">
	<tr>
	<td class="st_l">
	<input type="checkbox" name="EnableReceiveAllMail" <%
if ei.EnableReceiveAllMail = true then
	Response.Write "checked"
end if
%> onclick="javascript:modify_receiveall();">&nbsp;<%=b_lang_278 %>
	</td>
	</tr>
</table>
<br>

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white" style="border-bottom:1px solid #A5B6C8;">
	<tr>
	<td width="50%" class="st_l" style="padding-left:8px;">
	<%=b_lang_279 %><%=s_lang_mh %>
	<select name="SpamProcessMode" class=drpdwn onchange="javascript:SpamProcessMode_onchange();">
<%
if ei.SpamProcessMode = 0 then
%>
<option value="1"><%=b_lang_280 %></option>
<option value="0" selected><%=b_lang_281 %></option>
<%
else
%>
<option value="1" selected><%=b_lang_280 %></option>
<option value="0"><%=b_lang_281 %></option>
<%
end if
%>
</select>
	</td>
	<td class="st_r">
<select name="levelsel" class=drpdwn onchange="javascript:level_onchange();">
<option value="0">---<%=b_lang_282 %>---</option>
<option value="1"><%=b_lang_283 %></option>
<option value="2"><%=b_lang_284 %></option>
<option value="3"><%=b_lang_285 %></option>
</select>
	</td>
	</tr>
</table>
<br>

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white" style="border-bottom:1px solid #A5B6C8;">
	<tr>
	<td class="st_l">
	<input type="checkbox" name="EnableUser_SpamGuard" <%
if ei.EnableUser_SpamGuard = true then
	response.write "checked"
end if
%> onclick="javascript:modifyit();">
	<%=b_lang_286 %></td>
	</tr>

	<tr>
	<td class="st_l">
	&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="EnableUser_ReceiveLocal" <%
if ei.EnableUser_ReceiveLocal = true then
	response.write "checked"
end if

if ei.EnableUser_SpamGuard = false then
	response.write " disabled"
end if
%>>
	<%=b_lang_287 %></td>
	</tr>

	<tr>
	<td class="st_l">
	&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="EnableUser_ReceiveAddressBook" <%
if ei.EnableUser_ReceiveAddressBook = true then
	response.write "checked"
end if

if ei.EnableUser_SpamGuard = false then
	response.write " disabled"
end if
%>>
	<%=b_lang_288 %></td>
	</tr>

	<tr>
	<td class="st_l">
	&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="EnableUser_ReceiveFromOutEmails" <%
if ei.EnableUser_ReceiveFromOutEmails = true then
	response.write "checked"
end if

if ei.EnableUser_SpamGuard = false then
	response.write " disabled"
end if
%>>
	<%=b_lang_289 %></td>
	</tr>

	<tr>
	<td class="st_l">
	&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="EnableUser_ReceiveDomain" <%
if ei.EnableUser_ReceiveDomain = true then
	response.write "checked"
end if

if ei.EnableUser_SpamGuard = false then
	response.write " disabled"
end if
%>>
	<%=b_lang_290 %></td>
	</tr>
<%
if mam.Enable_NoSpam_Affirm = true then
%>
	<tr>
	<td class="st_l">
	&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="Enable_User_NoSpam_Affirm" <%
if ei.Enable_User_NoSpam_Affirm = true then
	response.write "checked"
end if

if ei.EnableUser_SpamGuard = false or ei.SpamProcessMode = 0 then
	response.write " disabled"
end if
%>>
	<%=b_lang_291 %></td>
	</tr>
<%
end if
%>
</table>
<br>

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white" style="border-bottom:1px solid #A5B6C8;">
	<tr>
	<td class="st_l">
	<input type="checkbox" name="EnableUser_Receive_MailToCc_MyEmail" <%
if ei.EnableUser_Receive_MailToCc_MyEmail = true then
	Response.Write "checked"
end if
%> onclick="javascript:modifyit();">
	<%=b_lang_292 %><font color="#901111"><%
if isamg = true and Len(showmail) > 0 then
	Response.Write showmail
else
	Response.Write Session("mail")
end if
%></font><%=b_lang_293 %></td>
	</tr>
</table>
<br>

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white" style="border-bottom:1px solid #A5B6C8;">
	<tr>
	<td width="40%" class="st_l">
	<input type="checkbox" name="EnableUser_OneDayMulRepeatReceive_Guard" <%
if ei.EnableUser_OneDayMulRepeatReceive_Guard = true then
	response.write "checked"
end if
%> onclick="javascript:modifyMulRepeatReceive();"> <%=b_lang_294 %></td>
	<td class="st_r"><%=b_lang_295 %><%=s_lang_mh %>
	<select name="OneDayMulRepeatReceive_Guard_ProcessMode" class=drpdwn<%
if ei.EnableUser_OneDayMulRepeatReceive_Guard = false then
	Response.Write " disabled"
end if
%>>
<%
if ei.OneDayMulRepeatReceive_Guard_ProcessMode = 0 then
%>
<option value="1"><%=b_lang_280 %></option>
<option value="0" selected><%=b_lang_281 %></option>
<%
else
%>
<option value="1" selected><%=b_lang_280 %></option>
<option value="0"><%=b_lang_281 %></option>
<%
end if
%>
</select>
	</td>
	</tr>
</table>
<br>

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white" style="border-bottom:1px solid #A5B6C8;">
	<tr>
	<td width="40%" class="st_l">
	<input type="checkbox" name="EnableTrashMsg" <%
if ei.EnableTrashMsg = true then
	response.write "checked"
end if
%> onclick="javascript:modifyEnableTrashMsg();"> <%=b_lang_296 %></td>
	<td class="st_r">
	<%=b_lang_297 %><%=s_lang_mh %>
		<select id="IntervalMin" name="IntervalMin" class="drpdwn">
		<option value="15">15 <%=b_lang_298 %></option>
		<option value="30">30 <%=b_lang_298 %></option>
		<option value="60">1 <%=b_lang_299 %></option>
		<option value="120">2 <%=b_lang_300 %></option>
		<option value="180">3 <%=b_lang_300 %></option>
		<option value="360">6 <%=b_lang_300 %></option>
		<option value="720">12 <%=b_lang_300 %></option>
		<option value="1440">1 <%=b_lang_301 %></option>
		<option value="2880">2 <%=b_lang_302 %></option>
		<option value="4320">3 <%=b_lang_302 %></option>
		<option value="5760">4 <%=b_lang_302 %></option>
		<option value="7200">5 <%=b_lang_302 %></option>
		<option value="8640">6 <%=b_lang_302 %></option>
		<option value="10080">7 <%=b_lang_302 %></option>
		</select>
	</td>
	</tr>
</table>
<br>

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white" style="border-bottom:1px solid #A5B6C8;">
	<tr>
	<td class="st_l">
	<input type="checkbox" name="EnableSizeLimitNoSpam" <%
if ei.EnableSizeLimitNoSpam = true then
	response.write "checked"
end if
%> onclick="javascript:modifyEnableSizeLimitNoSpam();"> <%=s_lang_0107 %>
	<select id="NoSpamSizeLimit" name="NoSpamSizeLimit" class="drpdwn">
		<option value="200">200K</option>
		<option value="300">300K</option>
		<option value="400">400K</option>
		<option value="500">500K</option>
		<option value="600">600K</option>
		<option value="700">700K</option>
		<option value="800">800K</option>
		<option value="900">900K</option>
		<option value="1000">1000K</option>
	</select>&nbsp;<%=s_lang_0108 %>
	</td>
	</tr>
</table>
<br>

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white" style="border-bottom:1px solid #A5B6C8;">
	<tr>
	<td width="16%" nowrap class="st_l" style="padding-left:8px;">
	<%=s_lang_0120 %><%=s_lang_mh %>
	</td>
	<td class="st_r" style="text-align:left">
	&nbsp;<input type="checkbox" name="EnableClearWhenFull" value="checkbox" <% if userweb.EnableClearWhenFull = true then response.write "checked"%> onclick="javascript:modifyEnableClearWhenFull();"><%=s_lang_0121 %><br>
	&nbsp;<input type="checkbox" name="EnableClearSendBox" value="checkbox" <% if userweb.EnableClearSendBox = true then response.write "checked"%>><%=s_lang_0122 %></td>
	</tr>
</table>
<br>

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white" style="border-bottom:1px solid #A5B6C8;">
	<tr>
	<td nowrap class="st_l">
	<input type="checkbox" name="enableAutoClear" value="checkbox" <% if userweb.useAutoClearTrashBox = true then response.write "checked"%>><%=b_lang_303 %>
	</td>
	</tr>

	<tr>
	<td class="st_l" style="padding-left:8px;">
	<%=b_lang_208 %><%=s_lang_mh %>
	<input type="text" name="autoClearDays" class='n_textbox' value="<%=userweb.autoClearTrashBoxDays %>" size="4" maxlength="4">&nbsp;<%=b_lang_230 %>
	</td>
    </tr>
</table>

</td></tr>
<tr><td colspan="2" align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px;">
<a class='wwm_btnDownload btn_blue' href="javascript:goback();"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:godef();"><%=b_lang_304 %></a>
</td></tr>
</table>

	<input type="hidden" name="gourl" value="<%=gourl %>">
	<input type="hidden" name="id" value="<%=userid %>">
	<input type="hidden" name="name" value="<%=name %>">
	<input type="hidden" name="domain" value="<%=domain %>">
	<input type="hidden" name="amg" value="<%=amg %>">
</FORM>

<table width="92%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px <%=MY_COLOR_1 %> solid; margin-top:50px; padding-bottom:10px;'>
	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:4px; color:#444444;">
	在启用"接收已发送至系统外邮件地址发来的邮件"功能后, 如果您写信(或回复, 转发)给 user@yahoo.com 后, 那么系统将会信任来自 user@yahoo.com 的邮件, 并将其直接放置到收件箱中.<br><br>
	</td>
	</tr>

	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:4px; color:#444444;">
	在启用"接收已发送至系统外域名发来的邮件"功能后, 如果您写信(或回复, 转发)给 user@yahoo.com 后, 那么系统将会信任来自 yahoo.com 域名的邮件, 并将所有来自 yahoo.com 的邮件直接放置到收件箱中.<br><br>
	</td>
	</tr>

<%
if mam.Enable_NoSpam_Affirm = true then
%>
	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:4px; color:#444444;">
	在启用"<font color="#901111">非垃圾邮件发送方确认功能</font>"后, 可以帮助您清除掉大量的垃圾邮件, 其工作方式是这样的:<br>
&nbsp;&nbsp;1. 有一个陌生地址给我的信箱写信时, 会先进入我的垃圾箱. 邮件系统会自动回一封信, 要求发信方回答一道随机生成的数学题.<br>
&nbsp;&nbsp;2. 收信方要在主题里填写这道数学题目的正确答案, 然后回复.<br>
&nbsp;&nbsp;3. 在收到正确的数学题答案后, 系统会将原来的邮件转入到我的收件箱里.<br><br>
原理是: 垃圾信通常是大量发送的邮件, 几乎不存在每一封邮件回答一道题的可能性.<br><br>
问: 正常的邮件, 发信人每次发信都要回答数学题不是太麻烦了吗?<br>
答: 不是的, 因为只要我回了一封信, 或是把他的邮址加入到我的地址薄里, 或是干脆先给对方发一封邮件的话, 那么这个地址发来的信就不再需要回答数学题而直接进入我的收件箱中了.<br><br>
	</td>
	</tr>
<%
end if
%>

	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:4px; color:#444444;">
	在启用"收件人地址判断功能"后, 您将会无法接收到暗送(Bcc)给您的电子邮件, 也有可能无法接收自动转发来的邮件.<br><br>
	</td>
	</tr>

	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:4px; color:#444444;">
	为避免将一些有用的邮件误判为垃圾邮件, 请使用<a href="cgfilter.asp?<%=getGRSN() %>">邮件分检助理</a>功能.<br><br>
	</td>
	</tr>

	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:4px; color:#444444;">
	在启用相同内容邮件过滤功能后, 系统会对24小时内接收到您收件箱内的邮件进行分析, 并根据您的设置, 除保留首封邮件外, 将其他重复邮件放入垃圾箱内或直接删除.<br>
	</td>
	</tr>
</table>
</BODY>
</HTML>

<%
set ei = nothing
set mam = nothing
set userweb = nothing
set tm = nothing
%>
