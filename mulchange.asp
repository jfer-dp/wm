<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
if isadmin() = false and isAccountsAdmin() = false then
	Response.Redirect "noadmin.asp"
end if
%>

<%
gourl = trim(request("gourl"))

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim msg
	msg = trim(request("mulusers"))
	dim item
	dim ss
	dim se

	if trim(request("onlyCleanMailBox")) = "1" then
		dim mam
		set mam = server.createobject("easymail.AdminManager")
		mam.Load

		if Len(msg) > 0 then
			ss = 1
			se = 1

			Do While 1
				se = InStr(ss, msg, Chr(12))

				If se <> 0 Then
					item = Mid(msg, ss, se - ss)
					mam.CleanMailBox(item)
				Else
					Exit Do
				End If

				ss = se + 1
			Loop
		end if

		set mam = nothing
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
	end if

	Set em = Application("em")


if trim(request("onlycloseIPLimit")) = "1" then
	if Len(msg) > 0 then
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(12))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				em.SetEnableIPLimit item, false
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	set em = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


if trim(request("onlychangesize")) = "1" then
	uSize = trim(request("uSize"))

	if trim(request("changeSize")) <> "" and IsNumeric(uSize) then
		uSize = CLng(uSize)

		if Len(msg) > 0 then
			ss = 1
			se = 1

			Do While 1
				se = InStr(ss, msg, Chr(12))

				If se <> 0 Then
					item = Mid(msg, ss, se - ss)
					em.SetMailBoxSize item, uSize
				Else
					Exit Do
				End If

				ss = se + 1
			Loop
		end if
	end if

	set em = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
elseif trim(request("onlychangesize")) = "2" then
	accessmode = trim(request("accessmode"))

	if IsNumeric(accessmode) = true then
		accessmode = CInt(accessmode)

		if Len(msg) > 0 then
			ss = 1
			se = 1

			Do While 1
				se = InStr(ss, msg, Chr(12))

				If se <> 0 Then
					item = Mid(msg, ss, se - ss)
					em.SetAccessMode item, accessmode
				Else
					Exit Do
				End If

				ss = se + 1
			Loop
		end if
	end if

	set em = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


if trim(request("onlychangelimitout")) = "1" then
	if trim(request("changeLimitOut")) <> "" then
		if trim(request("uLimitOut")) = "1" then
			uLimitOut = true
		else
			uLimitOut = false
		end if

		if Len(msg) > 0 then
			ss = 1
			se = 1

			Do While 1
				se = InStr(ss, msg, Chr(12))

				If se <> 0 Then
					item = Mid(msg, ss, se - ss)
					em.SetLimitOut item, uLimitOut
				Else
					Exit Do
				End If

				ss = se + 1
			Loop
		end if
	end if

	set em = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


if trim(request("onlyReceiveOutMail")) = "1" then
	if trim(request("changeReceiveOutMail")) <> "" then
		if trim(request("uReceiveOutMail")) = "1" then
			uReceiveOutMail = false
		else
			uReceiveOutMail = true
		end if

		if Len(msg) > 0 then
			ss = 1
			se = 1

			Do While 1
				se = InStr(ss, msg, Chr(12))

				If se <> 0 Then
					item = Mid(msg, ss, se - ss)
					em.SetReceiveOutMail item, uReceiveOutMail
				Else
					Exit Do
				End If

				ss = se + 1
			Loop
		end if
	end if

	set em = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


if trim(request("onlychangeMaxPerFoldersNumber")) = "1" then
	if IsNumeric(trim(request("uMaxFolders"))) = true then
		uMaxFolders = CInt(trim(request("uMaxFolders")))

		if Len(msg) > 0 then
			ss = 1
			se = 1

			Do While 1
				se = InStr(ss, msg, Chr(12))

				If se <> 0 Then
					item = Mid(msg, ss, se - ss)
					em.SetMaxPerFolderNumber item, uMaxFolders
				Else
					Exit Do
				End If

				ss = se + 1
			Loop
		end if
	end if

	set em = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


if trim(request("onlychangeExpiresDay")) = "1" then
	if trim(request("changeExpiresDay")) <> "" then
		t_date = trim(request("t_year")) & trim(request("t_month")) & trim(request("t_day"))

		if Len(msg) > 0 then
			ss = 1
			se = 1

			Do While 1
				se = InStr(ss, msg, Chr(12))

				If se <> 0 Then
					item = Mid(msg, ss, se - ss)
					em.SetExpires item, t_date
				Else
					Exit Do
				End If

				ss = se + 1
			Loop
		end if
	end if

	set em = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


if trim(request("onlychangeSSL")) = "1" then
	if trim(request("uSSL")) = "0" then
		uSSL = false
	else
		uSSL = true
	end if

	if Len(msg) > 0 then
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(12))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				em.SetSSL item, uSSL
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	set em = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


if trim(request("onlychangeMonitor")) = "1" then
	if trim(request("changeMonitor")) <> "" then
		if trim(request("uMonitor")) = "1" then
			uMonitor = true
		else
			uMonitor = false
		end if

		if Len(msg) > 0 then
			ss = 1
			se = 1

			Do While 1
				se = InStr(ss, msg, Chr(12))

				If se <> 0 Then
					item = Mid(msg, ss, se - ss)
					em.SetMonitor item, uMonitor
				Else
					Exit Do
				End If

				ss = se + 1
			Loop
		end if
	end if

	set em = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


if trim(request("onlychangespamguard")) = "1" then
	dim my_usg
	set my_usg = server.createobject("easymail.UserSpamGuard")
	my_usg.Load Session("wem")

	dim my_userweb
	set my_userweb = server.createobject("easymail.UserWeb")
	my_userweb.Load Session("wem")

	dim my_tm
	set my_tm = server.createobject("easymail.TrashMsg")
	my_tm.Load Session("wem")

	dim other_usg
	set other_usg = server.createobject("easymail.UserSpamGuard")

	dim other_userweb
	set other_userweb = server.createobject("easymail.UserWeb")

	dim other_tm
	set other_tm = server.createobject("easymail.TrashMsg")

	if Len(msg) > 0 then
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(12))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)

				other_usg.Load item
				other_userweb.Load item
				other_tm.Load item

				other_userweb.useAutoClearTrashBox = my_userweb.useAutoClearTrashBox
				other_userweb.autoClearTrashBoxDays = my_userweb.autoClearTrashBoxDays
				other_userweb.EnableClearWhenFull = my_userweb.EnableClearWhenFull
				other_userweb.EnableClearSendBox = my_userweb.EnableClearSendBox

				other_usg.EnableReceiveAllMail = my_usg.EnableReceiveAllMail
				other_usg.EnableTrashMsg = my_usg.EnableTrashMsg
				other_usg.EnableUser_Receive_MailToCc_MyEmail = my_usg.EnableUser_Receive_MailToCc_MyEmail
				other_usg.EnableUser_SpamGuard = my_usg.EnableUser_SpamGuard
				other_usg.EnableUser_ReceiveLocal = my_usg.EnableUser_ReceiveLocal
				other_usg.EnableUser_ReceiveAddressBook = my_usg.EnableUser_ReceiveAddressBook
				other_usg.EnableUser_ReceiveFromOutEmails = my_usg.EnableUser_ReceiveFromOutEmails
				other_usg.EnableUser_ReceiveDomain = my_usg.EnableUser_ReceiveDomain
				other_usg.Enable_User_NoSpam_Affirm = my_usg.Enable_User_NoSpam_Affirm
				other_usg.SpamProcessMode = my_usg.SpamProcessMode
				other_usg.EnableUser_OneDayMulRepeatReceive_Guard = my_usg.EnableUser_OneDayMulRepeatReceive_Guard
				other_usg.OneDayMulRepeatReceive_Guard_ProcessMode = my_usg.OneDayMulRepeatReceive_Guard_ProcessMode
				other_usg.EnableSizeLimitNoSpam = my_usg.EnableSizeLimitNoSpam
				other_usg.NoSpamSizeLimit = my_usg.NoSpamSizeLimit

				other_tm.IntervalMin = my_tm.IntervalMin

				other_usg.Save
				other_userweb.Save
				other_tm.Save
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	set my_usg = nothing
	set my_userweb = nothing
	set my_tm = nothing
	set other_usg = nothing
	set other_userweb = nothing
	set other_tm = nothing

	set em = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if

	set em = nothing
end if
%>

<html>
<head>
<%=s_lang_meta %>
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</head>

<SCRIPT LANGUAGE=javascript>
<!--
function modifyAccess()
{
	document.form1.onlychangesize.value = "2";
	document.form1.submit();
}

function changeAccess_onclick() {
	if (document.form1.accessmode.disabled == true)
	{
		document.form1.btaccesschange.disabled = false;
		document.form1.accessmode.disabled = false;
	}
	else
	{
		document.form1.btaccesschange.disabled = true;
		document.form1.accessmode.disabled = true;
	}
}

function modifySize()
{
	document.form1.onlychangesize.value = "1";
	document.form1.submit();
}

function changeSize_onclick() {
	if (document.form1.uSize.disabled == true)
	{
		document.form1.btchange.disabled = false;
		document.form1.uSize.disabled = false;
	}
	else
	{
		document.form1.btchange.disabled = true;
		document.form1.uSize.disabled = true;
	}
}

function changeLimitOut_onclick() {
	if (document.form1.uLimitOut.disabled == true)
	{
		document.form1.btLimitOut.disabled = false;
		document.form1.uLimitOut.disabled = false;
	}
	else
	{
		document.form1.btLimitOut.disabled = true;
		document.form1.uLimitOut.disabled = true;
	}
}

function modifyLimitOut(){
	document.form1.onlychangelimitout.value = "1";
	document.form1.submit();
}

function changeReceiveOutMail_onclick() {
	if (document.form1.uReceiveOutMail.disabled == true)
	{
		document.form1.btReceiveOutMail.disabled = false;
		document.form1.uReceiveOutMail.disabled = false;
	}
	else
	{
		document.form1.btReceiveOutMail.disabled = true;
		document.form1.uReceiveOutMail.disabled = true;
	}
}

function modifyReceiveOutMail(){
	document.form1.onlyReceiveOutMail.value = "1";
	document.form1.submit();
}

function changeExpiresDay_onclick() {
	if (document.form1.t_year.disabled == true)
	{
		document.form1.t_year.disabled = false;
		document.form1.t_month.disabled = false;
		document.form1.t_day.disabled = false;
		document.form1.btExpiresDay.disabled = false;
	}
	else
	{
		document.form1.t_year.disabled = true;
		document.form1.t_month.disabled = true;
		document.form1.t_day.disabled = true;
		document.form1.btExpiresDay.disabled = true;
	}
}

function modifyExpiresDay(){
	document.form1.onlychangeExpiresDay.value = "1";
	document.form1.submit();
}

function changeMonitor_onclick() {
	if (document.form1.uMonitor.disabled == true)
	{
		document.form1.btMonitor.disabled = false;
		document.form1.uMonitor.disabled = false;
	}
	else
	{
		document.form1.btMonitor.disabled = true;
		document.form1.uMonitor.disabled = true;
	}
}

function changeMaxFolders_onclick() {
	if (document.form1.uMaxFolders.disabled == true)
	{
		document.form1.mfnchange.disabled = false;
		document.form1.uMaxFolders.disabled = false;
	}
	else
	{
		document.form1.mfnchange.disabled = true;
		document.form1.uMaxFolders.disabled = true;
	}
}

function modifyMaxPerFolderNumber(){
	document.form1.onlychangeMaxPerFoldersNumber.value = "1";
	document.form1.submit();
}

function modifyMonitor(){
	document.form1.onlychangeMonitor.value = "1";
	document.form1.submit();
}

function changeSSL_onclick() {
	if (document.form1.uSSL.disabled == true)
	{
		document.form1.btSSL.disabled = false;
		document.form1.uSSL.disabled = false;
	}
	else
	{
		document.form1.btSSL.disabled = true;
		document.form1.uSSL.disabled = true;
	}
}

function modifySSL(){
	document.form1.onlychangeSSL.value = "1";
	document.form1.submit();
}

function changeComment_onclick() {
	if (document.form1.ucomment.disabled == true)
	{
		document.form1.ucomment.disabled = false;
		document.form1.ucmtchange.disabled = false;
	}
	else
	{
		document.form1.ucomment.disabled = true;
		document.form1.ucmtchange.disabled = true;
	}
}

function cleanmailbox(){
	if (confirm("确实要清空他们的邮箱吗?") == false)
		return ;

	document.form1.onlyCleanMailBox.value = "1";
	document.form1.submit();
}

function closeIPLimit(){
	document.form1.onlycloseIPLimit.value = "1";
	document.form1.submit();
}

function changespamguard(){
	document.form1.onlychangespamguard.value = "1";
	document.form1.submit();
}

function goback(){
	location.href = "<%=gourl %>";
}

function window_onload() {
	document.form1.mulusers.value = parent.f1.document.leftval.temp.value;
}
//-->
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
<br>
<br>
<div align="center">
<form name="form1" method="post" action="mulchange.asp">
	<input type="hidden" name="onlychangespamguard">
	<input type="hidden" name="mulusers">
	<input type="hidden" name="onlychangeMonitor">
	<input type="hidden" name="onlychangeExpiresDay">
	<input type="hidden" name="onlychangelimitout">
	<input type="hidden" name="onlyReceiveOutMail">
	<input type="hidden" name="onlychangesize">
	<input type="hidden" name="onlychangeMaxPerFoldersNumber">
	<input type="hidden" name="gourl" value="<%=gourl %>">
	<input type="hidden" name="onlychangeSSL">
	<input type="hidden" name="onlyCleanMailBox">
	<input type="hidden" name="onlycloseIPLimit">
    <table width="80%" border="0" align="center" cellspacing="0" style="border-top:1px <%=MY_COLOR_1 %> solid;">
    <tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%"><input type="checkbox" name="changeSize" LANGUAGE=javascript onclick="return changeSize_onclick()">修改用户邮箱大小&nbsp;&nbsp;
	<input type="text" name="uSize" maxlength="8" class="textbox" disabled>&nbsp;K
	</td><td><input type="button" id="btchange" value=" 修改 " onclick="modifySize()" class="Bsbttn" disabled>
	</td></tr></table></td>
    </tr>
	<tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%"><input type="checkbox" name="changeAccess" LANGUAGE=javascript onclick="return changeAccess_onclick()">修改用户访问方式&nbsp;&nbsp;
	<select name="accessmode" class=drpdwn size="1" disabled>
<%
anum = 0

do while anum < 7
	response.write "<option value=""" & anum & """>" & getaccessmode(anum) & "</option>"
	anum = anum + 1
loop
%>
	</select>
	</td><td><input type="button" id="btaccesschange" value=" 修改 " onclick="modifyAccess()" class="Bsbttn" disabled>
	</td></tr></table></td>
    </tr>
    <tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%"><input type="checkbox" name="changeLimitOut" LANGUAGE=javascript onclick="return changeLimitOut_onclick()">修改&nbsp;&nbsp;
	<select name="uLimitOut" class=drpdwn size="1" disabled>
	<option value="" selected>允许此帐号对系统外发信</option>
	<option value="1">禁止此帐号对系统外发信</option>
	</select>
	</td><td><input type="button" id="btLimitOut" value=" 修改 " onclick="modifyLimitOut()" class="Bsbttn" disabled>
	</td></tr></table></td>
    </tr>
    <tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%"><input type="checkbox" name="changeReceiveOutMail" LANGUAGE=javascript onclick="return changeReceiveOutMail_onclick()"><%=s_lang_modify %>&nbsp;&nbsp;
	<select name="uReceiveOutMail" class=drpdwn size="1" disabled>
	<option value="" selected><%=s_lang_0076 %></option>
	<option value="1"><%=s_lang_0077 %></option>
	</select>
	</td><td><input type="button" id="btReceiveOutMail" value=" <%=s_lang_modify %> " onclick="modifyReceiveOutMail()" class="Bsbttn" disabled>
	</td></tr></table></td>
    </tr>
    <tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%"><input type="checkbox" name="changeExpiresDay" LANGUAGE=javascript onclick="return changeExpiresDay_onclick()">修改期满日期&nbsp;&nbsp;

<select name="t_year" class="drpdwn" disabled>
<option value="">------</option>
<%
	now_temp = Year(Now())
	i = now_temp - 10

	if i < 2000 then
		i = 2000
	end if

	now_temp = i

	df_year = getYear(expiresday)

	if df_year = "" then
		df_year = "0"
	end if

	do while i < now_temp + 60
		if CInt(df_year) = i then
			response.write "<option value='" & i & "' selected>" & i & "年</option>" & Chr(13)
		else
			response.write "<option value='" & i & "'>" & i & "年</option>" & Chr(13)
		end if

		i = i + 1
	loop
%>
</select>&nbsp;
<select name="t_month" class="drpdwn" disabled>
<option value="">----</option>
<%
	now_temp = getMonth(expiresday)

	if now_temp = "" then
		now_temp = "0"
	end if

	i = 1
	do while i < 13
		if i <> CInt(now_temp) then
			if i < 10 then
				response.write "<option value='0" & i & "'>" & i & "月</option>"
			else
				response.write "<option value='" & i & "'>" & i & "月</option>"
			end if
		else
			if i < 10 then
				response.write "<option value='0" & i & "' selected>" & i & "月</option>"
			else
				response.write "<option value='" & i & "' selected>" & i & "月</option>"
			end if
		end if
		i = i + 1
	loop
%>
</select>&nbsp;
<select name="t_day" class="drpdwn" disabled>
<option value="">----</option>
<%
	now_temp = getDay(expiresday)

	if now_temp = "" then
		now_temp = "0"
	end if

	i = 1
	do while i < 32
		if i <> CInt(now_temp) then
			if i < 10 then
				response.write "<option value='0" & i & "'>" & i & "日</option>"
			else
				response.write "<option value='" & i & "'>" & i & "日</option>"
			end if
		else
			if i < 10 then
				response.write "<option value='0" & i & "' selected>" & i & "日</option>"
			else
				response.write "<option value='" & i & "' selected>" & i & "日</option>"
			end if
		end if
		i = i + 1
	loop
%>
</select>
	</td><td><input type="button" id="btExpiresDay" value=" 修改 " onclick="modifyExpiresDay()" class="Bsbttn" disabled>
	</td></tr></table></td>
    </tr>
    <tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%"><input type="checkbox" name="changeMonitor" LANGUAGE=javascript onclick="return changeMonitor_onclick()">修改是否进行域邮件监控&nbsp;&nbsp;
	<select name="uMonitor" class=drpdwn size="1" disabled>
	<option value="" selected>不监控</option>
	<option value="1">监控</option>
	</select>
	</td><td><input type="button" id="btMonitor" value=" 修改 " onclick="modifyMonitor()" class="Bsbttn" disabled>
	</td></tr></table></td>
    </tr>
    <tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%"><input type="checkbox" name="changeMaxFolders" LANGUAGE=javascript onclick="return changeMaxFolders_onclick()">修改最多允许创建的私人文件夹数&nbsp;&nbsp;
	<input type="text" name="uMaxFolders" size="5" maxlength="2" class="textbox" disabled>
	</td><td><input type="button" id="mfnchange" value=" 修改 " onclick="modifyMaxPerFolderNumber()" class="Bsbttn" disabled>
	</td></tr></table></td>
    </tr>
    <tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%"><input type="checkbox" name="changeSSL" LANGUAGE=javascript onclick="return changeSSL_onclick()">修改&nbsp;&nbsp;
	<select name="uSSL" class=drpdwn size="1" disabled>
	<option value="1" selected>启用SSL连接</option>
	<option value="0">禁用SSL连接</option>
	</select>
	</td><td><input type="button" id="btSSL" value=" 修改 " onclick="modifySSL()" class="Bsbttn" disabled>
	</td></tr></table></td>
    </tr>
    <tr>
	<td align="center" colspan="2" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td>
	<input type="button" value="关闭登录IP限制功能" onclick="closeIPLimit()" class="Bsbttn">
	</td></tr></table></td>
    </tr>
    <tr>
	<td align="center" colspan="2" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td>
	<input type="button" value="按我的防垃圾选项统一对他们进行修改" style="WIDTH: 260px" onclick="changespamguard()" class="Bsbttn">
	</td></tr></table></td>
    </tr>
    <tr>
	<td align="center" colspan="2" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td>
	<input type="button" value="清空他们的邮箱" onclick="cleanmailbox()" class="Bsbttn">
	</td></tr></table></td>
    </tr>
      <tr>
        <td colspan="2" align="right"><br>
          <input type="button" value=" 返回 " onclick="goback()" class="Bsbttn">
        </td>
      </tr>
</table>
</form>
</div>
<br><br>
  <div align="center">
    <table width="80%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10">
        </td>
      </tr>
      <tr>
        <td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
		<td width="94%">用户访问方式 (可以进行组合设置):
		<br>http: 允许用户通过浏览器访问WebMail
		<br>smtp: 允许用户使用smtp协议访问邮箱
		<br>pop3,imap4: 允许用户使用pop3以及imap4协议访问邮箱
		<br><br>在设定了用户的期满日期后, 除非当前日期已超过了所设的期满日期, 否则该用户将不会因其他原因(如: 长时间未使用邮箱)被系统自动禁用.
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
<br><br>
</body>
</html>

<%
function getYear(exday)
	getYear = Mid(Cstr(exday), 1, 4)
end function

function getMonth(exday)
	getMonth = Mid(Cstr(exday), 5, 2)
end function

function getDay(exday)
	getDay = Mid(Cstr(exday), 7, 2)
end function

function getaccessmode(amode)
	if amode = 0 then
		getaccessmode = "http/smtp/pop3,imap4"
	elseif amode = 1 then
		getaccessmode = "smtp/pop3,imap4"
	elseif amode = 2 then
		getaccessmode = "http/smtp"
	elseif amode = 3 then
		getaccessmode = "http/pop3,imap4"
	elseif amode = 4 then
		getaccessmode = "http"
	elseif amode = 5 then
		getaccessmode = "smtp"
	elseif amode = 6 then
		getaccessmode = "pop3,imap4"
	end if
end function
%>
