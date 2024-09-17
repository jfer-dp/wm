<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
fromother = trim(request("fo"))

userid = trim(request("id"))
if IsNumeric(userid) = true then
	userid = CInt(userid)
else
	response.redirect "noadmin.asp"
end if

gourl = trim(request("gourl"))

Set em = Application("em")
em.GetUserByIndex3 userid, name, domain, comment, forbid, lasttime, amode, limitout, expiresday, monitor

comment = NULL
forbid = NULL
lasttime = NULL

if isadmin() = false and isAccountsAdmin() = false then
	if LCase(name) = LCase(Application("em_SystemAdmin")) then
		set em = nothing

		domain = NULL
		name = NULL
		amode = NULL
		limitout = NULL
		expiresday = NULL
		monitor = NULL

		response.redirect "noadmin.asp"
	end if


	dim ed
	set ed = server.createobject("easymail.domain")
	ed.Load

	dim wem_user
	wem_user = Session("wem")

	if ed.GetUserManagerDomainCount(wem_user) < 1 then
		set ed = nothing
		set em = nothing
		domain = NULL
		name = NULL
		amode = NULL
		limitout = NULL
		expiresday = NULL
		monitor = NULL

		response.redirect "noadmin.asp"
	end if


	i = 0
	allnum = ed.GetUserManagerDomainCount(wem_user)

	dim isok
	isok = false

	do while i < allnum
		cdomainstr = ed.GetUserManagerDomain(wem_user, i)

		if cdomainstr = domain then
			isok = true
		end if

		cdomainstr = NULL

		i = i + 1
	loop

	set ed = nothing


	if isok = false then
		set em = nothing
		domain = NULL
		name = NULL
		amode = NULL
		limitout = NULL
		expiresday = NULL
		monitor = NULL

		response.redirect "noadmin.asp"
	end if
end if


if isadmin() = false then
	if LCase(name) = LCase(Application("em_SystemAdmin")) then
		set em = nothing

		domain = NULL
		name = NULL
		amode = NULL
		limitout = NULL
		expiresday = NULL
		monitor = NULL

		response.redirect "noadmin.asp"
	end if
end if


'-----------------------------------------
dim mam
set mam = server.createobject("easymail.AdminManager")
mam.Load

if trim(request("start_esum")) = "true" then
	mam.Enable_Show_User_Memo = true
	mam.Save

	set mam = nothing
	set em = nothing
	domain = NULL
	name = NULL
	amode = NULL
	limitout = NULL
	expiresday = NULL
	monitor = NULL

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if

Enable_DAdminAllotSize = mam.Enable_DAdminAllotSize
Enable_Show_User_Memo = mam.Enable_Show_User_Memo

if trim(request("onlyCleanMailBox")) = "1" then
	mam.CleanMailBox(name)

	set mam = nothing
	set em = nothing
	domain = NULL
	name = NULL
	amode = NULL
	limitout = NULL
	expiresday = NULL
	monitor = NULL

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if

set mam = nothing


if trim(request("onlycloseIPLimit")) = "1" then
	em.SetEnableIPLimit name, false

	set em = nothing
	domain = NULL
	name = NULL
	amode = NULL
	limitout = NULL
	expiresday = NULL
	monitor = NULL

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


if trim(request("onlychangesize")) = "1" and (isadmin() = true or isAccountsAdmin() = true or Enable_DAdminAllotSize = true) then
	uSize = trim(request("uSize"))

	if trim(request("changeSize")) <> "" and IsNumeric(uSize) then
		em.SetMailBoxSize name, CLng(uSize)
	end if

	set em = nothing
	domain = NULL
	name = NULL
	amode = NULL
	limitout = NULL
	expiresday = NULL
	monitor = NULL

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
elseif trim(request("onlychangesize")) = "2" then
	accessmode = trim(request("accessmode"))

	if IsNumeric(accessmode) = true then
		if amode <> CInt(accessmode) then
			em.SetAccessMode name, CInt(accessmode)
		end if
	end if

	set em = nothing
	domain = NULL
	name = NULL
	amode = NULL
	limitout = NULL
	expiresday = NULL
	monitor = NULL

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


if trim(request("onlychangelimitout")) = "1" then
	if trim(request("changeLimitOut")) <> "" then
		if trim(request("uLimitOut")) = "1" then
			em.SetLimitOut name, true
		else
			em.SetLimitOut name, false
		end if
	end if

	set em = nothing
	domain = NULL
	name = NULL
	amode = NULL
	limitout = NULL
	expiresday = NULL
	monitor = NULL

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


if trim(request("onlyReceiveOutMail")) = "1" then
	if trim(request("changeReceiveOutMail")) <> "" then
		if trim(request("uReceiveOutMail")) = "1" then
			em.SetReceiveOutMail name, false
		else
			em.SetReceiveOutMail name, true
		end if
	end if

	set em = nothing
	domain = NULL
	name = NULL
	amode = NULL
	limitout = NULL
	expiresday = NULL
	monitor = NULL

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


if trim(request("onlychangeMaxPerFoldersNumber")) = "1" then
	if IsNumeric(trim(request("uMaxFolders"))) = true then
		em.SetMaxPerFolderNumber name, CInt(trim(request("uMaxFolders")))
	end if

	set em = nothing
	domain = NULL
	name = NULL
	amode = NULL
	limitout = NULL
	expiresday = NULL
	monitor = NULL

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


if trim(request("onlychangeExpiresDay")) = "1" then
	if trim(request("changeExpiresDay")) <> "" then
		em.SetExpires name, trim(request("t_year")) & trim(request("t_month")) & trim(request("t_day"))
	end if

	set em = nothing
	domain = NULL
	name = NULL
	amode = NULL
	limitout = NULL
	expiresday = NULL
	monitor = NULL

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


if trim(request("onlychangeSSL")) = "1" then
	if trim(request("uSSL")) = "0" then
		em.SetSSL name, false
	else
		em.SetSSL name, true
	end if

	set em = nothing
	domain = NULL
	name = NULL
	amode = NULL
	limitout = NULL
	expiresday = NULL
	monitor = NULL

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


if trim(request("onlychangeARC")) = "1" then
	if IsNumeric(trim(request("Max_Archive"))) = true then
		set march = server.createobject("easymail.MailArchive")
		march.Load name
		march.Max_Archive = CLng(trim(request("Max_Archive")))
		set march = nothing
	end if

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


if trim(request("onlychangeComment")) = "1" then
	set mri = server.createobject("easymail.MoreRegInfo")
	mri.LoadRegInfo name
	mri.Comment = trim(request("ucomment"))
	mri.SaveRegInfo
	set mri = nothing

	set em = nothing
	domain = NULL
	name = NULL
	amode = NULL
	limitout = NULL
	expiresday = NULL
	monitor = NULL

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


dim sysinfo
set sysinfo = server.createobject("easymail.sysinfo")
sysinfo.Load

if trim(request("onlychangeMonitor")) = "1" and (isadmin() = true or isAccountsAdmin() = true or sysinfo.enableDomainMonitor = true) then
	if trim(request("changeMonitor")) <> "" then
		if trim(request("uMonitor")) = "1" then
			em.SetMonitor name, true
		else
			em.SetMonitor name, false
		end if
	end if

	set em = nothing
	set sysinfo = nothing

	domain = NULL
	name = NULL
	amode = NULL
	limitout = NULL
	expiresday = NULL
	monitor = NULL

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


pw = trim(request("pw1"))

if pw <> "" then
	em.ChangeUserPassWord name, pw

	set em = nothing
	set sysinfo = nothing

	domain = NULL
	name = NULL
	amode = NULL
	limitout = NULL
	expiresday = NULL
	monitor = NULL

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if
%>

<html>
<head>
<title>编辑</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
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

function checkpw(){
	if (document.form1.pw1.value != "" && document.form1.pw2.value != "")
	{
		if (document.form1.pw1.value != document.form1.pw2.value)
			alert("输入的密码不相同");
		else
			document.form1.submit();
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

function modifyARC(){
	document.form1.onlychangeARC.value = "1";
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

function modifyComment(){
	document.form1.onlychangeComment.value = "1";
	document.form1.submit();
}

function cleanmailbox(){
	if (confirm("确实要清空 [<%=server.htmlencode(name) %>] 的邮箱吗?") == false)
		return ;

	document.form1.onlyCleanMailBox.value = "1";
	document.form1.submit();
}

function closeIPLimit(){
	document.form1.onlycloseIPLimit.value = "1";
	document.form1.submit();
}

function changespamguard(){
	location.href = "userspamguard.asp?<%=getGRSN() %>&name=<%=Server.URLEncode(name) %>&domain=<%=Server.URLEncode(domain) %>&amg=1&id=<%=userid %>&gourl=<%=Server.URLEncode(gourl) %>";
}

<% if IsEnterpriseVersion = true then %>
function change_ldap(){
	location.href = "ldap.asp?<%=getGRSN() %>&name=<%=Server.URLEncode(name) %>&domain=<%=Server.URLEncode(domain) %>&amg=1&id=<%=userid %>&gourl=<%=Server.URLEncode(gourl) %>";
}
<% end if %>

function goback(){
<%
if fromother = "1" then
%>
	location.href = "<%=gourl %>";
<%
else
%>
	history.back();
<%
end if
%>
}

function start_Enable_Show_User_Memo(){
	document.form1.start_esum.value = "true";
	document.form1.submit();
}
//-->
</script>

<body>
<br>
<br>
<div align="center">
<form name="form1" method="post" action="changepw.asp">
	<input type="hidden" name="start_esum">
	<input type="hidden" name="onlychangeMonitor">
	<input type="hidden" name="onlychangeExpiresDay">
	<input type="hidden" name="onlychangelimitout">
	<input type="hidden" name="onlyReceiveOutMail">
	<input type="hidden" name="onlychangesize">
	<input type="hidden" name="onlychangeMaxPerFoldersNumber">
	<input type="hidden" name="gourl" value="<%=gourl %>">
	<input type="hidden" name="id" value="<%=trim(request("id")) %>">
	<input type="hidden" name="onlychangeSSL">
	<input type="hidden" name="onlychangeARC">
	<input type="hidden" name="onlychangeComment">
	<input type="hidden" name="onlyCleanMailBox">
	<input type="hidden" name="onlycloseIPLimit">
	<table width="80%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>">
        <td colspan="2" nowrap height="30" style="border-left:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>
          <div align="center">密码修改 - [<font color="#FF3333"><%=server.htmlencode(name) %></font>]</div>
        </td>
      </tr>
      <tr>
        <td width="28%" height="28" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
          <div align="right">新 密 码&nbsp;:&nbsp;</div>
        </td>
        <td style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>  
          <input type="password" name="pw1" maxlength="32" size="45" class="textbox">
          </td>
      </tr>
      <tr>
        <td width="28%" height="28" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
          <div align="right">密码确认&nbsp;:&nbsp;</div>
        </td>
        <td style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>  
          <input type="password" name="pw2" maxlength="32" size="45" class="textbox">
          </td>
      </tr>
      <tr>
        <td colspan="2" align="right"><br>
          <input type="button" value=" 确定 " onclick="checkpw()" class="Bsbttn">&nbsp;&nbsp;
          <input type="button" value=" 返回 " onclick="goback()" class="Bsbttn">
        </td>
      </tr>
    </table>
	<br><br><br>
    <table width="80%" border="0" align="center" cellspacing="0" style="border-top:1px <%=MY_COLOR_1 %> solid;">
<%
set mri = server.createobject("easymail.MoreRegInfo")
mri.LoadRegInfo name
%>
    <tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%"><input type="checkbox" name="changeComment" LANGUAGE=javascript onclick="return changeComment_onclick()">修改备注&nbsp;&nbsp;
	<input type="text" name="ucomment" value="<%=mri.Comment %>" maxlength="128" class="textbox" disabled size="50">
<%
if Enable_Show_User_Memo = false then
%>
	&nbsp;<input type="button" id="Enable_Show_User_Memo" name="Enable_Show_User_Memo" value="开启列表显示" style="WIDTH: 90px" onclick="start_Enable_Show_User_Memo()" class="Bsbttn"><br>
<%
end if
%>
	</td><td><input type="button" id="ucmtchange" value=" 修改 " onclick="modifyComment()" class="Bsbttn" disabled>
	</td></tr></table></td>
    </tr>
<%
set mri = nothing


if isadmin() = true or isAccountsAdmin() = true or Enable_DAdminAllotSize = true then
%>
    <tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%"><input type="checkbox" name="changeSize" LANGUAGE=javascript onclick="return changeSize_onclick()">修改用户邮箱大小&nbsp;&nbsp;
	<input type="text" name="uSize" value="<%=em.GetMailBoxSize(name) %>" maxlength="8" class="textbox" disabled>&nbsp;K
	</td><td><input type="button" id="btchange" value=" 修改 " onclick="modifySize()" class="Bsbttn" disabled>
	</td></tr></table></td>
    </tr>
<%
end if
%>
	<tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%"><input type="checkbox" name="changeAccess" LANGUAGE=javascript onclick="return changeAccess_onclick()">修改用户访问方式&nbsp;&nbsp;
	<select name="accessmode" class=drpdwn size="1" disabled>
<%
anum = 0

if wem_user <> name then
	do while anum < 7
		if amode = anum then
			response.write "<option value=""" & anum & """ selected>" & getaccessmode(anum) & "</option>"
		else
			response.write "<option value=""" & anum & """>" & getaccessmode(anum) & "</option>"
		end if
		anum = anum + 1
	loop
else
	do while anum < 7
		if anum <> 1 and anum <> 5 and anum <> 6 then
			if amode = anum then
				response.write "<option value=""" & anum & """ selected>" & getaccessmode(anum) & "</option>"
			else
				response.write "<option value=""" & anum & """>" & getaccessmode(anum) & "</option>"
			end if
		end if

		anum = anum + 1
	loop
end if
%>
	</select>
	</td><td><input type="button" id="btaccesschange" value=" 修改 " onclick="modifyAccess()" class="Bsbttn" disabled>
	</td></tr></table></td>
    </tr>
    <tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%"><input type="checkbox" name="changeLimitOut" LANGUAGE=javascript onclick="return changeLimitOut_onclick()">修改&nbsp;&nbsp;
	<select name="uLimitOut" class=drpdwn size="1" disabled>
<%
if limitout = false then
%>
	<option value="" selected>允许此帐号对系统外发信</option>
	<option value="1">禁止此帐号对系统外发信</option>
<%
else
%>
	<option value="">允许此帐号对系统外发信</option>
	<option value="1" selected>禁止此帐号对系统外发信</option>
<%
end if
%>
	</select>
	</td><td><input type="button" id="btLimitOut" value=" 修改 " onclick="modifyLimitOut()" class="Bsbttn" disabled>
	</td></tr></table></td>
    </tr>
    <tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%"><input type="checkbox" name="changeReceiveOutMail" LANGUAGE=javascript onclick="return changeReceiveOutMail_onclick()"><%=s_lang_modify %>&nbsp;&nbsp;
	<select name="uReceiveOutMail" class=drpdwn size="1" disabled>
<%
if em.GetReceiveOutMail(name) = true then
%>
	<option value="" selected><%=s_lang_0076 %></option>
	<option value="1"><%=s_lang_0077 %></option>
<%
else
%>
	<option value=""><%=s_lang_0076 %></option>
	<option value="1" selected><%=s_lang_0077 %></option>
<%
end if
%>
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
<%
if isadmin() = true or isAccountsAdmin() = true or sysinfo.enableDomainMonitor = true then
%>
    <tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%"><input type="checkbox" name="changeMonitor" LANGUAGE=javascript onclick="return changeMonitor_onclick()">修改是否进行域邮件监控&nbsp;&nbsp;
	<select name="uMonitor" class=drpdwn size="1" disabled>
<%
if monitor = false then
%>
	<option value="" selected>不监控</option>
	<option value="1">监控</option>
<%
else
%>
	<option value="">不监控</option>
	<option value="1" selected>监控</option>
<%
end if
%>
	</select>
	</td><td><input type="button" id="btMonitor" value=" 修改 " onclick="modifyMonitor()" class="Bsbttn" disabled>
	</td></tr></table></td>
    </tr>
<%
end if
%>
    <tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%"><input type="checkbox" name="changeMaxFolders" LANGUAGE=javascript onclick="return changeMaxFolders_onclick()">修改最多允许创建的私人文件夹数&nbsp;&nbsp;
	<input type="text" name="uMaxFolders" value="<%=em.GetMaxPerFolderNumber(name) %>" size="5" maxlength="2" class="textbox" disabled>
	</td><td><input type="button" id="mfnchange" value=" 修改 " onclick="modifyMaxPerFolderNumber()" class="Bsbttn" disabled>
	</td></tr></table></td>
    </tr>
    <tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%"><input type="checkbox" name="changeSSL" LANGUAGE=javascript onclick="return changeSSL_onclick()">修改&nbsp;&nbsp;
	<select name="uSSL" class=drpdwn size="1" disabled>
<%
if em.GetSSL(name) = true then
%>
	<option value="1" selected>启用SSL连接</option>
	<option value="0">禁用SSL连接</option>
<%
else
%>
	<option value="1">启用SSL连接</option>
	<option value="0" selected>禁用SSL连接</option>
<%
end if
%>
	</select>
	</td><td><input type="button" id="btSSL" value=" 修改 " onclick="modifySSL()" class="Bsbttn" disabled>
	</td></tr></table></td>
    </tr>

    <tr>
	<td width="62%" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td width="90%">&nbsp;文件归档最大限额:&nbsp;<input type="text" name="Max_Archive" id="Max_Archive" value="<%
set march = server.createobject("easymail.MailArchive")
march.Load name

Response.Write march.Max_Archive
set march = nothing
%>" size="10" maxlength="8" class="textbox">&nbsp;&nbsp;&nbsp;(此项为零时禁用文件归档功能)
	</td><td><input type="button" value=" 修改 " onclick="modifyARC()" class="Bsbttn">
	</td></tr></table></td>
    </tr>

    <tr>
	<td align="center" colspan="2" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td>
	<input type="button" value="关闭登录IP限制功能" onclick="closeIPLimit()" class="Bsbttn">
	</td></tr></table></td>
    </tr>
<% if IsEnterpriseVersion = true then %>
    <tr>
	<td align="center" colspan="2" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td>
	<input type="button" value="<%=s_lang_modify %> [<%=server.htmlencode(name) %>] <%=s_lang_0060 %>" onclick="change_ldap()" class="Bsbttn">
	</td></tr></table></td>
    </tr>
<% end if %>
    <tr>
	<td align="center" colspan="2" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td>
	<input type="button" value="修改 [<%=server.htmlencode(name) %>] 的防垃圾选项" onclick="changespamguard()" class="Bsbttn">
	</td></tr></table></td>
    </tr>
    <tr>
	<td align="center" colspan="2" bgcolor="<%=MY_COLOR_3 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table><tr><td>
	<input type="button" value="清空 [<%=server.htmlencode(name) %>] 的邮箱" onclick="cleanmailbox()" class="Bsbttn">
	</td></tr></table></td>
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
set em = nothing
set sysinfo = nothing

domain = NULL
name = NULL
amode = NULL
limitout = NULL
expiresday = NULL
monitor = NULL


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
