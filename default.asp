<%
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>

<%
un = trim(request("username"))
pw = trim(request("pwhidden"))
saveUser = trim(request("saveUser"))
cleancookies = trim(request("cleancookies"))

if cleancookies = "true" then
	SetHttpOnlyCookie "accounts", "", -1
	showaccounts = ""
else
	showaccounts = trim(request.Cookies("accounts"))
end if

SetHttpOnlyCookie "name", "", -1
Session("changepw") = ""

dim ei
dim errmsg
errmsg = trim(request("errstr"))


if IsEmpty(Application("em_MaxMPOP3")) and IsEmpty(Application("em_MaxSigns")) then
	dim mam
	set mam = server.createobject("easymail.AdminManager")

	tmp_num = 0
	do while tmp_num < 30
		mam.LoadExt

		if mam.IsLoadOK = true then
			Exit Do
		end if

		mam.Sleep 500
		tmp_num = tmp_num + 1
	loop

	if mam.IsLoadOK = true then
		Application("em_MaxMPOP3") = mam.MaxMPOP3
		Application("em_MaxSigns") = mam.MaxSigns
		Application("em_SystemAdmin") = LCase(mam.SystemAdmin)
		Application("em_EnableBBS") = mam.EnableBBS
		Application("em_Enable_SignHold") = mam.Enable_SignHold
		Application("em_Enable_FreeSign") = mam.Enable_FreeSign
		Application("em_Enable_SignWithDomainUser") = mam.Enable_SignWithDomainUser
		Application("em_Enable_SignNumberLimit") = mam.Enable_SignNumberLimit
		Application("em_SignNumberLimitDays") = mam.SignNumberLimitDays
		Application("em_Enable_ShareFolder") = mam.Enable_ShareFolder
		Application("em_Enable_SignEnglishName") = mam.Enable_SignEnglishName
		Application("em_LogPageKSize") = mam.LogPageKSize
		Application("em_TestAccounts") = LCase(mam.TestAccounts)
		Application("em_SignMode") = mam.SignMode
		Application("em_SignWaitDays") = mam.SignWaitDays
		Application("em_am_Name") = mam.am_Name
		Application("em_am_Accounts") = LCase(mam.am_Accounts)
		Application("em_AccountsAdmin") = LCase(mam.AccountsAdmin)
		Application("em_EnableEntAddress") = mam.Enable_Show_EntAddress
		Application("em_SpamAdmin") = LCase(mam.SpamAdmin)

		Application("em_EnableTrap") = mam.EnableTrap
		if mam.EnableTrap = true then
			Application("em_TrapMail") = mam.TrapMail
		end if

		set mam = nothing
	else
		set mam = nothing
		Response.Redirect "outerr.asp?errstr=" & Server.URLEncode("超时, 请重试") & "&" & getGRSN()
	end if
end if


if un <> "" and pw <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if Application("em_EnableVerification") = true then
		if trim(request.Cookies("zatt_checkcode")) <> trim(request("zck")) then
			Response.Redirect "outerr.asp?errstr=" & Server.URLEncode("验证码填写错误") & "&" & getGRSN()
		end if
	end if

	un = LCase(un)
	pw = strDecode(pw, trim(request("picnum")))

	if un <> Application("em_SystemAdmin") then
		dim webkill
		set webkill = server.createobject("easymail.WebKill")
		webkill.Load

		rip = Request.ServerVariables("REMOTE_ADDR")

		if webkill.IsKill(rip) = true then
			set webkill = nothing
			Response.Redirect "outerr.asp?errstr=" & Server.URLEncode("拒绝IP地址 " & rip & " 访问") & "&" & getGRSN()
		end if

		set webkill = nothing
	end if


	set ei = Application("em")
	Session("wem") = ""
	Session("mail") = ""
	Session("tid") = ""
	Session("SecEx") = ""
	Session("scpw") = ""
	Session("cert_ca") = ""
	Session("EnableSession") = ""
	Session("ReadOnlyUser") = 0


	dim tmp_un
	tmp_un = ei.GetRealUser(un)
	if IsNull(tmp_un) = false and Len(tmp_un) > 0 then
		un = LCase(tmp_un)
	end if

	rip = Request.ServerVariables("REMOTE_ADDR")
	if ei.CheckIPLimit(un, rip) = false then
		set ei = nothing
		Response.Redirect "outerr.asp?errstr=" & Server.URLEncode("拒绝IP地址 " & rip & " 访问") & "&" & getGRSN()
	end if

	dim checkret
	checkret = ei.CheckPassWordEx(un, pw, Request.ServerVariables("REMOTE_ADDR"))

	if checkret = 0 then
		if un = Application("em_SystemAdmin") and ei.CheckAdminIP(Request.ServerVariables("REMOTE_ADDR")) = false then
			set ei = nothing

			errmsg = "管理员登录IP地址错误。"
		else
			Session("tid") = ei.LoginEx(un, Request.ServerVariables("REMOTE_ADDR"))
			Session("wem") = un
			Session("mail") = ei.GetUserMail(un)
			set ei = nothing

			dim mri
			set mri = server.createobject("easymail.MoreRegInfo")
			mri.LoadRegInfo un
			mri.CurrentlyIP = Request.ServerVariables("REMOTE_ADDR")
			mri.SaveRegInfo
			set mri = nothing


			if saveUser = "true" then
				SetHttpOnlyCookie "accounts", un, 365
			end if

			SecEx = trim(request("SecEx"))
			if SecEx = "true" then
				Session("SecEx") = "1"
			else
				Session("SecEx") = "0"
			end if


			dim userweb
			set userweb = server.createobject("easymail.UserWeb")
			userweb.Load Session("wem")

			ShowLanguage = userweb.ShowLanguage

			set userweb = nothing

			dim ul
			set ul = server.createobject("easymail.UserLog")
			ul.Load Session("wem")
			ul.Add 1, Request.ServerVariables("REMOTE_ADDR")
			ul.Save
			set ul = nothing

			if ShowLanguage = 1 then
				Response.Redirect "en/welcome.asp"
			else
				Response.Redirect "welcome.asp"
			end if
		end if
	elseif checkret = 2 then
		dim pwwt
		pwwt = ei.PassWordWaitMinute
		set ei = nothing

		errmsg = "连续三次输入密码错误，请过" & pwwt & "分钟后再试。"
	else
		set ei = nothing

		errmsg = "错误的用户名或密码！请再次输入。"
	end if
end if

if trim(request("logout")) = "true" then
	if Session("wem") <> "" then
		Application("em").Logout Session("wem"), Session("tid")
	end if

	Session("wem") = ""
	Session("mail") = ""
	Session("tid") = ""
	Session("SecEx") = ""
	Session("scpw") = ""
	Session("cert_ca") = ""
	Session("EnableSession") = ""
	Session("ReadOnlyUser") = 0
end if
%>

<!DOCTYPE html>
<HTML>
<head>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<title>WinWebMail邮件系统</title>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<STYLE type=text/css>
<!--
a:hover {color:#b30008; text-decoration:underline;}
a		{color:#004276; text-decoration:none;}
.u_line:hover {color:#b30008; text-decoration:underline;}
.u_line	{color:#004276; text-decoration:underline;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.tdtext {padding-left:62px; padding-bottom:4px; _padding-bottom:2px;}
-->
</STYLE>
</head>

<script type="text/javascript">
<!--
if (top.location !== self.location) {
top.location=self.location;
}

function window_onload() {
<%
if errmsg <> "" then
%>
	alert("<%=server.htmlencode(errmsg) %>");
<%
end if

if showaccounts = "" then
%>
	var S = document.getElementById("usernameshow");
	S.focus();
<%
else
%>
	var S = document.getElementById("pwshow");
	S.focus();
<%
end if
%>
}

function gook() {
	var S;
<%
if showaccounts = "" then
%>
	S = document.getElementById("usernameshow");
	if (S.value == "")
	{
		alert("用户名不可为空");
		S.focus();
		return ;
	}
<%
end if
%>
	S = document.getElementById("pwshow");
	if (S.value == "")
	{
		alert("密码不可为空");
		S.focus();
		return ;
	}
<%
if Application("em_EnableVerification") = true then
%>
	S = document.getElementById("zck_out");
	if (S.value.length < 1)
	{
		S.focus();
		alert("验证码填写错误");
		return ;
	}
	else
		document.getElementById("zck").value = S.value;
<%
end if

if showaccounts = "" then
%>
	document.f1.username.value = document.getElementById("usernameshow").value;
	document.f1.saveUser.value = document.getElementById("showsaveUser").checked;
<%
else
%>
	document.f1.username.value = "<%=showaccounts %>";
<%
end if
%>
	S = document.getElementById("showSecEx");
	document.f1.SecEx.value = S.checked;

	S = document.getElementById("pwshow");
	document.f1.pwhidden.value = encode(S.value, parseInt(document.f1.picnum.value));

	document.f1.submit();
}

function encode(datastr, bassnum) {
	var tempstr;
	var tchar;
	var newdata = "";

	for (var i = 0; i < datastr.length; i++)
	{
		tchar = 65535 + bassnum - datastr.charCodeAt(i);
		tchar = tchar.toString();

		while(tchar.length < 5)
		{
			tchar = "0" + tchar;
		}

		newdata = newdata + tchar;
	}

	return newdata;
}
//-->
</script>

<body LANGUAGE=javascript onload="return window_onload()" style="margin-top:60px;">
<form name="f1" method="post" action="default.asp">
<input type="hidden" name="username">
<input type="hidden" name="pwhidden">
<input type="hidden" name="picnum" value="<%=createRnd() %>">
<input type="hidden" name="saveUser">
<input type="hidden" name="SecEx">
<input type="hidden" name="zck" id="zck">
</form>
<table cellspacing="0" cellpadding="0" width="420" align="center" border="0" bgcolor="white">
	<tr>
	<td align="center" rowspan="2">
	<table cellspacing="0" cellpadding="0" width="100%" border="0" style="border:#999999 1px solid;">
		<tr align="middle" bgcolor="#f3f3f3"><td height="35" style="border-bottom:#b0b0b0 1px solid; font-size:16px;color:#6699cc;"><b>
		欢迎使用WinWebMail邮件系统
		</b></td></tr>
		<tr><td class="block_top_td" style="height:20px;"></td></tr>
		<tr><td align="left" class="tdtext">用户名</td></tr>
		<tr><td align="left" style="padding-left:55px;">
<%
if showaccounts = "" then
%>
<input type="text" id="usernameshow" name="usernameshow" maxlength="64" size="24" class="usernameshow">
<%
else
%>
<input type="text" id="usernameshow" name="usernameshow" maxlength="64" size="24" class="usernameshow" value="<%=showaccounts %>">
<%
end if
%>
		</td></tr>

		<tr><td class="block_top_td" style="height:16px;"></td></tr>
		<tr><td align="left" class="tdtext">密码</td></tr>
		<tr><td align="left" style="padding-left:55px;">
		<input type="password" id="pwshow" name="pwshow" maxlength="32" size="24" class="pwshow">
		</td></tr>

<%
if Application("em_EnableVerification") = true then
%>
		<tr><td class="block_top_td" style="height:16px;"></td></tr>
		<tr><td nowrap align="left" class="tdtext"><img src="tu.asp" align="absmiddle" border="0">
		<input type="text" id="zck_out" class="n_textbox" size="2" maxlength="2">
		</td></tr>
<%
end if

if showaccounts = "" then
%>
		<tr><td class="block_top_td" style="height:12px;"></td></tr>
		<tr><td nowrap align="left" style="padding-left:60px;"><input type="checkbox" id="showSecEx" name="showSecEx">增强安全性&nbsp;&nbsp;&nbsp;
		<input type="checkbox" id="showsaveUser" name="showsaveUser">记住用户名
		</td></tr>
<%
else
%>
		<tr><td class="block_top_td" style="height:12px;"></td></tr>
		<tr><td nowrap align="left" style="padding-left:60px;"><input type="checkbox" id="showSecEx" name="showSecEx">增强安全性&nbsp;&nbsp;&nbsp;
		<a href="default.asp?cleancookies=true">改用其他身份登录</a></font>
		</td></tr>
<%
end if
%>
		<tr><td nowrap align="right" height="50" style="padding-right:18px; _padding-right:36px; padding-bottom:12px; _padding-bottom:6px;">
		<a class='wwm_btnDownload btn_gray' href="javascript:gook();">&nbsp;确 定&nbsp;</a><input type="submit" value="" onclick="javascript:gook();" style="filter:alpha(opacity=0); opacity:0; font-size:0pt; height:0px; width:0px; border:0px;">
		</td></tr>
	</table>

	</td>
	<td width=1 bgcolor=#ffffff height=5></td>
	<td width=1 bgcolor=#ffffff height=5></td>
	<td width=1 bgcolor=#ffffff height=5></td>
	</tr>
	<tr>
	<td width=1 bgcolor=#666666 height=120></td>
	<td width=1 bgcolor=#999999 height=120></td>
	<td width=1 bgcolor=#cccccc height=120></td>
	</tr>
	<tr valign=top align=right> 
	<td colspan=4> 
		<table cellspacing=0 cellpadding=0 width="417" border=0>
		<tr><td bgcolor=#666666 height=1></td></tr>
        <tr><td bgcolor=#999999 height=1></td></tr>
        <tr><td bgcolor=#cccccc height=1></td></tr>
		</table>
	</td></tr>
</table>
<table width="100%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr><td class="block_top_td" style="height:50px;"></td></tr>
	<tr><td align="center" nowrap>
<%
if Application("em_Enable_FreeSign") = true then
%>
[<a href="create.asp?<%=getGRSN() %>">申请邮箱</a>]&nbsp;&nbsp;
<%
end if
%>
[<a href="forgetbf.asp?<%=getGRSN() %>">忘记密码</a>]
	</td></tr>
	<tr><td height="15"></td></tr>
	<tr><td align="center" nowrap height="25">
	<a href="http://www.winwebmail.com" class="u_line" target="_blank">WinWebMail Server 邮件服务器</a>
	</td></tr>
	<tr><td align="center" nowrap>
	<a href="mailto:webeasymail@51webmail.com" class="u_line">版权所有:&nbsp;马坚</a>
	</td></tr>
</table>
<%
if Application("em_EnableTrap") = true then
%>
<div style="position:absolute; top:0; left:0; z-index:0; visibility:hidden">
<a href="mailto:<%=Application("em_TrapMail") %>"><%=Application("em_TrapMail") %></a>
</div>
<%
end if
%>
</body>
</html>

<%
function getGRSN()
	dim theGRSN
	Randomize
	theGRSN = Int((9999999 * Rnd) + 1)

	getGRSN = "GRSN=" & CStr(theGRSN)
end function

function createRnd()
	dim retval
	retval = getGRSN()

	if Len(retval) > 4 then
		retval = Right(retval, 4)
	end if

	if Left(retval, 1) = "0" then
		retval = "5" & Right(retval, 3)
	end if

	createRnd = retval
end function

function strDecode(sd_Data, sd_bassnum)
	dim sd_vChar
	dim sd_NewData
	dim sd_TempChar
	sd_vChar = 1

	do
		if sd_vChar > Len(sd_Data) then
			exit do
		end if

	    sd_TempChar = CLng(Mid(sd_Data, sd_vChar, 5))
		sd_TempChar = ChrW(65535 + sd_bassnum - sd_TempChar)

        sd_NewData = sd_NewData & sd_TempChar
		sd_vChar = sd_vChar + 5
	loop

	strDecode = sd_NewData
end function

function SetHttpOnlyCookie(cookieName, cookieValue, days)
	Dim cookie
	cookie=cookieName & "=" & Server.URLEncode(cookieValue) & "; path=/"
	cookie=cookie & "; expires=" & CStr(DateAdd("d", days, now()))
	cookie=cookie & "; domain=; HttpOnly"
	Response.AddHeader "Set-Cookie", cookie
end function
%>
