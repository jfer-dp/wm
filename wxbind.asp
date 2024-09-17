<%
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>

<%
wx = trim(request("wid"))
un = trim(request("username"))
pw = trim(request("pwhidden"))
picnum = trim(request("picnum"))

is_cf = true

if Request.Cookies("tmp_cf") <> picnum then
	is_cf = false
	Response.Cookies("tmp_cf") = picnum
end if

dim ei
dim errmsg

dim bind_ok
bind_ok = false

if is_cf = false and un <> "" and pw <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	un = LCase(un)
	pw = strDecode(pw, picnum)

	set ei = Application("em")

	dim tmp_un
	tmp_un = ei.GetRealUser(un)
	if IsNull(tmp_un) = false and Len(tmp_un) > 0 then
		un = LCase(tmp_un)
	end if

	if ei.CheckPassWord(un, pw) = true then
		dim wxset
		set wxset = server.createobject("easymail.WXSet")
		wxset.load
		bind_ok = wxset.Bind(un, wx)
		set wxset = nothing
	else
		errmsg = "错误的用户名或密码"
	end if

	set ei = nothing
end if
%>

<!doctype html>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<title>WinWebMail邮件系统 - 微信通知绑定功能</title>
<meta name="viewport" content="width=device-width,initial-scale=1.0,user-scalable=no">
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<style>
html, body {
	margin: 8px 6px 3px 5px;;
}

body {
	color: #4A4A4A ;
	background: #e6e6e6;
}

.wrap {
	margin-bottom: 18px;
	padding: 8px;
	background: #fff;
	border-radius: 2px 2px 2px 2px;
	-webkit-box-shadow: 0 10px 6px -6px #777;
	-moz-box-shadow: 0 10px 6px -6px #777;
	box-shadow: 0 10px 6px -6px #777;
}

.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
</style>
</head>

<script type="text/javascript">
<!--
if (top.location !== self.location) {
top.location=self.location;
}

javascript:window.history.forward(1); 

function window_onload() {
	var S = document.getElementById("usernameshow");
	if (S != null)
		S.focus();
}

function b_close() {
	self.close();
}

function bindother() {
	location.href = "wxbind.asp?wid=<%=Server.URLEncode(wx) %>"
}

function gook() {
	var S;
	S = document.getElementById("usernameshow");
	if (S != null && S.value == "")
	{
		S.focus();
		return ;
	}

	S = document.getElementById("pwshow");
	if (S != null && S.value == "")
	{
		S.focus();
		return ;
	}

	S = document.getElementById("usernameshow");
	document.f1.username.value = S.value;

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

<body LANGUAGE=javascript onload="return window_onload()" style="margin-top:20px;">
<div class="wrap">
<div style="padding-top:30px; min-height:60px; text-align:center;">
<font style="color:#5fa207; font-weight:bold; font-size:14px;"><%
if Len(wx) < 10 then
	Response.Write "链接不完整，请重试"
else
	if bind_ok = true then
		Response.Write "绑定帐号" & un & "成功"
	else
		if errmsg <> "" then
			Response.Write server.htmlencode(errmsg)
		else
			Response.Write "WinWebMail邮件系统 - 微信通知绑定功能"
		end if
	end if
end if
%></font>
</div>
</div>

<form name="f1" method="post" action="#">
<input type="hidden" name="wx" value="<%=wx %>">
<input type="hidden" name="username">
<input type="hidden" name="pwhidden">
<input type="hidden" name="picnum" value="<%=createRnd() %>">
</form>

<div style="padding-top:25px; text-align:center;">
<table cellspacing="0" cellpadding="0" width="20%" align="center" border="0">
	<tr>
	<td align="center" nowrap>
<%
if Len(wx) > 9 then
	if bind_ok = true then
%>
<a class='wwm_btnDownload btn_gray' href="javascript:bindother();" style="font-weight:bold;">&nbsp;继续绑定其他帐号&nbsp;</a><input type="button" value="" onclick="javascript:gook();" style="filter:alpha(opacity=0); opacity:0; font-size:0pt; height:0px; width:0px; border:0px;">
<%
	else
%>
	<tr><td align="left" nowrap style="font-size:14px; padding-bottom:1px; padding-left:2px;">
需要绑定的用户帐号
	</td></tr>
	<tr><td align="left" style="padding-bottom:8px;">
<input type="text" id="usernameshow" name="usernameshow" maxlength="64" size="18" class="usernameshow">
	</td></tr>
	<tr><td align="left" style="font-size:14px; padding-bottom:1px; padding-left:2px;">
密码
	</td></tr>
	<tr><td align="left">
<input type="password" id="pwshow" name="pwshow" maxlength="32" size="18" class="pwshow">
	</td></tr>
	<tr><td align="center" nowrap style="padding-top:20px;">
&nbsp;&nbsp;&nbsp;<a class='wwm_btnDownload btn_gray' href="javascript:gook();" style="font-weight:bold;">&nbsp;确 定&nbsp;</a><input type="submit" value="" onclick="javascript:gook();" style="filter:alpha(opacity=0); opacity:0; font-size:0pt; height:0px; width:0px; border:0px;">
	</td></tr>
<%
	end if
end if
%>
	</td>
	</tr>
</table>
</div>
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
%>
