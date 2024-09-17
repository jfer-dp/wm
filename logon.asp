<%
Response.CacheControl = "no-cache" 
%>

<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
pw = trim(request("pwhidden"))
oldpw = trim(request("oldpwhidden"))
isquestion = trim(request("isquestion"))
isiplimit = trim(request("isiplimit"))

dim ei
set ei = server.createobject("easymail.UserWeb")
ei.Load Session("wem")

Set em = Application("em")

dim errstr

if isiplimit = "1" then
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		isok = false

		if Session("wem") <> Application("em_TestAccounts") then
			if ei.QuestionInfo <> "" then
				if trim(request("EnableIPLimit")) <> "" then
					em.SetEnableIPLimit Session("wem"), true
				else
					em.SetEnableIPLimit Session("wem"), false
				end if

				em.SetIPLimit Session("wem"), trim(request("IPAddress"))

				isok = true
			end if
		else
			errstr = a_lang_272
		end if

		set em = nothing
		set ei = nothing

		if isok = true then
			Response.Redirect "ok.asp?gourl=" & Server.URLEncode("logon.asp") & "&" & getGRSN()
		else
			Response.Redirect "err.asp?errstr=" & Server.URLEncode(errstr) & "&" & getGRSN()
		end if
	end if
end if

if isquestion = "1" then
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		isok = false

		if em.CheckPassWord(Session("wem"), trim(request("qioldpw"))) = true then
			QuestionInfo = trim(request("QuestionInfo"))
			QuestionInfo = replace(QuestionInfo, """", "'")
			ei.QuestionInfo = QuestionInfo

			ei.AnswerInfo = trim(request("AnswerInfo"))

			HintInfo = trim(request("HintInfo"))
			HintInfo = replace(HintInfo, """", "'")
			ei.HintInfo = HintInfo

			ei.PasswordIsOK = true
			ei.Save
			isok = true
		end if

		set em = nothing
		set ei = nothing

		if isok = true then
			Response.Redirect "ok.asp?gourl=" & Server.URLEncode("logon.asp") & "&" & getGRSN()
		else
			Response.Redirect "err.asp?gourl=" & Server.URLEncode("logon.asp") & "&errstr=" & Server.URLEncode(a_lang_273) & "&" & getGRSN()
		end if
	end if
else
	if pw <> "" and (oldpw <> "" or Session("changepw") = Session("wem")) and Request.ServerVariables("REQUEST_METHOD") = "POST" then
		pw = strDecode(pw, trim(request("picnum")))
		oldpw = strDecode(oldpw, trim(request("picnum")))

		if Session("wem") <> Application("em_TestAccounts") then
			isok = false

			if Session("changepw") <> Session("wem") then
				if em.CheckPassWord(Session("wem"), oldpw) = true then
					em.ChangeUserPassWord Session("wem"), pw
					isok = true
				end if
			else
				em.ChangeUserPassWord Session("wem"), pw
				isok = true
			end if

			if isok = true then
				dim ul
				set ul = server.createobject("easymail.UserLog")
				ul.Load Session("wem")
				ul.Add 2, Request.ServerVariables("REMOTE_ADDR")
				ul.Save
				set ul = nothing
			end if

			set em = nothing
			set ei = nothing

			if Session("changepw") <> Session("wem") then
				if isok = true then
					Response.Redirect "ok.asp?" & getGRSN()
				else
					Response.Redirect "err.asp?errstr=" & Server.URLEncode(a_lang_274) & "&" & getGRSN()
				end if
			else
				Session("changepw") = ""
				Response.Redirect "ok.asp?gourl=" & Server.URLEncode("welcome.asp") & "&" & getGRSN()
			end if
		else
			set ei = nothing

			if Session("changepw") <> Session("wem") then
				Response.Redirect "err.asp?errstr=" & Server.URLEncode(a_lang_275) & "&" & getGRSN()
			else
				Response.Redirect "err.asp?gourl=" & Server.URLEncode("welcome.asp") & "&errstr=" & Server.URLEncode(a_lang_275) & "&" & getGRSN()
			end if
		end if
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

<STYLE type=text/css>
<!--
html {overflow: scroll; overflow-x: hidden; overflow-y: auto !important;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.cont_td {white-space:nowrap; height:26px; padding-left:5px; padding-right:5px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function checkpw(){
	if (now_rt < 2)
	{
		alert("<%=a_lang_276 %>");
		return ;
	}

	if (document.f1.pw1.value != "" && document.f1.pw2.value != "")
	{
		if (document.f1.pw1.value != document.f1.pw2.value)
			alert("<%=a_lang_019 %>");
		else
		{
<%
if Session("changepw") <> Session("wem") then
%>
			if (document.f1.oldpw.value.length == 0)
			{
				alert("<%=a_lang_277 %>");
				document.f1.oldpw.focus();
				return ;
			}

			document.form1.oldpwhidden.value = encode(document.f1.oldpw.value, parseInt(document.form1.picnum.value));
<%
end if
%>
			document.form1.pwhidden.value = encode(document.f1.pw1.value, parseInt(document.form1.picnum.value));
			document.form1.submit();
		}
	}
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

function gosub() {
	if (document.form2.qioldpw.value.length == 0)
	{
		alert("<%=a_lang_278 %>");
		document.form2.qioldpw.focus();
		return ;
	}

	if (document.form2.QuestionInfo.value == "")
	{
		alert("<%=a_lang_279 %>");
		document.form2.QuestionInfo.focus();
		return ;
	}

	if (document.form2.AnswerInfo.value == "" || document.form2.AnswerInfo.value == "******")
	{
		alert("<%=a_lang_280 %>");
		document.form2.AnswerInfo.focus();
		return ;
	}

	if (document.f1.oldpw.value.length > 0 || document.f1.pw1.value.length > 0 || document.f1.pw2.value.length > 0)
		alert("<%=a_lang_281 %>");

	document.form2.submit();
}

function isValidIP(ip) {
	var reg = /^(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])$/
	return reg.test(ip);
}

function goipsub() {
	if (document.getElementById("EnableIPLimit").checked == true)
	{
		if (isValidIP(document.getElementById("IPAddress").value) == false)
		{
			alert("<%=a_lang_350 %>");
			document.getElementById("IPAddress").focus();
			return ;
		}
	}
	else
	{
		document.getElementById("IPAddress").value = "";
	}

<%
if ei.QuestionInfo <> "" then
%>
	document.form3.submit();
<%
end if
%>
}

	ratingMsgs = new Array(4);
	ratingMsgColors = new Array(4);
	barColors = new Array(4);
	ratingMsgs[0] = "<%=a_lang_282 %>";
	ratingMsgs[1] = "<%=a_lang_283 %>";
	ratingMsgs[2] = "<%=a_lang_284 %>";
	ratingMsgs[3] = "<%=a_lang_285 %>";
	ratingMsgColors[0] = "#676767";
	ratingMsgColors[1] = "#aa0033";
	ratingMsgColors[2] = "#f5ac00";
	ratingMsgColors[3] = "#008000";
	barColors[0] = "#dddddd";
	barColors[1] = "#aa0033";
	barColors[2] = "#ffcc33";
	barColors[3] = "#008000";
	var now_rt = 0;

function CreateRatePasswdReq(pwd){
	if (!isBrowserCompatible) return;

	if(!pwd) return; 
	passwd=pwd.value;
	var min_passwd_len = 6; 
	if (passwd.length < min_passwd_len){
		if (passwd.length > 0)
			DrawBar(0);
		else
			ResetBar();
	} else {
		rating = checkPasswdRate(passwd);
		DrawBar(rating);
	}
}

function getElement(name){
	if (document.all)
		return document.all(name);

	return document.getElementById(name);
}

function DrawBar(rating){
	now_rt = rating;
	var posbar = getElement('posBar');
	var negbar = getElement('negBar');
	var passwdRating = getElement('passwdRating');
	var barLength = getElement('passwdBar').width;
	if (rating >= 0 && rating < 4) {
		posbar.style.width = barLength / 3 * rating + "px";
		negbar.style.width = barLength / 3 * (3 - rating) + "px";
	}

	posbar.style.background = barColors[rating];
 	passwdRating.innerHTML = "<font color='" + ratingMsgColors[rating] + "'>" + ratingMsgs[rating] + "</font>";
}

function ResetBar(){
	var posbar = getElement('posBar');
	var negbar = getElement('negBar');
	var passwdRating = getElement('passwdRating');
	var barLength = getElement('passwdBar').width;
	posbar.style.width = "0px";
	negbar.style.width = barLength + "px";
	passwdRating.innerHTML = "";
}

	var agt = navigator.userAgent.toLowerCase();
	var is_op = (agt.indexOf("opera") != -1);
	var is_ie = (agt.indexOf("msie") != -1) && document.all && !is_op;
	var is_mac = (agt.indexOf("mac") != -1);
	var is_gk = (agt.indexOf("gecko") != -1);
	var is_sf = (agt.indexOf("safari") != -1);

function gff(str, pfx){
	var i = str.indexOf(pfx);
	if (i != -1)
	{
		var v = parseFloat(str.substring(i + pfx.length));
		if (!isNaN(v))
			return v;
	}
	return null;
}

function Compatible(){
	if (is_ie && !is_op && !is_mac)
	{
		var v = gff(agt, "msie ");
		if (v != null)
			return (v >= 6.0);
	}

	if (is_gk && !is_sf)
	{
		var v = gff(agt, "rv:");
		if (v != null)
			return (v >= 1.4);
		else
		{
			v = gff(agt, "galeon/");
			if (v != null)
				return (v >= 1.3);
		}
	}

	if (is_sf)
	{
		var v = gff(agt, "applewebkit/");
		if (v != null)
			return (v >= 124);
	}
	return false;
}

var isBrowserCompatible = Compatible();

function CharMode(iN){
	if (iN>=48 && iN <=57)
		return 1;
	if ((iN>=65 && iN <=90) || (iN>=97 && iN <=122))
		return 2;
	else
		return 4;
}

function bitTotal(num){
	modes=0;
	for (i=0;i<4;i++)
	{
		if (num & 1) modes++;
			num>>>=1;
	}
	return modes; 
}

function checkPasswdRate(sPW){
	if (sPW.length < 6)
		return 0;

	Modes=0;
	for (i=0;i<sPW.length;i++){
		Modes|=CharMode(sPW.charCodeAt(i));
	}

	return bitTotal(Modes);
}
//-->
</script>

<body>
<form name="form1" method="post" action="logon.asp">
<input type="hidden" name="forget" value="<%=trim(request("forget")) %>">
<input type="hidden" name="oldpwhidden">
<input type="hidden" name="pwhidden">
<input type="hidden" name="picnum" value="<%=createRnd() %>">
</form>
<form name="f1">
<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td colspan="2" class="block_top_td" style="height:4px;"></td></tr>
<tr><td colspan="2" style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_286 %>
</td></tr>
<tr><td colspan="2" class="block_top_td" style="height:10px; _height:12px;"></td></tr>

<%
if Session("changepw") <> Session("wem") then
%>
	<tr><td width="11%" align="right" class="cont_td">
	<%=a_lang_021 %><%=s_lang_mh %>
	</td><td align="left">
	<input type="password" name="oldpw" maxlength="30" size="30" autocomplete="off" class="n_textbox">&nbsp;*
	</td></tr>
<%
end if
%>
	<tr><td width="11%" align="right" class="cont_td">
	<%=a_lang_022 %><%=s_lang_mh %>
	</td><td align="left">
	<table width="10" cellSpacing=0 cellPadding=0 border=0>
	<tr><td noWrap>
	<input type="password" name="pw1" maxlength="30" size="30" autocomplete="off" class="n_textbox" onkeyup="CreateRatePasswdReq(this)">&nbsp;*
	</td>
	<td>
	<table width="10" cellSpacing=0 cellPadding=0 border=0>
	<tr>
	<td vAlign=top noWrap width="0">&nbsp;<font color="#444444"><%=a_lang_287 %><%=s_lang_mh %> </font></td>
	<td vAlign=top noWrap><font color="#808080">
	<strong><div id=passwdRating></div></strong></font>
	</td></tr>
	<tr><td colspan="2">
		<table cellSpacing=0 cellPadding=0 border=0><tr>
		<td width="50">&nbsp;</td>
		<td>
			<table id=passwdBar cellSpacing=0 cellPadding=0 width="150" bgColor=#ffffff border=0>
			<tr>
			<td id=posBar width=0% bgColor=#e0e0e0 height=4></td>
			<td id=negBar width="100%" bgColor=#e0e0e0 height=4></td>
			</tr>
			</table>
		</td></tr>
		</table>
	</td></tr>
	</table>
	</td></tr>
	</table>
	</td></tr>

	<tr><td align="right" class="cont_td">
	<%=a_lang_023 %><%=s_lang_mh %>
	</td><td align="left">
	<input type="password" name="pw2" maxlength="30" size="30" autocomplete="off" class="n_textbox">&nbsp;*
	</td></tr>

	<tr><td colspan=2 class="block_top_td" style="height:8px;"></td></tr>

	<tr><td colspan="2" align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:checkpw();"><%=s_lang_save %></a>
	</td></tr>
</table>
</form>
<br>

<%
if Session("changepw") <> Session("wem") then
%>
<form name="form2" method="post" action="logon.asp">
<input type="hidden" name="isquestion" value="1">
<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td colspan="2" class="block_top_td" style="height:4px;"></td></tr>
<tr><td colspan="2" style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_288 %>
</td></tr>
<tr><td colspan="2" class="block_top_td" style="height:10px; _height:12px;"></td></tr>

	<tr><td width="36%" align="right" class="cont_td">
	<%=a_lang_289 %><%=s_lang_mh %>
	</td><td align="left">
	<input type="password" name="qioldpw" maxlength="32" autocomplete="off" class="n_textbox">&nbsp;*
	</td></tr>

	<tr><td align="right" class="cont_td">
	<%=a_lang_290 %> <font color="#444444">(<%=a_lang_291 %>)</font><%=s_lang_mh %>
	</td><td align="left">
	<input type="text" name="QuestionInfo" class='n_textbox' value="<%=ei.QuestionInfo %>" maxlength="256">&nbsp;*
	</td></tr>

	<tr><td align="right" class="cont_td">
	<%=a_lang_292 %> <font color="#444444">(<%=a_lang_293 %>)</font><%=s_lang_mh %>
	</td><td align="left">
	<input type="text" name="AnswerInfo" class='n_textbox' value="<%
if ei.QuestionInfo <> "" then
	Response.Write "******"
end if
%>" maxlength="256">&nbsp;*
	</td></tr>

	<tr><td align="right" class="cont_td">
	<%=a_lang_294 %> <font color="#444444">(<%=a_lang_295 %>)</font><%=s_lang_mh %>
	</td><td align="left">
	<input type="text" name="HintInfo" class='n_textbox' value="<%=ei.HintInfo %>" maxlength="256">
	</td></tr>

	<tr><td colspan=2 class="block_top_td" style="height:8px;"></td></tr>

	<tr><td colspan="2" align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
	</td></tr>
</table>
</form>
<br>

<form name="form3" method="post" action="logon.asp">
<input type="hidden" name="isiplimit" value="1">
<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=s_lang_0104 %>
</td></tr>
<tr><td class="block_top_td" style="height:10px; _height:12px;"></td></tr>

<tr><td class="cont_td">
&nbsp;<input type="checkbox" name="EnableIPLimit" id="EnableIPLimit" <%
em.GetIPLimit Session("wem"), EnableIPLimit, IPAddress

if EnableIPLimit = true then
	Response.Write "checked"
end if

if ei.QuestionInfo = "" then
	Response.Write " disabled"
end if
%>>&nbsp;<%=s_lang_0105 %>
</td></tr>

<tr><td class="cont_td">
&nbsp;&nbsp;<%=s_lang_0106 %><%=s_lang_mh %>
<input type="text" name="IPAddress" id="IPAddress" class='n_textbox' value="<%
Response.Write IPAddress

EnableIPLimit = NULL
IPAddress = NULL
%>" maxlength="32" size="32"<% if ei.QuestionInfo = "" then Response.Write " disabled" %>>
</td></tr>

<tr><td class="block_top_td" style="height:8px;"></td></tr>
<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<%
if ei.QuestionInfo <> "" then
%>
<a class='wwm_btnDownload btn_blue' href="javascript:goipsub();"><%=s_lang_save %></a>
<%
end if
%>
</td></tr>
</table>
</form>
<%
end if
%>

<div style="position:absolute; left:12px; top:10px;">
<a href="help.asp#style" target="_blank"><img src="images/help.gif" border="0" title="<%=s_lang_help %>"></a></div>
</body>
</html>


<%
set ei = nothing
set em = nothing


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
