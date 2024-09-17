<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
username = trim(request("username"))
foldername = trim(request("foldername"))
s_ia = trim(request("isatt"))

if s_ia = "" or s_ia = "False" then
	isatt = FALSE
else
	isatt = TRUE
end if

dim userweb
set userweb = server.createobject("easymail.UserWeb")
userweb.Load Session("wem")

if trim(request("save")) = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	pw1 = trim(request("pw1"))

	if userweb.HaveThisFriendFolder(username, foldername, isatt) = false then
		if Session("wem") <> username then
			userweb.AddFriendFolder username, foldername, pw1, isatt
			userweb.Save

			set userweb = nothing
			Response.Redirect "ff_showall.asp?" & getGRSN()
		else
			set userweb = nothing
			Response.Redirect "err.asp?gourl=ff_showall.asp?" & getGRSN()
		end if
	else
		Response.Redirect "ff_showall.asp?" & getGRSN()
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
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.cont_td {white-space:nowrap; height:28px; padding-left:5px; padding-right:5px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function gosub() {
	if (document.f1.username.value != "" && document.f1.foldername.value != "")
		document.f1.submit();
}

function select_mode_onchange()
{
	if (document.f1.s_mode.value == "0" || document.f1.foldername.value == "att")
		document.f1.isatt.value = "False";
	else
		document.f1.isatt.value = "True";
}
//-->
</SCRIPT>

<body>
<form name="f1" method="post" action="ff_addsharefolder.asp">
<input type="hidden" name="save" value="1">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_201 %><%
if foldername <> "" then
	if foldername <> "att" then
		Response.Write "&nbsp;(" & server.htmlencode(username) & "\" & server.htmlencode(foldername) & ")"
	else
		Response.Write "&nbsp;(" & server.htmlencode(username) & "\" & a_lang_202 & ")"
	end if
end if
%>
</td></tr>
<tr><td class="block_top_td" style="height:12px; _height:14px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr>
	<td width="10%" align="right" class="cont_td"><%=a_lang_203 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td"><input type="text" <%
if username <> "" and foldername <> "" then
	Response.Write "readonly "
end if
%>name="username" value="<%=username %>" class="n_textbox" size="30" maxlength="64"></td>
	</tr>
	<tr>
	<td align="right" class="cont_td"><%=a_lang_204 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td"><input type="text" <%
if username <> "" and foldername <> "" then
	Response.Write "readonly "
end if
%>name="foldername" value="<%=foldername %>" class="n_textbox" size="30" maxlength="128"></td>
	</tr>
	<tr>
	<td align="right" class="cont_td"><%=a_lang_205 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td"><input type="password" name="pw1" maxlength="32" class="n_textbox" size="30"></td>
	</tr><%
if username = "" or foldername = "" then
%>
	<tr>
	<td align="right" class="cont_td"><%=a_lang_206 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td">
<select name="s_mode" class=drpdwn LANGUAGE=javascript onchange="select_mode_onchange()">
<option value="0"><%=a_lang_207 %></option>
<option value="1"><%=a_lang_208 %></option>
</select>
	</td>
	</tr>
<%
end if
%>
	</td></tr>
</table>
</td></tr>
	<tr><td class="block_top_td" style="height:8px;"></td></tr>
	<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="ff_showall.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
	</td></tr>
</table>
<input type="hidden" name="isatt" value="<%=isatt %>">
</form>
</BODY>
</HTML>

<%
set userweb = nothing
%>
