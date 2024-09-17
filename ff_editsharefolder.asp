<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
username = trim(request("username"))
foldername = trim(request("foldername"))

isatt = "0"
if trim(request("isatt")) = "1" then
	isatt = "1"
end if

dim userweb
set userweb = server.createobject("easymail.UserWeb")
userweb.Load Session("wem")

if trim(request("save")) = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	pw1 = trim(request("pw1"))

	if Session("wem") <> username then
		if isatt = "1" then
			userweb.SetPassword username, foldername, pw1, true
		else
			userweb.SetPassword username, foldername, pw1, false
		end if

		userweb.Save

		set userweb = nothing
		Response.Redirect "ff_showall.asp?" & getGRSN()
	else
		set userweb = nothing
		Response.Redirect "err.asp?gourl=ff_showall.asp&" & getGRSN()
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
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function gosub() {
	document.f1.submit();
}
//-->
</script>

<body>
<form name="f1" method="post" action="ff_editsharefolder.asp">
<input type="hidden" name="username" value="<%=username %>">
<input type="hidden" name="foldername" value="<%=foldername %>">
<input type="hidden" name="isatt" value="<%=isatt %>">
<input type="hidden" name="save" value="1">

<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_209 %>&nbsp;(<%
response.write server.htmlencode(username) & "\"

if foldername <> "att" then
	response.write server.htmlencode(foldername)
else
	response.write a_lang_202
end if
%>)
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr>
	<td nowrap width="10%" align="right" height="16"><%=a_lang_205 %><%=s_lang_mh %></td>
	<td align="left">
	<input type="password" name="pw1" maxlength="32" class="n_textbox" size="30"></td>
	</tr>
</table>
</td></tr>
	<tr><td class="block_top_td" style="height:12px;"></td></tr>
	<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="javascript:history.back();"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
	</td></tr>
</table>
</form>
</BODY>
</HTML>

<%
set userweb = nothing
%>
