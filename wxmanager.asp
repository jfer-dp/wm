<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" -->

<%
if isadmin() = false then
	Response.Redirect "noadmin.asp"
end if
%>

<%
dim wx
set wx = server.createobject("easymail.WXSet")
wx.Load

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if trim(request("is_enabled")) <> "" then
		wx.is_enabled = true
	else
		wx.is_enabled = false
	end if

	wx.url_get_token = trim(request("url_get_token"))
	wx.app_id = trim(request("app_id"))
	wx.app_secret = trim(request("app_secret"))
	wx.url_template_send = trim(request("url_template_send"))
	wx.url_goto_mail = trim(request("url_goto_mail"))
	wx.url_mail_default = trim(request("url_mail_default"))
	wx.template_1_id = trim(request("template_1_id"))
	wx.template_2_id = trim(request("template_2_id"))

	wx.Save
	set wx = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("wxmanager.asp")
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
html {overflow: scroll; overflow-x: hidden; overflow-y: auto !important;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.td_line_l {white-space:nowrap; text-align:left; border-bottom:1px #d9d6c3 solid; color:#202020; padding:4px 0px 2px 6px;}
</STYLE>
</HEAD>

<script type="text/javascript">
function wximgerr() {
	document.getElementById("wx").style.display = "none";
}

function gosub() {
	var str = document.getElementById("url_mail_default").value;
	if (str.substr(0, 7).toLowerCase() == "http://")
		document.getElementById("url_mail_default").value = str.substr(7);

	if (str.substr(0, 8).toLowerCase() == "https://")
		document.getElementById("url_mail_default").value = str.substr(8);

	str = document.getElementById("url_goto_mail").value;
	if (str.substr(0, 7).toLowerCase() == "http://")
		document.getElementById("url_goto_mail").value = str.substr(7);

	if (str.substr(0, 8).toLowerCase() == "https://")
		document.getElementById("url_goto_mail").value = str.substr(8);

	document.f1.submit();
}
</script>

<body>
<form method="post" action="#" name="f1">
<table width="92%" border="0" align="center" cellspacing="0">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_393 %>
</td></tr>
<tr><td class="block_top_td" style="height:12px; _height:14px;">
<div id="wx" style="text-align:left; padding-left:16px;">
<table width="60%" border="0" align="left" cellspacing="0">
<tr>
<td width="5%" style="padding:6px 2px 6px 2px;"><img src="wx.jpg" align='absmiddle' border="0" onerror="wximgerr();"></td>
<td nowrap align="left" style="padding-left:20px;"><%=b_lang_392 %></td>
</tr>
</table>
</div>
</td></tr>
</table>

<table width="88%" border="0" align="center" cellspacing="0">
	<tr><td width="86%" class="td_line_l"><input type="checkbox" name="is_enabled" value="checkbox" <% if wx.is_enabled = true then Response.Write "checked"%>>
	<%=b_lang_394 %>&nbsp;&nbsp;&nbsp;[<a href="http://www.winwebmail.com/article-wx.html" target="_blank"><%=b_lang_395 %></a>]
	</td></tr>
	<tr><td class="td_line_l">
	<%=b_lang_396 %>&nbsp;<input type="text" name="url_get_token" class='n_textbox' size="100" value="<%=wx.url_get_token %>">
	</td></tr>
	<tr><td class="td_line_l">
	<%=b_lang_397 %>&nbsp;<input type="text" name="app_id" class='n_textbox' size="60" value="<%=wx.app_id %>">
	</td></tr>
	<tr><td class="td_line_l">
	<%=b_lang_398 %>&nbsp;<input type="text" name="app_secret" class='n_textbox' size="60" value="<%=wx.app_secret %>">
	</td></tr>
	<tr><td class="td_line_l">
	<%=b_lang_399 %>&nbsp;<input type="text" name="url_template_send" class='n_textbox' size="100" value="<%=wx.url_template_send %>">
	</td></tr>
	<tr><td class="td_line_l">
	<%=b_lang_400 %>&nbsp;<input type="text" id="url_mail_default" name="url_mail_default" class='n_textbox' size="100" value="<%=wx.url_mail_default %>">
	</td></tr>
	<tr><td class="td_line_l">
	<%=b_lang_401 %>&nbsp;<input type="text" id="url_goto_mail" name="url_goto_mail" class='n_textbox' size="100" value="<%=wx.url_goto_mail %>">
	</td></tr>
	<tr><td class="td_line_l">
	<%=b_lang_402 %>&nbsp;<input type="text" name="template_1_id" class='n_textbox' size="80" value="<%=wx.template_1_id %>">
	</td></tr>
	<tr><td class="td_line_l">
	<%=b_lang_403 %>&nbsp;<input type="text" name="template_2_id" class='n_textbox' size="80" value="<%=wx.template_2_id %>">
	</td></tr>
</table>

<table width="92%" border="0" align="center" cellspacing="0">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">&nbsp;</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="right">
<a class='wwm_btnDownload btn_blue' href="right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
&nbsp;&nbsp;<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
</td></tr>
</table>
<br>
</form>
</body>
</HTML>

<%
set wx = nothing
%>
