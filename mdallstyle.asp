<%
Response.CacheControl = "no-cache" 
%>

<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" -->

<%
if isadmin() = false then
	Response.Redirect "noadmin.asp"
end if
%>

<%
if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim ei
	set ei = server.createobject("easymail.UserWeb")

	dim euser
	set euser = Application("em")
	allnum = euser.GetUsersCount
	i = 0

	do while i < allnum
		euser.GetUserByIndex i, name, domain, comment

		ei.Load name

		if trim(request("EnableBackupAllSendMail")) = "" then
			ei.EnableBackupAllSendMail = false
		else
			ei.EnableBackupAllSendMail = true
		end if

		if trim(request("EnableClearWhenFull")) = "" then
			ei.EnableClearWhenFull = false
		else
			ei.EnableClearWhenFull = true
		end if

		if trim(request("EnableClearSendBox")) = "" then
			ei.EnableClearSendBox = false
		else
			ei.EnableClearSendBox = true
		end if

		if trim(request("enableRichEditer")) = "" then
			ei.useRichEditer = false
		else
			ei.useRichEditer = true
		end if

		if trim(request("EnableShowHtmlMail")) = "" then
			ei.EnableShowHtmlMail = false
		else
			ei.EnableShowHtmlMail = true
		end if

		if trim(request("Enable_AutoReply_OneDay")) = "" then
			ei.Enable_AutoReply_OneDay = false
		else
			ei.Enable_AutoReply_OneDay = true
		end if

		if trim(request("enableAutoClear")) = "" then
			ei.useAutoClearTrashBox = false
		else
			ei.useAutoClearTrashBox = true
		end if

		if trim(request("autoClearDays")) <> "" and IsNumeric(trim(request("autoClearDays"))) = true then
			ei.autoClearTrashBoxDays = CInt(trim(request("autoClearDays")))
		else
			ei.autoClearTrashBoxDays = 15
		end if

		ei.save

		name = NULL
		domain = NULL
		comment = NULL

		i = i + 1
	loop

	set euser = nothing
	set ei = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=mdallstyle.asp"
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
.td_line_l {text-align:right; white-space:nowrap; background-color:#EFF7FF; border-bottom:1px #A5B6C8 solid; height:30px; color:#303030;}
.td_line_r {text-align:left; background-color:white; border-bottom:1px #A5B6C8 solid; height:30px; padding-left:6px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function window_onload() {
	EnableClearWhenFull_onclick();
}

function EnableClearWhenFull_onclick() {
	if (document.fm1.EnableClearWhenFull.checked == true)
		document.fm1.EnableClearSendBox.disabled = false;
	else
		document.fm1.EnableClearSendBox.disabled = true;
}

function gosub() {
	if (confirm("<%=b_lang_370 %>\r\n\r\n<%=b_lang_371 %>") == false)
		return ;

	document.fm1.submit();
}
//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload()">
<form method="post" action="mdallstyle.asp" name="fm1">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_368 %>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
</table>
<table width="88%" border="0" align="center" cellspacing="0">
	<tr>
	<td valign=center class="td_line_l"><%=b_lang_199 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="checkbox" name="EnableShowHtmlMail" value="checkbox"></td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_200 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="checkbox" name="enableRichEditer" value="checkbox"></td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=s_lang_0119 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="checkbox" name="EnableBackupAllSendMail" id="EnableBackupAllSendMail" value="checkbox"></td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=s_lang_0120 %><%=s_lang_mh %></td>
	<td class="td_line_r">
		<input type="checkbox" name="EnableClearWhenFull" value="checkbox" LANGUAGE=javascript onclick="return EnableClearWhenFull_onclick()">
		<%=s_lang_0121 %><br>
		<input type="checkbox" name="EnableClearSendBox" value="checkbox">
		<%=s_lang_0122 %>
	</td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_367 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="checkbox" name="Enable_AutoReply_OneDay" checked value="checkbox"></td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_207 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="checkbox" name="enableAutoClear" value="checkbox"></td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_208 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="text" name="autoClearDays" class='n_textbox' value="15" size="4" maxlength="4"> <%=b_lang_230 %></td>
	</tr>

	<tr><td colspan="2" align="left" height="28"><br>
	<a class='wwm_btnDownload btn_blue' href="right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
	<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=b_lang_369 %></a>
	</td></tr>
</table>
<br>
</Form>
</BODY>
</HTML>
