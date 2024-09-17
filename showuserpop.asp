<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" -->

<%
dim ei
set ei = server.createobject("easymail.UserMessages")
ei.Load Session("wem")

allnum = ei.GetMulPop3Count
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
.sbttn {font-family:<%=s_lang_font %>;font-size:9pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5;color:#000066;height:19px;text-decoration:none;cursor:pointer}
.title_tr {white-space:nowrap; background:#f2f4f6; height:24px;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.st_l, .st_r {height:24px; text-align:center; white-space:nowrap; border-left:1px solid #A5B6C8; border-top:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8;}
.st_r {border-right:1px solid #A5B6C8;}
.cont_tr {background:white; height:24px;}
.cont_td {white-space:nowrap; height:24px; border-bottom:1px solid #A5B6C8; padding-left:4px; padding-right:4px;}
.cont_td_word {height:24px; border-bottom:1px solid #A5B6C8; padding-left:4px; padding-right:4px; word-break:break-all; word-wrap:break-word;}
.ctd {white-space:nowrap; height:24px; padding-left:4px; padding-right:4px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function save() {
	document.f1.mode.value = "save";
	document.f1.submit();
}

function del() {
	if (ischeck() == true)
	{
		document.f1.mode.value = "del";
		document.f1.submit();
	}
}

function add() {
	if (document.f1.uname.value != "" && document.f1.userver.value != "" && document.f1.uport.value != "" && document.f1.uusername.value != "" && document.f1.upassword.value != "")
	{
		document.f1.mode.value = "add";
		document.f1.submit();
	}
	else
		alert("<%=b_lang_163 %>");
}

function ischeck() {
	var i = 0;
	var theObj;

	for(; i<<%=allnum %>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}

function m_over(tag_obj) {
	tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj) {
	tag_obj.style.backgroundColor = "white";
}
//-->
</SCRIPT>

<BODY>
<FORM ACTION="saveuserpop.asp" METHOD="POST" NAME="f1">
<input type="hidden" name="mode">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_164 %>
</td></tr>
<tr><td class="block_top_td" style="height:10px; _height:12px;"></td></tr>
<tr><td align="left">
	&nbsp;<input type="checkbox" name="checkpop" <% if ei.POP3Support = true then Response.Write "checked"%>>
	<%=b_lang_165 %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="button" value=" <%=s_lang_save %> " onclick="javascript:save()" class="sbttn">
</td></tr>
</table>

<%
if allnum < CInt(Application("em_MaxMPOP3")) then
%>
<br><br>
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_166 %>
</td></tr>
<tr><td class="block_top_td" style="height:10px; _height:12px;"></td></tr>
<tr><td align="left">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr><td width="15%" align="right" class="ctd">
	<%=b_lang_167 %><%=s_lang_mh %>
	</td><td width="35%" align="left">
	<input type="text" name="uname" size="20" maxlength="64" class="n_textbox">
	</td>
	<td width="15%" align="right" class="ctd">
	<%=b_lang_168 %><%=s_lang_mh %>
	</td><td width="35%" align="left">
	<input type="text" name="userver" size="20" maxlength="64" class="n_textbox">
	</td></tr>

	<tr><td align="right" class="ctd">
	<%=b_lang_169 %><%=s_lang_mh %>
	</td><td align="left">
	<input type="text" name="uport" size="5" maxlength="5" value="110" class="n_textbox">
	</td>
	<td align="right" class="ctd">
	<%=b_lang_170 %><%=s_lang_mh %>
	</td><td align="left">
	<input type="text" name="uusername" size="20" maxlength="64" class="n_textbox">
	</td></tr>

	<tr><td align="right" class="ctd">
	<%=b_lang_171 %><%=s_lang_mh %>
	</td><td align="left">
	<input type="password" name="upassword" size="16" maxlength="64" class="n_textbox">
	</td>
	<td align="right" class="ctd">
	<%=b_lang_172 %><%=s_lang_mh %>
	</td><td align="left">
	<input type="checkbox" name="uisdel" checked>
	</td></tr>
</table>

<tr><td class="block_top_td" style="height:6px;"></td></tr>

<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:add();"><%=s_lang_add %></a>
	</td></tr>
</table>
<%
end if
%>

<br><br>
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_173 %>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="left">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr class="title_tr">
	<td width="4%" class="st_l">&nbsp;</td>
	<td width="25%" class="st_l"><%=b_lang_167 %></td>
	<td width="23%" class="st_l"><%=b_lang_168 %></td>
	<td width="8%" class="st_l"><%=b_lang_169 %></td>
	<td width="25%" class="st_l"><%=b_lang_170 %></td>
	<td width="15%" class="st_r"><%=b_lang_174 %></td>
	</tr>
<%
i = 0

do while i < allnum
	ei.GetMulPop3 i, name, sev, port, username, isdel

	Response.Write "<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);'><td align='center' class='cont_td'><input type='checkbox' name='check" & i & "'>"
	Response.Write "</td><td align='center' class='cont_td_word'>" & server.htmlencode(name)
	Response.Write "</td><td align='center' class='cont_td'>" & server.htmlencode(sev)
	Response.Write "</td><td align='center' class='cont_td'>" & port
	Response.Write "</td><td align='center' class='cont_td_word'>" & server.htmlencode(username)

	if isdel = true then
		Response.Write "</td><td align='center' class='cont_td'>" & s_lang_del
	else
		Response.Write "</td><td align='center' class='cont_td'>" & b_lang_175
	end if

	Response.Write "</td></tr>"

	name = NULL
	sev = NULL
	port = NULL
	username = NULL
	isdel = NULL

	i = i + 1
loop
%>
</table>
	</td></tr>

<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px;">
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:del();"><%=s_lang_del %></a>
	</td></tr>
</table>
</FORM>

<div style="position:absolute; left:12px; top:10px;">
<a href="help.asp#showuserpop" target="_blank"><img src="images/help.gif" border="0" title="<%=s_lang_help %>"></a></div>
</BODY>
</HTML>

<%
set ei = nothing
%>
