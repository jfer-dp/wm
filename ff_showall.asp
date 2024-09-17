<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
dim spf
set spf = server.createobject("easymail.ShareFolder")

dim ei
set ei = Application("em")

dim userweb
set userweb = server.createobject("easymail.UserWeb")
userweb.Load Session("wem")

allnum = userweb.Count
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
.title_tr {white-space:nowrap; background:#f2f4f6; height:24px;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.st_l, .st_r {height:24px; text-align:center; white-space:nowrap; border-left:1px solid #A5B6C8; border-top:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8;}
.st_r {border-right:1px solid #A5B6C8;}
.cont_tr {background:white; height:24px;}
.cont_td {height:24px; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px;}
.cont_td_word {height:24px; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px; word-break:break-all; word-wrap:break-word;}
-->
</STYLE>
</HEAD>

<script type="text/javascript" src="images/sc_left.js"></script>

<script type="text/javascript">
<!--
function m_over(tag_obj) {
	tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj) {
	tag_obj.style.backgroundColor = "white";
}
//-->
</script>

<body>
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_226 %>&nbsp;(<%=allnum %>)
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr class="title_tr">
	<td width="5%" class="st_l"><%=a_lang_071 %></td>
	<td width="37%" class="st_l"><%=a_lang_227 %></td>
	<td width="20%" class="st_l"><%=a_lang_203 %></td>
	<td width="26%" class="st_l"><%=a_lang_228 %></td>
	<td width="6%" class="st_l"><%=s_lang_modify %></td>
	<td width="6%" class="st_r"><%=a_lang_229 %></td>
	</tr>
<%
i = 0

dim username
dim foldername
dim isatt

do while i < allnum
	userweb.GetFriendFolderInfo i, username, foldername, isatt

	Response.Write "<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);'><td align='center' class='cont_td'>" & i + 1 & "</td>"

	if foldername <> "att" then
		if isatt = false then
			Response.Write "<td align='center' class='cont_td_word'><a href=""listmail.asp?sname=" & Server.URLEncode(username) & "&sfname=" & Server.URLEncode(foldername) & "&" & getGRSN() & """>" & server.htmlencode(foldername) & "</a></td>"
		else
			Response.Write "<td align='center' class='cont_td_word'><a href=""listatt.asp?mb=" & Server.URLEncode(foldername) & "&sname=" & Server.URLEncode(username) & "&sfname=" & Server.URLEncode(foldername) & "&" & getGRSN() & """>" & server.htmlencode(foldername & " (" & a_lang_230 & ")") & "</a></td>"
		end if
	else
		Response.Write "<td align='center' class='cont_td_word'><a href=""listatt.asp?sname=" & Server.URLEncode(username) & "&sfname=" & Server.URLEncode(foldername) & "&" & getGRSN() & """>" & a_lang_231 & "</a></td>"
	end if

	Response.Write "<td align='center' class='cont_td_word'>" & server.htmlencode(username) & "</td>"
	Response.Write "<td align='center' class='cont_td_word'>" & server.htmlencode(ei.GetUserMail(username)) & "&nbsp;</td>"

	if isatt = false then
		Response.Write "<td align='center' class='cont_td'><a href='ff_editsharefolder.asp?isatt=0&username=" & Server.URLEncode(username) & "&foldername=" & Server.URLEncode(foldername) & "&" & getGRSN() & "'><img src='images/edit.gif' border='0' title='" & s_lang_modify & "'></a></td>"
	else
		Response.Write "<td align='center' class='cont_td'><a href='ff_editsharefolder.asp?isatt=1&username=" & Server.URLEncode(username) & "&foldername=" & Server.URLEncode(foldername) & "&" & getGRSN() & "'><img src='images/edit.gif' border='0' title='" & s_lang_modify & "'></a></td>"
	end if

	Response.Write "<td align='center' class='cont_td'><a href='ff_delsharefolder.asp?index=" & i & "&" & getGRSN() & "'><img src='images/disshareff.gif' border='0' title='" & a_lang_232 & "'></a></td>"

	Response.Write "</tr>"

	username = NULL
	foldername = NULL
	isatt = NULL

    i = i + 1
loop

allnum = spf.Count
%>
	<tr><td bgcolor="#ffffff" colspan="6" align="right"><br>
<a class='wwm_btnDownload btn_gray' href="ff_addsharefolder.asp?<%=getGRSN() %>"><%=a_lang_233 %></a>
	</td></tr>
	</table>
</td></tr>
</table>

<br>
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_234 %>&nbsp;(<%=allnum %>)
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr class="title_tr">
	<td width="5%" class="st_l"><%=a_lang_071 %></td>
	<td width="43%" class="st_l"><%=a_lang_227 %></td>
	<td width="20%" class="st_l"><%=a_lang_203 %></td>
	<td width="26%" class="st_l"><%=a_lang_228 %></td>
	<td width="6%" class="st_r"><%=s_lang_add %></td>
	</tr>
<%
i = 0

do while i < allnum
	spf.GetInfo i, username, foldername, isatt

	if IsNull(username) = false then
		userstate = ei.GetUserState(username)
	else
		userstate = 1
	end if


	if userstate = 0 then
		if foldername <> "att" then
			if isatt = false then
				showfoldername = server.htmlencode(foldername)
			else
				showfoldername = server.htmlencode(foldername & " (" & a_lang_230 & ")")
			end if
		else
			showfoldername = a_lang_231
		end if

		Response.Write "<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);'><td align='center' class='cont_td'>" & i + 1 & "</td>"
		Response.Write "<td align='center' class='cont_td_word'>" & showfoldername & "</td>"
		Response.Write "<td align='center' class='cont_td_word'>" & server.htmlencode(username) & "</td>"
		Response.Write "<td align='center' class='cont_td_word'>" & server.htmlencode(ei.GetUserMail(username)) & "</td>"

		if Session("wem") <> username and userweb.HaveThisFriendFolder(username, foldername, isatt) = false then
			Response.Write "<td align='center' class='cont_td'><a href='ff_addsharefolder.asp?isatt=" & isatt & "&username=" & Server.URLEncode(username) & "&foldername=" & Server.URLEncode(foldername) & "&" & getGRSN() & "'><img src='images/shareff.gif' border='0' title='" & a_lang_235 & "'></a></td>"
		else
			Response.Write "<td align='center' class='cont_td'>&nbsp;</td>"
		end if

		Response.Write "</tr>"
	end if

    i = i + 1

	if userstate = 2 then
		spf.Remove username, foldername, isatt
		i = i - 1
	end if

	username = NULL
	foldername = NULL
	isatt = NULL
loop
%>
	</table>
</td></tr>
</table>
<div style="position:absolute; left:12px; top:10px;">
<a href="help.asp#ff_showall" target="_blank"><img src="images/help.gif" border="0" title="<%=s_lang_help %>"></a></div>
</BODY>
</HTML>

<%
set userweb = nothing
set ei = nothing
set spf = nothing
%>
