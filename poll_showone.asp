<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" --> 

<%
fid = trim(request("fid"))
gourl = trim(request("gourl"))
curdomain = Mid(Session("mail"), InStr(Session("mail"), "@") + 1)

ismanager = false
if isadmin() = true then
	ismanager = true
end if

dim poll
set poll = server.createobject("easymail.Poll")
poll.LoadOne fid

if poll.PI_HaveThisDomain(curdomain) = false and ismanager = false then
	set poll = nothing
	response.redirect "noadmin.asp"
end if

PI_Can_Poll = poll.PI_Can_Poll(Session("wem"))
allnum = poll.PI_ChooseCount
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
.font_g {font-size:12px; color:#444444; font-weight:normal;}
.cont_td {height:27px; bgcolor:white; border-left:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px;}
.cont_td_word {height:27px; bgcolor:white; border-left:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px; word-break:break-all; word-wrap:break-word;}
-->
</STYLE>
</HEAD>

<BODY>
<FORM ACTION="poll_vote.asp" METHOD="POST" NAME="f1">
<body>
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px; padding-right:6px; word-break:break-all; word-wrap:break-word;">
<%=b_lang_048 %><%=server.htmlencode(poll.PI_Title) %><%=b_lang_049 %><%
if PI_Can_Poll = false then
	Response.Write " <font class='font_g'>" & b_lang_050 & "</font>"
end if
%>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="center">

<table width="96%" border="0" align="center" cellspacing="0" bgcolor="white">
<%
i = 0
allvotenum = poll.PI_VoteSum

do while i < allnum
	poll.PI_GetNameAndNumber i, v_name, v_num

	if allvotenum > 0 then
		bf = CLng((100 * v_num) / allvotenum)
	else
		bf = 0
	end if

	if bf > 100 then
		bf = 100
	end if

	if bf < 0 then
		bf = 0
	end if
%>
	<tr><td width="40%" align="left" class="cont_td_word"<%
if i = 0 then
	Response.Write " style='border-top:1px solid #A5B6C8;'"
end if
%>>&nbsp;<%=server.htmlencode(v_name) %></td>
	<td nowrap width="48%" align="left" class="cont_td"<%
if i = 0 then
	Response.Write " style='border-top:1px solid #A5B6C8;'"
end if
%>>&nbsp;<img height="10" width="<%=bf * 2 %>" src="images/bar.gif" align='absmiddle'>&nbsp;<%=bf %>%</td>
	<td nowrap width="12%" align="right" class="cont_td" style="border-right:1px solid #A5B6C8;<%
if i = 0 then
	Response.Write " border-top:1px solid #A5B6C8;'"
end if
%>"><%=v_num %><%=b_lang_051 %>&nbsp;</td>
	</tr>
<%
	i = i + 1

	v_name = NULL
	v_num = NULL
loop
%>
	<tr>
	<td colspan="3" align="left" height="26"><br>
<a class='wwm_btnDownload btn_blue' href="javascript:history.back();"><< <%=s_lang_return %></a>
<%
if PI_Can_Poll = true and poll.PI_IsEnd = false then
%>
<a class='wwm_btnDownload btn_blue' href="poll_vote.asp?fid=<%=fid %>&<%=getGRSN() %>&gourl=<%=Server.URLEncode(gourl) %>"><%=b_lang_044 %></a>
<%
end if

if poll.PI_Poll_BBS <> "" then
%>
<a class='wwm_btnDownload btn_blue' href="showpf.asp?backurl=<%=Server.URLEncode(gourl) %>&fileid=<%=Server.URLEncode(poll.PI_Poll_BBS) %>&<%=getGRSN() %>"><%=b_lang_046 %></a>
<%
end if
%>
</td></tr>
</table>

</td></tr>
</table>

</form>
</BODY>
</HTML>

<%
set poll = nothing
%>
