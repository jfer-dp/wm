<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" --> 

<%
showisend = trim(request("showisend"))
if LCase(showisend) <> "true" then
	showisend = false
else
	showisend = true
end if

ismanager = false
if isadmin() = true then
	ismanager = true
end if

dim poll
set poll = server.createobject("easymail.Poll")
poll.LoadAll

if showisend = false then
	allnum = poll.UnEnd_Count
else
	allnum = poll.End_Count
end if

curdomain = Mid(Session("mail"), InStr(Session("mail"), "@") + 1)
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
function delone(fid)
{
	if (confirm("<%=b_lang_036 %>") == false)
		return ;

<%
	if showisend = false then
%>
	location.href = "poll_del.asp?<%=getGRSN() %>&fid=" + fid + "&gourl=<%=Server.URLEncode("poll_showall.asp") %>";
<%
else
%>
	location.href = "poll_del.asp?<%=getGRSN() %>&fid=" + fid + "&gourl=<%=Server.URLEncode("poll_showall.asp?showisend=true") %>";
<%
end if
%>
}

function m_over(tag_obj) {
	tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj) {
	tag_obj.style.backgroundColor = "white";
}
//-->
</script>

<body>
<table width="90%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr>
	<td align="left" height="28" width="55%" nowrap style="padding-left:4px;">
<%
if ismanager = true then
%>
<a class='wwm_btnDownload btn_blue' href="poll_create.asp?<%=getGRSN() %>"><%=b_lang_037 %></a>
<%
end if

if showisend = false then
%>
<a class='wwm_btnDownload btn_blue' href="poll_showall.asp?<%=getGRSN() %>&showisend=True"><%=b_lang_038 %></a>
	<td align="right" nowrap style="padding-right:8px; color:#444444;"><%=b_lang_039 %></td>
<%
else
%>
<a class='wwm_btnDownload btn_blue' href="poll_showall.asp?<%=getGRSN() %>&showisend=False"><%=b_lang_039 %></a>
	<td align="right" nowrap style="padding-right:8px; color:#444444;"><%=b_lang_038 %></td>
<%
end if
%>
	</tr>
</table>
<br>

<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr class="title_tr">
	<td width="6%" class="st_l"><%=b_lang_040 %></td>
	<td width="47%" class="st_l"><%=b_lang_041 %></td>
	<td width="10%" class="st_l"><%=b_lang_042 %></td>
	<td width="15%" class="st_l"><%=b_lang_043 %></td>
<%
if showisend = false then
%>
	<td width="6%" class="st_l"><%=b_lang_044 %></td>
<%
else
%>
	<td width="6%" class="st_l"><%=b_lang_045 %></td>
<%
end if

if ismanager = true then
%>
	<td width="6%" class="st_l"><%=b_lang_046 %></td>
	<td width="5%" class="st_l"><%=b_lang_047 %></td>
	<td width="5%" class="st_r"><%=s_lang_del %></td>
<%
else
%>
	<td width="6%" class="st_r"><%=b_lang_046 %></td>
<%
end if
%>
	</tr>
<%
i = 0
shownum = 0

do while i < allnum
	poll.MoveTo i, showisend

	if poll.PI_HaveThisDomain(curdomain) = true or ismanager = true then
		Response.Write "<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);'><td height='24' align='center' class='cont_td'>" & shownum + 1 & "</td>"

		if showisend = false then
			Response.Write "<td align='center' class='cont_td_word'><a href='poll_showone.asp?gourl=" & Server.URLEncode("poll_showall.asp") & "&fid=" & poll.PI_Filename & "&" & getGRSN() & "'>" & server.htmlencode(poll.PI_Title) & "</a></td>"
		else
			Response.Write "<td align='center' class='cont_td_word'><a href='poll_showone.asp?gourl=" & Server.URLEncode("poll_showall.asp?showisend=true") & "&fid=" & poll.PI_Filename & "&" & getGRSN() & "'>" & server.htmlencode(poll.PI_Title) & "</a></td>"
		end if

		Response.Write "<td align='right' class='cont_td'>" & poll.PI_VoteSum & "</td>"
		Response.Write "<td align='center' class='cont_td'>" & getYear(poll.PI_CreateTime) & "-" & getMonth(poll.PI_CreateTime) & "-" & getDay(poll.PI_CreateTime) & "</td>"

		if showisend = false then
			Response.Write "<td align='center' class='cont_td'><a href='poll_vote.asp?gourl=" & Server.URLEncode("poll_showall.asp") & "&fid=" & poll.PI_Filename & "&" & getGRSN() & "'>" & b_lang_044 & "</a></td>"
		else
			Response.Write "<td align='center' class='cont_td'><a href='poll_showone.asp?gourl=" & Server.URLEncode("poll_showall.asp?showisend=true") & "&fid=" & poll.PI_Filename & "&" & getGRSN() & "'>" & b_lang_045 & "</a></td>"
		end if

		if poll.PI_Poll_BBS <> "" then
			if showisend = false then
				Response.Write "<td align='center' class='cont_td'><a href='showpf.asp?backurl=" & Server.URLEncode("poll_showall.asp") & "&fileid=" & Server.URLEncode(poll.PI_Poll_BBS) & "&" & getGRSN() & "'>" & b_lang_046 & "</a></td>"
			else
				Response.Write "<td align='center' class='cont_td'><a href='showpf.asp?backurl=" & Server.URLEncode("poll_showall.asp?showisend=true") & "&fileid=" & Server.URLEncode(poll.PI_Poll_BBS) & "&" & getGRSN() & "'>" & b_lang_046 & "</a></td>"
			end if
		else
			Response.Write "<td align='center' class='cont_td'>&nbsp;</td>"
		end if

		if ismanager = true then
			if showisend = false then
				Response.Write "<td align='center' class='cont_td'><a href='poll_edit.asp?gourl=" & Server.URLEncode("poll_showall.asp") & "&fid=" & poll.PI_Filename & "&" & getGRSN() & "'><img src='images/edit.gif' border='0' title='" & s_lang_modify & "'></a></td>"
			else
				Response.Write "<td align='center' class='cont_td'><a href='poll_edit.asp?gourl=" & Server.URLEncode("poll_showall.asp?showisend=true") & "&fid=" & poll.PI_Filename & "&" & getGRSN() & "'><img src='images/edit.gif' border='0' title='" & s_lang_modify & "'></a></td>"
			end if

			Response.Write "<td align='center' class='cont_td'><a href=""javascript:delone('" & poll.PI_Filename & "')""><img src='images/del.gif' border='0' title='" & s_lang_del & "'></a></td>"
		end if

		Response.Write "</tr>" & Chr(13)
		shownum = shownum + 1
	end if

    i = i + 1
loop
%>
</table>
</BODY>
</HTML>

<%
set poll = nothing


function getYear(exday)
	getYear = Mid(Cstr(exday), 1, 4)
end function

function getMonth(exday)
	getMonth = Mid(Cstr(exday), 5, 2)
end function

function getDay(exday)
	getDay = Mid(Cstr(exday), 7, 2)
end function
%>
