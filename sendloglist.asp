<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
user = ""
if isadmin() = true then
	user = trim(request("user"))
end if

dim slm
set slm = server.createobject("easymail.SendLogManager")

if Len(user) < 1 then
	slm.LoadList Session("wem")
else
	slm.LoadList user
end if

allnum = slm.ListCount

if trim(request("page")) = "" then
	page = 0
else
	page = CInt(request("page"))
end if

allpage = CInt((allnum - (allnum mod pageline))/ pageline)

if allnum mod pageline <> 0 then
	allpage = allpage + 1
end if

if page >= allpage then
	page = allpage - 1
end if

if page < 0 then
	page = 0
end if

if allpage = 0 then
	allpage = 1
end if

dim t_day
t_day = Year(date)

if Month(date) < 10 then
	t_day = t_day & "0"
end if

t_day = t_day & Month(date)

if Day(date) < 10 then
	t_day = t_day & "0"
end if

t_day = t_day & Day(date)

gourl = "sendloglist.asp?user=" & Server.URLEncode(user) & "&page=" & page
returl = trim(request("returl"))
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/hwem.css">
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<STYLE type=text/css>
<!--
.EX_TITLE {border-left:1px solid #d1d8e2; border-right:1px solid #d1d8e2; border-bottom:1px solid #d1d8e2; BACKGROUND-COLOR: #F8F8D2; padding-left:4px; padding-top:3px; white-space:nowrap; height:19px;}
.EX_TITLE_FONT {FONT-WEIGHT:bold; COLOR:#666666;}

.title_tr {white-space:nowrap; background:#f2f4f6; height:24px;}
.st_1,.st_2,.st_3,.st_4 {text-align:center; white-space:nowrap; border-left:1px solid #c1c8d2; border-top:1px solid #c1c8d2; border-bottom:1px solid #c1c8d2;}

.st_1 {width:10%;}
.st_2 {width:35%;}
.st_3 {width:49%;}
.st_4 {width:6%; border-right:1px solid #c1c8d2;}

.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.table_min_width {width:660px; font-size:0pt; height:0px; width:0px; border:0px;}
.cont_tr {background:white; height:26px; cursor:pointer;}
.cont_td {border-bottom:1px solid #e3e6eb; padding-left:8px; padding-right:8px;}
-->
</STYLE>
</HEAD>

<%
if Len(user) < 1 then
%>
	<script type="text/javascript" src="images/sc_left.js"></script>
<%
end if
%>

<SCRIPT LANGUAGE=javascript>
<!--
function back() {
<% if returl = "" then %>
	location.href = "listmail.asp?mode=sed&<%=getGRSN() %>";
<% else %>
	location.href = "<%=returl %>&<%=getGRSN() %>";
<% end if %>
}

function selectpage_onchange() {
	location.href = "sendloglist.asp?user=<%=Server.URLEncode(user) %>&page=" + document.f1.page.value + "&<%=getGRSN() %>&returl=<%=Server.URLEncode(returl) %>";
}

function m_over(tag_obj) {
	tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj) {
	tag_obj.style.backgroundColor = "white";
}

function showonelog(s_url) {
	location.href = "sendonelog.asp?" + s_url;
}
//-->
</SCRIPT>

<BODY>
<FORM name="f1">
<table width="90%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr>
	<td align="left" height="28" width="30%" nowrap style="padding-left:15px; color:#444;"><%=s_lang_0124 %><font color='#901111'><%=allnum %></font><%=s_lang_0272 %></td>
	<td align="center" width="40%" nowrap><%
if page - 1 < 0 then
	Response.Write "<img src='images\gfirstp.gif' border='0' align='absmiddle'>&nbsp;"
	Response.Write "<img src='images\gprep.gif' border='0' align='absmiddle'>&nbsp;&nbsp;"
else
	Response.Write "<a href=""sendloglist.asp?user=" & Server.URLEncode(user) & "&page=" & 0 & "&" & getGRSN() & "&returl=" & Server.URLEncode(returl) & """><img src='images\firstp.gif' border='0' align='absmiddle'></a>&nbsp;"
	Response.Write "<a href=""sendloglist.asp?user=" & Server.URLEncode(user) & "&page=" & page - 1 & "&" & getGRSN() & "&returl=" & Server.URLEncode(returl) & """><img src='images\prep.gif' border='0' align='absmiddle'></a>&nbsp;&nbsp;"
end if
%><select name="page" class="drpdwn" size="1" LANGUAGE=javascript onchange="selectpage_onchange()">
<%
i = 0

do while i < allpage
	if i <> page then
		Response.Write "<option value=""" & i & """>" & i + 1 & "</option>"
	else
		Response.Write "<option value=""" & i & """ selected>" & i + 1 & "</option>"
	end if
	i = i + 1
loop
%></select>
<%
if ((page+1) * pageline) => allnum then
	Response.Write "<img src='images\gnextp.gif' border='0' align='absmiddle'>&nbsp;"
else
	Response.Write "<a href=""sendloglist.asp?user=" & Server.URLEncode(user) & "&page=" & page + 1 & "&" & getGRSN() & "&returl=" & Server.URLEncode(returl) & """><img src='images\nextp.gif' border='0' align='absmiddle'></a>&nbsp;"
end if

if page + 1 >= allpage then
	Response.Write "<img src='images\gendp.gif' border='0' align='absmiddle'>&nbsp;"
else
	Response.Write "<a href=""sendloglist.asp?user=" & Server.URLEncode(user) & "&page=" & allpage - 1 & "&" & getGRSN() & "&returl=" & Server.URLEncode(returl) & """><img src='images\endp.gif' border='0' align='absmiddle'></a>&nbsp;"
end if
%>
	</td>
	<td align="right" width="30%" nowrap style="padding-right:15px;">
<% if returl <> "" then %>
	<a class='wwm_btnDownload btn_blue' href="javascript:back();"><%=s_lang_return %></a>
<% end if %>
	</td></tr>
</table>
<br>
<table width="90%" border="0" align="center" cellspacing="0">
	<tr class="title_tr">
	<td class="st_1"><%=s_lang_0128 %></td>
	<td class="st_2"><%=s_lang_0149 %></td>
	<td class="st_3"><%=s_lang_0127 %></td>
	<td class="st_4"><%=s_lang_0126 %></td>
	</tr>
<%
i = page * pageline
li = 0
dim cur_show_day

do while i < allnum and li < pageline
	slm.GetListInfo i, filename, subject, recv_names, is_end, datestr

	dim show_day_str
	show_day_str = false

	if cur_show_day <> mid(datestr, 1, 8) then
		cur_show_day = mid(datestr, 1, 8)
		show_day_str = true
	end if

	if li = 0 or show_day_str = true then
		Response.Write "<tr><td colspan=7 height='26' nowrap class='EX_TITLE'><font class='EX_TITLE_FONT'>&nbsp;"
		if t_day = mid(datestr, 1, 8) then
			Response.Write s_lang_0267
		else
			Response.Write get_date_showstr(datestr)
		end if
		Response.Write "</font></td></tr>"
	end if
%>
    <tr class="cont_tr" onmouseover='m_over(this);' onmouseout='m_out(this);' onclick="showonelog('<%="user=" & Server.URLEncode(user) & "&filename=" & filename & "&" & getGRSN() & "&gourl=" & Server.URLEncode(gourl) & "&returl=" & Server.URLEncode(returl) %>');">
	<td nowrap class="cont_td"><%=get_time_showstr(datestr) %></td>
	<td class="cont_td" style="word-break:break-all; word-wrap:break-word;"><%=server.htmlencode(recv_names) %>&nbsp;</td>
	<td class="cont_td" style="word-break:break-all; word-wrap:break-word;"><%=server.htmlencode(subject) %>&nbsp;</td>
	<td align="center" class="cont_td"><img src="images/<%
if is_end = true then
	Response.Write "rc_end.gif"" title=""" & s_lang_0273
else
	Response.Write "rc_noend.gif"" title=""" & s_lang_0274
end if
%>" border="0" align="absmiddle"></a></td>
	</tr>
<%	
    li = li + 1

	filename = NULL
	subject = NULL
	recv_names = NULL
	is_end = NULL
	datestr = NULL

	i = i + 1
loop
%>
	<tr><td class="block_top_td" colspan="4"><div class="table_min_width"></div></td></tr>
	</table>
</FORM>
</BODY>
</HTML>

<%
set slm = nothing

function get_date_showstr(show_date_str)
	if Len(show_date_str) = 14 or Len(show_date_str) = 12 then
		tmp_month = Mid(show_date_str, 5, 2)
		if Mid(tmp_month, 1, 1) = "0" then
			tmp_month = Mid(tmp_month, 2, 1)
		end if

		tmp_day = Mid(show_date_str, 7, 2)
		if Mid(tmp_day, 1, 1) = "0" then
			tmp_day = Mid(tmp_day, 2, 1)
		end if

		get_date_showstr = Mid(show_date_str, 1, 4) & s_lang_0139 & tmp_month & s_lang_0140 & tmp_day & s_lang_0141
	else
		get_date_showstr = ""
	end if
end function

function get_time_showstr(show_date_str)
	if Len(show_date_str) = 14 or Len(show_date_str) = 12 then
		t_time_hour = CInt(Mid(show_date_str, 9, 2))

		dim t_hour_name
		if t_time_hour >= 0 and t_time_hour < 6 then
			t_hour_name = s_lang_0268
		elseif t_time_hour >= 6 and t_time_hour < 12 then
			t_hour_name = s_lang_0269
		elseif t_time_hour >= 12 and t_time_hour < 14 then
			t_hour_name = s_lang_0283
			if t_time_hour > 12 then
				t_time_hour = t_time_hour - 12
			end if
		elseif t_time_hour >= 14 and t_time_hour < 18 then
			t_hour_name = s_lang_0270
			t_time_hour = t_time_hour - 12
		elseif t_time_hour >= 18 then
			t_hour_name = s_lang_0271
			t_time_hour = t_time_hour - 12
		end if

		ts_time_hour = t_time_hour

		get_time_showstr = t_hour_name & ts_time_hour& ":" & Mid(show_date_str, 11, 2)
	else
		get_time_showstr = ""
	end if
end function
%>
