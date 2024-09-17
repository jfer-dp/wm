<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
dim uw
set uw = server.createobject("easymail.UserWeb")
uw.Load Session("wem")
Show_EC_Date_Style = uw.EnableShowDateECMailList
set uw = nothing

exstr = request("exstr")
if IsEmpty(exstr) = true then
	exstr = "00000"
end if

addsortstr = "&exstr=" & exstr

dim ei
set ei = server.createobject("easymail.RecallListManager")

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	doid = trim(request("deloneid"))

	if doid <> "" then
		if Len(doid) > 5 then
			ei.Del Session("wem"), doid
		elseif doid = "no" then
			ei.Load Session("wem")

			if ei.count > pageline then
				themax = pageline
			else
				themax = ei.count
			end if

			i = 0
			do while i <= themax
				if trim(request("check" & i)) <> "" then
					ei.Del Session("wem"), trim(request("check" & i))
				end if 
	
			    i = i + 1
			loop
		end if
	end if
end if

ei.Load Session("wem")
ei.SortByDate

allnum = ei.Count

if trim(request("page")) = "" then
	page = 0
else
	page = CInt(request("page"))
end if

if Show_EC_Date_Style = true then
	allnum = ei.GetDateExRecalls(exstr)
end if
show_allnum = allnum

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

if Show_EC_Date_Style = true then
	allnum = ei.Count
end if

dim show_EC_mode
show_EC_mode = -1

dim is_show_EC

gourl = "recalllist.asp?page=" & page
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
.EX_TITLE {border-left:1px solid #d1d8e2; border-right:1px solid #d1d8e2; border-bottom:1px solid #d1d8e2; BACKGROUND-COLOR: #F8F8D2; padding-left:4px; padding-top:3px; white-space:nowrap; height:19px;}
.EX_TITLE_FONT {FONT-WEIGHT:bold; COLOR:#666666;}

.title_tr {white-space:nowrap; background:#f2f4f6; height:24px;}
.st_1,.st_2,.st_3,.st_4,.st_5,.st_6 {text-align:center; white-space:nowrap; border-left:1px solid #c1c8d2; border-top:1px solid #c1c8d2; border-bottom:1px solid #c1c8d2;}

.st_1 {width:4%;}
.st_2 {width:5%;}
.st_3 {width:5%;}
.st_4 {width:59%;}
.st_5 {width:20%;}
.st_6 {width:7%; border-right:1px solid #c1c8d2;}

.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.table_min_width {width:660px; font-size:0pt; height:0px; width:0px; border:0px;}
.cont_tr {background:white; height:26px; cursor:pointer;}
.cont_td {border-bottom:1px solid #e3e6eb; padding-left:8px; padding-right:8px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript" src="images/sc_left.js"></script>

<BODY>
<FORM ACTION="recalllist.asp" METHOD="POST" name="f1">
<INPUT NAME="deloneid" TYPE="hidden">
<INPUT NAME="exstr" TYPE="hidden">
<table width="90%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr>
	<td align="left" height="28" width="30%" nowrap style="padding-left:15px; color:#444;"><%=s_lang_0124 %><font color='#901111'><%=allnum %></font><%=s_lang_0125 %></td>
	<td align="center" width="40%" nowrap><%
if page - 1 < 0 then
	Response.Write "<img src='images/gfirstp.gif' border='0' align='absmiddle'>&nbsp;"
	Response.Write "<img src='images/gprep.gif' border='0' align='absmiddle'>&nbsp;&nbsp;"
else
	Response.Write "<a href=""recalllist.asp?page=" & 0 & addsortstr & "&" & getGRSN() & """><img src='images/firstp.gif' border='0' align='absmiddle'></a>&nbsp;"
	Response.Write "<a href=""recalllist.asp?page=" & page - 1 & addsortstr & "&" & getGRSN() & """><img src='images/prep.gif' border='0' align='absmiddle'></a>&nbsp;&nbsp;"
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
if ((page+1) * pageline) => show_allnum then
	Response.Write "<img src='images/gnextp.gif' border='0' align='absmiddle'>&nbsp;"
else
	Response.Write "<a href=""recalllist.asp?page=" & page + 1 & addsortstr & "&" & getGRSN() & """><img src='images/nextp.gif' border='0' align='absmiddle'></a>&nbsp;"
end if

if page + 1 >= allpage then
	Response.Write "<img src='images/gendp.gif' border='0' align='absmiddle'>&nbsp;"
else
	Response.Write "<a href=""recalllist.asp?page=" & allpage - 1 & addsortstr & "&" & getGRSN() & """><img src='images/endp.gif' border='0' align='absmiddle'></a>&nbsp;"
end if
%>
	</td>
	<td align="right" width="30%" nowrap style="padding-right:15px;">
	<a class='wwm_btnDownload btn_blue' href="javascript:del();"><%=s_lang_del %></a>
	</td></tr>
</table>
<br>
<table width="90%" border="0" align="center" cellspacing="0">
	<tr class="title_tr">
	<td class="st_1"><img src='images/high.gif' border='0' align='absmiddle'></td>
	<td class="st_2"><%=s_lang_0126 %></td>
	<td class="st_3"><input type="checkbox" onclick="checkall(this)"></td>
	<td class="st_4"><%=s_lang_0127 %></td>
	<td class="st_5"><%=s_lang_0128 %></td>
	<td class="st_6"><%=s_lang_del %></td>
	</tr>
<%
i = page * pageline
li = 0

show_page_ex_head = true

if Show_EC_Date_Style = true and page > 0 then
	ei.GetDateExMode exstr, i - 1, exmode, isshow
	i = i + getBeforeLines(exmode)

	exmode = NULL
	isshow = NULL

	ei.GetDateExMode exstr, i, bf_exmode, bf_isshow
	ei.GetDateExMode exstr, i - 1, exmode, isshow

	if bf_exmode > -1 and exmode > -1 and bf_exmode = exmode then
		show_page_ex_head = false
	end if

	if show_page_ex_head = true then
		writeBeforeHiddenTitle(bf_exmode)
	end if

	bf_exmode = NULL
	bf_isshow = NULL

	exmode = NULL
	isshow = NULL
end if

do while i < allnum and li < pageline
	ei.Get i, rc_filename, rc_is_end, rc_time, rc_priority, rc_from_name, rc_from_email, rc_subject

if Show_EC_Date_Style = false or (Show_EC_Date_Style = true and writeDateEC(i) = true) then
	if rc_subject = "" then
		rc_subject = s_lang_0129
	end if

	if rc_priority = 2 then
		xmsp = "<img src='images/high.gif' border='0' title='" & s_lang_0130 & "'>"
	elseif rc_priority = 1 then
		xmsp = "<img src='images/low.gif' border='0' title='" & s_lang_0131 & "'>"
	else
		xmsp = "&nbsp;"
	end if
%>
	<tr id="tr_<%=li %>" class="cont_tr" onmouseover='m_over(this);' onmouseout='m_out(this);' onclick="showone('<%="filename=" & rc_filename & "&" & getGRSN() & "&gourl=" & Server.URLEncode(gourl & addsortstr) %>');">
	<td align="center" nowrap class="cont_td"><%=xmsp %></td>
	<td align="center" nowrap class="cont_td"><img src="images/<%
if rc_is_end = true then
	Response.Write "rc_end.gif"" title=""" & s_lang_0132
else
	Response.Write "rc_noend.gif"" title=""" & s_lang_0133
end if
%>" border=0></td>
	<td align="center" class="cont_td" style="cursor:default;" onclick="no_click(event)"><input type="checkbox" id="check<%=li %>" name="check<%=li %>" value="<%=rc_filename %>" onclick="ck_select(this);" style="cursor:pointer;"></td>
	<td align="left" class="cont_td" style='word-break:break-all; word-wrap:break-word;'><%=server.htmlencode(rc_subject) %>&nbsp;</td>
	<td align="left" nowrap class="cont_td"><%=get_date_showstr(rc_time) %></td>
	<td align="center" class="cont_td" style="cursor:default;" onclick="no_click(event)"><a href="javascript:delone('<%=Server.URLEncode(rc_filename) %>')"><img src='images/del.gif' border='0' title='<%=s_lang_del %>'></a></td>
	</tr>
<%
    li = li + 1
end if

	rc_filename = NULL
	rc_is_end = NULL
	rc_time = NULL
	rc_priority = NULL
	rc_from_name = NULL
	rc_from_email = NULL
	rc_subject = NULL

	i = i + 1
loop

if Show_EC_Date_Style = true then
	last_exmode = -1
	do while i < allnum
		ei.GetDateExMode exstr, i, bf_exmode, bf_isshow
		tmp_bf_exmode = bf_exmode

		if bf_isshow = true or canshowlast(tmp_bf_exmode) = false then
			bf_exmode = NULL
			bf_isshow = NULL

			exit do
		end if

		if last_exmode <> bf_exmode then
			writeDateEC(i)
		end if

		bf_exmode = NULL
		bf_isshow = NULL

		i = i + 1
	loop
end if
%>
</table>
</FORM>

<script type="text/javascript">
function checkall(tgobj) {
	var theObj;
	for(var i = 0; i < <%=li %>; i++)
	{
		theObj = document.getElementById("check" + i);
		if (theObj != null)
		{
			theObj.checked = tgobj.checked;
			ck_select(theObj);
		}
	}
}

function del() {
	if (ischeck() == true)
	{
		if (confirm("<%=s_lang_0115 %>") == false)
			return ;

		document.f1.exstr.value = "<%=exstr %>";
		document.f1.deloneid.value = "no";
		document.f1.submit();
	}
}

function delone(id) {
	if (confirm("<%=s_lang_0115 %>") == false)
		return ;

	document.f1.exstr.value = "<%=exstr %>";
	document.f1.deloneid.value = id;
	document.f1.submit();
}

function ischeck() {
	var theObj;

	for(var i = 0; i < <%=li %>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}

function selectpage_onchange()
{
	location.href = "recalllist.asp?page=" + document.f1.page.value + "<%=addsortstr & "&" & getGRSN() %>";
}

function m_over(tag_obj)
{
	if (document.getElementById("check" + tag_obj.id.substr(3)).checked == false)
		tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj)
{
	if (document.getElementById("check" + tag_obj.id.substr(3)).checked == false)
		tag_obj.style.backgroundColor = "white";
}

function ck_select(tag_obj) {
	if (tag_obj.checked == true)
		document.getElementById("tr_" + tag_obj.id.substr(5)).style.background = "#93BEE2";
	else
		document.getElementById("tr_" + tag_obj.id.substr(5)).style.background = "white";
}

function no_click(e) {
	stop_event(e);
}

function stop_event(e) {
	if (!e) var e = window.event;
	e.cancelBubble = true;
	if (e.stopPropagation)
	e.stopPropagation();
}

function showone(s_url) {
	location.href = "recallinfo.asp?" + s_url;
}

<%
if Show_EC_Date_Style = true then
%>
function click_EC(ex_num)
{
	var url_exstr;
	var tmp_exstr = "<%=exstr %>";
	if (tmp_exstr.charAt(ex_num) == '0')
	{
		if (ex_num > 0)
			url_exstr = tmp_exstr.substring(0, ex_num) + "1" + tmp_exstr.substring(ex_num + 1);
		else
			url_exstr = "1" + tmp_exstr.substring(ex_num + 1);
	}
	else
	{
		if (ex_num > 0)
			url_exstr = tmp_exstr.substring(0, ex_num) + "0" + tmp_exstr.substring(ex_num + 1);
		else
			url_exstr = '0' + tmp_exstr.substring(ex_num + 1);
	}

	location.href = "<%
Response.Write gourl

mid_len = InStr(addsortstr, "&exstr=")
if mid_len > 0 then
	Response.Write Mid(addsortstr, 1, mid_len - 1)
	Response.Write Mid(addsortstr, mid_len + 14)
end if
%>" + "&exstr=" + url_exstr + "&<%=getGRSN() %>";
}
<%
end if
%>
</script>

</BODY>
</HTML>

<%
function writeDateEC(mail_index)
	writeDateEC = true

	if Show_EC_Date_Style = true then
		ei.GetDateExMode exstr, mail_index, exmode, isshow

		if exmode = 0 then
			if show_EC_mode <> 0 then 
				show_EC_mode = 0

				if show_page_ex_head = true then
					Response.Write "<tr valign='bottom'><td colspan=7 height='21' nowrap class='EX_TITLE'><a href='javascript:click_EC(0)'><img src='images/"
					if isshow = true then
						Response.Write "listclose.gif' border='0' title='" & s_lang_0057 & "'"
					else
						Response.Write "listopen.gif' border='0' title='" & s_lang_0058 & "'"
						writeDateEC = false
					end if

					Response.Write "></a> <font class='EX_TITLE_FONT'>" & s_lang_0134 & "</font></td></tr>"
				else
					show_page_ex_head = true
				end if

				is_show_EC = writeDateEC
			else
				writeDateEC = is_show_EC
			end if
		elseif exmode = 1 then
			if show_EC_mode <> 1 then 
				show_EC_mode = 1

				if show_page_ex_head = true then
					Response.Write "<tr valign='bottom'><td colspan=7 height='21' nowrap class='EX_TITLE'><a href='javascript:click_EC(1)'><img src='images/"
					if isshow = true then
						Response.Write "listclose.gif' border='0' title='" & s_lang_0057 & "'"
					else
						Response.Write "listopen.gif' border='0' title='" & s_lang_0058 & "'"
						writeDateEC = false
					end if

					Response.Write "></a> <font class='EX_TITLE_FONT'>" & s_lang_0135 & "</font></td></tr>"
				else
					show_page_ex_head = true
				end if

				is_show_EC = writeDateEC
			else
				writeDateEC = is_show_EC
			end if
		elseif exmode = 2 then
			if show_EC_mode <> 2 then 
				show_EC_mode = 2

				if show_page_ex_head = true then
					Response.Write "<tr valign='bottom'><td colspan=7 height='21' nowrap class='EX_TITLE'><a href='javascript:click_EC(2)'><img src='images/"
					if isshow = true then
						Response.Write "listclose.gif' border='0' title='" & s_lang_0057 & "'"
					else
						Response.Write "listopen.gif' border='0' title='" & s_lang_0058 & "'"
						writeDateEC = false
					end if

					Response.Write "></a> <font class='EX_TITLE_FONT'>" & s_lang_0136 & "</font></td></tr>"
				else
					show_page_ex_head = true
				end if

				is_show_EC = writeDateEC
			else
				writeDateEC = is_show_EC
			end if
		elseif exmode = 3 then
			if show_EC_mode <> 3 then 
				show_EC_mode = 3

				if show_page_ex_head = true then
					Response.Write "<tr valign='bottom'><td colspan=7 height='21' nowrap class='EX_TITLE'><a href='javascript:click_EC(3)'><img src='images/"
					if isshow = true then
						Response.Write "listclose.gif' border='0' title='" & s_lang_0057 & "'"
					else
						Response.Write "listopen.gif' border='0' title='" & s_lang_0058 & "'"
						writeDateEC = false
					end if

					Response.Write "></a> <font class='EX_TITLE_FONT'>" & s_lang_0137 & "</font></td></tr>"
				else
					show_page_ex_head = true
				end if

				is_show_EC = writeDateEC
			else
				writeDateEC = is_show_EC
			end if
		elseif exmode = 4 then
			if show_EC_mode <> 4 then 
				show_EC_mode = 4

				if show_page_ex_head = true then
					Response.Write "<tr valign='bottom'><td colspan=7 height='21' nowrap class='EX_TITLE'><a href='javascript:click_EC(4)'><img src='images/"
					if isshow = true then
						Response.Write "listclose.gif' border='0' title='" & s_lang_0057 & "'"
					else
						Response.Write "listopen.gif' border='0' title='" & s_lang_0058 & "'"
						writeDateEC = false
					end if

					Response.Write "></a> <font class='EX_TITLE_FONT'>" & s_lang_0138 & "</font></td></tr>"
				else
					show_page_ex_head = true
				end if

				is_show_EC = writeDateEC
			else
				writeDateEC = is_show_EC
			end if
		end if

		exmode = NULL
		isshow = NULL
	end if
end function


function getBeforeLines(tmp_mode)
	before_num = 0
	tmp_tmp_mode = tmp_mode

	do while tmp_mode >= 0
		if Mid(exstr, tmp_mode + 1, 1) = "1" then
			if tmp_mode = 4 then
				before_num = before_num + ei.DateEX_4
			elseif tmp_mode = 3 then
				before_num = before_num + ei.DateEX_3
			elseif tmp_mode = 2 then
				before_num = before_num + ei.DateEX_2
			elseif tmp_mode = 1 then
				before_num = before_num + ei.DateEX_1
			elseif tmp_mode = 0 then
				before_num = before_num + ei.DateEX_0
			end if
		end if
		tmp_mode = tmp_mode - 1
	loop

	if Mid(exstr, tmp_tmp_mode + 1, 1) = "1" then
		tmp_mode = tmp_tmp_mode + 1
		do while tmp_mode < 5
			if Mid(exstr, tmp_mode + 1, 1) = "1" then
				if tmp_mode = 4 then
					before_num = before_num + ei.DateEX_4
				elseif tmp_mode = 3 then
					before_num = before_num + ei.DateEX_3
				elseif tmp_mode = 2 then
					before_num = before_num + ei.DateEX_2
				elseif tmp_mode = 1 then
					before_num = before_num + ei.DateEX_1
				elseif tmp_mode = 0 then
					before_num = before_num + ei.DateEX_0
				end if
			else
				exit do
			end if
			tmp_mode = tmp_mode + 1
		loop
	end if

	getBeforeLines = before_num
end function


function canshowlast(tmp_last_exmode)
	canshowlast = true

	do while tmp_last_exmode < 5
		if Mid(exstr, tmp_last_exmode + 1, 1) = "0" and getDateMails(tmp_last_exmode) > 0 then
			canshowlast = false
			exit do
		end if
		tmp_last_exmode = tmp_last_exmode + 1
	loop
end function


function getDateMails(sea_ex_mode)
	getDateMails = 0

	if sea_ex_mode = 0 then
		getDateMails = ei.DateEX_0
	elseif sea_ex_mode = 1 then
		getDateMails = ei.DateEX_1
	elseif sea_ex_mode = 2 then
		getDateMails = ei.DateEX_2
	elseif sea_ex_mode = 3 then
		getDateMails = ei.DateEX_3
	elseif sea_ex_mode = 4 then
		getDateMails = ei.DateEX_4
	end if
end function


function writeBeforeHiddenTitle(start_exmode)
	tmp_start_exmode = -1
	tmp_end_exmode = start_exmode

	do while start_exmode >= 0
		if Mid(exstr, start_exmode + 1, 1) = "0" then
			exit do
		else
			tmp_start_exmode = start_exmode
		end if
		start_exmode = start_exmode - 1
	loop

	if tmp_start_exmode > 0 and tmp_start_exmode <= tmp_end_exmode then
		do while tmp_start_exmode <= tmp_end_exmode
			writeHiddenTitle(tmp_start_exmode)
			tmp_start_exmode = tmp_start_exmode + 1
		loop
	end if
end function


function writeHiddenTitle(write_exmode)
	Response.Write "<tr valign='bottom'><td colspan=7 height='21' nowrap class='EX_TITLE'><a href='javascript:click_EC(" & write_exmode & ")'><img src='images/"
	Response.Write "listopen.gif' border='0' title='" & s_lang_0058 & "'></a> <font class='EX_TITLE_FONT'>"

	if write_exmode = 0 then
		Response.Write s_lang_0134
	elseif write_exmode = 1 then
		Response.Write s_lang_0135
	elseif write_exmode = 2 then
		Response.Write s_lang_0136
	elseif write_exmode = 3 then
		Response.Write s_lang_0137
	elseif write_exmode = 4 then
		Response.Write s_lang_0138
	end if

	Response.Write "</font></td></tr>"
end function


function get_date_showstr(show_date_str)
	if Len(show_date_str) = 14 then
		tmp_month = Mid(show_date_str, 5, 2)
		if Mid(tmp_month, 1, 1) = "0" then
			tmp_month = Mid(tmp_month, 2, 1)
		end if

		tmp_day = Mid(show_date_str, 7, 2)
		if Mid(tmp_day, 1, 1) = "0" then
			tmp_day = Mid(tmp_day, 2, 1)
		end if

		get_date_showstr = Mid(show_date_str, 1, 4) & s_lang_0139 & tmp_month & s_lang_0140 & tmp_day & s_lang_0141 & " " & Mid(show_date_str, 9, 2) & ":" & Mid(show_date_str, 11, 2) & ":" & Mid(show_date_str, 13, 2)
	else
		get_date_showstr = ""
	end if
end function

set ei = nothing
%>
