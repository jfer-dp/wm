<!--#include file="passinc.asp" -->

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

viewwho = trim(request("viewwho"))

if Len(Request.QueryString) < 13 then
	Session("svcal") = ""
end if

if Len(viewwho) > 0 then
	Session("svcal") = viewwho
end if

sy = trim(request("sy"))
sm = trim(request("sm"))
sd = trim(request("sd"))

start_year = trim(request("stay"))
start_month = trim(request("stam"))

if Len(sy) < 1 or Len(sm) < 1 or Len(sd) < 1 then
	my_Date = Now

	sy = Year(my_Date)
	sm = Month(my_Date)
	sd = Day(my_Date)
end if

if Len(start_year) < 1 or Len(start_month) < 1 then
	start_year = sy
	start_month = sm
end if


dim ecalset
set ecalset = server.createobject("easymail.CalOptions")
ecalset.Load Session("wem")

if Len(Session("svcal")) > 0 then
	if ecalset.haveitFavorite(Session("svcal")) = false then
		Session("svcal") = ""
	else
		dim ecalset_tmp
		set ecalset_tmp = server.createobject("easymail.CalOptions")

		if ecalset_tmp.Load(Session("svcal")) = false then
			Session("svcal") = ""
		else
			if ecalset_tmp.MyCalendarViewState = 0 then
				Session("svcal") = ""
			end if

			if Len(Session("svcal")) > 0 and ecalset_tmp.MyCalendarViewState = 1 and ecalset_tmp.haveitFriend(Session("wem")) = false then
				Session("svcal") = ""
			end if
		end if

		set ecalset_tmp = nothing
	end if
end if

if Len(viewwho) > 0 and Session("svcal") = "" then
	set ecalset = nothing
	Response.Redirect "err.asp?" & getGRSN() & "&errstr=" & Server.URLEncode("查看用户 " & viewwho & " 的效率手册失败")
end if


show_APM = false
if ecalset.Show24Hour = false then
	show_APM = true
end if

dim tab_selected_num
tab_selected_num = "0"

tab_selected_num = trim(request("tsn"))
if Len(tab_selected_num) <> 1 then
	tab_selected_num = "0"
end if

sortby = trim(request("sortby"))

if tab_selected_num = "4" then
	if Len(sortby) > 0 then
		if IsNumeric(sortby) = true then
			sortby = CLng(sortby)

			if sortby < 0 then
				sortby = 0
			end if
		else
			sortby = 0
		end if
	else
		sortby = 0
	end if
elseif tab_selected_num = "5" then
	if Len(sortby) > 0 then
		if IsNumeric(sortby) = true then
			sortby = CLng(sortby)

			if sortby < 0 then
				sortby = 3
			end if
		else
			sortby = 3
		end if
	else
		sortby = 3
	end if
end if

sortmode = trim(request("sortmode"))
if sortmode <> "1" then
	sortmode = "0"
end if

dim ecal
set ecal = server.createobject("easymail.Calendar")

if tab_selected_num = "4" then
	if Len(Session("svcal")) < 1 then
		ecal.SortListMode = sortby
	else
		ecal.SortListMode = 0
	end if
end if

if Len(Session("svcal")) < 1 then
	ecal.Load Session("wem")
else
	ecal.Load Session("svcal")
	ecal.HidePrivate
end if


dim ectk
set ectk = server.createobject("easymail.CalTask")
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
body {font-family:<%=s_lang_font %>; font-size:9pt;color:#000000;margin-top:5px;margin-left:10px;margin-right:10px;margin-bottom:2px;background-color:#ffffff}
.sbttn {font-family:<%=s_lang_font %>; font-size:9pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5;color:#000066;height:19px;text-decoration:none;cursor:pointer}
.textbox {BORDER:1px #555555 solid;}
.b_td {white-space:nowrap; text-align:left; padding-left:14px;}
.st_l,.st_r {text-align:center; white-space:nowrap; border-left:1px solid #A5B6C8; border-top:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8;}
.st_r {border-right:1px solid #A5B6C8;}
.cont_tr {background:white; height:22px;}
.cont_td {height:22px; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px;}
.urf {color:black;}
.urf:hover {color:black;}
.ttf1 {text-align:center;  white-space:nowrap; background-color:#93BEE2; color:black; font-weight:bold;}
.ttf2 {text-align:center;  white-space:nowrap; background-color:#DBEAF5; color:black; cursor:pointer;}

a:hover {text-decoration:underline;}

.mjNoLine {
 	text-decoration: none; 
}
.mjLinkLeft {
	color: #447172;
 	text-decoration: none; 
<%
if isMSIE = true then
	Response.Write "CURSOR: Hand;"
else
	Response.Write "CURSOR: pointer;"
end if
%>
}
.mjLink {
	color: #002f72;
 	text-decoration: none; 
<%
if isMSIE = true then
	Response.Write "CURSOR: Hand;"
else
	Response.Write "CURSOR: pointer;"
end if
%>
}
.calendar_dayname {
	BORDER-TOP: #ffffc0 7px solid;
	BORDER-LEFT: #ffffc0 5px solid;
	BORDER-BOTTOM: #ffffc0 3px solid;
	FONT-WEIGHT: normal;
	color: #202020;
	BACKGROUND-COLOR: #ffffc0;
}
.mjCalTitleBottom {
	BORDER-TOP: #c0c0c0 1px solid;
}
.calendar_dayname_dm {
	BORDER-TOP: #f0f0f0 3px solid;
	COLOR: #000000;
	BACKGROUND-COLOR: #f0f0f0;
}
.calendar_today_left {
	COLOR: #000000; BACKGROUND-COLOR: #e0f0f0; 
}
.calendar_today {
	FONT-WEIGHT: bold; COLOR: #000000; BACKGROUND-COLOR: #e0f0f0; 
}
.calendar_weekline {
	 COLOR: #002f72; BACKGROUND-COLOR: #c3d7d7
}
.mjWeekLink {
	color: #447172;
 	text-decoration: none; 
<%
if isMSIE = true then
	Response.Write "CURSOR: Hand;"
else
	Response.Write "CURSOR: pointer;"
end if
%>
}
.mjPrevLink {
	color: #c3d7d7;
 	text-decoration: none; 
<%
if isMSIE = true then
	Response.Write "CURSOR: Hand;"
else
	Response.Write "CURSOR: pointer;"
end if
%>
}
.calendar_main {
	BACKGROUND-COLOR: #e0e0e0;
}
.calendar_main_y {
	BACKGROUND-COLOR: #ffffff;
}
.calendar_main_bm {
	BACKGROUND-COLOR: #c0c0c0;
}
.calendar_title {
	FONT-WEIGHT: bold; COLOR: #000000; BACKGROUND-COLOR: #ffffff
}
.calendar_nav {
<%
if isMSIE = true then
	Response.Write "CURSOR: Hand;"
else
	Response.Write "CURSOR: pointer;"
end if
%>
}
.calendar_dayname_left {
	COLOR: #000000;
}
.calendar_dayname_y {
	COLOR: #000000; BACKGROUND-COLOR: #f0f0f0;
}
.calendar_daybg_left {
	 COLOR: #002f72; BACKGROUND-COLOR: #ffffff
}
.calendar_daybg {
	 FONT-WEIGHT: bold; COLOR: #002f72; BACKGROUND-COLOR: #ffffff
}
.calendar_daybg {
	 COLOR: #002f72; BACKGROUND-COLOR: #ffffff
}
.calendar_today {
	FONT-WEIGHT: bold; COLOR: #000000; BACKGROUND-COLOR: #e0f0f0; 
}
.calendar_grid {
	FONT-WEIGHT: normal;
}
.mjAddLink {
	font-size: 9pt;
	FONT-WEIGHT: normal;
	color: #447172;
 	text-decoration: none; 
<%
if isMSIE = true then
	Response.Write "CURSOR: Hand;"
else
	Response.Write "CURSOR: pointer;"
end if
%>
}
.mjEL {
	font-size: 9pt;
	color: #447172;
 	text-decoration: none; 
<%
if isMSIE = true then
	Response.Write "CURSOR: Hand;"
else
	Response.Write "CURSOR: pointer;"
end if
%>
}
.NLshow{
	font-size: 9pt;
	FONT-WEIGHT: normal;
	color: #c0c0c0;
 	text-decoration: none;
}
.calmy {
	background-color:<%=MY_COLOR_3 %>;
	border-color:#99b5b6;
}
.calendar_weekbg {
	FONT-WEIGHT: normal; COLOR: #002f72; BACKGROUND-COLOR: #c3d7d7;
	BORDER-TOP: #c0c0c0 1px solid;
}
.mjEvent {
	BACKGROUND-COLOR:#edf9f9;
	font-size: 9pt;
	FONT-WEIGHT: normal;
	color: #101010;
 	text-decoration: none;
}
.mjMarkLink {
	color: #ff6633;
 	text-decoration: none; 
<%
if isMSIE = true then
	Response.Write "CURSOR: Hand;"
else
	Response.Write "CURSOR: pointer;"
end if
%>
}
.mjYGrid {
	border-top:1px #c0c0c0 solid;
	border-right:1px #c0c0c0 solid;
}
.mjDMLine {
	BORDER-TOP: #a0a0a0 1px solid;
}
-->
</STYLE>
</head>

<script type="text/javascript" src="images/sc_left.js"></script>
<script type="text/javascript" src="images/mglobal.js"></script>

<script language="JavaScript">
<!--
if (ie == 6)
	document.execCommand("BackgroundImageCache", false, true);

var sFtv = new Array(
"0101 元旦",
"0214 情人节",
"0308 妇女节",
"0312 植树节",
"0401 愚人节",
"0405 清明节",
"0501 劳动节",
"0504 青年节",
"0512 护士节",
"0601 儿童节",
"0701 建党节",
"0801 建军节",
"0910 教师节",
"0928 孔子诞辰",
"1001 国庆节",
"1006 老人节",
"1225 圣诞节"<%
i = 0
allnum = ecalset.Count
do while i < allnum
	ecalset.Get i, mm, dd, fname
	Response.Write "," & Chr(13) & """" & convFeast(mm, dd, fname) & """"

	mm = NULL
	dd = NULL
	fname = NULL

	i = i + 1
loop


dim ecalsysfet
set ecalsysfet = server.createobject("easymail.CalSystemFeast")
ecalsysfet.Load

i = 0
allnum = ecalsysfet.Count
do while i < allnum
	ecalsysfet.Get i, mm, dd, fname
	Response.Write "," & Chr(13) & """" & convFeast(mm, dd, fname) & """"

	mm = NULL
	dd = NULL
	fname = NULL

	i = i + 1
loop


dim ecaldomainfet
set ecaldomainfet = server.createobject("easymail.CalDomainFeasts")
ecaldomainfet.Load Mid(Session("mail"), InStr(Session("mail"), "@") + 1)

i = 0
allnum = ecaldomainfet.Count
do while i < allnum
	ecaldomainfet.Get i, mm, dd, fname
	Response.Write "," & Chr(13) & """" & convFeast(mm, dd, fname) & """"

	mm = NULL
	dd = NULL
	fname = NULL

	i = i + 1
loop
%>);


var lFtv = new Array(
"0101 春节",
"0115 元宵节",
"0505 端午节",
"0707 七夕情人节",
"0815 中秋节",
"0909 重阳节",
"1224 小年"<%
i = 0
allnum = ecalset.CountNL
do while i < allnum
	ecalset.GetNL i, mm, dd, fname
	Response.Write "," & Chr(13) & """" & convFeast(mm, dd, fname) & """"

	mm = NULL
	dd = NULL
	fname = NULL

	i = i + 1
loop


i = 0
allnum = ecalsysfet.CountNL
do while i < allnum
	ecalsysfet.GetNL i, mm, dd, fname
	Response.Write "," & Chr(13) & """" & convFeast(mm, dd, fname) & """"

	mm = NULL
	dd = NULL
	fname = NULL

	i = i + 1
loop

set ecalsysfet = nothing


i = 0
allnum = ecaldomainfet.CountNL
do while i < allnum
	ecaldomainfet.GetNL i, mm, dd, fname
	Response.Write "," & Chr(13) & """" & convFeast(mm, dd, fname) & """"

	mm = NULL
	dd = NULL
	fname = NULL

	i = i + 1
loop

set ecaldomainfet = nothing
%>);

var show_ext = <%=LCase(CStr(ecalset.ShowDayExt)) %>;
var show_g_feast = <%=LCase(CStr(ecalset.ShowFeasts)) %>;
var show_n_feast = <%=LCase(CStr(ecalset.ShowNLFeasts)) %>;
var show_nl = <%=LCase(CStr(ecalset.ShowNL)) %>;
var show_APM = <%=LCase(CStr(show_APM)) %>;

function get_APM(vsh, vsm)
{
	var t_str = ""
	if (show_APM == false)
	{
		t_str = vsh + ":";

		if (vsm < 10)
			t_str = t_str + "0";
		t_str = t_str + vsm;
	}
	else
	{
		if (vsm < 10)
			t_str = "0";
		t_str = t_str + vsm;

		if (vsh == 0)
			t_str = "12:" + t_str + "AM";
		else if (vsh == 12)
			t_str = "12:" + t_str + "PM";
		else if (vsh < 12)
			t_str = vsh + ":" + t_str + "AM";
		else
			t_str = vsh + ":" + t_str + "PM";
	}

	return t_str;
}

function getShowStartStr(vy, vm, vd, vh, vmin, vnt)
{
	if (vnt == 0)
		return "全天"

	var s_str = "";
	currentDate = new Date(vy, vm - 1, vd, vh, vmin);

	s_str = get_APM(currentDate.getHours(), currentDate.getMinutes()) + "-";
	currentDate.setTime(currentDate.getTime() + (vnt * 1000));
	s_str = s_str + get_APM(currentDate.getHours(), currentDate.getMinutes());

	return s_str;
}

<%
if tab_selected_num = "0" then
%>
var strDatesEvents_date = new Array(25);
var strDatesEvents_msg = new Array(25);
var strDatesAllDayEvents_msg = ""

var i = 0;
for (i = 0; i < 25; i++)
{
	strDatesEvents_msg[i] = "";
	strDatesEvents_date[i] = "";
}
<%
	date_str = sy

	if sm < 10 then
		date_str = date_str & "0"
	end if

	date_str = date_str & sm

	if sd < 10 then
		date_str = date_str & "0"
	end if

	date_str = date_str & sd

	if ecal.GetEvents(date_str, date_str, true, false) = true then
		if Len(Session("svcal")) > 0 then
			ecal.Event_HidePrivate
		end if

		allnum = ecal.Event_Count
		i = 0
		tmp_hour = 0
		tmp_allday_str = ""
		tmp_hour_str = ""

		do while i < allnum
			ecal.Event_MoveTo i
			ecal.get_ev_bi_start_date b_year, b_month, b_day, b_hour, b_minute

			if ecal.ev_bi_needtime > 0 then
				if tmp_hour = b_hour then
					if Len(tmp_hour_str) > 1 then
						tmp_hour_str = tmp_hour_str & "<td>&nbsp;</td>"
					end if
				else
					if Len(tmp_hour_str) > 20 then
						Response.Write "strDatesEvents_date[" & tmp_hour & "] = """ & tmp_hour & """;" & Chr(13)
						Response.Write "strDatesEvents_msg[" & tmp_hour & "] = """ & tmp_hour_str & """;" & Chr(13)
					end if

					tmp_hour = b_hour
					tmp_hour_str = "<td>&nbsp;</td>"
				end if

				if Len(Session("svcal")) < 1 or ecal.ev_bi_shareMode = 2 then
					tmp_hour_str = tmp_hour_str & "<td class=calmy>"" + getShowStartStr(" & b_hour & "," & b_minute & "," & b_day & "," & b_hour & "," & b_minute & "," & ecal.ev_bi_needtime & ") + ""&nbsp;"" + getShowIconStr(" & ecal.ev_bi_mode & "," & LCase(CStr(ecal.ev_bi_remind)) & "," & LCase(CStr(ecal.ev_bi_isRepeat)) & ") + ""<a class=mjLink href=\""javascript:showevent('" & ecal.ev_id & "')\"">" & server.htmlencode(ecal.ev_bi_name) & "</a>&nbsp;"" + getShowUsersIconStr('" & ecal.ev_id & "'," & ecal.ev_Yes_User & "," & ecal.ev_Wait_User & "," & ecal.ev_No_User & ") + ""[<a class=mjEL href=\""javascript:delevent('" & ecal.ev_id & "')\"">删除</a>]</td>"
				else
					tmp_hour_str = tmp_hour_str & "<td class=calmy>"" + getShowStartStr(" & b_hour & "," & b_minute & "," & b_day & "," & b_hour & "," & b_minute & "," & ecal.ev_bi_needtime & ") + ""&nbsp;忙碌</td>"
				end if
			else
				if Len(tmp_allday_str) > 1 then
					tmp_allday_str = tmp_allday_str & "<td>&nbsp;</td>"
				else
					tmp_allday_str = "<table border=0 cellspacing=0 cellpadding=1><tr>"
				end if

				if Len(Session("svcal")) < 1 or ecal.ev_bi_shareMode = 2 then
					tmp_allday_str = tmp_allday_str & "<td class=calmy>"" + getShowStartStr(" & b_hour & "," & b_minute & "," & b_day & "," & b_hour & "," & b_minute & "," & ecal.ev_bi_needtime & ") + ""&nbsp;"" + getShowIconStr(" & ecal.ev_bi_mode & "," & LCase(CStr(ecal.ev_bi_remind)) & "," & LCase(CStr(ecal.ev_bi_isRepeat)) & ") + ""<a class=mjLink href=\""javascript:showevent('" & ecal.ev_id & "')\"">" & server.htmlencode(ecal.ev_bi_name) & "</a>&nbsp;"" + getShowUsersIconStr('" & ecal.ev_id & "'," & ecal.ev_Yes_User & "," & ecal.ev_Wait_User & "," & ecal.ev_No_User & ") + ""[<a class=mjEL href=\""javascript:delevent('" & ecal.ev_id & "')\"">删除</a>]</td>"
				else
					tmp_allday_str = tmp_allday_str & "<td class=calmy>"" + getShowStartStr(" & b_hour & "," & b_minute & "," & b_day & "," & b_hour & "," & b_minute & "," & ecal.ev_bi_needtime & ") + ""&nbsp;忙碌</td>"
				end if
			end if

			b_year = NULL
			b_month = NULL
			b_day = NULL
			b_hour = NULL
			b_minute = NULL

			i = i + 1
		loop

		if Len(tmp_hour_str) > 20 then
			Response.Write "strDatesEvents_date[" & tmp_hour & "] = """ & tmp_hour & """;" & Chr(13)
			Response.Write "strDatesEvents_msg[" & tmp_hour & "] = """ & tmp_hour_str & """;" & Chr(13)
		end if

		if Len(tmp_allday_str) > 1 then
			tmp_allday_str = tmp_allday_str & "</tr></table>"
			Response.Write "strDatesAllDayEvents_msg = """ & tmp_allday_str & """;" & Chr(13)
		end if
	end if
%>
function getEventIndex(sdate)
{
	var gi = 0;
	for (gi = 0; gi < 25; gi++)
	{
		if (sdate == strDatesEvents_date[gi])
			return gi;
	}

	return -1;
}

<%
elseif tab_selected_num = "1" then
%>
var strDatesEvents_date = new Array(8);
var strDatesEvents_msg = new Array(8);

var i = 0;
for (i = 0; i < 8; i++)
{
	strDatesEvents_msg[i] = "";
	strDatesEvents_date[i] = "";
}

<%
	date_str = sy

	if sm < 10 then
		date_str = date_str & "0"
	end if

	date_str = date_str & sm

	if sd < 10 then
		date_str = date_str & "0"
	end if

	date_str = date_str & sd

	if ecal.GetEvents(date_str, "", true, false) = true then
		if Len(Session("svcal")) > 0 then
			ecal.Event_HidePrivate
		end if

		allnum = ecal.Event_Count
		i = 0
		tmp_index = 0
		tmp_date_str = sy & "_" & sm & "_" & sd
		tmp_str = ""

		do while i < allnum
			ecal.Event_MoveTo i
			ecal.get_ev_bi_start_date b_year, b_month, b_day, b_hour, b_minute
			tmp_date_str_1 = b_year & "_" & b_month & "_" & b_day

			if tmp_date_str = tmp_date_str_1 then
				tmp_str = tmp_str & "<td>&nbsp;</td>"
			else
				if Len(tmp_str) > 20 then
					Response.Write "strDatesEvents_date[" & tmp_index & "] = """ & tmp_date_str & """;" & Chr(13)
					Response.Write "strDatesEvents_msg[" & tmp_index & "] = """ & tmp_str & """;" & Chr(13)
				end if

				tmp_date_str = tmp_date_str_1
				tmp_str = "<td>&nbsp;</td>"

				tmp_index = tmp_index + 1
			end if

			if Len(Session("svcal")) < 1 or ecal.ev_bi_shareMode = 2 then
				tmp_str = tmp_str & "<td class=calmy>"" + getShowStartStr(" & b_hour & "," & b_minute & "," & b_day & "," & b_hour & "," & b_minute & "," & ecal.ev_bi_needtime & ") + ""&nbsp;"" + getShowIconStr(" & ecal.ev_bi_mode & "," & LCase(CStr(ecal.ev_bi_remind)) & "," & LCase(CStr(ecal.ev_bi_isRepeat)) & ") + ""<a class=mjLink href=\""javascript:showevent('" & ecal.ev_id & "')\"">" & server.htmlencode(ecal.ev_bi_name) & "</a>&nbsp;"" + getShowUsersIconStr('" & ecal.ev_id & "'," & ecal.ev_Yes_User & "," & ecal.ev_Wait_User & "," & ecal.ev_No_User & ") + ""[<a class=mjEL href=\""javascript:delevent('" & ecal.ev_id & "')\"">删除</a>]</td>"
			else
				tmp_str = tmp_str & "<td class=calmy>"" + getShowStartStr(" & b_hour & "," & b_minute & "," & b_day & "," & b_hour & "," & b_minute & "," & ecal.ev_bi_needtime & ") + ""&nbsp;忙碌</td>"
			end if

			b_year = NULL
			b_month = NULL
			b_day = NULL
			b_hour = NULL
			b_minute = NULL

			i = i + 1
		loop

		if Len(tmp_str) > 20 then
			Response.Write "strDatesEvents_date[" & tmp_index & "] = """ & tmp_date_str_1 & """;" & Chr(13)
			Response.Write "strDatesEvents_msg[" & tmp_index & "] = """ & tmp_str & """;" & Chr(13)
		end if
	end if
%>

function getEventIndex(sdate)
{
	var gi = 0;
	for (gi = 0; gi < 8; gi++)
	{
		if (sdate == strDatesEvents_date[gi])
			return gi;
	}

	return -1;
}
<%
elseif tab_selected_num = "2" then
%>
var strDatesEvents_date = new Array(50);
var strDatesEvents_msg = new Array(50);

var i = 0;
for (i = 0; i < 50; i++)
{
	strDatesEvents_msg[i] = "";
	strDatesEvents_date[i] = "";
}

<%
	if Len(start_year) > 3 and Len(start_month) > 0 then
		date_str = start_year

		if Len(start_month) = 1 then
			date_str = date_str & "0"
		end if

		date_str = date_str & start_month
	else
		date_str = sy

		if sm < 10 then
			date_str = date_str & "0"
		end if

		date_str = date_str & sm
	end if

	if ecal.GetEvents(date_str, "", true, true) = true then
		if Len(Session("svcal")) > 0 then
			ecal.Event_HidePrivate
		end if

		allnum = ecal.Event_Count
		i = 0
		tmp_index = 0
		tmp_date_str = sy & "_" & sm & "_" & sd
		tmp_str = ""

		do while i < allnum
			ecal.Event_MoveTo i
			ecal.get_ev_bi_start_date b_year, b_month, b_day, b_hour, b_minute
			tmp_date_str_1 = b_year & "_" & b_month & "_" & b_day

			if tmp_date_str = tmp_date_str_1 then
				if Len(tmp_str) < 1 then
					tmp_str = "<table width='100%' cellspacing=0 cellpadding=1 border=0>"
				end if
			else
				if Len(tmp_str) > 50 then
					Response.Write "strDatesEvents_date[" & tmp_index & "] = """ & tmp_date_str & """;" & Chr(13)
					Response.Write "strDatesEvents_msg[" & tmp_index & "] = ""<div class=mjEvent>" & tmp_str & "</table></div>"";" & Chr(13)
				end if

				tmp_date_str = tmp_date_str_1
				tmp_str = "<table width='100%' cellspacing=0 cellpadding=1 border=0>"

				tmp_index = tmp_index + 1
			end if

			if Len(Session("svcal")) < 1 or ecal.ev_bi_shareMode = 2 then
				tmp_str = tmp_str & "<tr><td class=mjDMLine>"" + getShowStartStr(" & b_hour & "," & b_minute & "," & b_day & "," & b_hour & "," & b_minute & "," & ecal.ev_bi_needtime & ") + ""&nbsp;"" + getShowIconStr(" & ecal.ev_bi_mode & "," & LCase(CStr(ecal.ev_bi_remind)) & "," & LCase(CStr(ecal.ev_bi_isRepeat)) & ") + ""<a class=mjLink href=\""javascript:showevent('" & ecal.ev_id & "')\"">" & server.htmlencode(ecal.ev_bi_name) & "</a>&nbsp;"" + getShowUsersIconStr('" & ecal.ev_id & "'," & ecal.ev_Yes_User & "," & ecal.ev_Wait_User & "," & ecal.ev_No_User & ") + ""[<a class=mjEL href=\""javascript:delevent('" & ecal.ev_id & "')\"">删除</a>]</td></tr>"
			else
				tmp_str = tmp_str & "<tr><td class=mjDMLine>"" + getShowStartStr(" & b_hour & "," & b_minute & "," & b_day & "," & b_hour & "," & b_minute & "," & ecal.ev_bi_needtime & ") + ""&nbsp;忙碌</tr>"
			end if

			b_year = NULL
			b_month = NULL
			b_day = NULL
			b_hour = NULL
			b_minute = NULL

			i = i + 1
		loop

		if Len(tmp_str) > 50 then
			Response.Write "strDatesEvents_date[" & tmp_index & "] = """ & tmp_date_str_1 & """;" & Chr(13)
			Response.Write "strDatesEvents_msg[" & tmp_index & "] = ""<div class=mjEvent>" & tmp_str & "</table></div>"";" & Chr(13)
		end if
	end if
%>

function getEventIndex(sdate)
{
	var gi = 0;
	for (gi = 0; gi < 50; gi++)
	{
		if (sdate == strDatesEvents_date[gi])
			return gi;
	}

	return -1;
}
<%
elseif tab_selected_num = "3" then
%>
var monstep = 1;
<%
	tmp_y_sy = sy
	if Len(start_year) > 3 then
		tmp_y_sy = start_year
	end if

	if ecal.GetEvents(tmp_y_sy, "", false, false) = true then
		if Len(Session("svcal")) > 0 then
			ecal.Event_HidePrivate
		end if

		allnum = ecal.Event_Count
		i = 0
		tmp_str = ","

		do while i < allnum
			ecal.Event_MoveTo i
			ecal.get_ev_bi_start_date b_year, b_month, b_day, b_hour, b_minute
			tmp_str = tmp_str & b_year & "_" & b_month & "_" & b_day & "," 

			b_year = NULL
			b_month = NULL
			b_day = NULL
			b_hour = NULL
			b_minute = NULL

			i = i + 1
		loop
	end if

	Response.Write "var strDatesHasPost = """ & tmp_str & """" & Chr(13)
end if
%>


var Stag;

var curyear = <%=Year(Now) %>;
var curmon = <%=Month(Now) %>;
var curday = <%=Day(Now) %>;

var selyear = <%=sy %>;
var selmon = <%=sm %>;
var selday = <%=sd %>;

var start_year = <%=start_year %>;
var start_month = <%=start_month %>;

var showWeekLine = false;
<%
if tab_selected_num = "1" then
%>
showWeekLine = true;
<%
end if
%>

var StartWeekDay = <%=ecalset.StartWeekDay %>;

<%
gourl = "cal_index.asp?" & getGRSN() & "&tsn=" & tab_selected_num
bakurl = "cal_index.asp?" & getGRSN() & "&tsn=" & tab_selected_num & "&sy=" & sy & "&sm=" & sm & "&sd=" & sd & "&stay=" & start_year & "&stam=" & start_month
%>
function Selected_One(selected_year, selected_month, selected_day, selected_tab)
{
	if (selected_tab < 0 || selected_tab > 3)
	{
		if (showWeekLine == false)
			location.href = "cal_index.asp?<%=getGRSN() %>&tsn=0&sy=" + selected_year + "&sm=" + selected_month + "&sd=" + selected_day;
		else
			location.href = "<%=gourl %>&sy=" + selected_year + "&sm=" + selected_month + "&sd=" + selected_day;
	}
	else
	{
		if (selected_tab == 2 || selected_tab == 3)
		{
			if (selected_year < 1972)
			{
				start_year = 1972;
				return ;
			}

			location.href = "cal_index.asp?<%=getGRSN() %>&tsn=" + selected_tab + "&sy=" + selyear + "&sm=" + selmon + "&sd=" + selday + "&stay=" + selected_year + "&stam=" + selected_month;
		}
		else
			location.href = "cal_index.asp?<%=getGRSN() %>&tsn=" + selected_tab + "&sy=" + selected_year + "&sm=" + selected_month + "&sd=" + selected_day;
	}
}


function change_tab(t_tab_num)
{
	if (t_tab_num == 2 || t_tab_num == 3)
		location.href = "cal_index.asp?<%=getGRSN() %>&tsn=" + t_tab_num + "&sy=" + selyear + "&sm=" + selmon + "&sd=" + selday + "&stay=" + start_year + "&stam=" + start_month;
	else
		location.href = "cal_index.asp?<%=getGRSN() %>&tsn=" + t_tab_num + "&sy=" + selyear + "&sm=" + selmon + "&sd=" + selday;
}


function getShowUsersIconStr(eid, byes, bwait, bno)
{
	if (byes < 1 && bwait < 1 && bno < 1)
		return "";

	var s_str = "<br>";

	s_str = s_str + "<font face='Arial, Helvetica, sans-serif'><img src='images/cal/a.gif' border=0 title='参加'>&nbsp;" + byes.toString();
	s_str = s_str + "&nbsp;&nbsp;<img src='images/cal/u.gif' border=0 title='未决定的'>&nbsp;" + bwait.toString();
	s_str = s_str + "&nbsp;&nbsp;<img src='images/cal/d.gif' border=0 title='婉言拒绝'>&nbsp;" + bno.toString();

	s_str = s_str + "&nbsp;-&nbsp;</font>[<a class=mjEL href='javascript:viewInv(\"" + eid + "\")'>查看请柬</a>]";

	return s_str;
}

function write_getShowUsersIconStr(eid, byes, bwait, bno)
{
	document.write(getShowUsersIconStr(eid, byes, bwait, bno));
}

function getShowIconStr(bmode, bremind, brp)
{
	var s_str = "";

	if (bmode == 3)
		s_str = s_str + "<img src='images/cal/bdc.gif' border=0 align='absmiddle' title='生日'>";

	if (bremind == true)
		s_str = s_str + "<img src='images/cal/bell.gif' border=0 align='absmiddle' title='提醒'>";

	if (brp == true)
		s_str = s_str + "<img src='images/cal/repeat.gif' border=0 align='absmiddle' title='重复'>";
<%
if tab_selected_num <> "4" then
%>
	if (s_str.length > 0)
		s_str = s_str + "<br>";
<%
end if
%>
	return s_str;
}

function write_getShowIconStr(bmode, bremind, brp)
{
	document.write(getShowIconStr(bmode, bremind, brp));
}
<%
if tab_selected_num = "4" then
%>
function showtab4(p_page)
{
	if (p_page < 0)
		location.href = "<%=bakurl %>&page=" + document.getElementById("page").value + "&vsd=" + document.getElementById("view_search_date").value + "&vmd=" + document.getElementById("bi_mode").value + "&sortby=<%=sortby %>&sortmode=<%=sortmode %>";
	else
		location.href = "<%=bakurl %>&page=" + p_page + "&vsd=" + document.getElementById("view_search_date").value + "&vmd=" + document.getElementById("bi_mode").value + "&sortby=<%=sortby %>&sortmode=<%=sortmode %>";
}

function selectpage_onchange()
{
	showtab4(-1);
}
<%
end if

if tab_selected_num = "5" then
%>
function showtab5(p_page)
{
	if (p_page < 0)
		location.href = "<%=bakurl %>&page=" + document.getElementById("page").value + "&vsd=" + document.getElementById("view_state").value + "&sortby=<%=sortby %>&sortmode=<%=sortmode %>";
	else
		location.href = "<%=bakurl %>&page=" + p_page + "&vsd=" + document.getElementById("view_state").value + "&sortby=<%=sortby %>&sortmode=<%=sortmode %>";
}

function selectpage_onchange()
{
	showtab5(-1);
}

function speedadd()
{
	var theDOM = document.getElementById("ti_title");
	if (theDOM.value.length < 1)
	{
		alert("请输入“任务名称”项");
		theDOM.focus();
		return ;
	}

	document.f1.action = "cal_tasknew.asp";

	document.f1.sp_title.value = theDOM.value;
	document.f1.sp_start_year.value = document.getElementById("bi_start_year").value;
	document.f1.sp_start_month.value = document.getElementById("bi_start_month").value;
	document.f1.sp_start_day.value = document.getElementById("bi_start_day").value;
	document.f1.sp_level.value = document.getElementById("ti_level").value;

	if (document.f1.ti_is_set_end_true.checked == true)
		document.f1.sp_ti_is_set_end.value = "1";
	else
		document.f1.sp_ti_is_set_end.value = "0";

	document.f1.submit();
}
<%
else
%>
function speedadd()
{
	var theDOM = document.getElementById("bi_name");
	if (theDOM.value.length < 1)
	{
		alert("请输入“活动名称”项");
		theDOM.focus();
		return ;
	}

	document.f1.sp_title.value = theDOM.value;
	document.f1.sp_start_year.value = document.getElementById("bi_start_year").value;
	document.f1.sp_start_month.value = document.getElementById("bi_start_month").value;
	document.f1.sp_start_day.value = document.getElementById("bi_start_day").value;
	document.f1.sp_start_hour.value = document.getElementById("bi_start_hour").value;
	document.f1.sp_start_minute.value = document.getElementById("bi_start_minute").value;

	document.f1.submit();
}
<%
end if
%>

function viewinvited()
{
	location.href = "cal_listinvited.asp?<%=getGRSN() %>";
}

function m_over(tag_obj) {
	tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj) {
	tag_obj.style.backgroundColor = "white";
}
//-->
</script>

<body LANGUAGE=javascript onload="return window_onload()">
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
	<td width="50%" nowrap height="36" align="left" bgcolor="#c1d9f3" style="padding-left:6px;">
<%
if Len(Session("svcal")) < 1 then
%>
<a class='wwm_btnDownload btn_gray' href="javascript:viewinvited();">查看活动邀请</a>
<a class='wwm_btnDownload btn_gray' href="javascript:newevent();">添加活动</a>
<a class='wwm_btnDownload btn_gray' href="javascript:newtask();">添加待办事项</a>
&nbsp;&nbsp;[<a href="javascript:setcal()" class=mjNoLine>选项</a>]
<%
else
%>
	<input type="button" value="返回“我的效率手册”" style="WIDTH: 150px" onclick="javascript:viewme();" class="Bsbttn">
<%
end if
%>
    </td>
    <td align="right" height="36" bgcolor="#c1d9f3">
	<select id="ViewWho" name="ViewWho" class="drpdwn" LANGUAGE=javascript onchange="viewwho_onchange()">
	<option value="">我的效率手册</option>
	<option value="-1">[编辑我收藏的效率手册]</option>
	<option value="">--------------------</option>
<%
allnum = ecalset.CountFavorites
i = 0

do while i < allnum
	tmsg = server.htmlencode(ecalset.GetFavorite(i))
	Response.Write "<option value=""" & tmsg & """>查看:" & tmsg & "</option>" & Chr(13)

	tmsg = NULL

	i = i + 1
loop
%>
	</select>&nbsp;&nbsp;
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="152" bgcolor="<%=MY_COLOR_3 %>" valign="top">
      <table width="152" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="50" align="center" style="border-top:4px <%=MY_COLOR_3 %> solid; border-left:4px <%=MY_COLOR_3 %> solid; border-right:4px <%=MY_COLOR_3 %> solid; border-bottom:4px <%=MY_COLOR_3 %> solid;">
<%
if tab_selected_num = "0" or tab_selected_num = "1" or tab_selected_num = "4" or tab_selected_num = "5" then
%>
<SPAN id=calendar_container_left style="WIDTH: 100%"></SPAN>
<SCRIPT language=javascript src="images/cal/calendar.js"></SCRIPT>
<SCRIPT language=javascript>
showCalendar_left(selyear, selmon);
</SCRIPT>
		</td>
        <tr>
          <td height="16" align="center"><SPAN id=s_cur_day_str></SPAN></td>
        </tr>
        <tr>
          <td height="0" align="center">&nbsp;</td>
        </tr>
<%
elseif tab_selected_num = "2" then
%>
<table bgcolor="#fffff0" width="100%" border="0" cellspacing="3" cellpadding="1">
<tr>
<td colspan="3" align="center">
<table width=80% cellspacing=0 cellpadding=0>
<tr><td align=left width=32>
<span class=calendar_nav title='上一年' onclick="javascript:Selected_One(--start_year, start_month, 0, 2)"><img src="images/lefts1.gif"></span>
</td>
<td align=center nowrap class=calendar_grid><font class=calendar_title><%=start_year %>年</font></td>
<td align=right width=32>
<span class=calendar_nav title='下一年' onclick="javascript:Selected_One(++start_year, start_month, 0, 2)"><img src="images/rights1.gif"></span>
</td></tr></table>
</td>
</tr>
<tr>
<%
i = 1
do while i < 13
	if CInt(start_month) <> i then
		Response.Write "<td width='33%' align='center'><a href='javascript:Selected_One(start_year, " & i & ", 0, 2)'>" & i & "月</a></td>" & Chr(13)
	else
		Response.Write "<td width='33%' align='center'><b>" & i & "月</b></td>"
	end if

	if i = 3 or i = 6 or i = 9 or i = 12 then
		Response.Write "</tr><tr>"
	end if

	i = i + 1
loop
%>
</tr>
</table>
<%
elseif tab_selected_num = "3" then
%>
<table bgcolor="#fffff0" width="100%" border="0" cellspacing="7" cellpadding="1">
<tr>
<td colspan="3" align="center">
<table width=80% cellspacing=0 cellpadding=0>
<tr><td align=left width=32>
<span class=calendar_nav title='上一年' onclick="javascript:Selected_One(--start_year, start_month, 0, 3)"><img src="images/lefts1.gif"></span>
</td>
<td align=center nowrap class=calendar_grid><font class=calendar_title><%=start_year %>年</font></td>
<td align=right width=32>
<span class=calendar_nav title='下一年' onclick="javascript:Selected_One(++start_year, start_month, 0, 3)"><img src="images/rights1.gif"></span>
</td></tr></table>
</td>
</tr>
<tr>
<%
if start_year > 1972 then
%>
<td width="33%" align="center"><a href="javascript:Selected_One(--start_year, start_month, 0, 3)"><%=start_year - 1 %>年</a></td>
<td width="33%" align="center"><%=start_year %>年</td>
<td width="34%" align="center"><a href="javascript:Selected_One(++start_year, start_month, 0, 3)"><%=start_year + 1 %>年</a></td>
<%
else
%>
<td width="33%" align="center"><%=start_year %>年</td>
<td width="33%" align="center"><a href="javascript:Selected_One(start_year + 1, start_month, 0, 3)"><%=start_year + 1 %>年</a></td>
<td width="34%" align="center"><a href="javascript:Selected_One(start_year + 2, start_month, 0, 3)"><%=start_year + 2 %>年</a></td>
<%
end if
%>
</tr>
</table>
<%
end if
%>
          </td>
        </tr>
<%
if tab_selected_num <> "5" then
	ectk.SortListMode = 3

	if Len(Session("svcal")) < 1 then
		ectk.Load Session("wem")
	else
		ectk.Load Session("svcal")
	end if

	ectk.Search "", false, false, 0

	if Len(Session("svcal")) > 0 then
		ectk.HidePrivate
	end if

	allnum = ectk.Count
%>
        <tr>
          <td height="24" valign="top" bgcolor="#fffff0" style="border-left:4px <%=MY_COLOR_3 %> solid; border-right:4px <%=MY_COLOR_3 %> solid; border-bottom:4px <%=MY_COLOR_3 %> solid;">
<table width="100%">
<tr><td align="center" colspan="2">
<b>待办事项</b>
<%
if Len(Session("svcal")) < 1 then
%>
&nbsp;[<a href="javascript:newtask()">添加</a>]
<%
end if
%>
</td></tr>
<%
i = allnum - 1

do while i >= 0
	ectk.MoveTo i

	Response.Write "<tr><td align='left'><font color='#505050'>" & ectk.ti_level & "</font></td><td align='left'><a class=mjLinkLeft href=""javascript:showtask('" & ectk.ti_id & "')"">" & server.htmlencode(ectk.ti_title) & "</a></td></tr>" & Chr(13)

	i = i - 1
loop
%>
</table>
          </td>
        </tr>
<%
end if
%>
        <tr>
          <td height="30" bgcolor="#fffff0" style="border-left:4px <%=MY_COLOR_3 %> solid; border-right:4px <%=MY_COLOR_3 %> solid; border-bottom:4px <%=MY_COLOR_3 %> solid;">
<table width="100%">
<tr><td align="center">
<b>检索活动</b>
</td></tr>
<tr><td align="center">
<input type="text" id="searchtitle" name="searchtitle" size="12" class='textbox' maxlength="50">
<input type="button" value="搜索" style="WIDTH: 40px" onclick="javascript:searchit()" class="sbttn">
</td></tr>
<tr><td align="left">
&nbsp;<font style="font-size: 8pt;"><a href="javascript:tosearch()">高级搜索</a></font>
</td></tr>
</table>
          </td>
        </tr>
      </table>
    </td>
    <td valign="top">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="1" align="left">
          &nbsp;
          </td>
        </tr>
        <tr> 
          <td align="left">
            <table width="100%" height="25" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF" style="border-bottom:1px <%=MY_COLOR_5 %> solid;">
              <tr>
                <td width="10">&nbsp;</td>
<%
if tab_selected_num = "0" then
	Response.Write "<td width='40' class='ttf1'>日</td>"
else
	Response.Write "<td width='40' class='ttf2' onclick='change_tab(0);'>日</td>"
end if
%>
                <td width="16">&nbsp;</td>
<%
if tab_selected_num = "1" then
	Response.Write "<td width='40' class='ttf1'>周</td>"
else
	Response.Write "<td width='40' class='ttf2' onclick='change_tab(1);'>周</td>"
end if
%>
                <td width="16">&nbsp;</td>
<%
if tab_selected_num = "2" then
	Response.Write "<td width='40' class='ttf1'>月</td>"
else
	Response.Write "<td width='40' class='ttf2' onclick='change_tab(2);'>月</td>"
end if
%>
                <td width="16">&nbsp;</td>
<%
if tab_selected_num = "3" then
	Response.Write "<td width='40' class='ttf1'>年</td>"
else
	Response.Write "<td width='40' class='ttf2' onclick='change_tab(3);'>年</td>"
end if
%>
                <td width="16">&nbsp;</td>
<%
if tab_selected_num = "4" then
	Response.Write "<td width='80' class='ttf1'>活动列表</td>"
else
	Response.Write "<td width='80' class='ttf2' onclick='change_tab(4);'>活动列表</td>"
end if
%>
                <td width="16">&nbsp;</td>
<%
if tab_selected_num = "5" then
	Response.Write "<td width='80' class='ttf1'>待办事项</td>"
else
	Response.Write "<td width='80' class='ttf2' onclick='change_tab(5);'>待办事项</td>"
end if
%>
                <td>&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="border-left:7px #ffffff solid; border-right:5px #ffffff solid;">
<%
if tab_selected_num = "0" then
%>
<SPAN id=calendar_container style="WIDTH: 100%"></SPAN>
<SCRIPT language=javascript src="images/cal/dcalendar.js"></SCRIPT>
<SCRIPT language=javascript src="images/cal/nlcal.js"></SCRIPT>
<SCRIPT language=javascript>
var tmpDate = new Date(selyear, selmon - 1, selday);
var pDate = new Date();
var nDate = new Date();
pDate.setTime(tmpDate.getTime() - (86400 * 1000));
nDate.setTime(tmpDate.getTime() + (86400 * 1000));
var prevShowButton = "<a title='上一天' class=calendar_nav onclick=\"javascript:Selected_One(pDate.getFullYear(), pDate.getMonth() + 1, pDate.getDate(), 0)\"><img src=\"images/lefts1.gif\"></a>";
var nextShowButton = "<a title='下一天' class=calendar_nav onclick=\"javascript:Selected_One(nDate.getFullYear(), nDate.getMonth() + 1, nDate.getDate(), 0)\"><img src=\"images/rights1.gif\"></a>";
showCalendar(selyear, selmon, selday);
</SCRIPT>
<%
elseif tab_selected_num = "1" then
%>
<SPAN id=calendar_container style="WIDTH: 100%"></SPAN>
<SCRIPT language=javascript src="images/cal/wcalendar.js"></SCRIPT>
<SCRIPT language=javascript src="images/cal/nlcal.js"></SCRIPT>
<SCRIPT language=javascript>
var tmpDate = new Date(selyear, selmon - 1, selday);
var pDate = new Date();
var nDate = new Date();
pDate.setTime(tmpDate.getTime() - (86400 * 1000 * 7));
nDate.setTime(tmpDate.getTime() + (86400 * 1000 * 7));
var prevShowButton = "<a title='上一周' class=calendar_nav onclick=\"javascript:Selected_One(pDate.getFullYear(), pDate.getMonth() + 1, pDate.getDate(), 1)\"><img src=\"images/lefts1.gif\"></a>";
var nextShowButton = "<a title='下一周' class=calendar_nav onclick=\"javascript:Selected_One(nDate.getFullYear(), nDate.getMonth() + 1, nDate.getDate(), 1)\"><img src=\"images/rights1.gif\"></a>";
showCalendar(selyear, selmon, selday);
</SCRIPT>
<%
elseif tab_selected_num = "2" then
%>
<SPAN id=calendar_container style="WIDTH: 100%"></SPAN>
<SCRIPT language=javascript src="images/cal/mcalendar.js"></SCRIPT>
<SCRIPT language=javascript src="images/cal/nlcal.js"></SCRIPT>
<SCRIPT language=javascript>
showCalendar(start_year, start_month);
</SCRIPT>
<%
elseif tab_selected_num = "3" then
%>
<table width=100% border=0 cellspacing=0 cellpadding=0>
<tr><td class=calendar_dayname>
<a class=calendar_nav title='上一年' onclick="javascript:Selected_One(--start_year, start_month, 0, 3)"><img src="images/lefts1.gif"></a>&nbsp;<a 
class=calendar_nav title='下一年' onclick="javascript:Selected_One(++start_year, start_month, 0, 3)"><img src="images/rights1.gif"></a>&nbsp;&nbsp;<%
Response.Write start_year
%>年
</td></tr>
</table>
<SCRIPT language=javascript src="images/cal/ycalendar.js"></SCRIPT>
<table width="100%" border="0" cellspacing="0" cellpadding="3" style="border-left:1px #c0c0c0 solid; border-bottom:1px #c0c0c0 solid;">
<tr valign="top" align=center>
	<td width="33%" class="mjYGrid">
<SPAN id=calendar_container1 style="WIDTH: 100%"></SPAN>
<SCRIPT language=javascript>
showCalendar(start_year, monstep++);
</SCRIPT>
	</td>
	<td width="33%" class="mjYGrid">
<SPAN id=calendar_container2 style="WIDTH: 100%"></SPAN>
<SCRIPT language=javascript>
showCalendar(start_year, monstep++);
</SCRIPT>
	</td>
	<td width="34%" class="mjYGrid">
<SPAN id=calendar_container3 style="WIDTH: 100%"></SPAN>
<SCRIPT language=javascript>
showCalendar(start_year, monstep++);
</SCRIPT>
	</td>
</tr>
<tr valign="top" align=center>
	<td width="33%" class="mjYGrid">
<SPAN id=calendar_container4 style="WIDTH: 100%"></SPAN>
<SCRIPT language=javascript>
showCalendar(start_year, monstep++);
</SCRIPT>
	</td>
	<td width="33%" class="mjYGrid">
<SPAN id=calendar_container5 style="WIDTH: 100%"></SPAN>
<SCRIPT language=javascript>
showCalendar(start_year, monstep++);
</SCRIPT>
	</td>
	<td width="34%" class="mjYGrid">
<SPAN id=calendar_container6 style="WIDTH: 100%"></SPAN>
<SCRIPT language=javascript>
showCalendar(start_year, monstep++);
</SCRIPT>
	</td>
</tr>
<tr valign="top" align=center>
	<td width="33%" class="mjYGrid">
<SPAN id=calendar_container7 style="WIDTH: 100%"></SPAN>
<SCRIPT language=javascript>
showCalendar(start_year, monstep++);
</SCRIPT>
	</td>
	<td width="33%" class="mjYGrid">
<SPAN id=calendar_container8 style="WIDTH: 100%"></SPAN>
<SCRIPT language=javascript>
showCalendar(start_year, monstep++);
</SCRIPT>
	</td>
	<td width="34%" class="mjYGrid">
<SPAN id=calendar_container9 style="WIDTH: 100%"></SPAN>
<SCRIPT language=javascript>
showCalendar(start_year, monstep++);
</SCRIPT>
	</td>
</tr>
<tr valign="top" align=center>
	<td width="33%" class="mjYGrid">
<SPAN id=calendar_container10 style="WIDTH: 100%"></SPAN>
<SCRIPT language=javascript>
showCalendar(start_year, monstep++);
</SCRIPT>
	</td>
	<td width="33%" class="mjYGrid">
<SPAN id=calendar_container11 style="WIDTH: 100%"></SPAN>
<SCRIPT language=javascript>
showCalendar(start_year, monstep++);
</SCRIPT>
	</td>
	<td width="34%" class="mjYGrid">
<SPAN id=calendar_container12 style="WIDTH: 100%"></SPAN>
<SCRIPT language=javascript>
showCalendar(start_year, monstep++);
</SCRIPT>
	</td>
</tr>
</table>
<%
elseif tab_selected_num = "4" then
	vsd = trim(request("vsd"))
	if Len(vsd) < 1 then
		vsd = "0"
	end if

	if vsd = "1" then
		ecal.SelFuture
	elseif vsd = "2" then
		ecal.SelPast
	end if

	if Len(Session("svcal")) < 1 then
		vmd = trim(request("vmd"))
		if Len(vmd) > 0 then
			ecal.Search "", false, false, CLng(vmd)
		end if
	end if

	if Len(Session("svcal")) > 0 then
		ecal.HidePrivate
	end if

	allnum = ecal.Count
	allnb = ecal.Count

	if trim(request("page")) = "" then
		page = 0
	else
		page = CInt(trim(request("page")))
	end if

	if page < 0 then
		page = 0
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

	add_bakurl = "&page=" & page & "&vsd=" & vsd & "&vmd=" & vmd & "&sortby=" & sortby & "&sortmode=" & sortmode
%>
<table width=100% border=0 cellspacing=0 cellpadding=0>
<tr><td class=calendar_dayname>
<input type="button" value="删除" style="WIDTH: 40px" onclick="javascript:delmulevent()" class="sbttn">&nbsp;&nbsp;
<select id="view_search_date" name="view_search_date" class="drpdwn" LANGUAGE=javascript onchange="selectpage_onchange()">
<option value="0">全部</option>
<option value="1"<% if vsd = "1" then Response.Write " selected" %>>即将来临的</option>
<option value="2"<% if vsd = "2" then Response.Write " selected" %>>过去的</option>
</select>&nbsp;&nbsp;&nbsp;[<a href="javascript:setcal()" class=mjNoLine>选项</a>]&nbsp;&nbsp;
<%
if page > 0 then
	Response.Write "<a href='javascript:showtab4(" & page - 1 & ")'><img src='images/prep.gif' border='0' align='absmiddle'></a>&nbsp;"
else
	Response.Write "<img src='images/gprep.gif' border='0' align='absmiddle'>&nbsp;"
end if
%>
<select id="page" name="page" class="drpdwn" size="1" LANGUAGE=javascript onchange="selectpage_onchange()">
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
if page < allpage - 1 then
	Response.Write "&nbsp;<a href='javascript:showtab4(" & page + 1 & ")'><img src='images/nextp.gif' border='0' align='absmiddle'></a>"
else
	Response.Write "&nbsp;<img src='images/gnextp.gif' border='0' align='absmiddle'>"
end if

if Len(Session("svcal")) < 1 then
%>
&nbsp;&nbsp;
<select id="bi_mode" name="bi_mode" class="drpdwn" LANGUAGE=javascript onchange="selectpage_onchange()">
<option value="" selected>显示所有类型</option>
<option value="100">提醒功能</option>
<option value="101">重复</option>
<option value="102">请柬</option>
<option value="">------------</option>
<%
i = 0

do while i < 28
	Response.Write "<option value=""" & i & """>" & getModeName(i) & "</option>" & Chr(13)

	i = i + 1
loop
%>
</select>
<%
else
%>
<input type="hidden" id="bi_mode" name="bi_mode">
<%
end if
%>
</td></tr>
<tr><td>
<form method="post" action="cal_del.asp" name="f2">
<input type="hidden" name="returl" value="<%=bakurl & add_bakurl %>">
<input type="hidden" name="calmode" value="2">
<table width="100%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr bgcolor="#f2f4f6">
    <td width="6%" height="25" class="st_l"><input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()"></td>
	<td width="23%" class="st_l"><%
if sortby = "0" then
	if sortmode = "0" then
		Response.Write "<a class='urf' href=""javascript:setsort('0', '1')"">日期</a>&nbsp;<a href=""javascript:setsort('0', '1')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "<a class='urf' href=""javascript:setsort('0', '0')"">日期</a>&nbsp;<a href=""javascript:setsort('0', '0')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	end if
else
	Response.Write "<a class='urf' href=""javascript:setsort('0', '0')"">日期</a>"
end if
%></td>
	<td width="13%"  class="st_l">时间</td>
	<td width="42%" colspan="2" class="st_l"><%
if Len(Session("svcal")) < 1 then
	if sortby = "2" then
		if sortmode = "0" then
			Response.Write "<a class='urf' href=""javascript:setsort('2', '1')"">事件</a>&nbsp;<a href=""javascript:setsort('2', '1')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
		else
			Response.Write "<a class='urf' href=""javascript:setsort('2', '0')"">事件</a>&nbsp;<a href=""javascript:setsort('2', '0')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
		end if
	else
		Response.Write "<a class='urf' href=""javascript:setsort('2', '0')"">事件</a>"
	end if
else
	Response.Write "事件"
end if
%></td>
	<td width="16%" class="st_r"><%
if Len(Session("svcal")) < 1 then
	if sortby = "1" then
		if sortmode = "0" then
			Response.Write "<a class='urf' href=""javascript:setsort('1', '1')"">类型</a>&nbsp;<a href=""javascript:setsort('1', '1')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
		else
			Response.Write "<a class='urf' href=""javascript:setsort('1', '0')"">类型</a>&nbsp;<a href=""javascript:setsort('1', '0')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
		end if
	else
		Response.Write "<a class='urf' href=""javascript:setsort('1', '0')"">类型</a>"
	end if
else
	Response.Write "类型"
end if
%></td>
  </tr>
<%
si = 0
i = 0
ecal.timeMode = 0
set ecalext = server.createobject("easymail.CalendarExtend")

do while i < ((page + 1) * pageline) and i < allnum
	if i >= page * pageline then
		if sortmode = "1" then
			showi = i
		else
			showi = allnum - i - 1
		end if

		ecal.MoveTo showi
		show_bi_start_date = ecal.show_bi_start_date

		Response.Write "<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);'>"
		if Len(Session("svcal")) < 1 or ecal.bi_shareMode = 2 then
			Response.Write "	<td align='center' class='cont_td'><input type='checkbox' name='check" & si & "' value='" & ecal.bi_id & "'></td>"
		else
			Response.Write "	<td align='center' class='cont_td'><input type='checkbox' name='check" & si & "'></td>"
		end if

		Response.Write "    <td align='left' nowrap class='cont_td'><a href='" & get_tab_day_url_by_date(show_bi_start_date) & "'>" & get_show_date(show_bi_start_date) & "</a></td>"
		Response.Write "    <td align='left' nowrap class='cont_td'>" & get_show_time(show_bi_start_date) & "</td>"

		if Len(Session("svcal")) < 1 or ecal.bi_shareMode = 2 then
			Response.Write "    <td align='left' class='cont_td'><a href=""javascript:showevent('" & ecal.bi_id & "')"">" & server.htmlencode(ecal.bi_name) & "</a>"

			if ecal.bi_has_invitation = true then
				LightLoad_isok = false

				if Len(ecal.bi_host_account) < 1 then
					if Len(Session("svcal")) < 1 then
						LightLoad_isok = ecalext.LightLoad(Session("wem"), ecal.bi_id)
					else
						LightLoad_isok = ecalext.LightLoad(Session("svcal"), ecal.bi_id)
					end if
				else
					LightLoad_isok = ecalext.LightLoad(ecal.bi_host_account, ecal.bi_id)
				end if

				if LightLoad_isok = true then
					Response.Write "<script>write_getShowUsersIconStr('" & ecal.bi_id & "'," & ecalext.Yes_User & "," & ecalext.Wait_User & "," & ecalext.No_User & ")</script></td>"
				end if
			end if

			Response.Write "    <td align='right' width='1%' nowrap class='cont_td'>&nbsp;<script>write_getShowIconStr(" & ecal.bi_mode & "," & LCase(CStr(ecal.bi_remind)) & "," & LCase(CStr(ecal.bi_isRepeat)) & ")</script></td>"

			Response.Write "    <td align='center' nowrap class='cont_td'>" & server.htmlencode(getModeName(ecal.bi_mode)) & "</td>"
		else
			Response.Write "    <td align='left' class='cont_td'>忙碌</a>"
			Response.Write "    <td align='right' width='1%' nowrap class='cont_td'>&nbsp;</td>"
			Response.Write "    <td align='center' nowrap class='cont_td'>---</td>"
		end if

		Response.Write "  </tr>" & chr(13)

		si = si + 1
	end if

	show_bi_start_date = NULL

	i = i + 1
loop

set ecalext = nothing
%>
</table>
</form>
</td></tr>
</table>
<%
elseif tab_selected_num = "5" then
	ectk.SortListMode = sortby

	if Len(Session("svcal")) < 1 then
		ectk.Load Session("wem")
	else
		ectk.Load Session("svcal")
	end if

	vsd = trim(request("vsd"))
	if Len(vsd) > 0 and IsNumeric(vsd) = true then
		ectk.Search "", false, false, CLng(vsd)
	end if

	if Len(Session("svcal")) > 0 then
		ectk.HidePrivate
	end if

	allnum = ectk.Count
	allnb = ectk.Count

	if trim(request("page")) = "" then
		page = 0
	else
		page = CInt(trim(request("page")))
	end if

	if page < 0 then
		page = 0
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

	add_bakurl = "&page=" & page & "&vsd=" & vsd & "&sortby=" & sortby & "&sortmode=" & sortmode
%>
<table width=100% border=0 cellspacing=0 cellpadding=0>
<tr><td class=calendar_dayname>
<input type="button" value="删除" style="WIDTH: 40px" onclick="javascript:delmulevent()" class="sbttn">&nbsp;
<input type="button" value="标记为完成" style="WIDTH: 76px" onclick="javascript:setover()" class="sbttn">&nbsp;&nbsp;&nbsp;
<select id="view_state" name="view_state" class="drpdwn" LANGUAGE=javascript onchange="selectpage_onchange()">
<option value="-1">全部</option>
<option value="0"<% if vsd = "0" then Response.Write " selected" %>>未完成</option>
<option value="1"<% if vsd = "1" then Response.Write " selected" %>>完成</option>
</select>&nbsp;&nbsp;&nbsp;[<a href="javascript:setcal()" class=mjNoLine>选项</a>]&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%
if page > 0 then
	Response.Write "<a href='javascript:showtab5(" & page - 1 & ")'><img src='images/prep.gif' border='0' align='absmiddle'></a>&nbsp;"
else
	Response.Write "<img src='images/gprep.gif' border='0' align='absmiddle'>&nbsp;"
end if
%>
<select id="page" name="page" class="drpdwn" size="1" LANGUAGE=javascript onchange="selectpage_onchange()">
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
if page < allpage - 1 then
	Response.Write "&nbsp;<a href='javascript:showtab5(" & page + 1 & ")'><img src='images/nextp.gif' border='0' align='absmiddle'></a>"
else
	Response.Write "&nbsp;<img src='images/gnextp.gif' border='0' align='absmiddle'>"
end if
%>
</td></tr>
<tr><td>
<form method="post" action="cal_del.asp" name="f2">
<input type="hidden" name="returl" value="<%=bakurl & add_bakurl %>">
<input type="hidden" name="calmode" value="3">
<table width="100%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr bgcolor="#f2f4f6">
    <td width="6%" height="25" class="st_l"><input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()"></td>
	<td width="12%" nowrap class="st_l"><%
if sortby = "3" then
	if sortmode = "0" then
		Response.Write "<a class='urf' href=""javascript:setsort('3', '1')"">优先级</a>&nbsp;<a href=""javascript:setsort('3', '1')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "<a class='urf' href=""javascript:setsort('3', '0')"">优先级</a>&nbsp;<a href=""javascript:setsort('3', '0')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	end if
else
	Response.Write "<a class='urf' href=""javascript:setsort('3', '0')"">优先级</a>"
end if
%></td>
	<td width="55%" class="st_l"><%
if sortby = "2" then
	if sortmode = "0" then
		Response.Write "<a class='urf' href=""javascript:setsort('2', '1')"">待办事项</a>&nbsp;<a href=""javascript:setsort('2', '1')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "<a class='urf' href=""javascript:setsort('2', '0')"">待办事项</a>&nbsp;<a href=""javascript:setsort('2', '0')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	end if
else
	Response.Write "<a class='urf' href=""javascript:setsort('2', '0')"">待办事项</a>"
end if
%></td>
	<td width="7%" nowrap class="st_l">完成</td>
	<td width="20%" nowrap class="st_r"><%
if sortby = "0" then
	if sortmode = "0" then
		Response.Write "<a class='urf' href=""javascript:setsort('0', '1')"">完成期限</a>&nbsp;<a href=""javascript:setsort('0', '1')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "<a class='urf' href=""javascript:setsort('0', '0')"">完成期限</a>&nbsp;<a href=""javascript:setsort('0', '0')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	end if
else
	Response.Write "<a class='urf' href=""javascript:setsort('0', '0')"">完成期限</a>"
end if
%></td>
  </tr>
<%
si = 0
i = 0
ectk.timeMode = 0

do while i < ((page + 1) * pageline) and i < allnum
	if i >= page * pageline then
		if sortmode = 1 then
			showi = i
		else
			showi = allnum - i - 1
		end if

		ectk.MoveTo showi

		Response.Write "<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);'>"
		Response.Write "	<td align='center' class='cont_td'><input type='checkbox' name='check" & si & "' value='" & ectk.ti_id & "'></td>"

		Response.Write "	<td align='center' nowrap class='cont_td'>" & ectk.ti_level & "</td>"
		Response.Write "	<td align='left' class='cont_td'><a href=""javascript:showtask('" & ectk.ti_id & "')"">" & server.htmlencode(ectk.ti_title) & "</a></td>"
		Response.Write "	<td align='center' nowrap class='cont_td'>"

		if ectk.ti_state = true then
			Response.Write "<img src='images/cal/ok.gif' border=0 align='absmiddle'></td>"
		else
			Response.Write "&nbsp;</td>"
		end if

		if Len(ectk.show_ti_end_date) > 0 then
			Response.Write "	<td align='left' nowrap class='cont_td'><a href='" & get_tab_day_url_by_date(ectk.show_ti_end_date) & "'>" & server.htmlencode(ectk.show_ti_end_date) & "</a></td>"
		else
			Response.Write "	<td align='left' nowrap class='cont_td'>&nbsp;</td>"
		end if
		Response.Write "  </tr>" & chr(13)

		si = si + 1
	end if

	i = i + 1
loop
%>
</table>
</form>
</td></tr>
</table>
<%
end if
%>
          </td>
        </tr>
        <tr>
		<td height="10">&nbsp;</td>
        </tr>
        <tr>
		<td height="<%
if Len(Session("svcal")) < 1 then
	Response.Write "50"
else
	Response.Write "0"
end if
%>" align="center" style="border-top:4px <%=MY_COLOR_3 %> solid; border-right:4px <%=MY_COLOR_3 %> solid; border-bottom:4px <%=MY_COLOR_3 %> solid;">
<%
if tab_selected_num = "5" then
	if Len(Session("svcal")) < 1 then
%>
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
	<td class="b_td">
	<b nowrap>快速添加任务</b>
	</td>
	<td class="b_td">
	<b nowrap>完成期限</b>
	</td>
	<td class="b_td">
	<b>优先级</b>
	</td>
	</tr>
	<tr>
	<td class="b_td">
		<input type="text" id="ti_title" name="ti_title" class='textbox' size="12" maxlength="50">
	</td>
	<td class="b_td">
		<input type=radio value="1" name="ti_is_set_end" id="ti_is_set_end_true" checked>
		<select id="bi_start_year" name="bi_start_year" class="drpdwn">
<%
curDate = Now
i = Year(curDate) - 1

do while i < Year(curDate) + 7
	if Year(curDate) = i then
		Response.Write "<option value=""" & i & """ selected>" & i & "年</option>" & Chr(13)
	else
		Response.Write "<option value=""" & i & """>" & i & "年</option>" & Chr(13)
	end if

	i = i + 1
loop
%>
		</select><select id="bi_start_month" name="bi_start_month" class="drpdwn">
<%
i = 1

do while i < 13
	if Month(curDate) = i then
		Response.Write "<option value=""" & i & """ selected>" & i & "月</option>" & Chr(13)
	else
		Response.Write "<option value=""" & i & """>" & i & "月</option>" & Chr(13)
	end if

	i = i + 1
loop
%>
		</select><select id="bi_start_day" name="bi_start_day" class="drpdwn">
<%
i = 1

do while i < 32
	if Day(curDate) = i then
		Response.Write "<option value=""" & i & """ selected>" & i & "日</option>" & Chr(13)
	else
		Response.Write "<option value=""" & i & """>" & i & "日</option>" & Chr(13)
	end if

	i = i + 1
loop
%>
		</select>
		<input type=radio value="0" name="ti_is_set_end" id="ti_is_set_end_false">未设到期日
	</td>
	<td class="b_td">
		<select id="ti_level" name="ti_level" class="drpdwn">
<%
i = 1

do while i < 10
	if i <> 5 then
		Response.Write "<option value=""" & i & """>" & i & "</option>" & Chr(13)
	else
		Response.Write "<option value=""" & i & """ selected>" & i & "</option>" & Chr(13)
	end if

	i = i + 1
loop
%>
		</select>
	</td>
	<td class="b_td">
<input type="button" value="添加" style="WIDTH: 42px" onclick="javascript:speedadd()" class="sbttn">
	</td>
	</tr>
</table>
<%
	end if
else
	if Len(Session("svcal")) < 1 then
%>
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
	<td class="b_td">
	<b>快速添加活动</b>
	</td>
	<td class="b_td">
	<b>日期</b>
	</td>
	<td class="b_td">
	<b>时间</b>
	</td>
	</tr>
	<tr>
	<td class="b_td">
		<input type="text" id="bi_name" name="bi_name" class='textbox' size="16" maxlength="40">
	</td>
	<td class="b_td">
		<select id="bi_start_year" name="bi_start_year" class="drpdwn">
<%
curDate = Now
i = Year(curDate) - 1

do while i < Year(curDate) + 7
	if Year(curDate) = i then
		Response.Write "<option value=""" & i & """ selected>" & i & "年</option>" & Chr(13)
	else
		Response.Write "<option value=""" & i & """>" & i & "年</option>" & Chr(13)
	end if

	i = i + 1
loop
%>
		</select><select id="bi_start_month" name="bi_start_month" class="drpdwn">
<%
i = 1

do while i < 13
	if Month(curDate) = i then
		Response.Write "<option value=""" & i & """ selected>" & i & "月</option>" & Chr(13)
	else
		Response.Write "<option value=""" & i & """>" & i & "月</option>" & Chr(13)
	end if

	i = i + 1
loop
%>
		</select><select id="bi_start_day" name="bi_start_day" class="drpdwn">
<%
i = 1

do while i < 32
	if Day(curDate) = i then
		Response.Write "<option value=""" & i & """ selected>" & i & "日</option>" & Chr(13)
	else
		Response.Write "<option value=""" & i & """>" & i & "日</option>" & Chr(13)
	end if

	i = i + 1
loop
%>
		</select>
	</td>
	<td class="b_td">
		<select id="bi_start_hour" name="bi_start_hour" class="drpdwn">
<%
i = 0

do while i < 24
	is_this_hour = false

	if Hour(curDate) = i then
		is_this_hour = true
	end if

if show_APM = true then
	if i = 0 then
		if is_this_hour = false then
			Response.Write "<option value=""" & i & """>" & 12 & " am</option>" & Chr(13)
		else
			Response.Write "<option value=""" & i & """ selected>" & 12 & " am</option>" & Chr(13)
		end if
	elseif i = 12 then
		if is_this_hour = false then
			Response.Write "<option value=""" & i & """>" & 12 & " pm</option>" & Chr(13)
		else
			Response.Write "<option value=""" & i & """ selected>" & 12 & " pm</option>" & Chr(13)
		end if
	elseif i < 12 then
		if is_this_hour = false then
			Response.Write "<option value=""" & i & """>" & i & " am</option>" & Chr(13)
		else
			Response.Write "<option value=""" & i & """ selected>" & i & " am</option>" & Chr(13)
		end if
	else
		if is_this_hour = false then
			Response.Write "<option value=""" & i & """>" & i - 12 & " pm</option>" & Chr(13)
		else
			Response.Write "<option value=""" & i & """ selected>" & i - 12 & " pm</option>" & Chr(13)
		end if
	end if
else
	if is_this_hour = false then
		Response.Write "<option value=""" & i & """>" & i & "</option>" & Chr(13)
	else
		Response.Write "<option value=""" & i & """ selected>" & i & "</option>" & Chr(13)
	end if
end if

	i = i + 1
loop
%>
		</select><select id="bi_start_minute" name="bi_start_minute" class="drpdwn">
		<option value="0">:00</option>
		<option value="15">:15</option>
		<option value="30">:30</option>
		<option value="45">:45</option>
		</select>
	</td>
	<td class="b_td">
<input type="button" value="添加" style="WIDTH: 42px" onclick="javascript:speedadd()" class="sbttn">
	</td>
	</tr>
</table>
<%
	end if
end if
%>
		</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
	<td>&nbsp;</td>
  </tr>
</table>
<form method="post" action="cal_new.asp" name="f1">
<input type="hidden" name="returl" value="<%=bakurl & add_bakurl %>">
<input type="hidden" name="isspd" value="1">
<input type="hidden" name="sp_title">
<input type="hidden" name="sp_start_year">
<input type="hidden" name="sp_start_month">
<input type="hidden" name="sp_start_day">
<input type="hidden" name="sp_start_hour">
<input type="hidden" name="sp_start_minute">
<input type="hidden" name="sp_level">
<input type="hidden" name="sp_ti_is_set_end">
<input type="hidden" name="searchstr">
<input type="hidden" name="st" value="1">
</form>
</body>


<SCRIPT language=javascript>
<!--
function window_onload()
{
<%
if tab_selected_num = "0" or tab_selected_num = "1" or tab_selected_num = "4" or tab_selected_num = "5" then
%>
	Stag = document.getElementById("s_cur_day_str");
	Stag.innerHTML = "今天: <a class=mjLink href=\"JavaScript:Selected_One(" + curyear + "," + curmon + "," + curday + ",0)\">" + curyear + "年" + curmon + "月" + curday + "日</a>";
<%
end if

if tab_selected_num = "4" then
	if Len(vmd) > 0 then
		Response.Write "document.getElementById('bi_mode').value = """ & vmd & """;" & Chr(13)
	end if
end if

if Len(Session("svcal")) > 0 then
	Response.Write "ViewWho.value = """ & Session("svcal") & """;" & Chr(13)
end if
%>
}

function showevent(evid)
{
	location.href = "cal_new.asp?<%=getGRSN() %>&editcal=1&calid=" + evid + "&returl=<%=Server.URLEncode(bakurl & add_bakurl) %>";
}

function showtask(evid)
{
	location.href = "cal_tasknew.asp?<%=getGRSN() %>&editcal=1&calid=" + evid + "&returl=<%=Server.URLEncode(bakurl & add_bakurl) %>";
}

<%
if tab_selected_num = "4" then
%>
function setsort(p_sortby, p_sortmode)
{
	location.href = "<%=bakurl %>&page=" + document.getElementById("page").value + "&vsd=" + document.getElementById("view_search_date").value + "&vmd=" + document.getElementById("bi_mode").value + "&sortby=" + p_sortby + "&sortmode=" + p_sortmode;
}
<%
end if

if tab_selected_num = "5" then
%>
function setsort(p_sortby, p_sortmode)
{
	location.href = "<%=bakurl %>&page=" + document.getElementById("page").value + "&vsd=" + document.getElementById("view_state").value + "&sortby=" + p_sortby + "&sortmode=" + p_sortmode;
}
<%
end if

if tab_selected_num = "4" or tab_selected_num = "5" then
%>
function allcheck_onclick() {
	if (f2.allcheck.checked == true)
		checkall(true);
	else
		checkall(false);
}

function checkall(check) {
	var i = 0;
	var theObj;

	for(; i<<%
if pageline < allnb then
	Response.Write pageline
else
	Response.Write allnb
end if
%>; i++)
	{
		theObj = eval("f2.check" + i);

		if (theObj != null)
			theObj.checked = check;
	}
}

function ischeck() {
	var i = 0;
	var theObj;

	for(; i<<%
if pageline < allnb then
	Response.Write pageline
else
	Response.Write allnb
end if
%>; i++)
	{
		theObj = eval("f2.check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}

function delmulevent()
{
	if (ischeck() == true)
	{
<%
if Len(Session("svcal")) < 1 then
%>
		if (confirm("确实要删除吗?") == false)
			return ;

		f2.submit();
<%
else
%>
		alert("您的权限不足.");
<%
end if
%>
	}
}

function setover()
{
	if (ischeck() == true)
	{
<%
if Len(Session("svcal")) < 1 then
%>
		if (confirm("确实要标记为完成吗?") == false)
			return ;

		f2.calmode.value = "4";
		f2.submit();
<%
else
%>
		alert("您的权限不足.");
<%
end if
%>
	}
}
<%
end if
%>
function tosearch()
{
	location.href = "cal_search.asp?<%=getGRSN() %>&returl=<%=Server.URLEncode(bakurl & add_bakurl) %>";
}

function searchit()
{
	var theDOM = document.getElementById("searchtitle")
	if (theDOM.value.length > 0)
	{
		document.f1.action = "cal_search.asp";
		document.f1.searchstr.value = theDOM.value;
		document.f1.submit();
	}
}

function newevent()
{
	location.href="cal_new.asp?<%=getGRSN() %>&returl=<%=Server.URLEncode(bakurl & add_bakurl) %>";
}

function newtask()
{
	location.href="cal_tasknew.asp?<%=getGRSN() %>&returl=<%=Server.URLEncode(bakurl & add_bakurl) %>";
}

function delevent(evid)
{
<%
if Len(Session("svcal")) < 1 then
%>
	if (confirm("确实要删除吗?") == false)
		return ;

	location.href = "cal_del.asp?<%=getGRSN() %>&calmode=1&calid=" + evid + "&returl=<%=Server.URLEncode(bakurl & add_bakurl) %>";
<%
else
%>
	alert("您的权限不足.");
<%
end if
%>
}

function gonewurl(bsy, bsm, bsd, bsh)
{
<%
if Len(Session("svcal")) < 1 then
%>
	location.href="cal_new.asp?<%=getGRSN() %>&bsy=" + bsy + "&bsm=" + bsm + "&bsd=" + bsd + "&bsh=" + bsh + "&returl=<%=Server.URLEncode(bakurl & add_bakurl) %>";
<%
else
%>
	alert("您的权限不足.");
<%
end if
%>
}

function viewInv(evid)
{
	location.href = "cal_showinvite.asp?<%=getGRSN() %>&fmcal=1&calid=" + evid + "&returl=<%=Server.URLEncode(bakurl & add_bakurl) %>";
}

function setcal()
{
<%
if Len(Session("svcal")) < 1 then
%>
	location.href="cal_set.asp?<%=getGRSN() %>&returl=<%=Server.URLEncode(bakurl & add_bakurl) %>";
<%
else
%>
	alert("您的权限不足.");
<%
end if
%>
}

function viewwho_onchange()
{
	theDOM = document.getElementById("ViewWho");
	var msi = theDOM.value;
	theDOM.selectedIndex = 0;

<%
if Len(Session("svcal")) < 1 then
%>
	if (msi.length == 0)
		return ;
<%
else
%>
	if (msi.length == 0)
	{
		viewme();
		return ;
	}
<%
end if
%>

	if (msi == "-1")
	{
		location.href = "cal_favorite.asp?<%=getGRSN() %>&returl=<%=Server.URLEncode(bakurl & add_bakurl) %>";
		return ;
	}

	location.href = "cal_index.asp?<%=getGRSN() %>&viewwho=" + msi;
}

function viewme()
{
	location.href = "cal_index.asp?<%=getGRSN() %>";
}
//-->
</SCRIPT>
</html>

<%
set ecalset = nothing
set ectk = nothing
set ecal = nothing


function getModeName(mdnum)
	temp_mode_str = ""
	if mdnum = "0" then
		temp_mode_str = "纪念日"
	elseif mdnum = "1" then
		temp_mode_str = "约会"
	elseif mdnum = "2" then
		temp_mode_str = "支付帐单"
	elseif mdnum = "3" then
		temp_mode_str = "生日"
	elseif mdnum = "4" then
		temp_mode_str = "早餐"
	elseif mdnum = "5" then
		temp_mode_str = "访问"
	elseif mdnum = "6" then
		temp_mode_str = "聊天"
	elseif mdnum = "7" then
		temp_mode_str = "课程"
	elseif mdnum = "8" then
		temp_mode_str = "Club 事件"
	elseif mdnum = "9" then
		temp_mode_str = "音乐会"
	elseif mdnum = "10" then
		temp_mode_str = "晚饭"
	elseif mdnum = "11" then
		temp_mode_str = "毕业"
	elseif mdnum = "12" then
		temp_mode_str = "Happy Hour"
	elseif mdnum = "13" then
		temp_mode_str = "节日"
	elseif mdnum = "14" then
		temp_mode_str = "会见"
	elseif mdnum = "15" then
		temp_mode_str = "午餐"
	elseif mdnum = "16" then
		temp_mode_str = "会议"
	elseif mdnum = "17" then
		temp_mode_str = "电影"
	elseif mdnum = "18" then
		temp_mode_str = "网络事件"
	elseif mdnum = "19" then
		temp_mode_str = "其他"
	elseif mdnum = "20" then
		temp_mode_str = "宴会"
	elseif mdnum = "21" then
		temp_mode_str = "表演"
	elseif mdnum = "22" then
		temp_mode_str = "亲友重聚"
	elseif mdnum = "23" then
		temp_mode_str = "运动比赛"
	elseif mdnum = "24" then
		temp_mode_str = "旅行"
	elseif mdnum = "25" then
		temp_mode_str = "电视节目"
	elseif mdnum = "26" then
		temp_mode_str = "假期"
	elseif mdnum = "27" then
		temp_mode_str = "婚礼"
	end if

	getModeName = temp_mode_str
end function


function get_show_date(s_date_str)
	tmp_s_date_str = InStr(s_date_str, " ")
	if tmp_s_date_str > 0 then
		get_show_date = Left(s_date_str, tmp_s_date_str - 1)
	end if
end function


function get_conv_24hour_2_apm(hstr)
	temp_hstr = hstr
	if Left(temp_hstr, 1) = "0" then
		temp_hstr = Right(temp_hstr, Len(temp_hstr) - 1)
	end if

	tmp_fg_p = InStr(temp_hstr, ":")
	if tmp_fg_p > 0 then
		tmp_fg_hour = CInt(Left(temp_hstr, tmp_fg_p - 1))

		if tmp_fg_hour = 0 then
			get_conv_24hour_2_apm = "12:" & Right(temp_hstr, Len(temp_hstr) - tmp_fg_p) & "AM"
		elseif tmp_fg_hour = 12 then
			get_conv_24hour_2_apm = "12:" & Right(temp_hstr, Len(temp_hstr) - tmp_fg_p) & "PM"
		elseif tmp_fg_hour < 12 then
			get_conv_24hour_2_apm = temp_hstr & "AM"
		else
			get_conv_24hour_2_apm = CStr(tmp_fg_hour - 12) & ":" & Right(temp_hstr, Len(temp_hstr) - tmp_fg_p) & "PM"
		end if
	end if
end function


function get_show_time(s_date_str)
	tmp_s_date_str = InStr(s_date_str, " ")
	if tmp_s_date_str > 0 then
		if show_APM = false then
			get_show_time = Right(s_date_str, Len(s_date_str) - tmp_s_date_str)
		else
			get_show_time = get_conv_24hour_2_apm(Right(s_date_str, Len(s_date_str) - tmp_s_date_str))
		end if
	end if
end function


function get_tab_day_url_by_date(show_date_str)
	tmp_jp_day = 9
	tmp_month = Mid(show_date_str, 6, 2)

	if IsNumeric(tmp_month) = false then
		tmp_month = Mid(show_date_str, 6, 1)
		tmp_jp_day = 8
	end if

	tmp_day = Mid(show_date_str, tmp_jp_day, 2)

	if IsNumeric(tmp_day) = false then
		tmp_day = Mid(show_date_str, tmp_jp_day, 1)
	end if

	get_tab_day_url_by_date = "cal_index.asp?" & getGRSN() & "&tsn=0&sy=" & Mid(show_date_str, 1, 4) & "&sm=" & tmp_month & "&sd=" & tmp_day
end function


function convFeast(mm, dd, fname)
	tmpstr = ""
	if mm < 10 then
		tmpstr = "0" & mm
	else
		tmpstr = mm
	end if

	if dd < 10 then
		tmpstr = tmpstr & "0" & dd
	else
		tmpstr = tmpstr & dd
	end if

	convFeast = tmpstr & " " & server.htmlencode(fname)
end function
%>
