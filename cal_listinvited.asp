<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
fmeml = trim(request("fmeml"))

dim ecalset
set ecalset = server.createobject("easymail.CalOptions")
ecalset.Load Session("wem")

show_APM = false
if ecalset.Show24Hour = false then
	show_APM = true
end if

sortby = trim(request("sortby"))

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

sortmode = trim(request("sortmode"))
if sortmode <> "1" then
	sortmode = "0"
end if

dim ecal
set ecal = server.createobject("easymail.CalendarNotice")

ecal.SortListMode = sortby

ecal.Load Session("wem")
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
.st_l,.st_r {text-align:center; white-space:nowrap; border-left:1px solid #A5B6C8; border-top:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8;}
.st_r {border-right:1px solid #A5B6C8;}
.cont_tr {background:white; height:22px;}
.cont_td {height:22px; white-space:nowrap; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px;}
.urf {color:black;}
.urf:hover {color:black;}

a:hover {text-decoration:underline;}

.mjNoLine {
 	text-decoration: none; 
}
.mjLinkLeft {
	color: #447172;
 	text-decoration: none; 
	CURSOR: pointer;
}
.mjLink {
	color: #002f72;
 	text-decoration: none; 
	CURSOR: pointer;
}
.calendar_dayname {
	BORDER-TOP: #ffffc0 7px solid;
	BORDER-LEFT: #ffffc0 5px solid;
	BORDER-BOTTOM: #ffffc0 3px solid;
	color: #202020;
	BACKGROUND-COLOR: #ffffc0;
}
.mjEL {
	font-size: 9pt;
	color: #447172;
 	text-decoration: none; 
	CURSOR: pointer;
}
-->
</STYLE>
</head>

<script type="text/javascript">
<!--
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
bakurl = "cal_listinvited.asp?" & getGRSN()
%>

function delevent(evid)
{
	if (confirm("确实要删除吗?") == false)
		return ;

	location.href = "cal_del.asp?<%=getGRSN() %>&calmode=1&calid=" + evid + "&returl=<%=Server.URLEncode(bakurl) %>";
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

	if (s_str.length > 0)
		s_str = s_str + "<br>";

	return s_str;
}

function write_getShowIconStr(bmode, bremind, brp)
{
	document.write(getShowIconStr(bmode, bremind, brp));
}

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

function vieweff()
{
	location.href = "cal_index.asp?<%=getGRSN() %>";
}
<%
if fmeml = "1" then
%>
function goback()
{
	history.back();
}
<%
end if
%>

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
	<td nowrap height="36" align="left"<%
if fmeml <> "1" then
	Response.Write " colspan=2"
end if
%> bgcolor="#c1d9f3" style="padding-left:6px;">
<a class='wwm_btnDownload btn_gray' href="javascript:vieweff();">查看我的效率手册</a>
	</td>
<%
if fmeml = "1" then
%>
	<td height="36" align="right" bgcolor="#c1d9f3" style="padding-right:6px;">
<a class='wwm_btnDownload btn_blue' href="javascript:goback();"><%=s_lang_return %></a>
	</td>
<%
end if
%>
	</tr>
	<tr>
	<td width="80%" colspan="2" style="border-left:3px #ffffff solid; border-right:3px #ffffff solid;">
<%
	vsd = trim(request("vsd"))
	if Len(vsd) < 1 then
		vsd = "0"
	end if

	if vsd = "1" then
		ecal.SelFuture
	elseif vsd = "2" then
		ecal.SelPast
	end if

	vmd = trim(request("vmd"))
	if Len(vmd) > 0 then
		ecal.Search "", false, false, CLng(vmd)
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
</select>&nbsp;&nbsp;&nbsp;
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
%>
&nbsp;&nbsp;
<select id="bi_mode" name="bi_mode" class="drpdwn" LANGUAGE=javascript onchange="selectpage_onchange()">
<option value="" selected>显示所有类型</option>
<option value="100">提醒功能</option>
<option value="101">重复</option>
<option value="">------------</option>
<%
i = 0

do while i < 28
	Response.Write "<option value=""" & i & """>" & getModeName(i) & "</option>" & Chr(13)

	i = i + 1
loop
%>
</select>
</td>
<td nowrap align="right" class=calendar_dayname>
[<b>邀请我参加的活动</b>]&nbsp;
</td></tr>
<tr><td colspan="2">
<form method="post" action="cal_del.asp" name="f2">
<input type="hidden" name="returl" value="<%=bakurl & add_bakurl %>">
<input type="hidden" name="calmode" value="7">
<table width="100%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr bgcolor="#f2f4f6">
    <td width="4%" height="24" class="st_l"><input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()"></td>
	<td width="18%" class="st_l"><%
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
	<td width="11%" class="st_l">时间</td>
	<td width="53%" colspan="2" class="st_l"><%
if sortby = "2" then
	if sortmode = "0" then
		Response.Write "<a class='urf' href=""javascript:setsort('2', '1')"">事件</a>&nbsp;<a href=""javascript:setsort('2', '1')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "<a class='urf' href=""javascript:setsort('2', '0')"">事件</a>&nbsp;<a href=""javascript:setsort('2', '0')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	end if
else
	Response.Write "<a class='urf' href=""javascript:setsort('2', '0')"">事件</a>"
end if
%></td>
	<td width="14%" class="st_r"><%
if sortby = "1" then
	if sortmode = "0" then
		Response.Write "<a class='urf' href=""javascript:setsort('1', '1')"">类型</a>&nbsp;<a href=""javascript:setsort('1', '1')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "<a class='urf' href=""javascript:setsort('1', '0')"">类型</a>&nbsp;<a href=""javascript:setsort('1', '0')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	end if
else
	Response.Write "<a class='urf' href=""javascript:setsort('1', '0')"">类型</a>"
end if
%></td>
	</tr>
<%
si = 0
i = 0
ecal.timeMode = 0

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
		Response.Write "	<td align='center' class='cont_td'><input type='checkbox' name='check" & si & "' value='" & ecal.bi_id & "'></td>"

		Response.Write "	<td align='left' class='cont_td'><a href='" & get_tab_day_url_by_date(show_bi_start_date) & "'>" & get_show_date(show_bi_start_date) & "</a></td>"
		Response.Write "	<td align='left' class='cont_td'>" & get_show_time(show_bi_start_date) & "</td>"
		Response.Write "	<td align='left' class='cont_td'><a href=""javascript:showevent('" & ecal.bi_id & "')"">" & server.htmlencode(ecal.bi_name) & "</a></td>"

		Response.Write "	<td align='right' width='1%' class='cont_td'>&nbsp;<script>write_getShowIconStr(" & ecal.bi_mode & "," & LCase(CStr(ecal.bi_remind)) & "," & LCase(CStr(ecal.bi_isRepeat)) & ")</script></td>"

		Response.Write "	<td align='center' class='cont_td'>" & server.htmlencode(getModeName(ecal.bi_mode)) & "</td>"
		Response.Write "</tr>" & Chr(13)

		si = si + 1
	end if

	show_bi_start_date = NULL
	i = i + 1
loop
%>
</table>
</form>
	</td></tr>
</table>

</td></tr>
</table>
</body>

<SCRIPT language=javascript>
<!--
function window_onload()
{
<%
	if Len(vmd) > 0 then
		Response.Write "document.getElementById(""bi_mode"").value = """ & vmd & """"
	end if
%>
}

function showevent(evid)
{
	location.href = "cal_showinvite.asp?<%=getGRSN() %>&calid=" + evid + "&returl=<%=Server.URLEncode(bakurl & add_bakurl) %>";
}

function setsort(p_sortby, p_sortmode)
{
	location.href = "<%=bakurl %>&page=" + document.getElementById("page").value + "&vsd=" + document.getElementById("view_search_date").value + "&vmd=" + document.getElementById("bi_mode").value + "&sortby=" + p_sortby + "&sortmode=" + p_sortmode;
}

function allcheck_onclick() {
	if (document.f2.allcheck.checked == true)
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
		theObj = eval("document.f2.check" + i);

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
		theObj = eval("document.f2.check" + i);

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
		if (confirm("确实要删除吗?") == false)
			return ;

		document.f2.submit();
	}
}
//-->
</SCRIPT>
</html>

<%
set ecal = nothing
set ecalset = nothing


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
%>
