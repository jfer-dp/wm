<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
dim ecalset
set ecalset = server.createobject("easymail.CalOptions")
ecalset.Load Session("wem")

show_APM = false
if ecalset.Show24Hour = false then
	show_APM = true
end if

dim ecal
set ecal = server.createobject("easymail.Calendar")

editbak = trim(request("editbak"))
returl = trim(request("returl"))
searchstr = trim(request("searchstr"))
st = trim(request("st"))
sn = trim(request("sn"))
sm = trim(request("sm"))
if Len(st) > 0 then
	st = "1"
end if

if Len(sn) > 0 then
	sn = "1"
end if
sortmode = trim(request("sortmode"))

if Len(sm) < 1 then
	sm = "-1"
end if

bakurl = "cal_search.asp?" & getGRSN() & "&searchstr=" & Server.URLEncode(searchstr) & "&st=" & st & "&sn=" & sn & "&sm=" & sm & "&editbak=1" & "&returl=" & Server.URLEncode(returl)
allnum = 0

if editbak = "" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	if st = "" and sn = "" then
		st = "1"
	end if
end if

if (st = "1" or sn = "1") and Len(searchstr) > 0 and (Request.ServerVariables("REQUEST_METHOD") = "POST" or editbak = "1") then
	ecal.SortListMode = 0

	if Len(Session("svcal")) < 1 then
		ecal.Load Session("wem")
	else
		ecal.Load Session("svcal")
	end if

	search_title = false
	if st = "1" then
		search_title = true
	end if

	search_note = false
	if sn = "1" then
		search_note = true
	end if

	if Len(sm) > 0 then
		if IsNumeric(sm) = true then
			sm = CLng(sm)
		else
			sm = -1
		end if
	else
		sm = -1
	end if

	if sm < 0 or sm > 27 then
		sm = -1
	end if

	ecal.Search searchstr, search_title, search_note, sm

	if Len(Session("svcal")) > 0 then
		ecal.HidePrivate
		ecal.HideBusy
	end if

	allnum = ecal.Count
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
body {font-family:<%=s_lang_font %>; font-size:9pt;color:#000000;margin-top:5px;margin-left:10px;margin-right:10px;margin-bottom:2px;background-color:#ffffff}
.sbttn {font-family:<%=s_lang_font %>; font-size:9pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5;color:#000066;height:19px;text-decoration:none;cursor:pointer}
.textbox {BORDER:1px #555555 solid;}
.st_l,.st_r {text-align:center; white-space:nowrap; border-left:1px solid #A5B6C8; border-top:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8;}
.st_r {border-right:1px solid #A5B6C8;}
.cont_tr {background:white; height:22px;}
.cont_td {height:22px; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px;}

.top_del {
	BORDER-TOP: #c0c0c0 1px solid;
	BORDER-LEFT: #ffffc0 1px solid;
	BORDER-BOTTOM: #ffffc0 0px solid;
	FONT-WEIGHT: normal;
	color: #202020;
	BACKGROUND-COLOR: #ffffc0;
}
.bottom_del {
	BORDER-TOP: #ffffc0 0px solid;
	BORDER-LEFT: #ffffc0 1px solid;
	BORDER-BOTTOM: #c0c0c0 1px solid;
	FONT-WEIGHT: normal;
	color: #202020;
	BACKGROUND-COLOR: #ffffc0;
}
-->
</STYLE>
</head>

<script language="JavaScript">
<!--
function window_onload()
{
	document.f1.sm.value = "<%=sm %>";
}

function goback()
{
	if (document.f1.returl.value.length < 3)
		history.back();
	else
		location.href=document.f1.returl.value;
}

function searchit()
{
	if (document.f1.searchstr.value.length > 0)
	{
		if (document.f1.st.checked == false && document.f1.sn.checked == false)
			alert("请选定检索范围");
		else
			document.f1.submit();
	}
	else
	{
		alert("请输入“检索词”项");
		document.f1.searchstr.focus();
	}
}

function getShowIconStr(bmode, bremind, brp)
{
	var s_str = "";

	if (bmode == 3)
		s_str = s_str + "<img src='images/cal/bdc.gif' border=0 align='absmiddle'>";

	if (bremind == true)
		s_str = s_str + "<img src='images/cal/bell.gif' border=0 align='absmiddle'>";

	if (brp == true)
		s_str = s_str + "<img src='images/cal/repeat.gif' border=0 align='absmiddle'>";

	if (s_str.length > 0)
		s_str = s_str + "<br>";

	return s_str;
}

function write_getShowIconStr(bmode, bremind, brp)
{
	document.write(getShowIconStr(bmode, bremind, brp));
}

function showevent(evid)
{
	location.href = "cal_new.asp?<%=getGRSN() %>&editcal=1&calid=" + evid + "&returl=<%=Server.URLEncode(bakurl) %>";
}

function allcheck_onclick() {
	if (document.f1.allcheck.checked == true)
		checkall(true);
	else
		checkall(false);
}

function checkall(check) {
	var i = 0;
	var theObj;

	for(; i<<%=allnum %>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			theObj.checked = check;
	}
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

function delmulevent()
{
	if (ischeck() == true)
	{
<%
if Len(Session("svcal")) < 1 then
%>
		if (confirm("确实要删除吗?") == false)
			return ;

		document.f1.action = "cal_del.asp";
		document.f1.returl.value = "<%=bakurl %>";
		document.f1.submit();
<%
else
%>
		alert("您的权限不足.");
<%
end if
%>
	}
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
<form method="post" action="cal_search.asp" name="f1">
<input type="hidden" name="returl" value="<%=returl %>">
<input type="hidden" name="calmode" value="2">
<table width="90%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr>
	<td align="left" height="28" width="55%" nowrap style="padding-left:4px;">
<a class='wwm_btnDownload btn_blue' href="javascript:goback();"><< <%=s_lang_return %></a>
	</td>
	<td align="right" width="25%" nowrap style="padding-right:8px; color:#444444;">高级搜索</td>
	</tr>
</table>
<br>

<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr> 
	<td>
		<table width="100%" border="0" align="center" cellspacing="0" bgcolor="white" style="border-top:1px #8CA5B5 solid;">
		<tr>
		<td colspan=2 height="24" valign=center align=left bgcolor="#DBEAF5" style='border-bottom:1px #8CA5B5 solid;'> 
		&nbsp;<b><%
if Len(searchstr) > 0 then
	Response.Write "[" & server.htmlencode(searchstr) & "]的检索结果"
else
	Response.Write "检索结果"
end if
%></b>
		</td>
		</tr>

		<tr>
		<td valign=center width="20%" align=right style='height:60px; border-bottom:1px #8CA5B5 solid;'>
		<b>检索词</b><%=s_lang_mh %>
		</td>
		<td align=left style='border-bottom:1px #8CA5B5 solid;'>
		<input type="text" name="searchstr" class='textbox' value="<%=searchstr %>" size="30" maxlength="50">
<input type="button" value="搜索" style="WIDTH: 40px" onclick="javascript:searchit()" class="sbttn"><br><font color="#444444">(最多 50 个字符)</font><br>
<input type="checkbox" name="st" <% if st = "1" then Response.Write "checked" %>>在[名称]栏中检索&nbsp;&nbsp;
<input type="checkbox" name="sn" <% if sn = "1" then Response.Write "checked" %>>在[便笺]栏中检索
		</td>
		</tr>

		<tr>
		<td valign=center height=32 align=right> 
<b>活动类型</b><%=s_lang_mh %>
		</td>
		<td align=left>
<select name="sm" class="drpdwn">
<option value="-1" selected>所有类型</option>
<%
i = 0

do while i < 28
	Response.Write "<option value=""" & i & """>" & getModeName(i) & "</option>" & Chr(13)

	i = i + 1
loop
%>
</select>
		</td>
		</tr>
<%
if allnum < 1 then
%>
		<tr bgcolor="#FFF8D3"> 
		<td valign=center colspan=2 height="28" style="border:#ff0000 solid 1px;">
		<div style="font-size:9pt;">&nbsp;请输入更多的关键字后再试一次.</div>
		</td>
		</tr>
<%
else
%>
		<tr>
<td class=top_del colspan=2 height="27" valign=center>
&nbsp;<input type="button" value="删除" style="WIDTH: 40px" onclick="javascript:delmulevent()" class="sbttn">
</td></tr>
		<tr>
		<td bgcolor="#ffffff" valign=center colspan=2>
<table width="100%" border="0" align="center" bgcolor="#ffffff" cellspacing="0">
<tr bgcolor="#f2f4f6">
    <td width="5%" height="25" class="st_l"><input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()"></td>
	<td width="20%" class="st_l">日期</td>
	<td width="13%" class="st_l">时间</td>
	<td width="46%" colspan="2" class="st_l">事件</td>
    <td width="16%" class="st_r">类型</td>
  </tr>
<%
i = 0
ecal.timeMode = 0

do while i < allnum
	showi = i
	ecal.MoveTo showi
	show_bi_start_date = ecal.show_bi_start_date

	Response.Write "<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);'>"
	Response.Write "	<td align='center' class='cont_td'><input type='checkbox' name='check" & showi & "' value='" & ecal.bi_id & "'></td>"

	Response.Write "	<td align='left' nowrap class='cont_td'><a href='" & get_tab_day_url_by_date(show_bi_start_date) & "'>" & get_show_date(show_bi_start_date) & "</a></td>"
	Response.Write "	<td align='left' nowrap class='cont_td'>" & get_show_time(show_bi_start_date) & "</td>"
	Response.Write "	<td align='left' class='cont_td'><a href=""javascript:showevent('" & ecal.bi_id & "')"">" & server.htmlencode(ecal.bi_name) & "</a></td>"
	Response.Write "	<td align='right' nowrap class='cont_td'>&nbsp;<script>write_getShowIconStr(" & ecal.bi_mode & "," & LCase(CStr(ecal.bi_remind)) & "," & LCase(CStr(ecal.bi_isRepeat)) & ")</script></a></td>"

	Response.Write "	<td align='center' nowrap class='cont_td'>" & server.htmlencode(getModeName(ecal.bi_mode)) & "</td>"
	Response.Write "</tr>" & Chr(13)

	show_bi_start_date = NULL
	i = i + 1
loop
%>
</table>
				</td>
			</tr>
<tr>
<td class=bottom_del colspan=2 height="27" valign=center>
&nbsp;<input type="button" value="删除" style="WIDTH: 40px" onclick="javascript:delmulevent()" class="sbttn">
</td></tr>
<%
end if
%>
		</table>
	</td></tr>
</table>
</form>
</body>
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
