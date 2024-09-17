<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
dim ecal
set ecal = server.createobject("easymail.CalTask")

if Len(Session("svcal")) < 1 then
	ecal.Load Session("wem")
else
	ecal.Load Session("svcal")
end if

returl = trim(request("returl"))
editcal = trim(request("editcal"))
calid = trim(request("calid"))
is_edit = false

isspd = trim(request("isspd"))

if isspd = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if Len(Session("svcal")) > 0 then
		Response.Redirect "noadmin.asp"
	end if

	sp_title = trim(request("sp_title"))
	sp_start_year = trim(request("sp_start_year"))
	sp_start_month = trim(request("sp_start_month"))
	sp_start_day = trim(request("sp_start_day"))
	sp_level = trim(request("sp_level"))
	sp_ti_is_set_end = trim(request("sp_ti_is_set_end"))

	isok = false

	if Len(sp_title) > 0 then
		ecal.Load Session("wem")
		ecal.ti_title = sp_title

		if sp_ti_is_set_end = "" or sp_ti_is_set_end = "0" then
			ecal.ti_is_set_end = false
		else
			ecal.ti_is_set_end = true
			ecal.set_ti_end_date CLng(sp_start_year), CLng(sp_start_month), CLng(sp_start_day)
		end if

		ecal.ti_level = CLng(sp_level)
		ecal.ti_state = false
		ecal.ti_share_mode = 0

		isok = true
		if ecal.CreateNew() = false then
			isok = false
		end if

		if isok = true then
			isok = ecal.Save()
		end if
	end if

	set ecal = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	end if
end if


if Len(calid) > 10 and editcal = "1" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	is_edit = true
	isok = false

	if ecal.MoveToID(calid) = true then
		if Len(Session("svcal")) < 1 or ecal.ti_share_mode = 1 then
			isok = true
		end if
	end if

	if isok = false then
		set ecal = nothing

		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_tasknew.asp")
		end if
	end if
end if


ti_title = trim(request("ti_title"))
if Len(ti_title) > 0 and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if Len(Session("svcal")) > 0 then
		Response.Redirect "noadmin.asp"
	end if

	ecal.ti_title = ti_title
	ti_end_date = trim(request("ti_end_date"))

	ti_is_set_end = trim(request("ti_is_set_end"))
	if ti_is_set_end = "" or ti_is_set_end = "0" then
		ecal.ti_is_set_end = false
	else
		ecal.ti_is_set_end = true

		if Len(ti_end_date) = 10 then
			bi_start_date_year = Clng(Mid(ti_end_date, 1, 4))
			bi_start_date_month = Clng(Mid(ti_end_date, 6, 2))
			bi_start_date_day = Clng(Mid(ti_end_date, 9, 2))
			ecal.set_ti_end_date bi_start_date_year, bi_start_date_month, bi_start_date_day
		else
			ecal.ti_is_set_end = false
		end if
	end if

	ecal.ti_level = CLng(trim(request("ti_level")))

	ti_state = trim(request("ti_state"))
	if ti_state = "" or ti_state = "0" then
		ecal.ti_state = false
	else
		ecal.ti_state = true
	end if

	ecal.ti_share_mode = CLng(trim(request("ti_share_mode")))
	ecal.ti_note = trim(request("ti_note"))

	isok = true

	if editcal <> "2" then
		if ecal.CreateNew() = false then
			isok = false
		end if
	else
		isok = ecal.Set(calid)
	end if

	if isok = true then
		isok = ecal.Save()
	end if

	set ecal = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_tasknew.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_tasknew.asp")
		end if
	end if
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
.tl {height:24px; text-align:right; border-bottom:1px #8CA5B5 solid;}
.trr {text-align:left; border-bottom:1px #8CA5B5 solid;}
-->
</STYLE>
</head>

<SCRIPT language=javascript src="images/cal/popcalendar.js"></SCRIPT>
<SCRIPT Language="JavaScript">dateFormat='yyyy-mm-dd'</SCRIPT>

<script type="text/javascript">
<!--
function window_onload() {
	init();
<%
if is_edit = false then
	Response.Write "document.f1.ti_is_set_end_true.checked = true;" & Chr(13)
	Response.Write "document.f1.ti_level.value = ""5"";" & Chr(13)
else
	if ecal.ti_is_set_end = true then
		Response.Write "document.f1.ti_is_set_end_true.checked = true;" & Chr(13)
	else
		Response.Write "document.f1.ti_is_set_end_false.checked = true;" & Chr(13)
	end if

	Response.Write "document.f1.ti_level.value = """ & ecal.ti_level & """;" & Chr(13)
end if
%>
}

function goback()
{
	if (document.f1.returl.value.length < 3)
		history.back();
	else
		location.href=document.f1.returl.value;
}

function gosub()
{
	if (document.f1.ti_title.value.length < 1)
	{
		alert("请输入“名称”项");
		document.f1.ti_title.focus();
		return ;
	}

	document.f1.submit();
}
<%
if is_edit = true then
%>
function godel()
{
	if (confirm("确实要删除吗?") == false)
		return ;

	location.href = "cal_del.asp?<%=getGRSN() %>&calmode=5&calid=<%=calid %>&returl=<%=Server.URLEncode(returl) %>";
}

function setover()
{
	if (confirm("确实要标记为完成吗?") == false)
		return ;

	location.href = "cal_del.asp?<%=getGRSN() %>&calmode=6&calid=<%=calid %>&returl=<%=Server.URLEncode(returl) %>";
}
<%
end if
%>
//-->
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
<br>
<form method="post" action="cal_tasknew.asp" name="f1">
<input type="hidden" name="returl" value="<%=returl %>">
<input type="hidden" name="editcal" value="<%
if editcal = "1" then
	Response.Write "2"
end if
%>">
<input type="hidden" name="calid" value="<%=calid %>">

<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td colspan="2" style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%
if is_edit = false then
	Response.Write "添加待办事项"
else
	Response.Write "查看待办事项"
end if
%>
</td></tr>
<tr><td colspan=2 class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" cellspacing="0" cellpadding="0">
	<tr> 
	<td align="center">
	<table width="100%" border="0" align="center" cellspacing="0" style="border-top:1px #8CA5B5 solid;">
		<tr>
		<td colspan=2 height="24" valign=center align=left bgcolor="#DBEAF5" style='border-bottom:1px #8CA5B5 solid;'> 
		&nbsp;<b>基本信息</b>
		</td>
		</tr>

		<tr>
		<td valign=center width="15%" class="tl"> 
		<b>名称</b><%=s_lang_mh %>
		</td>
		<td class="trr">
		<input type="text" name="ti_title" class='n_textbox' value="<%=ecal.ti_title %>" size="50" maxlength="40">
		</td>
		</tr>

		<tr>
		<td valign=center class="tl"> 
		<b>到期日</b><%=s_lang_mh %>
		</td>
		<td class="trr">
		<input type=radio value="1" name="ti_is_set_end" id="ti_is_set_end_true">
		<input type="text" name="ti_end_date" class='n_textbox' value="<%
if is_edit = false then
	curDate = Now
	bs_year = Year(curDate)
	bs_month = Month(curDate)
	bs_day = Day(curDate)

	pbs_year = bs_year
	pbs_month = bs_month
	pbs_day = bs_day

	if Len(bs_month) = 1 then
		bs_month = "0" & bs_month
	end if

	if Len(bs_day) = 1 then
		bs_day = "0" & bs_day
	end if

	if Len(bs_year) > 3 and Len(bs_month) > 0 and Len(bs_day) > 0 then
		Response.Write bs_year & "-" & bs_month & "-" & bs_day
	end if
else
	ecal.get_ti_end_date b_year, b_month, b_day
	start_date_isok = false

	if b_year > 1971 and b_month > 0 and b_day > 0 then
		start_date_isok = true
	end if

	if start_date_isok = true then
		s_b_month = b_month
		if b_month < 10 then
			s_b_month = "0" & b_month
		end if

		s_b_day = b_day
		if b_day < 10 then
			s_b_day = "0" & b_day
		end if

		Response.Write b_year & "-" & s_b_month & "-" & s_b_day
	end if
end if
%>" readonly size="20" maxlength="16">
<script language='javascript'> 
<!--
if (!document.layers) {
	document.write("<img align=absmiddle style='CURSOR:pointer' src='images/cal/calendar.gif' onclick='popUpCalendar(this, document.f1.ti_end_date, dateFormat,-1,-1)' title='Select'>");
}
//-->
</script>
		&nbsp;<span id="Out_Week_Name"><%
if is_edit = true and start_date_isok = true then
	my_Date = DateSerial(b_year, b_month, b_day)
	Response.Write "<b>" & getWeekName3(Weekday(my_Date) - 1) & "</b>"
	my_Date = NULL
else
	if Len(pbs_year) > 0 and Len(pbs_month) > 0 and Len(pbs_day) > 0 then
		my_Date = DateSerial(pbs_year, pbs_month, pbs_day)
		Response.Write "<b>" & getWeekName3(Weekday(my_Date) - 1) & "</b>"
		my_Date = NULL
	end if
end if
%></span>
		<br>
		<input type=radio value="0" name="ti_is_set_end" id="ti_is_set_end_false">未设到期日
				</td>
			</tr>
			<tr>
				<td valign=center class="tl"> 
		<b>优先级</b><%=s_lang_mh %>
				</td>
				<td class="trr">
		<select name="ti_level" class="drpdwn">
<%
i = 1

do while i < 10
	Response.Write "<option value=""" & i & """>" & i & "</option>" & Chr(13)

	i = i + 1
loop
%>
		</select>
				</td>
			</tr>
			<tr>
				<td valign=center class="tl"> 
		<b>状态</b><%=s_lang_mh %>
				</td>
				<td class="trr">
		<input type=radio <% if ecal.ti_state = false then response.write "checked"%> value="0" name="ti_state" id="ti_state_false">未完成
		<br>
		<input type=radio <% if ecal.ti_state = true then response.write "checked"%> value="1" name="ti_state" id="ti_state_true">完成
				</td>
			</tr>
			<tr>
				<td valign=center class="tl"> 
		<b>共享</b><%=s_lang_mh %>
				</td>
				<td class="trr">
<%
if is_edit = true then
%>
		<input type=radio <% if ecal.ti_share_mode = 0 then response.write "checked"%> value="0" name="ti_share_mode" id="ti_share_mode_pri">私人的&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<a href="help.asp#calshare">什么是共享?</a>]
		<br>
		<input type=radio <% if ecal.ti_share_mode = 1 then response.write "checked"%> value="1" name="ti_share_mode" id="ti_share_mode_pub">公开的
<%
else
	dim ecalset
	set ecalset = server.createobject("easymail.CalOptions")
	ecalset.Load Session("wem")
	tmp_ti_share_mode = ecalset.TaskShareDefault
	set ecalset = nothing
%>
		<input type=radio <% if tmp_ti_share_mode = 0 then response.write "checked"%> value="0" name="ti_share_mode" id="ti_share_mode_pri">私人的&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<a href="help.asp#calshare">什么是共享?</a>]
		<br>
		<input type=radio <% if tmp_ti_share_mode = 1 then response.write "checked"%> value="1" name="ti_share_mode" id="ti_share_mode_pub">公开的
<%
end if
%>
				</td>
			</tr>
			<tr>
				<td valign=center class="tl"> 
		<b>便笺</b><%=s_lang_mh %>
				</td>
				<td class="trr">
		<textarea name="ti_note" cols="55" rows="5" class='n_textarea'><%=ecal.ti_note %></textarea><br>
		最多 800 个字符
				</td>
			</tr>
		</table>
	</td></tr>
	</table>
</td></tr>

<tr><td colspan="2" align="left" style="background-color:white; padding-right:16px; padding-top:12px; padding-bottom:10px;">
<a class='wwm_btnDownload btn_blue' href="javascript:goback();"><< <%=s_lang_return %></a>
<%
if Len(Session("svcal")) < 1 then
%>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();">保存</a>
<%
	if is_edit = true then
		if ecal.ti_state = false then
%>
<a class='wwm_btnDownload btn_blue' href="javascript:setover();">标记为完成</a>
<%
		end if
%>
<a class='wwm_btnDownload btn_blue' href="javascript:godel();">删除</a>
<%
	end if
end if
%>
</td></tr>
</table>
</form>
</body>
</html>

<%
b_year = NULL
b_month = NULL
b_day = NULL

set ecal = nothing


function getWeekName3(wknum)
	temp_wk_str = ""

	if wknum = "0" then
		temp_wk_str = "星期日"
	elseif wknum = "1" then
		temp_wk_str = "星期一"
	elseif wknum = "2" then
		temp_wk_str = "星期二"
	elseif wknum = "3" then
		temp_wk_str = "星期三"
	elseif wknum = "4" then
		temp_wk_str = "星期四"
	elseif wknum = "5" then
		temp_wk_str = "星期五"
	elseif wknum = "6" then
		temp_wk_str = "星期六"
	end if

	getWeekName3 = temp_wk_str
end function
%>
