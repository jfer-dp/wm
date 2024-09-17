<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
fmeml = trim(request("fmeml"))
fmcal = trim(request("fmcal"))
returl = trim(request("returl"))
calid = trim(request("calid"))

set ecalnt = server.createobject("easymail.CalendarNotice")

if fmcal <> "1" then
	ecalnt.Load Session("wem")
	ecalnt.MoveToID calid
end if

dim host_name
dim host_account
dim is_host_account
is_host_account = false

isok = true
dim ecal
set ecal = server.createobject("easymail.Calendar")

if fmcal = "1" then
	if Len(Session("svcal")) < 1 then
		if ecal.Load(Session("wem")) = false then
			isok = false
		end if
	else
		if ecal.Load(Session("svcal")) = false then
			isok = false
		end if
	end if

	if isok = true then
		isok = ecal.MoveToID(calid)
	end if

	host_account = ecal.bi_host_account
	if Len(host_account) < 1 then
		if Len(Session("svcal")) < 1 then
			host_account = Session("wem")
		else
			host_account = Session("svcal")
		end if
	end if
else
	host_account = ecalnt.bi_host_account

	if ecal.Load(host_account) = false then
		isok = false
	end if

	if isok = true then
		isok = ecal.MoveToID(calid)
	end if
end if


if Len(Session("svcal")) > 0 and ecal.bi_shareMode <> 2 then
	set ecal = nothing
	set ecalnt = nothing

	Response.Redirect "noadmin.asp"
end if


if isok = false then
	set ecal = nothing
	set ecalnt = nothing

	Response.Redirect "err.asp?" & getGRSN() & "&errstr=" & Server.URLEncode("打开活动失败, 活动发起人可能已删除此活动")
end if


if LCase(host_account) = LCase(Session("wem")) then
	is_host_account = true
end if

set ecalext = server.createobject("easymail.CalendarExtend")

if Len(host_account) > 0 then
	ecalext.Load host_account, ecal.bi_id
end if

ishave = ecalext.haveit(Session("mail"))

if ishave >= 0 then
	ecalext.MoveTo ishave
end if

if trim(request("RemoveOwn")) = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if Len(Session("svcal")) > 0 then
		Response.Redirect "noadmin.asp"
	end if

	ecalext.ce_join = -1
	ecalext.ce_withGuest = 0
	ecalext.ce_askRemove = true

	isok = false
	if ecalext.Set(Session("mail")) = true then
		if ecalext.Save() = true then
			isok = true
		end if
	end if

	set ecalext = nothing
	set ecal = nothing
	set ecalnt = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_listinvited.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_listinvited.asp")
		end if
	end if
end if


if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if Len(Session("svcal")) > 0 then
		Response.Redirect "noadmin.asp"
	end if

	ce_join = CLng(trim(request("joinmd")))

	ecalext.ce_email = Session("mail")
	ecalext.ce_myname = trim(request("myname"))
	ecalext.ce_join = ce_join
	ecalext.ce_remark = trim(request("mynote"))
	ecalext.ce_username = Session("wem")
	ecalext.ce_askRemove = false

	ecalext.ce_withGuest = 0
	if ce_join >= 0 then
		myghostnum = trim(request("myghostnum"))
		if IsNumeric(myghostnum) = true then
			ecalext.ce_withGuest = CLng(myghostnum)
		end if
	end if

	isok = false
	if ecalext.Set(Session("mail")) = true then
		if ecalext.Save() = true then
			isok = true
		end if
	end if

	if isok = true and fmcal <> "1" and ce_join = 1 then
		if ecalnt.SaveToCalendar(Session("wem")) = true then
			if ecalnt.RemoveByID(calid) = true then
				ecalnt.Save
			end if
		end if
	end if

	set ecalext = nothing
	set ecal = nothing
	set ecalnt = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_listinvited.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_listinvited.asp")
		end if
	end if
end if


dim ecalset
set ecalset = server.createobject("easymail.CalOptions")
ecalset.Load Session("wem")

show_APM = false
if ecalset.Show24Hour = false then
	show_APM = true
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
.sbttn {font-family:<%=s_lang_font %>; font-size:9pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5;color:#000066;height:19px;text-decoration:none;cursor:pointer}
a:hover {text-decoration:underline;}

.mjNoLine {
	text-decoration: none;
}
.mjRemove {
	text-decoration: line-through;
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

<script type="text/javascript" src="images/mglobal.js"></script>

<script type="text/javascript">
<!--
if (ie == 6)
	document.execCommand("BackgroundImageCache", false, true);

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

function delevent(evid)
{
<%
if Len(Session("svcal")) < 1 then
%>
	if (confirm("确实要删除吗?") == false)
		return ;

	location.href = "cal_del.asp?<%=getGRSN() %>&calmode=1&calid=" + evid + "&returl=<%=Server.URLEncode(returl) %>";
<%
else
%>
	alert("您的权限不足.");
<%
end if
%>
}

function getShowUsersIconStr(eid, byes, bwait, bno)
{
	if (byes < 1 && bwait < 1 && bno < 1)
		return "";

	var s_str = "<br>";

	s_str = s_str + "<font face='Arial, Helvetica, sans-serif'><img src='images/cal/a.gif' border=0 title='参加'>&nbsp;" + byes.toString();
	s_str = s_str + "&nbsp;&nbsp;<img src='images/cal/u.gif' border=0 title='未决定的'>&nbsp;" + bwait.toString();
	s_str = s_str + "&nbsp;&nbsp;<img src='images/cal/d.gif' border=0 title='婉言拒绝'>&nbsp;" + bno.toString();

	s_str = s_str + "&nbsp;-&nbsp;</font>[<a class=mjEL href='viewInv(\"" + eid + "\")'>查看请柬</a>]";

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

	if (s_str.length > 0)
		s_str = s_str + "&nbsp;";

	return s_str;
}

function write_getShowIconStr(bmode, bremind, brp)
{
	document.write(getShowIconStr(bmode, bremind, brp));
}

function vieweff()
{
<%
if fmeml <> "1" then
	if Len(returl) > 0 then
%>
	location.href = "<%=returl %>";
<%
	else
%>
	location.href = "cal_listinvited.asp?<%=getGRSN() %>";
<%
	end if
else
%>
	history.back();
<%
end if
%>
}

function showevent(evid)
{
	location.href = "cal_new.asp?<%=getGRSN() %>&editcal=1&calid=" + evid + "&returl=<%=Server.URLEncode("cal_showinvite.asp?" & getGRSN() & "&fmcal=" & fmcal & "&calid=" & calid) %>&purl=<%=Server.URLEncode(returl) %>";
}

function updateinv()
{
	document.f1.submit();
}

function fRemoveOwn()
{
<%
if ecalext.ce_askRemove = false then
%>
	document.f1.RemoveOwn.value = "1";
<%
else
%>
	document.f1.RemoveOwn.value = "2";
<%
end if
%>
	document.f1.submit();
}

function setover()
{
	if (ischeck() == true)
	{
		if (confirm("确实要标记为完成吗?") == false)
			return ;

		f2.calmode.value = "4";
		f2.submit();
	}
}

function write_ShowDateUrl(by, bm, bd)
{
	var showS_Str = "";
	currentDate = new Date(by, bm - 1, bd);

	showS_Str = "<a href=\"cal_index.asp?<%=getGRSN() %>&tsn=0&sy=" + by + "&sm=" + bm + "&sd=" + bd + "\">";
	showS_Str = showS_Str + by + "年" + bm + "月" + bd + "日 " + convWeeekName(currentDate.getDay());
	showS_Str = showS_Str + "</a>";

	document.write(showS_Str);
}

function write_ShowEventTime(vy, vm, vd, vh, vmin, vnt)
{
	document.write(getShowStartStr(vy, vm, vd, vh, vmin, vnt));
}

function convWeeekName(wnum) {
	if (wnum > 6)
		wnum = wnum - 7;

	if (wnum == 0)
		return "星期日";
	else if (wnum == 1)
		return "星期一";
	else if (wnum == 2)
		return "星期二";
	else if (wnum == 3)
		return "星期三";
	else if (wnum == 4)
		return "星期四";
	else if (wnum == 5)
		return "星期五";
	else if (wnum == 6)
		return "星期六";
}

function join_onclick(joinnum)
{
<%
if Len(Session("svcal")) < 1 then
	if fmcal <> "1" and is_host_account = false and ecalext.ce_askRemove = false then
%>
	if (joinnum == 1)
	{
		document.f1.updatebutton.value = "更新并加入我的效率手册";
		document.f1.updatebutton.style.width = "160px";
	}
	else
	{
		document.f1.updatebutton.value = "更新";
		document.f1.updatebutton.style.width = "50px";
	}
<%
	end if
end if
%>
}
//-->
</script>

<body LANGUAGE=javascript onload="return window_onload()">
<form method="post" action="cal_showinvite.asp" name="f1">
<input type="hidden" name="returl" value="<%=returl %>">
<input type="hidden" name="calid" value="<%=calid %>">
<input type="hidden" name="RemoveOwn">
<input type="hidden" name="fmcal" value="<%=fmcal %>">
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-top:6px;">
	<tr>
	<td height="33" width="75%" align="left" bgcolor="#ffffff" style="border-left:1px #8CA5B5 solid; border-top:1px #8CA5B5 solid; border-bottom:1px #8CA5B5 solid;">
&nbsp;<font style="FONT-SIZE: 15px; color:#104A7B"><b><%
ev_name = server.htmlencode(ecal.bi_name)
Response.Write ev_name
%></b></font>
	</td>
	<td align="right" bgcolor="#ffffff" style="border-right:1px #8CA5B5 solid; border-top:1px #8CA5B5 solid; border-bottom:1px #8CA5B5 solid; padding-right:6px;">
<%
if fmcal <> "1" then
%>
<a class='wwm_btnDownload btn_blue' href="javascript:delnt();">删除</a>
<%
end if
%>
<a class='wwm_btnDownload btn_blue' href="javascript:vieweff();"><%=s_lang_return %></a>
	</td>
	</tr>
	<tr>
	<td width="100%" colspan="2">
<table width=100% border=0 cellspacing=0 cellpadding=0>
<tr><td height="3" colspan="2"></td></tr>
<tr>
<%
if Len(Session("svcal")) < 1 then
%>
<td width="50%" style="border:5px #ffffff solid;" valign="top">
	<table width=100% border=0 cellspacing=0 cellpadding=0 bgcolor="#EFF7FF" style="border:1px #8CA5B5 solid;">
	<tr>
	<td height="26" colspan="2">
&nbsp;<font class=s color="#104A7B"><b>您是否参加?</b></font>
	</td>
	</tr>
	<tr>
	<td colspan="2">
&nbsp;<input type=radio value="1" name="joinmd" id="joinmd1"<% if is_host_account = true then Response.Write " checked" %> LANGUAGE=javascript onclick="return join_onclick(1)"><b>是</b>&nbsp;&nbsp;
<input type=radio checked value="0" name="joinmd" id="joinmd0"<% if is_host_account = true then Response.Write " disabled" %> LANGUAGE=javascript onclick="return join_onclick(0)"><b>未定</b>&nbsp;&nbsp;
<input type=radio value="-1" name="joinmd" id="joinmd2"<% if is_host_account = true then Response.Write " disabled" %> LANGUAGE=javascript onclick="return join_onclick(-1)"><b>不</b><%
if is_host_account = false then
	Response.Write "&nbsp;&nbsp;[<a href=""javascript:fRemoveOwn()"" class=mjNoLine>"

	if ecalext.ce_askRemove = false then
		Response.Write "将我从请柬中删除"
	else
		Response.Write "将我重新加入请柬"
	end if

	Response.Write "</a>]"
end if
%>
	</td>
	</tr>
<tr><td height="15" colspan="2"></td></tr>
	<tr>
	<td nowrap width="35%">
&nbsp;我的邮件地址是:
	</td>
	<td align="left">
<%=Session("mail") %>
	</td>
	</tr>
<tr><td height="10" colspan="2"></td></tr>
	<tr>
	<td nowrap>
&nbsp;我的姓名是:
	</td>
	<td align="left">
<input type="text" name="myname" class='textbox' size="28" value="<%
dim ei
set ei = server.createobject("easymail.UserWeb")
ei.Load Session("wem")

UserWeb_showname = ""

if Len(ei.MailName) < 1 then
	Response.Write Session("wem")
else
	Response.Write ei.MailName
	UserWeb_showname = ei.MailName
end if

set ei = nothing
%>" maxlength="64">
	</td>
	</tr>
<tr><td height="10" colspan="2"></td></tr>
	<tr>
	<td nowrap>
&nbsp;备注:(选填)
	</td>
	<td align="left">
<input type="text" name="mynote" class='textbox' size="28" maxlength="200">
	</td>
	</tr>
<tr><td height="10" colspan="2"></td></tr>
	<tr>
	<td nowrap>
&nbsp;我将带
	</td>
	<td align="left">
<input type="text" name="myghostnum" class='textbox' size="3" maxlength="3">&nbsp;名客人
	</td>
	</tr>
<tr><td height="17" colspan="2"></td></tr>
	<tr>
	<td colspan="2" align="center">
	<input type="button" name="updatebutton" value="更新" style="WIDTH: 50px" onclick="javascript:updateinv()" class="sbttn"<%
if ecalext.ce_askRemove = true then
	Response.Write " disabled"
end if
%>>
	</td>
	</tr>
<tr><td height="8" colspan="2"></td></tr>
	</table>
</td>
<%
else
%>
<input type="hidden" name="joinmd0">
<input type="hidden" name="joinmd1">
<input type="hidden" name="joinmd2">
<input type="hidden" name="myname">
<input type="hidden" name="mynote">
<input type="hidden" name="myghostnum">
<%
end if
%>
<td valign="top">
<table width=100% border=0 cellspacing=0 cellpadding=0>
<tr><td width="100%" style="border:5px #ffffff solid;">
	<table width=100% border=0 cellspacing=0 cellpadding=0 bgcolor="#ffffff" style="border-top:1px #b0b0b0 solid;">
	<tr>
	<td height="23" colspan="2" bgcolor="#eeeeee" style="border-bottom:1px #b0b0b0 solid;">
&nbsp;<b>事件详细信息</b>
<%
if fmcal = "1" then
%>
&nbsp;
[<a href="javascript:showevent('<%=ecal.bi_id %>')" class=mjNoLine><%
if is_host_account = true then
	Response.Write "编辑"
else
	Response.Write "查看"
end if
%>详细信息</a> - <a href="javascript:delevent('<%=ecal.bi_id %>')" class=mjNoLine>删除活动</a>]
<%
end if
%>
</td>
	</tr>
<tr><td height="3" colspan="2"></td></tr>
<tr><td height="6" colspan="2" bgcolor="#ffffc0"></td></tr>
	<tr>
	<td colspan="2" bgcolor="#ffffc0" style="border-left:7px #ffffc0 solid; border-right:7px #ffffc0 solid;">
<font class=s color="#104A7B"><script>write_getShowIconStr(<%=ecal.bi_mode %>,<%=LCase(CStr(ecal.bi_remind)) %>,<%=LCase(CStr(ecal.bi_isRepeat)) %>)</script><b><%=server.htmlencode(ecal.bi_name) %></b></font><br>
<font class=s color="#104A7B"><%
ht = server.htmlencode(ecal.bi_note)
ht = replace(ht, Chr(13), "<br>")
ht = replace(ht, Chr(32), "&nbsp;")
ht = replace(ht, Chr(9), "&nbsp;&nbsp;&nbsp;&nbsp;")
Response.Write ht
%></font>
	</td>
	</tr>
<tr><td height="3" colspan="2" bgcolor="#ffffc0"></td></tr>
<tr><td height="4" colspan="2"></td></tr>
	<tr>
	<td width="<%
if Len(Session("svcal")) < 1 then
	Response.Write "30"
else
	Response.Write "16"
end if
%>%">
&nbsp;<b>活动发起人</b>:
	</td>
	<td id=theObj>
	</td>
	</tr>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b>发起人邮址</b>:
	</td>
	<td>
<%
if LCase(ecal.bi_notice_email) = Session("mail") then
	Response.Write server.htmlencode(ecal.bi_notice_email)
else
	Response.Write "<a href=""mailto:" & server.htmlencode(ecal.bi_notice_email) & "?subject=" & ev_name & """>" & server.htmlencode(ecal.bi_notice_email) & "</a>"
end if
%>
	</td>
	</tr>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b>日期</b>:
	</td>
	<td>
<%
ecal.get_bi_start_date b_year, b_month, b_day, b_hour, b_minute
Response.Write "<script>write_ShowDateUrl(" & b_year & "," & b_month & "," & b_day & ")</script>"
%>
	</td>
	</tr>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b>时间</b>:
	</td>
	<td>
<%
Response.Write "<script>write_ShowEventTime(" & b_year & "," & b_month & "," & b_day & "," & b_hour & "," & b_minute & "," & ecal.bi_needtime & ")</script>"
%>
	</td>
	</tr>
<%
if ecal.bi_isRepeat = true then
%>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b>重复</b>:
	</td>
	<td>
此活动有设定重复功能
	</td>
	</tr>
<%
end if

if Len(ecal.bi_place) > 0 then
%>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b>位置</b>:
	</td>
	<td>
<%=server.htmlencode(ecal.bi_place) %>
	</td>
	</tr>
<%
end if

if Len(ecal.bi_city) > 0 then
%>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b>城市</b>:
	</td>
	<td>
<%=server.htmlencode(ecal.bi_city) %>
	</td>
	</tr>
<%
end if

if Len(ecal.bi_address) > 0 then
%>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b>地址</b>:
	</td>
	<td>
<%=server.htmlencode(ecal.bi_address) %>
	</td>
	</tr>
<%
end if

if Len(ecal.bi_phone) > 0 then
%>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b>电话</b>:
	</td>
	<td>
<%=server.htmlencode(ecal.bi_phone) %>
	</td>
	</tr>
<%
end if

if Len(ecal.bi_other_phone) > 0 then
%>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b>其他</b>:
	</td>
	<td>
<%=server.htmlencode(ecal.bi_other_phone) %>
	</td>
	</tr>
<%
end if

if fmcal <> "1" then
	show_bi_notice_datetime = ecalnt.show_bi_notice_datetime
else
	show_bi_notice_datetime = ecal.show_bi_notice_datetime
end if

if Len(show_bi_notice_datetime) > 0 then
%>
<tr><td height="6" colspan="2"></td></tr>
	<tr>
	<td colspan="2">
&nbsp;<b>您的请柬发送于&nbsp;<%=server.htmlencode(show_bi_notice_datetime) %></b>
	</td>
	</tr>
<%
end if
%>
</table>
<br>
<table width=100% border=0 cellspacing=0 cellpadding=0>
<tr><td width="100%" style="border:5px #ffffff solid;">
	<table width=100% border=0 cellspacing=0 cellpadding=0 bgcolor="#ffffff" style="border-top:1px #b0b0b0 solid;">
	<tr>
	<td height="23" colspan="2" bgcolor="#eeeeee" style="border-bottom:1px #b0b0b0 solid;">
&nbsp;<b>客人名单</b>(共<%=ecalext.All_User - ecalext.Remove_Own_User %>)<%
if is_host_account = true then
%>
&nbsp;&nbsp;[<a href="javascript:showguest()" class=mjNoLine>编辑客人</a> - <a href="javascript:sendmsg()" class=mjNoLine>发送电子邮件</a>]
<%
end if
%></td>
	</tr>
<tr><td height="9" colspan="2"></td></tr>
	<tr>
	<td colspan="2">
&nbsp;<font class=s color="green"><b><%=ecalext.Yes_User %>&nbsp;参加</b></font>
	</td>
	</tr>
<tr><td height="3" colspan="2"></td></tr>
<%
i = 0
allnum = ecalext.Count
findit = false

do while i < allnum
	ecalext.MoveTo i

	if ecalext.ce_join = 1 then
		Response.Write "<tr><td width='"
		if Len(Session("svcal")) < 1 then
			Response.Write "7"
		else
			Response.Write "3"
		end if
		Response.Write "%'>&nbsp;<img src='images/cal/a.gif' border=0 title='参加'></td><td>"

		show_name = server.htmlencode(ecalext.ce_myname)
		if Len(show_name) = 0 then
			show_name = server.htmlencode(ecalext.ce_email)
		end if

		if LCase(ecalext.ce_email) <> Session("mail") then
			Response.Write "<a href=""mailto:" & server.htmlencode(ecalext.ce_email) & "?subject=" & ev_name & """ class=mjNoLine>" & show_name & "</a>"
		else
			Response.Write "<a href=""mailto:" & server.htmlencode(ecalext.ce_email) & "?subject=" & ev_name & """ class=mjNoLine><b>" & show_name & "</b></a>"
		end if

		if LCase(ecalext.ce_username) = LCase(host_account) then
			host_name = ecalext.ce_myname
			Response.Write "(活动发起人)"
		end if

		if ecalext.ce_withGuest > 0 then
			Response.Write "&nbsp;(+" & ecalext.ce_withGuest & ")"
		end if

		if Len(ecalext.ce_remark) > 0 then
			Response.Write "&nbsp;-&nbsp;" & server.htmlencode(ecalext.ce_remark)
		end if

		Response.Write "</td></tr>" & Chr(13)
		findit = true
	end if

    i = i + 1
loop

if findit = false then
	Response.Write "<tr><td colspan='2'>&nbsp;<i>没有客人参加</i></td></tr>" & Chr(13)
end if
%>
<tr><td height="13" colspan="2"></td></tr>
	<tr>
	<td colspan="2">
&nbsp;<font class=s color="gray"><b><%=ecalext.Wait_User %>&nbsp;未做决定</b></font>
	</td>
	</tr>
<tr><td height="3" colspan="2"></td></tr>
<%
i = 0
findit = false

do while i < allnum
	ecalext.MoveTo i

	if ecalext.ce_join = 0 then
		Response.Write "<tr><td width='"
		if Len(Session("svcal")) < 1 then
			Response.Write "7"
		else
			Response.Write "3"
		end if
		Response.Write "%'>&nbsp;<img src='images/cal/u.gif' border=0 title='未决定的'></td><td>"

		show_name = server.htmlencode(ecalext.ce_myname)
		if Len(show_name) = 0 then
			show_name = server.htmlencode(ecalext.ce_email)
		end if

		if LCase(ecalext.ce_email) <> Session("mail") then
			Response.Write "<a href=""mailto:" & server.htmlencode(ecalext.ce_email) & "?subject=" & ev_name & """ class=mjNoLine>" & show_name & "</a>"
		else
			Response.Write "<a href=""mailto:" & server.htmlencode(ecalext.ce_email) & "?subject=" & ev_name & """ class=mjNoLine><b>" & show_name & "</b></a>"
		end if

		if LCase(ecalext.ce_username) = LCase(host_account) then
			host_name = ecalext.ce_myname
			Response.Write "(活动发起人)"
		end if

		if ecalext.ce_withGuest > 0 then
			Response.Write "&nbsp;(+" & ecalext.ce_withGuest & ")"
		end if

		if Len(ecalext.ce_remark) > 0 then
			Response.Write "&nbsp;-&nbsp;" & server.htmlencode(ecalext.ce_remark)
		end if

		Response.Write "</td></tr>" & Chr(13)

		findit = true
	end if

    i = i + 1
loop

if findit = false then
	Response.Write "<tr><td colspan='2'>&nbsp;<i>没有客人未定</i></td></tr>" & Chr(13)
end if
%>
<tr><td height="13" colspan="2"></td></tr>
	<tr>
	<td colspan="2">
&nbsp;<font class=s color="red"><b><%=ecalext.No_User %>&nbsp;谢绝邀请</b></font>
	</td>
	</tr>
<tr><td height="3" colspan="2"></td></tr>
<%
i = 0
findit = false

do while i < allnum
	ecalext.MoveTo i

	if ecalext.ce_join = -1 and ecalext.ce_askRemove = false then
		Response.Write "<tr><td width='"
		if Len(Session("svcal")) < 1 then
			Response.Write "7"
		else
			Response.Write "3"
		end if
		Response.Write "%'>&nbsp;<img src='images/cal/d.gif' border=0 title='婉言拒绝'></td><td>"

		show_name = server.htmlencode(ecalext.ce_myname)
		if Len(show_name) = 0 then
			show_name = server.htmlencode(ecalext.ce_email)
		end if

		if LCase(ecalext.ce_email) <> Session("mail") then
			Response.Write "<a href=""mailto:" & server.htmlencode(ecalext.ce_email) & "?subject=" & ev_name & """ class=mjNoLine>" & show_name & "</a>"
		else
			Response.Write "<a href=""mailto:" & server.htmlencode(ecalext.ce_email) & "?subject=" & ev_name & """ class=mjNoLine><b>" & show_name & "</b></a>"
		end if

		if LCase(ecalext.ce_username) = LCase(host_account) then
			host_name = ecalext.ce_myname
			Response.Write "(活动发起人)"
		end if

		if ecalext.ce_withGuest > 0 then
			Response.Write "&nbsp;(+" & ecalext.ce_withGuest & ")"
		end if

		if Len(ecalext.ce_remark) > 0 then
			Response.Write "&nbsp;-&nbsp;" & server.htmlencode(ecalext.ce_remark)
		end if

		Response.Write "</td></tr>" & Chr(13)

		findit = true
	end if

    i = i + 1
loop

if findit = false then
	Response.Write "<tr><td colspan='2'>&nbsp;<i>没有客人谢绝</i></td></tr>" & Chr(13)
end if

if is_host_account = true and ecalext.Remove_Own_User > 0 then
%>
<tr><td height="13" colspan="2"></td></tr>
	<tr>
	<td colspan="2">
&nbsp;<font class=s color="red"><b><%=ecalext.Remove_Own_User %></b>&nbsp;从请柬中删除的客人</font>
	</td>
	</tr>
<tr><td height="3" colspan="2"></td></tr>
<%
	i = 0
	findit = false

	do while i < allnum
		ecalext.MoveTo i

		if ecalext.ce_join = -1 and ecalext.ce_askRemove = true then
			Response.Write "<tr><td width='"
			if Len(Session("svcal")) < 1 then
				Response.Write "7"
			else
				Response.Write "3"
			end if
			Response.Write "%'>&nbsp;<img src='images/cal/d.gif' border=0 title='已删除'></td><td>"

			show_name = server.htmlencode(ecalext.ce_myname)
			if Len(show_name) = 0 then
				show_name = server.htmlencode(ecalext.ce_email)
			end if

			if LCase(ecalext.ce_email) <> Session("mail") then
				Response.Write "<a href=""mailto:" & server.htmlencode(ecalext.ce_email) & "?subject=" & ev_name & """ class=mjRemove>" & show_name & "</a>"
			else
				Response.Write "<a href=""mailto:" & server.htmlencode(ecalext.ce_email) & "?subject=" & ev_name & """ class=mjRemove><b>" & show_name & "</b></a>"
			end if

			if LCase(ecalext.ce_username) = LCase(host_account) then
				host_name = ecalext.ce_myname
				Response.Write "(活动发起人)"
			end if

			if ecalext.ce_withGuest > 0 then
				Response.Write "&nbsp;(+" & ecalext.ce_withGuest & ")"
			end if

			if Len(ecalext.ce_remark) > 0 then
				Response.Write "&nbsp;-&nbsp;" & server.htmlencode(ecalext.ce_remark)
			end if

			Response.Write "</td></tr>" & Chr(13)

			findit = true
		end if

	    i = i + 1
	loop

	if findit = false then
		Response.Write "<tr><td colspan='2'>&nbsp;<i>没有从请柬中删除的客人</i></td></tr>" & Chr(13)
	end if
end if
%>
</table>
<%
if is_host_account = true then
%>
<br>
<table width=100% border=0 cellspacing=0 cellpadding=0>
<tr><td width="100%" style="border:5px #ffffff solid;">
	<table width=100% border=0 cellspacing=0 cellpadding=0 bgcolor="#ffffff" style="border-top:1px #b0b0b0 solid;">
	<tr>
	<td height="23" colspan="2" bgcolor="#eeeeee" style="border-bottom:1px #b0b0b0 solid;">
&nbsp;<b>主人编辑选项</b>
	</td>
	</tr>
<tr><td height="9"></td></tr>
	<tr>
	<td>
&nbsp;<img src="aw.gif" border="0">&nbsp;<a href="javascript:showguest()" class=mjNoLine>编辑客人</a>
	</td>
	</tr>
<tr><td height="3"></td></tr>
	<tr>
	<td>
&nbsp;<img src="aw.gif" border="0">&nbsp;<a href="javascript:sendmsg()" class=mjNoLine>发送电子邮件</a>
	</td>
	</tr>
<tr><td height="3"></td></tr>
	<tr>
	<td>
&nbsp;<img src="aw.gif" border="0">&nbsp;<a href="javascript:moreguest()" class=mjNoLine>邀请更多客人</a>
	</td>
	</tr>
<tr><td height="3"></td></tr>
	<tr>
	<td>
&nbsp;<img src="aw.gif" border="0">&nbsp;<a href="javascript:newaddress()" class=mjNoLine>将客人添加到我的地址簿</a>
	</td>
	</tr>
</table>
</td>
</tr>
</table>
<%
end if
%>
    </td>
  </tr>
</table>
</form>
</body>

<script language="JavaScript">
<!--
function window_onload()
{
<%
if ishave >= 0 then
	ecalext.MoveTo ishave

	if ecalext.ce_join = 0 then
		Response.Write "document.f1.joinmd0.checked = true;" & Chr(13)
	elseif ecalext.ce_join = 1 then
		Response.Write "document.f1.joinmd1.checked = true;" & Chr(13)
	elseif ecalext.ce_join = -1 then
		Response.Write "document.f1.joinmd2.checked = true;" & Chr(13)
	end if

	Response.Write "join_onclick(" & ecalext.ce_join & ");" & Chr(13)

	if Len(ecalext.ce_myname) > 0 then
		Response.Write "document.f1.myname.value=""" & ecalext.ce_myname & """;" & Chr(13)
	end if

	Response.Write "document.f1.mynote.value=""" & ecalext.ce_remark & """;" & Chr(13)
	Response.Write "document.f1.myghostnum.value=""" & ecalext.ce_withGuest & """;" & Chr(13)
end if

if Len(host_name) > 0 then
	Response.Write "theObj.innerHTML = """ & server.htmlencode(host_account) & "&nbsp;(" & server.htmlencode(host_name) & ")"";" & Chr(13)
else
	Response.Write "theObj.innerHTML = """ & server.htmlencode(host_account) & """;" & Chr(13)
end if
%>
}

function sendmsg()
{
	location.href = "cal_sendmsg.asp?<%=getGRSN() %>&calid=<%=calid %>&msgname=<%

if Len(host_name) > 0 then
	UserWeb_showname = host_name
end if

Response.Write Server.URLEncode(UserWeb_showname)

%>&preturl=<%=Server.URLEncode("cal_showinvite.asp?" & getGRSN() & "&fmcal=" & fmcal & "&calid=" & calid) %>&ppreturl=<%=Server.URLEncode(returl) %>";
}

function moreguest()
{
	location.href = "cal_moreguest.asp?<%=getGRSN() %>&calid=<%=calid %>&msgname=<%

if Len(host_name) > 0 then
	UserWeb_showname = host_name
end if

Response.Write Server.URLEncode(UserWeb_showname)

%>&preturl=<%=Server.URLEncode("cal_showinvite.asp?" & getGRSN() & "&fmcal=" & fmcal & "&calid=" & calid) %>&ppreturl=<%=Server.URLEncode(returl) %>";
}

function showguest()
{
	location.href = "cal_showguest.asp?<%=getGRSN() %>&calid=<%=calid %>&msgname=<%

if Len(host_name) > 0 then
	UserWeb_showname = host_name
end if

Response.Write Server.URLEncode(UserWeb_showname)

%>&preturl=<%=Server.URLEncode("cal_showinvite.asp?" & getGRSN() & "&fmcal=" & fmcal & "&calid=" & calid) %>&ppreturl=<%=Server.URLEncode(returl) %>";
}

function newaddress()
{
	location.href = "cal_newaddress.asp?<%=getGRSN() %>&calid=<%=calid %>&preturl=<%=Server.URLEncode("cal_showinvite.asp?" & getGRSN() & "&fmcal=" & fmcal & "&calid=" & calid) %>&ppreturl=<%=Server.URLEncode(returl) %>";
}

function delnt()
{
<%
if Len(Session("svcal")) < 1 then
%>
	if (confirm("确实要删除吗?") == false)
		return ;

	location.href = "cal_del.asp?<%=getGRSN() %>&calmode=10&calid=<%=calid %>&returl=<%=Server.URLEncode(returl) %>";
<%
else
%>
	alert("您的权限不足.");
<%
end if
%>
}
//-->
</script>
</html>

<%
b_year = NULL
b_month = NULL
b_day = NULL
b_hour = NULL
b_minute = NULL


set ecalset = nothing
set ecalext = nothing
set ecal = nothing
set ecalnt = nothing
%>
