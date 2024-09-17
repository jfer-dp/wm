<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
isMobile = false
dim http_user_agent
http_user_agent = LCase(Request.ServerVariables("HTTP_User-Agent"))
if InStr(http_user_agent, "applewebkit") > 0 or InStr(http_user_agent, "mobile") > 0 then
	if InStr(http_user_agent, "iphone") > 0 or InStr(http_user_agent, "ipod") > 0 or InStr(http_user_agent, "android") > 0 or InStr(http_user_agent, "ios") > 0 or InStr(http_user_agent, "ipad") > 0 then
		isMobile = true
	end if
end if

dim ecal
set ecal = server.createobject("easymail.Calendar")

returl = trim(request("returl"))
editcal = trim(request("editcal"))
calid = trim(request("calid"))
is_edit = false

purl = trim(request("purl"))

isspd = trim(request("isspd"))

dim host_account
host_account = ""

dim ext_invitation_emails
yes_User = 0
no_User = 0
wait_User = 0

if isspd = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if Len(Session("svcal")) > 0 then
		set ecal = nothing
		Response.Redirect "noadmin.asp"
	end if

	sp_title = trim(request("sp_title"))
	sp_start_year = trim(request("sp_start_year"))
	sp_start_month = trim(request("sp_start_month"))
	sp_start_day = trim(request("sp_start_day"))
	sp_start_hour = trim(request("sp_start_hour"))
	sp_start_minute = trim(request("sp_start_minute"))

	isok = false

	if Len(sp_title) > 0 then
		ecal.Load Session("wem")
		ecal.bi_name = sp_title
		ecal.bi_mode = 1
		ecal.bi_place = sp_title
		ecal.bi_note = sp_title

		ecal.set_bi_start_date CLng(sp_start_year), CLng(sp_start_month), CLng(sp_start_day), CLng(sp_start_hour), CLng(sp_start_minute)

		isok = true
		if ecal.CreateNew("") = -1 then
			isok = false
		end if
	end if

	set ecal = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl & "&returl=" & purl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl & "&returl=" & purl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	end if
end if

if Len(calid) > 10 and editcal = "1" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	is_edit = true

	if Len(Session("svcal")) < 1 then
		ecal.Load Session("wem")
	else
		ecal.Load Session("svcal")
	end if

	isok = false
	if ecal.MoveToID(calid) = true then
		if Len(Session("svcal")) < 1 or ecal.bi_shareMode = 2 then
			host_account = ecal.bi_host_account
			isok = true
		end if
	end if

	if isok = false then
		set ecal = nothing

		if Len(purl) < 1 then
			purl = returl
		end if

		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(purl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	end if


	set ecalext = server.createobject("easymail.CalendarExtend")

	if Len(host_account) < 1 then
		if Len(Session("svcal")) < 1 then
			LightLoad_isok = ecalext.LightLoad(Session("wem"), ecal.bi_id)
		else
			LightLoad_isok = ecalext.LightLoad(Session("svcal"), ecal.bi_id)
		end if
	else
		LightLoad_isok = ecalext.LightLoad(host_account, ecal.bi_id)
	end if

	if LightLoad_isok = true then
		yes_User = ecalext.Yes_User
		no_User = ecalext.No_User
		wait_User = ecalext.Wait_User
	end if

	set ecalext = nothing
end if

bi_start_date = trim(request("bi_start_date"))
if Len(bi_start_date) = 10 and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if Len(Session("svcal")) > 0 then
		set ecal = nothing
		Response.Redirect "noadmin.asp"
	end if

	ecal.Load Session("wem")

	if Len(calid) > 10 and editcal = "2" then
		ecal.MoveToID calid
	end if

	ecal.bi_name = trim(request("bi_name"))
	ecal.bi_mode = CLng(trim(request("bi_mode")))

	bi_start_date_year = Clng(Mid(bi_start_date, 1, 4))
	bi_start_date_month = Clng(Mid(bi_start_date, 6, 2))
	bi_start_date_day = Clng(Mid(bi_start_date, 9, 2))
	bi_start_date_hour = 0
	bi_start_date_minute = 0

	bi_needtime = CLng(trim(request("need_hour"))) * 3600 + CLng(trim(request("need_minute"))) * 60

	if trim(request("need_allday")) = 0 then
		bi_start_date_hour = Clng(trim(request("start_hour")))
		bi_start_date_minute = Clng(trim(request("start_minute")))
	else
		bi_needtime = 0
	end if

	ecal.set_bi_start_date bi_start_date_year, bi_start_date_month, bi_start_date_day, bi_start_date_hour, bi_start_date_minute

	ecal.bi_needtime = bi_needtime
	ecal.bi_place = trim(request("bi_place"))
	ecal.bi_note = trim(request("bi_note"))

	ecal.bi_shareMode = CLng(trim(request("bi_shareMode")))

	if trim(request("bi_isRepeat")) = "1" then
		ecal.bi_isRepeat = true

		ecal.bi_rp_jiange = CLng(trim(request("bi_rp_jiange")))
		ecal.bi_rp_jump_dwmy = CLng(trim(request("bi_rp_jump_dwmy")))
		ecal.bi_rp_done_dwm = CLng(trim(request("bi_rp_done_dwm")))

		i = 0
		date_in_str = ""
		if ecal.bi_rp_done_dwm = 1 or (ecal.bi_rp_done_dwm > 2 and ecal.bi_rp_done_dwm < 7) then
			do while i < 7
				if trim(request("week_check" & i)) <> "" then
					date_in_str = date_in_str & "1"
				else
					date_in_str = date_in_str & "0"
				end if

			    i = i + 1
			loop

			ecal.bi_rp_week_str = date_in_str
		elseif ecal.bi_rp_done_dwm = 2 then
			do while i < 31
				if trim(request("month_check" & i)) <> "" then
					date_in_str = date_in_str & "1"
				else
					date_in_str = date_in_str & "0"
				end if

			    i = i + 1
			loop

			ecal.bi_rp_month_str = date_in_str
		end if
	else
		ecal.bi_isRepeat = false
	end if

	if trim(request("bi_rp_have_end")) = "" or trim(request("bi_rp_have_end")) = "0" then
		ecal.bi_rp_have_end = false
	else
		ecal.bi_rp_have_end = true
		ecal.set_bi_rp_end_date CLng(trim(request("bi_rp_end_date_year"))), CLng(trim(request("bi_rp_end_date_month"))), CLng(trim(request("bi_rp_end_date_day")))
	end if

	if trim(request("bi_has_invitation")) = "1" then
		ecal.bi_has_invitation = true

		ecal.bi_notice_name = Session("wem")
		ecal.bi_notice_email = Session("mail")
	else
		ecal.bi_has_invitation = false
	end if

	if trim(request("bi_remind")) = "1" then
		ecal.bi_remind = true
		ecal.bi_remind_sec = CLng(trim(request("bi_remind_sec")))
	else
		ecal.bi_remind = false
	end if

	ecal.bi_address = trim(request("bi_address"))
	ecal.bi_city = trim(request("bi_city"))
	ecal.bi_phone = trim(request("bi_phone"))
	ecal.bi_other_phone = trim(request("bi_other_phone"))

	invitation_emails = trim(request("invitation_emails"))
	if Len(invitation_emails) > 2 then
		ecal.bi_has_invitation = true

		ecal.bi_notice_name = Session("wem")
		ecal.bi_notice_email = Session("mail")
	end if

	isok = true

	if editcal <> "2" then
		if ecal.CreateNew(invitation_emails) = -1 then
			isok = false
		end if
	else
		isok = ecal.Set(calid)

		if isok = true then
			isok = ecal.Save()

			if isok = true then
				if Len(invitation_emails) > 2 then
					if ecal.AddEmails(calid, invitation_emails) = -1 then
						isok = false
					end if
				end if
			end if
		end if
	end if

	set ecal = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl & "&returl=" & purl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_new.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl & "&returl=" & purl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_new.asp")
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
html {overflow: scroll; overflow-x: hidden; overflow-y: auto !important;}
body {font-family:<%=s_lang_font %>; font-size:9pt;color:#000000;margin-top:5px;margin-left:10px;margin-right:10px;margin-bottom:2px;background-color:#ffffff}
.tl {height:24px; text-align:right; border-bottom:1px #8CA5B5 solid;}
-->
</STYLE>
</head>

<SCRIPT language=javascript src="images/cal/popcalendar.js"></SCRIPT>
<SCRIPT Language="JavaScript">dateFormat='yyyy-mm-dd'</SCRIPT>

<script type="text/javascript">
<!--
var repeat_is_show = <%
if ecal.bi_isRepeat = false then
	Response.Write "false"
else
	Response.Write "true"
end if
%>;

var repeat_done_is_show = <%
if ecal.bi_rp_done_dwm = 2 then
	Response.Write "true"
else
	Response.Write "false"
end if
%>;

var repeat_done_dwm = <%=ecal.bi_rp_done_dwm %>;

var invitation_is_show = <%
has_invitation = false

if ecal.bi_has_invitation = false then
	Response.Write "false"
else
	if yes_User > 0 or no_User > 0 or wait_User > 0 then
		has_invitation = true
		Response.Write "true"
	else
		Response.Write "false"
	end if
end if
%>;

var remind_is_show = <%
if ecal.bi_remind = false then
	Response.Write "false"
else
	Response.Write "true"
end if
%>;

function window_onload() {
	init();
	document.f1.bi_mode.value = "1";

<%
if is_edit = false then
%>
	document.f1.week_check1.checked = true;
	document.f1.month_check0.checked = true;
<%
bs_hour = trim(request("bsh"))
if Len(bs_hour) > 0 then
	Response.Write "document.f1.need_allday_false.checked = true;" & Chr(13)
	Response.Write "document.f1.start_hour.value = " & bs_hour & ";"
else
	Response.Write "document.f1.start_hour.value = 9;"
end if
%>
	document.f1.need_hour.value = 1;
<%
else
	Response.Write "document.f1.bi_mode.value = " & ecal.bi_mode & ";" & Chr(13)

	ecal.get_bi_start_date bi_start_date_year, bi_start_date_month, bi_start_date_day, bi_start_date_hour, bi_start_date_minute
	Response.Write "document.f1.start_hour.value = " & bi_start_date_hour & ";" & Chr(13)
	Response.Write "document.f1.start_minute.value = " & bi_start_date_minute & ";" & Chr(13)

	if ecal.bi_needtime > 0 then
		if ecal.bi_needtime < 3600 then
			Response.Write "document.f1.need_hour.value = 0;" & Chr(13)
		else
			Response.Write "document.f1.need_hour.value = " & CLng(ecal.bi_needtime / 3600) & ";" & Chr(13)
		end if

		Response.Write "document.f1.need_minute.value = " & CLng(CLng(ecal.bi_needtime Mod 3600) / 60) & ";" & Chr(13)
	else
		Response.Write "document.f1.start_hour.value = 9;" & Chr(13)
		Response.Write "document.f1.need_hour.value = 1;" & Chr(13)
	end if

	Response.Write "document.f1.bi_rp_jiange.value = " & ecal.bi_rp_jiange & ";" & Chr(13)
	Response.Write "document.f1.bi_rp_jump_dwmy.value = " & ecal.bi_rp_jump_dwmy & ";" & Chr(13)


	if ecal.bi_rp_done_dwm = 1 or (ecal.bi_rp_done_dwm > 2 and ecal.bi_rp_done_dwm < 7) then
		i = 0
		date_in_str = ecal.bi_rp_week_str
		do while i < 7
			if Mid(date_in_str, i + 1, 1) = "1" then
				Response.Write "document.f1.week_check" & i & ".checked = true;" & Chr(13)
			else
				Response.Write "document.f1.week_check" & i & ".checked = false;" & Chr(13)
			end if

		    i = i + 1
		loop
	end if

	if ecal.bi_rp_done_dwm = 2 then
		i = 0
		date_in_str = ecal.bi_rp_month_str
		do while i < 31
			if Mid(date_in_str, i + 1, 1) = "1" then
				Response.Write "document.f1.month_check" & i & ".checked = true;" & Chr(13)
			else
				Response.Write "document.f1.month_check" & i & ".checked = false;" & Chr(13)
			end if

		    i = i + 1
		loop
	end if

	ecal.get_bi_rp_end_date bi_start_date_year, bi_start_date_month, bi_start_date_day
	if bi_start_date_year > 1971 and bi_start_date_month > 0 and bi_start_date_month < 13 and bi_start_date_day > 0 and bi_start_date_day < 32 then
		Response.Write "document.f1.bi_rp_end_date_year.value = " & bi_start_date_year & ";" & Chr(13)
		Response.Write "document.f1.bi_rp_end_date_month.value = " & bi_start_date_month & ";" & Chr(13)
		Response.Write "document.f1.bi_rp_end_date_day.value = " & bi_start_date_day & ";" & Chr(13)
	end if

	Response.Write "document.f1.bi_remind_sec.value = " & ecal.bi_remind_sec & ";" & Chr(13)


	bi_start_date_year = NULL
	bi_start_date_month = NULL
	bi_start_date_day = NULL
	bi_start_date_hour = NULL
	bi_start_date_minute = NULL
end if
%>

	bi_isRepeat_onclick();
	hide_repeat();
	hide_repeat_done();

<%
if is_edit = true then
	Response.Write "repeat_done_dwm = " & ecal.bi_rp_done_dwm & ";" & Chr(13)

	if ecal.bi_rp_done_dwm > 0 and ecal.bi_rp_done_dwm < 7 then
		Response.Write "document.f1.bi_rp_done_dwm.value = " & ecal.bi_rp_done_dwm & ";" & Chr(13)
	end if
end if
%>
	show_done_dwm();

	hide_invitation();
	hide_remind();
	bi_remind_onclick();
}

var Stag;
function hide_repeat()
{
	Stag = document.getElementById("repeat_bt");
	if (repeat_is_show == true)
	{
		Stag.innerHTML = "隐藏";
		document.getElementById("repeat_div").style.display = "inline";
	}
	else
	{
		Stag.innerHTML = "展示";
		document.getElementById("repeat_div").style.display = "none";
	}

	repeat_is_show = !repeat_is_show;
}

function hide_invitation()
{
	Stag = document.getElementById("invitation_bt");
	if (invitation_is_show == true)
	{
		Stag.innerHTML = "隐藏";
		document.getElementById("invitation_div").style.display = "inline";
	}
	else
	{
		Stag.innerHTML = "展示";
		document.getElementById("invitation_div").style.display = "none";
	}

	invitation_is_show = !invitation_is_show;
}

function hide_remind()
{
	Stag = document.getElementById("remind_bt");
	if (remind_is_show == true)
	{
		Stag.innerHTML = "隐藏";
		document.getElementById("remind_div").style.display = "inline";
	}
	else
	{
		Stag.innerHTML = "展示";
		document.getElementById("remind_div").style.display = "none";
	}

	remind_is_show = !remind_is_show;
}

function hide_repeat_done()
{
	if (repeat_done_is_show == true)
		document.getElementById("repeat_done_div").style.display = "inline";
	else
		document.getElementById("repeat_done_div").style.display = "none";
}

function select_jump_dwmy_onchange()
{
	if (document.f1.bi_rp_jump_dwmy.selectedIndex == 2)
	{
		repeat_done_dwm = 2;
		document.f1.bi_rp_done_dwm.value = "2";
		repeat_done_is_show = true;
	}
	else if (document.f1.bi_rp_jump_dwmy.selectedIndex == 1)
	{
		repeat_done_dwm = 1;
		document.f1.bi_rp_done_dwm.value = "1";
		repeat_done_is_show = false;
	}
	else
	{
		repeat_done_dwm = 0;
		repeat_done_is_show = false;
	}

	hide_repeat_done();
	show_done_dwm();
}

function enable_week(is_enable)
{
	var theObj;
	if (is_enable == true)
	{
		for (i = 0; i < 7; i++)
		{
			theObj = eval("document.f1.week_check" + i);
			theObj.disabled = false;
		}
	}
	else
	{
		for (i = 0; i < 7; i++)
		{
			theObj = eval("document.f1.week_check" + i);
			theObj.disabled = true;
		}
	}
}

function enable_month(is_enable)
{
	var theObj;
	if (is_enable == true)
	{
		for (i = 0; i < 31; i++)
		{
			theObj = eval("document.f1.month_check" + i);
			theObj.disabled = false;
		}
	}
	else
	{
		for (i = 0; i < 31; i++)
		{
			theObj = eval("document.f1.month_check" + i);
			theObj.disabled = true;
		}
	}
}

function select_done_dwm_onchange()
{
	repeat_done_dwm = document.f1.bi_rp_done_dwm.value;
	show_done_dwm();
}

function show_done_dwm()
{
	if (repeat_done_dwm == 1 || (repeat_done_dwm > 2 && repeat_done_dwm < 7))
	{
		enable_week(true);
		enable_month(false)
	}
	else if (repeat_done_dwm == 2)
	{
		enable_week(false);
		enable_month(true)
	}
	else
	{
		enable_week(false);
		enable_month(false)
	}
}

function bi_isRepeat_onclick()
{
	var rd_rep_value = 0;
	Stag = document.getElementsByName("bi_isRepeat");
	if (Stag != null)
	{
		for (i = 0; i < Stag.length; i++)
		{
			if (Stag[i].checked == true)
			{
				rd_rep_value = Stag[i].value;
				break;
			}
		}
	}

	if (rd_rep_value == 0)
	{
		document.f1.bi_rp_jiange.disabled = true;
		document.f1.bi_rp_jump_dwmy.disabled = true;

		select_jump_dwmy_onchange()

		enable_week(false);
		enable_month(false);

		repeat_done_is_show = false;
		hide_repeat_done();

		Stag = document.getElementsByName("bi_rp_have_end");
		if (Stag != null)
		{
			for (i = 0; i < Stag.length; i++)
				Stag[i].disabled = true;
		}

		document.f1.bi_rp_end_date_year.disabled = true;
		document.f1.bi_rp_end_date_month.disabled = true;
		document.f1.bi_rp_end_date_day.disabled = true;
	}
	else if (rd_rep_value == 1)
	{
		document.f1.bi_rp_jiange.disabled = false;
		document.f1.bi_rp_jump_dwmy.disabled = false;

		select_jump_dwmy_onchange()

		Stag = document.getElementsByName("bi_rp_have_end");
		if (Stag != null)
		{
			for (i = 0; i < Stag.length; i++)
				Stag[i].disabled = false;
		}

		bi_rp_have_end_onclick();
	}
}

function bi_rp_have_end_onclick()
{
	var rd_rep_value = 0;
	Stag = document.getElementsByName("bi_rp_have_end");
	if (Stag != null)
	{
		for (i = 0; i < Stag.length; i++)
		{
			if (Stag[i].checked == true)
			{
				rd_rep_value = Stag[i].value;
				break;
			}
		}
	}

	if (rd_rep_value == 0)
	{
		document.f1.bi_rp_end_date_year.disabled = true;
		document.f1.bi_rp_end_date_month.disabled = true;
		document.f1.bi_rp_end_date_day.disabled = true;
	}
	else
	{
		document.f1.bi_rp_end_date_year.disabled = false;
		document.f1.bi_rp_end_date_month.disabled = false;
		document.f1.bi_rp_end_date_day.disabled = false;
	}
}

function bi_remind_onclick()
{
	var rd_rep_value = 0;
	Stag = document.getElementsByName("bi_remind");
	if (Stag != null)
	{
		for (i = 0; i < Stag.length; i++)
		{
			if (Stag[i].checked == true)
			{
				rd_rep_value = Stag[i].value;
				break;
			}
		}
	}

	if (rd_rep_value == 0)
		document.f1.bi_remind_sec.disabled = true;
	else
		document.f1.bi_remind_sec.disabled = false;
}

function gosub()
{
	if (document.f1.bi_name.value.length < 1)
	{
		alert("请输入“名称”项");
		document.f1.bi_name.focus();
		return ;
	}

	if (document.f1.bi_start_date.value.length != 10)
	{
		alert("请输入“日期”项");
		document.f1.bi_start_date.focus();
		return ;
	}

	if (document.f1.bi_place.value.length < 1)
	{
		alert("请输入“联系地址”项");
		document.f1.bi_place.focus();
		return ;
	}

	if (document.f1.bi_note.value.length < 1)
	{
		alert("请输入“便笺”项");
		document.f1.bi_note.focus();
		return ;
	}

	document.f1.submit();
}

function goback()
{
	if (document.f1.returl.value.length < 3)
		history.back();
	else
	{
		if (document.f1.purl.value.length > 0)
			location.href=document.f1.returl.value + "&returl=" + document.f1.purl.value;
		else
			location.href=document.f1.returl.value;
	}
}
<%
if is_edit = true then
%>
function godel()
{
	if (confirm("确实要删除吗?") == false)
		return ;

	location.href = "cal_del.asp?<%=getGRSN() %>&calmode=1&calid=<%=calid %>&returl=<%
if Len(purl) > 0 then
	Response.Write Server.URLEncode(purl)
else
	Response.Write Server.URLEncode(returl)
end if
%>";
}
<%
end if
%>

function viewInv(evid)
{
	location.href = "cal_showinvite.asp?<%=getGRSN() %>&fmcal=1&calid=" + evid + "&returl=<%
if Len(purl) > 0 then
	Response.Write Server.URLEncode(purl)
else
	Response.Write Server.URLEncode(returl)
end if
%>";
}

function popaddress()
{
	var remote = null;
	remote = window.open("selectadd.asp?mode=To&ofm=<%=Server.URLEncode("document.f1.invitation_emails") %>&<%=getGRSN() %>", "", "top=80; left=150; height=345,width=496,scrollbars=yes,resizable=yes,status=no,toolbar=no,menubar=no,location=no");
}

<%
if Application("em_EnableEntAddress") = true then
%>
function eapop() {
	window.open("ea_pop.asp?mode=To&ofm=<%=Server.URLEncode("document.f1.invitation_emails") %>&<%=getGRSN() %>", "", "top=80; left=130; height=330,width=510,scrollbars=yes,resizable=yes,status=no,toolbar=no,menubar=no,location=no");
}
<%
end if
%>
//-->
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
<br>
<form method="post" action="cal_new.asp" name="f1">
<input type="hidden" name="bi_has_invitation" value="<%
if ecal.bi_has_invitation = true then
	Response.Write "1"
else
	Response.Write "0"
end if
%>">
<input type="hidden" name="returl" value="<%=returl %>">
<input type="hidden" name="editcal" value="<%
if editcal = "1" then
	Response.Write "2"
end if
%>">
<input type="hidden" name="calid" value="<%=calid %>">
<input type="hidden" name="purl" value="<%=Server.URLEncode(purl) %>">

<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td colspan="2" style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%
if is_edit = false then
	Response.Write "添加活动"
else
	Response.Write "查看活动"
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
		名称<%=s_lang_mh %>
		</td>
		<td align=left style='border-bottom:1px #8CA5B5 solid;'>
		<input type="text" name="bi_name" class='n_textbox' value="<%=ecal.bi_name %>" size="50" maxlength="40">
		</td>
		</tr>

		<tr>
		<td valign=center class="tl"> 
		事件类型<%=s_lang_mh %>
		</td>
		<td align=left style='border-bottom:1px #8CA5B5 solid;'>
		<select name="bi_mode" class="drpdwn">
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

		<tr>
		<td valign=center class="tl"> 
		日期<%=s_lang_mh %>
		</td>
		<td align=left style='border-bottom:1px #8CA5B5 solid;'>
		<input type="text" name="bi_start_date" class='n_textbox' value="<%
ecal.get_bi_start_date b_year, b_month, b_day, b_hour, b_minute
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
else
	bs_year = trim(request("bsy"))
	bs_month = trim(request("bsm"))
	bs_day = trim(request("bsd"))

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
end if
%>" readonly size="20" maxlength="16">
<script language='javascript'> 
<!--
if (!document.layers) {
	document.write("<img align=absmiddle style='CURSOR:pointer' src='images/cal/calendar.gif' onclick='popUpCalendar(this, document.f1.bi_start_date, dateFormat,-1,-1)' alt='Select'>");
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
		</td>
		</tr>

		<tr>
		<td valign=center class="tl"> 
		时间<%=s_lang_mh %>
		</td>
		<td align=left style='border-bottom:1px #8CA5B5 solid;'>
		<input type=radio <% if ecal.bi_needtime = 0 then response.write "checked"%> value="1" name="need_allday" id="need_allday_true">这是一个<b>全天</b>活动。
		<br>
		<input type=radio <% if ecal.bi_needtime > 0 then response.write "checked"%> value="0" name="need_allday" id="need_allday_false">开始于:
		<select name="start_hour" class="drpdwn">
<%
i = 0

do while i < 24
	if show_APM = true then
		if i = 0 then
			Response.Write "<option value=""" & i & """>" & 12 & " am</option>" & Chr(13)
		elseif i = 12 then
			Response.Write "<option value=""" & i & """>" & 12 & " pm</option>" & Chr(13)
		elseif i < 12 then
			Response.Write "<option value=""" & i & """>" & i & " am</option>" & Chr(13)
		else
			Response.Write "<option value=""" & i & """>" & i - 12 & " pm</option>" & Chr(13)
		end if
	else
		Response.Write "<option value=""" & i & """>" & i & "</option>" & Chr(13)
	end if

	i = i + 1
loop
%>
		</select>
		<select name="start_minute" class="drpdwn">
		<option value="0">:00</option>
		<option value="15">:15</option>
		<option value="30">:30</option>
		<option value="45">:45</option>
		</select>&nbsp;&nbsp;&nbsp;
		需要时间<%=s_lang_mh %><select name="need_hour" class="drpdwn">
<%
i = 0

do while i < 13
	Response.Write "<option value=""" & i & """>" & i & " 小时</option>" & Chr(13)

	i = i + 1
loop
%>
		</select>
		<select name="need_minute" class="drpdwn">
		<option value="0">0 分钟</option>
		<option value="15">15 分钟</option>
		<option value="30">30 分钟</option>
		<option value="45">45 分钟</option>
		</select>
		</td>
		</tr>

		<tr>
		<td valign=center class="tl"> 
		联系地址<%=s_lang_mh %>
		</td>
		<td align=left style='border-bottom:1px #8CA5B5 solid;'>
		<input type="text" name="bi_place" class='n_textbox' value="<%=ecal.bi_place %>" size="50" maxlength="50">
		</td>
		</tr>

		<tr>
		<td valign=center class="tl"> 
		便笺<%=s_lang_mh %>
		</td>
		<td align=left style='border-bottom:1px #8CA5B5 solid;'>
		<textarea name="bi_note" cols="70" rows="3" class='n_textarea'><%=ecal.bi_note %></textarea><br>
		至多 500 个字符
		</td>
		</tr>

		<tr>
		<td valign=center class="tl"> 
		分享<%=s_lang_mh %>
		</td>
		<td align=left style='border-bottom:1px #8CA5B5 solid;'>
<%
if is_edit = true then
%>
		<input type=radio <% if ecal.bi_shareMode = 0 then response.write "checked"%> value="0" name="bi_shareMode">私人的&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<a href="help.asp#calshare">什么是共享?</a>]
		<br>
		<input type=radio <% if ecal.bi_shareMode = 1 then response.write "checked"%> value="1" name="bi_shareMode">显示为忙碌中
		<br>
		<input type=radio <% if ecal.bi_shareMode = 2 then response.write "checked"%> value="2" name="bi_shareMode">公开的
<%
else
	tmp_bi_shareMode = ecalset.EventShareDefault
%>
		<input type=radio <% if tmp_bi_shareMode = 0 then response.write "checked"%> value="0" name="bi_shareMode">私人的&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<a href="help.asp#calshare">什么是共享?</a>]
		<br>
		<input type=radio <% if tmp_bi_shareMode = 1 then response.write "checked"%> value="1" name="bi_shareMode">显示为忙碌中
		<br>
		<input type=radio <% if tmp_bi_shareMode = 2 then response.write "checked"%> value="2" name="bi_shareMode">公开的
<%
end if
%>
		</td>
		</tr>

		<tr>
		<td colspan=2 height="24" valign=center align=left bgcolor="#DBEAF5" style='border-bottom:1px #8CA5B5 solid;'> 
		&nbsp;<b>重复</b>&nbsp;[<a href="javascript:hide_repeat()"><span id="repeat_bt"></span></a>]
		<div id="repeat_div">
		<table width="100%" border=0 cellspacing=0 cellpadding=0>
			<tr>
			<td valign=center width="15%" height="24" align=right> 
			&nbsp;
			</td>
			<td align=left style='border-bottom:1px #8CA5B5 solid;'>
		<input type=radio <% if ecal.bi_isRepeat = false then response.write "checked"%> value="0" name="bi_isRepeat" LANGUAGE=javascript onclick="return bi_isRepeat_onclick()">该活动不重复出现。
		<br>
		<input type=radio <% if ecal.bi_isRepeat = true then response.write "checked"%> value="1" name="bi_isRepeat" LANGUAGE=javascript onclick="return bi_isRepeat_onclick()">重复&nbsp;&nbsp;
		<select name="bi_rp_jiange" class="drpdwn">
<%
i = 0

do while i < 100
	Response.Write "<option value=""" & i & """>每 " & i+1 & "</option>"

	i = i + 1
loop
%>
		</select><select name="bi_rp_jump_dwmy" class="drpdwn" LANGUAGE=javascript onchange="select_jump_dwmy_onchange()">
		<option value="0">日</option>
		<option value="1">周</option>
		<option value="2">月</option>
		<option value="3">年</option>
		</select>
		<div id="repeat_done_div">
		&nbsp;&nbsp;&nbsp;执行对象<%=s_lang_mh %></select><select name="bi_rp_done_dwm" class="drpdwn" LANGUAGE=javascript onchange="select_done_dwm_onchange()">
		<option value="2">月</option>
		<option value="1">第1周</option>
		<option value="3">第2周</option>
		<option value="4">第3周</option>
		<option value="5">第4周</option>
		<option value="6">最后一周</option>
		</select>
		</div>
			</td>
			</tr>

			<tr>
			<td>
			</td>
			<td style='border-bottom:1px #8CA5B5 solid;'>
				<table width="100%" border=0 cellspacing=0 cellpadding=0>
				<tr>
				<td valign=center width="3%" height="24" align=right> 
				&nbsp;
				</td>
				<td align=left>
<%
i = 1

do while i < 7
	Response.Write "<input type='checkbox' name='week_check" & i & "' value=""" & i & """>" & server.htmlencode(getWeekName2(i)) & "&nbsp;&nbsp;"

	i = i + 1
loop

Response.Write "<input type='checkbox' name='week_check0' value='0'>" & server.htmlencode(getWeekName2(0)) & Chr(13)
%>
				</td>
				</tr>
				</table>
			</td>
			</tr>
			<tr>
			<td>
			</td>
			<td>
				<table width="100%" border=0 cellspacing=0 cellpadding=0 style='border-bottom:1px #8CA5B5 solid;'>
				<tr>
				<td valign=center width="3%" height="24" align=right> 
				&nbsp;
				</td>
				<td align=left>
<%
i = 0

do while i < 31
	Response.Write "<input type='checkbox' name='month_check" & i & "' value=""" & i & """>" & i+1 & "&nbsp;&nbsp;"

	if i = 10 or i = 20 then
		Response.Write "<br>" & Chr(13)
	end if

	i = i + 1
loop
%>
				</td>
				</tr>
				</table>
			</td>
			</tr>
			<tr>
			<td>&nbsp;
			</td>
			</tr>
			<tr><td></td>
			<td>
		<b>结束日期:</b>
			</td>
			</tr>
			<tr><td></td>
			<td>
		<input type=radio <% if ecal.bi_rp_have_end = false then response.write "checked"%> value=0 name="bi_rp_have_end" LANGUAGE=javascript onclick="return bi_rp_have_end_onclick()">没有结束日期。
		<br>
		<input type=radio <% if ecal.bi_rp_have_end = true then response.write "checked"%> value=1 name="bi_rp_have_end" LANGUAGE=javascript onclick="return bi_rp_have_end_onclick()">直到
		<select name="bi_rp_end_date_year" class="drpdwn">
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
		</select><select name="bi_rp_end_date_month" class="drpdwn">
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
		</select><select name="bi_rp_end_date_day" class="drpdwn">
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
			</table>
			</div>
			</td>
			</tr>
			<tr>
			<td colspan=2 height="24" valign=center align=left bgcolor="#DBEAF5" style='border-bottom:1px #8CA5B5 solid;'> 
		&nbsp;<b>请柬</b>&nbsp;[<a href="javascript:hide_invitation()"><span id="invitation_bt"></span></a>]
			<div id="invitation_div">
			<table width="100%" border=0 cellspacing=0 cellpadding=0>
			<tr>
				<td valign=center width="15%" height="24" align=right>
		&nbsp;
				</td>
<%
if ecal.bi_has_invitation = true and has_invitation = true then
%>
<td width="100" align=left valign="top">
<table width="100%" border=0 cellspacing=0 cellpadding=0>
<tr><td>
<b>客人概况<b>
</td></tr>
<tr>
<td nowrap>
<br>
<table width="100%" border=0 cellspacing=0 cellpadding=0>
<tr><td>
<img src='images/cal/a.gif' border=0>&nbsp;</td><td nowrap><%=yes_User %>&nbsp;</td><td nowrap>参加
</td></tr>
<tr><td>
<img src='images/cal/u.gif' border=0>&nbsp;</td><td nowrap><%=wait_User %>&nbsp;</td><td nowrap>未决定的
</td></tr>
<tr><td>
<img src='images/cal/d.gif' border=0>&nbsp;</td><td nowrap><%=no_User %>&nbsp;</td><td nowrap>婉言拒绝&nbsp;
</td></tr>
</table>
<br>
[<a href="javascript:viewInv('<%=ecal.bi_id %>')">查看请柬</a>]
</td>
</td>
</table>
</td>
<%
end if
%>
				<td align=left>
<%
if ecal.bi_has_invitation = false then
%>
		请输入您希望发给电子请柬的人的本系统电子邮件地址（给多人发送电子邮件请柬，每个人之间请用逗号隔开）。
<%
else
%>
		<b>邀请其他客人</b><br>
输入您想要发送电子请帖的电子邮件地址，用逗号间隔。  
<%
end if
%>
		<br><br>
<%
if isMobile = false then
%>
		[<a href="javascript:popaddress()">从地址簿选择收件人</a>]
<%
	if Application("em_EnableEntAddress") = true then
%>
&nbsp;[<a href="javascript:eapop()">从企业地址簿选择收件人</a>]
<%
	end if
%>
		<br><br>
<%
end if
%>
		<textarea name="invitation_emails" cols="70" rows="4" class='n_textarea'></textarea>
		<br>&nbsp;
				</td>
			</tr>
			</table>
			</div>
			</td>
			</tr>
			<tr>
			<td colspan=2 height="24" valign=center align=left bgcolor="#DBEAF5" style='border-bottom:1px #8CA5B5 solid;'> 
		&nbsp;<b>提醒功能</b>&nbsp;[<a href="javascript:hide_remind()"><span id="remind_bt"></span></a>]
			<div id="remind_div">
			<table width="100%" border=0 cellspacing=0 cellpadding=0>
			<tr>
				<td valign=center width="15%" height="24" align=right> 
		&nbsp;
				</td>
				<td align=left>
		<input type=radio <% if ecal.bi_remind = false then response.write "checked"%> value=0 name="bi_remind" LANGUAGE=javascript onclick="return bi_remind_onclick()">不发送提醒。
		<br>
		<input type=radio <% if ecal.bi_remind = true then response.write "checked"%> value=1 name="bi_remind" LANGUAGE=javascript onclick="return bi_remind_onclick()">活动前
		<select name="bi_remind_sec" class="drpdwn">
		<option value="0" selected>----</option>
		<option value="1800">30 分</option>
		<option value="3600">1 小时</option>
		<option value="7200">2 小时</option>
		<option value="10800">3 小时</option>
		<option value="21600">6 小时</option>
		<option value="43200">12 小时</option>
		<option value="86400">1 天</option>
		<option value="172800">2 天</option>
		<option value="259200">3 天</option>
		<option value="345600">4 天</option>
		<option value="432000">5 天</option>
		<option value="518400">6 天</option>
		<option value="604800">7 天</option>
		<option value="691200">8 天</option>
		<option value="777600">9 天</option>
		<option value="864000">10 天</option>
		<option value="950400">11 天</option>
		<option value="1036800">12 天</option>
		<option value="1123200">13 天</option>
		<option value="1209600">14 天</option>
		</select>&nbsp;发送提醒。
				</td>
			</tr>
			</table>
			</div>
			</td>
			</tr>
			<tr>
			<td colspan=2 height="24" valign=center align=left bgcolor="#DBEAF5" style='border-bottom:1px #8CA5B5 solid;'> 
		&nbsp;<b>可以选择的信息</b>
			</td>
			</tr>
			<tr>
				<td valign=center width="15%" class="tl"> 
		地址<%=s_lang_mh %>
				</td>
				<td align=left style='border-bottom:1px #8CA5B5 solid;'>
		<input type="text" name="bi_address" class='n_textbox' value="<%=ecal.bi_address %>" size="50" maxlength="100">&nbsp;街道
		<br><span style="padding-top:3px; display:inline-block;">
		<input type="text" name="bi_city" class='n_textbox' value="<%=ecal.bi_city %>" size="40" maxlength="100">&nbsp;城市/州/省/邮政编码
		</span>
				</td>
			</tr>
			<tr>
				<td valign=center width="15%" class="tl"> 
		电话<%=s_lang_mh %>
				</td>
				<td align=left style='border-bottom:1px #8CA5B5 solid;'>
		<input type="text" name="bi_phone" class='n_textbox' value="<%=ecal.bi_phone %>" size="30" maxlength="100">
		<br><span style="padding-top:3px; display:inline-block;">
		<input type="text" name="bi_other_phone" class='n_textbox' value="<%=ecal.bi_other_phone %>" size="40" maxlength="100">&nbsp;其他联系方式
		</span>
				</td>
			</tr>
		</table>
	</td></tr>
	</table>
</td></tr>

<tr><td colspan="2" align="left" style="background-color:white; padding-right:16px; padding-top:10px; padding-bottom:10px;">
<a class='wwm_btnDownload btn_blue' href="javascript:goback();"><< <%=s_lang_return %></a>
<%
if Len(Session("svcal")) < 1 then
	if Len(host_account) < 1 then
%>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();">保存</a>
<%
	end if

	if is_edit = true then
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
b_hour = NULL
b_minute = NULL

set ecalset = nothing
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


function getWeekName2(wknum)
	temp_wk_str = ""

	if wknum = "0" then
		temp_wk_str = "周日"
	elseif wknum = "1" then
		temp_wk_str = "周一"
	elseif wknum = "2" then
		temp_wk_str = "周二"
	elseif wknum = "3" then
		temp_wk_str = "周三"
	elseif wknum = "4" then
		temp_wk_str = "周四"
	elseif wknum = "5" then
		temp_wk_str = "周五"
	elseif wknum = "6" then
		temp_wk_str = "周六"
	end if

	getWeekName2 = temp_wk_str
end function


function getWeekName1(wknum)
	temp_wk_str = ""

	if wknum = "0" then
		temp_wk_str = "日"
	elseif wknum = "1" then
		temp_wk_str = "一"
	elseif wknum = "2" then
		temp_wk_str = "二"
	elseif wknum = "3" then
		temp_wk_str = "三"
	elseif wknum = "4" then
		temp_wk_str = "四"
	elseif wknum = "5" then
		temp_wk_str = "五"
	elseif wknum = "6" then
		temp_wk_str = "六"
	end if

	getWeekName1 = temp_wk_str
end function
%>
