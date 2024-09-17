<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
if trim(request("set")) = "1" or Session("EnableSession") = "" then
	dim uwes
	set uwes = server.createobject("easymail.UserWeb")
	uwes.Load Session("wem")

	if Session("EnableSession") = "" then
		Session("EnableSession") = uwes.EnableSession
	else
		uwes.EnableSession = true
		Session("EnableSession") = true
		uwes.Save
	end if

	set uwes = nothing
end if

if Session("EnableSession") = false then
	Response.Redirect "listmail.asp?mode=in&" & getGRSN()
end if

sortstr = request("sortstr")
sortmode = request("sortmode")

dim ei
set ei = server.createobject("easymail.InfoList")

dim addsortstr

exstr = request("exstr")
if IsEmpty(exstr) = true then
	exstr = "0000000"
end if

if sortmode = 1 then
	addsortstr = "&sortstr=" & sortstr & "&sortmode=1" & "&exstr=" & exstr
	sortmode = true
else
	addsortstr = "&sortstr=" & sortstr & "&sortmode=0" & "&exstr=" & exstr
	sortmode = false
end if

if sortstr <> "" then
	ei.SetSort sortstr, sortmode
else
	sortstr = "Date"
end if

ei.LoadSession Session("wem")

dim exMailInfo
set exMailInfo = server.createobject("easymail.ExMailInfo")

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	set_mode = trim(Request.Form("setmode"))
	tmp_ck = trim(Request.Form("maxck"))
	i = 0
	dim is_ok
	is_ok = true
	dim ret_show_str
	ret_show_str = ""

	if IsNull(tmp_ck) = false and IsNumeric(tmp_ck) = true then
		themax = CLng(tmp_ck)
	end if

	if set_mode = "1" then
		isload_emi = false
		lbid = trim(Request.Form("lbid"))

		if Len(lbid) = 8 and themax > 0 then
			do while i <= themax
				tmp_ck = trim(Request.Form("ck_" & i))
				if tmp_ck <> "" then
					if isload_emi = false then
						exMailInfo.Load Session("wem"), tmp_ck
						isload_emi = true
						if exMailInfo.AddLabel(lbid) = false then
							is_ok = false
						end if
					else
						if exMailInfo.Add_Mail_Label(tmp_ck, lbid) = false then
							is_ok = false
						end if
					end if
				end if 

			    i = i + 1
			loop
		end if
	elseif set_mode = "2" then
		lbid = trim(Request.Form("lbid"))
		msid = trim(Request.Form("msid"))
		mailid = trim(Request.Form("mailid"))

		if Len(lbid) = 8 and Len(mailid) > 10 then
			exMailInfo.Load Session("wem"), mailid
			if exMailInfo.DelLabel(lbid) = false then
				is_ok = false
			else
				ei.Get_Session_IsRead_IsStar_Labels Session("wem"), msid, ret_IsRead, ret_IsStar

				if ei.LabelCount > 0 then
					lball = ei.LabelCount
					lbi = 0
					do while lbi < lball
						if ei.GetLabel(lbi) = lbid then
							ret_show_str = "2"
							Exit Do
						end if

						lbi = lbi + 1
					loop
				end if

				ret_IsRead = NULL
				ret_IsStar = NULL
			end if
		end if
	elseif set_mode = "3" then
		bj = trim(Request.Form("bj"))
		if themax > 0 then
			do while i <= themax
				tmp_ck = trim(Request.Form("ck_" & i))
				tmp_msid = trim(Request.Form("msid_" & i))
				if tmp_ck <> "" then
					exMailInfo.Load Session("wem"), tmp_ck
					if bj = "1" then
						exMailInfo.SetRead(true)
					else
						exMailInfo.SetRead(false)
					end if

					ei.Get_Session_IsRead_IsStar_Labels Session("wem"), tmp_msid, ret_IsRead, ret_IsStar
					if ret_IsRead = true then
						ret_show_str = ret_show_str & "1"
					else
						ret_show_str = ret_show_str & "0"
					end if

					ret_IsRead = NULL
					ret_IsStar = NULL
				end if 

			    i = i + 1
			loop
		end if

		if Len(ret_show_str) < 1 then
			ret_show_str = "x"
		end if
	elseif set_mode = "4" then
		bj = trim(Request.Form("bj"))
		if themax > 0 then
			do while i <= themax
				tmp_ck = trim(Request.Form("ck_" & i))
				tmp_msid = trim(Request.Form("msid_" & i))
				if tmp_ck <> "" then
					exMailInfo.Load Session("wem"), tmp_ck
					if bj = "1" then
						if exMailInfo.AddLabel("--star--") = true then
							exMailInfo.SetStar(true)
						end if
					else
						if exMailInfo.DelLabel("--star--") = true then
							exMailInfo.SetStar(false)
						end if
					end if

					ei.Get_Session_IsRead_IsStar_Labels Session("wem"), tmp_msid, ret_IsRead, ret_IsStar
					if ret_IsStar = true then
						ret_show_str = ret_show_str & "1"
					else
						ret_show_str = ret_show_str & "0"
					end if

					ret_IsRead = NULL
					ret_IsStar = NULL
				end if 

			    i = i + 1
			loop
		end if

		if Len(ret_show_str) < 1 then
			ret_show_str = "x"
		end if
	end if

	set exMailInfo = nothing
	set ei = nothing

	if Len(ret_show_str) > 0 then
		Response.Write ret_show_str
	else
		if is_ok = true then
			Response.Write "1"
		else
			Response.Write "0"
		end if
	end if
	Response.End
end if

Show_EC_Date_Style = false
if sortstr = "" or LCase(sortstr) = "date" then
	dim uw
	set uw = server.createobject("easymail.UserWeb")
	uw.Load Session("wem")
	Show_EC_Date_Style = uw.EnableShowDateECMailList
	set uw = nothing
end if

dim mlb
set mlb = server.createobject("easymail.Labels")
mlb.Load Session("wem")

dim pf
set pf = server.createobject("easymail.PerFolders")
pf.Load Session("wem")

allnum = ei.getMailsCount


if trim(request("page")) = "" then
	page = 0
else
	page = CInt(request("page"))
end if


if Show_EC_Date_Style = true then
	allnum = ei.GetDateEXmails(exstr)
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
	allnum = ei.getMailsCount
end if

dim show_EC_mode
show_EC_mode = -1

dim is_show_EC
dim bottom_bar

dim is_already_show
is_already_show = false

gourl = "listsesmail.asp?page=" & page & "&" & getGRSN()
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
<link rel="stylesheet" type="text/css" href="images/listsesmail.css">

<STYLE type=text/css>
<!--
.st_1 {width:3%;}
.st_2 {width:5%;}
.st_3 {width:3%;}
.st_4 {width:17%;}
.st_5 {width:41%;}
.st_6 {width:21%;}
.st_7 {width:10%; border-right:1px solid #c1c8d2;}
-->
</STYLE>
</HEAD>

<script type="text/javascript" src="images/sc_left.js"></script>
<script type="text/javascript" src="images/mglobal.js"></script>
<script type="text/javascript" src="images/jquery.min.js"></script>
<script type="text/javascript" src="images/jquery-powerFloat-min.js"></script>

<SCRIPT LANGUAGE=javascript>
<!--
if (ie == 6)
	document.execCommand("BackgroundImageCache", false, true);

parent.f1.document.leftval.tgname.value = "";
parent.f1.document.leftval.sortinfo.value = "<%
if sortmode = true then
	Response.Write "&sortstr=" & sortstr & "&sortmode=1"
else
	Response.Write "&sortstr=" & sortstr & "&sortmode=0"
end if
%>";
parent.f1.document.leftval.temp.value = "";

function setsort(addsortstr){
	if ("<%=sortstr %>" != addsortstr)
		location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=0&exstr=<%=exstr %>";
	else
<% if sortmode = false then %>
		location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=1&exstr=<%=exstr %>";
<% else %>
		location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=0&exstr=<%=exstr %>";
<% end if %>
}

function ck_select(tag_obj)
{
	if (tag_obj.checked == true)
		document.getElementById("tr_" + tag_obj.id.substr(3)).style.background = "#93BEE2";
	else
		document.getElementById("tr_" + tag_obj.id.substr(3)).style.background = "white";
}

function m_over(tag_obj)
{
	if (document.getElementById("ck_" + tag_obj.id.substr(3)).checked == false)
		tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj)
{
	if (document.getElementById("ck_" + tag_obj.id.substr(3)).checked == false)
		tag_obj.style.backgroundColor = "white";
}

function showsession(s_url) {
	parent.f1.document.leftval.purl.value = "<%=gourl & addsortstr %>";
	location.href = "showsession.asp?" + s_url;
}

parent.f1.document.leftval.purl.value = "";
function showlabel(lb_id, e) {
	if (!e) var e = window.event;
	e.cancelBubble = true;
	if (e.stopPropagation)
	e.stopPropagation();

	parent.f1.document.leftval.purl.value = "<%=gourl & addsortstr %>";
	location.href = "listlabel.asp?<%=getGRSN() %>&lbid=" + lb_id;
}

function set_one_star(tgli, bj)
{
	ck_array = [];
	ck_array.push(tgli);

	var post_date = "setmode=4&<%=getGRSN() %>&bj=" + bj + "&maxck=" + (tgli + 1) + "&ck_" + tgli + "=" + document.getElementById("ck_" + tgli).value + "&msid_" + tgli + "=" + ms_array[tgli];
	send_star(post_date);
}

function set_star(bj)
{
	var ck_date = get_sel_check(true);
	if (ck_date.length < 1)
		return ;

	var post_date = "setmode=4&<%=getGRSN() %>&bj=" + bj + "&" + ck_date;
	send_star(post_date);
}

function send_star(post_date)
{
$.ajax({
	type:"POST",
	url:"listsesmail.asp",
	data:post_date,
	success:function(data){
		if (data == "x")
			document.location.reload(true);
		else
		{
			var theObj;
			for (var i = 0; i < ck_array.length; i++)
			{
				theObj = document.getElementById("ck_" + ck_array[i]);
				if (theObj != null)
				{
					if (theObj.checked == true)
					{
						theObj.checked = false;
						ck_select(theObj);
					}
				}

				theObj = document.getElementById("icn_star_" + ck_array[i]);
				if (theObj != null)
				{
					if (data.charAt(i) == '1')
					{
						theObj.src = "images/star_yes.gif";
						theObj.onclick = new Function("set_one_star(" + ck_array[i] + ", '0')");
					}
					else
					{
						theObj.src = "images/star_no.gif";
						theObj.onclick = new Function("set_one_star(" + ck_array[i] + ", '1')");
					}
				}
			}
		}
		ck_array = [];
	},
	error:function(){
		ck_array = [];
	}
});
}

function set_read(bj)
{
	var ck_date = get_sel_check(true);
	if (ck_date.length < 1)
		return ;

	var post_date = "setmode=3&<%=getGRSN() %>&bj=" + bj + "&" + ck_date;

$.ajax({
	type:"POST",
	url:"listsesmail.asp",
	data:post_date,
	success:function(data){
		if (data == "x")
			document.location.reload(true);
		else
		{
			var theObj;
			for (var i = 0; i < ck_array.length; i++)
			{
				theObj = document.getElementById("ck_" + ck_array[i]);
				if (theObj != null)
				{
					if (theObj.checked == true)
					{
						theObj.checked = false;
						ck_select(theObj);
					}
				}

				theObj = document.getElementById("tr_" + ck_array[i]);
				if (data.charAt(i) == '1')
					theObj.className = "cont_tr";
				else
					theObj.className = "cont_tr_b";

				theObj = document.getElementById("icn_b_" + ck_array[i]);
				if (theObj != null)
				{
					if (data.charAt(i) == '1')
					{
						theObj.src = "mail.gif";
						theObj.title = "<%=s_lang_0420 %>";
					}
					else
					{
						theObj.src = "newmail.gif";
						theObj.title = "<%=s_lang_0421 %>";
					}
				}

				theObj = document.getElementById("icn_e_" + ck_array[i]);
				if (theObj != null)
				{
					if (data.charAt(i) == '1')
					{
						if (theObj.src.indexOf("s0.gif") != -1)
							theObj.src = "images/s0-1.gif";

						if (theObj.src.indexOf("e0.gif") != -1)
							theObj.src = "images/e0-1.gif";
					}
					else
					{
						if (theObj.src.indexOf("s0-1.gif") != -1)
							theObj.src = "images/s0.gif";

						if (theObj.src.indexOf("e0-1.gif") != -1)
							theObj.src = "images/e0.gif";
					}
				}
			}
		}
		ck_array = [];
	},
	error:function(){
		ck_array = [];
	}
});
}

function out_label(tgli, tglbid, isolb)
{
	if (isolb == "0")
		return ;

	var tgobj = document.getElementById("lbc_" + tgli + "_" + tglbid);
	if (tgobj != null)
	{
		tgobj.style.background = "";
		tgobj.style.width = "0px";
		tgobj.style.display = "";
	}
}

function up_label(tgli, tglbid, isolb)
{
	if (isolb == "0")
		return ;

	var tgobj = document.getElementById("lbc_" + tgli + "_" + tglbid);
	if (tgobj != null)
	{
		tgobj.style.background = "url('images/lbclose.gif')";
		tgobj.style.backgroundRepeat = "no-repeat";
		tgobj.style.backgroundPosition = "right center";
		tgobj.style.width = "19px";
		tgobj.style.display = "inline-block";
	}
}

var ms_array = [];
function push_msid(tgmsid)
{
	ms_array.push(tgmsid);
}

function lbclose(tgli, tglbid, e, isolb)
{
	if (isolb == "0")
		return ;

	if (!e) var e = window.event;
	e.cancelBubble = true;
	if (e.stopPropagation)
	e.stopPropagation();

	var post_date = "setmode=2&<%=getGRSN() %>&lbid=" + tglbid + "&msid=" + ms_array[tgli] + "&mailid=" + document.getElementById("ck_" + tgli).value;

$.ajax({
	type:"POST",
	url:"listsesmail.asp",
	data:post_date,
	success:function(data){
		if (data == "0")
			document.location.reload(true);
		else if (data == "1")
		{
			$("#lba_" + tgli + "_" + tglbid).remove();
			$("#sp_lb_bk_" + tgli + "_" + tglbid).remove();
		}
		else if (data == "2")
		{
			var theObj = document.getElementById("lb_" + tgli + "_" + tglbid);
			if (theObj != null)
			{
				out_label(tgli, tglbid, "1");
				theObj.onmouseover = "";
				theObj.onmouseout = "";
			}
		}
	},
	error:function(){
	}
});
}

function get_lb_str(mailid, lb_id, lb_title, lb_color)
{
	return "<span id=\"lba_" + mailid + "_" + lb_id + "\"><span id=\"lb_" + mailid + "_" + lb_id + "\" class=\"wwm_lb_box\" style=\"background:#" + lb_color + ";\"\
 onclick=\"showlabel('" + lb_id + "', event)\" onmouseover=\"up_label('" + mailid + "', '" + lb_id + "', 1)\" onmouseout=\"out_label('" + mailid + "', '" + lb_id + "', 1)\"><span class=\"wwm_lb_text\">" + htmlEscape(lb_title) + "</span>\
<span id=\"lbc_" + mailid + "_" + lb_id + "\" class=\"wwm_lb_close\" onclick=\"lbclose('" + mailid + "', '" + lb_id + "', event, 1)\">&nbsp;</span></span></span>";
}

var ck_array = [];

function set_lb(tgid, tgcolor, tgtitle)
{
	var ck_date = get_sel_check(false);
	if (ck_date.length < 1)
		return ;

	var post_date = "setmode=1&<%=getGRSN() %>&lbid=" + tgid + "&" + ck_date;
	var theObj;

$.ajax({
	type:"POST",
	url:"listsesmail.asp",
	data:post_date,
	success:function(data){
		if (data != "1")
			document.location.reload(true);
		else
		{
			for (var i = 0; i < ck_array.length; i++)
			{
				theObj = document.getElementById("ck_" + ck_array[i]);
				if (theObj != null)
				{
					if (theObj.checked == true)
					{
						theObj.checked = false;
						ck_select(theObj);
					}
				}

				theObj = document.getElementById("lba_" + ck_array[i] + "_" + tgid);
				if (theObj == null)
				{
					if ($("#sp_lb_" + ck_array[i]).text().length > 0)
						$("#sp_lb_" + ck_array[i]).append("<span id='sp_lb_bk_" + ck_array[i] + "_" + tgid + "' style='float:right; width:3px;'>&nbsp;</span>" + get_lb_str(ck_array[i], tgid, tgtitle, tgcolor));
					else
						$("#sp_lb_" + ck_array[i]).append(get_lb_str(ck_array[i], tgid, tgtitle, tgcolor));
				}
			}
		}
		ck_array = [];
	},
	error:function(){
		ck_array = [];
	}
});
}

function del() {
	if (ischeck() == true)
	{
		document.f1.mto.value = "del";
		document.f1.gourl.value = "<%=gourl & addsortstr %>";
		document.f1.mode.value = "semove";
		change_sel_check();
		document.f1.submit();
	}
}

function killspam() {
	if (ischeck() == true)
	{
		if (confirm("<%=s_lang_0422 %>") == false)
			return ;

		document.f1.gourl.value = "<%=gourl & addsortstr %>";
		document.f1.mode.value = "spam";
		document.f1.submit();
	}
}

function realdel() {
	if (ischeck() == true)
	{
		if (confirm("<%=s_lang_0443 %>") == false)
			return ;

		document.f1.gourl.value = "<%=gourl & addsortstr %>";
		document.f1.mode.value = "sedel";
		document.f1.isremove.value = "1";
		change_sel_check();
		document.f1.submit();
	}
}

function move(tgname) {
	if (ischeck() == false)
	{
		document.f1.mto.value = "";
		return ;
	}

	document.f1.mto.value = tgname;
	document.f1.gourl.value = "<%=gourl & addsortstr %>";
	document.f1.mode.value = "semove";
	change_sel_check();
	document.f1.submit();
}

function allcheck_onclick() {
	document.body.focus();
	return false;
}

function selectpage_onchange()
{
	location.href = "listsesmail.asp?page=" + document.f1.page.value + "<%=addsortstr & "&" & getGRSN() %>";
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
%>" + "&exstr=" + url_exstr;
}
<%
end if
%>
//-->
</SCRIPT>

<BODY>
<FORM ACTION="mulmail.asp" METHOD="POST" name="f1" id="f1">
<INPUT NAME="mode" TYPE="hidden">
<INPUT NAME="mto" TYPE="hidden">
<INPUT NAME="gourl" TYPE="hidden">
<table id="table_main" class="table_main" align="center" cellspacing="0" cellpadding="0">
  <tr>
	<td colspan="7" class="box_title_td">
<font class="font_top_title"><%=s_lang_0327 %></font><%
if allnum > 0 then
	Response.Write " " & s_lang_0401 & " <font color='#901111'>" & allnum & "</font>" & s_lang_0423
end if
%>
	</td></tr>
	<tr><td class="block_top_td" colspan="7"><div class="table_min_width"></div></td></tr>
  <tr>
    <td colspan="7" class="tool_top_td">
<span class="st_span">
<a class='wwm_btnDownload btn_gray' href='javascript:del()'><%=s_lang_del %></a>
</span>
<span style='float:left; width:3px;'>&nbsp;</span>

<span class="st_span">
<a class='wwm_btnDownload btn_gray' href='javascript:realdel()'><%=s_lang_0424 %></a>
</span>
<span style='float:left; width:3px;'>&nbsp;</span>

<span class="st_span"><span id="pm_moveto" class="menu_pop">
<div class='attbg'><img style='margin: 6px 0pt 0pt;' src='images/popshow.gif'></div>
<div class='menu_pop_text'><%=s_lang_0404 %>...</div>
</span></span>
<span style='float:left; width:3px;'>&nbsp;</span>

<span class="st_span"><span id="pm_bj" class="menu_pop">
<div class='attbg'><img style='margin: 6px 0pt 0pt;' src='images/popshow.gif'></div>
<div class='menu_pop_text'><%=s_lang_0425 %>...</div>
</span></span>
<span style='float:left; width:3px;'>&nbsp;</span>

<span class="st_span"><span id="pm_more" class="menu_pop" style="width:48px; +width:51px; _width:48px;">
<div class='attbg'><img style='margin: 6px 0pt 0pt;' src='images/popshow.gif'></div>
<div class='menu_pop_text'><%=s_lang_0426 %></div>
</span></span>

<span class="st_r1_span"><select name="page" class="drpdwn" size="1" LANGUAGE=javascript onchange="selectpage_onchange()">
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
%></select>/<%=allpage %>
</span>

<span class="st_r2_span"><%
if page - 1 < 0 then
	bottom_bar = "<img src='images/gfirstp.gif' border='0' align='absmiddle'>&nbsp;"
	bottom_bar = bottom_bar & "<img src='images/gprep.gif' border='0' align='absmiddle'>&nbsp;"
	Response.Write "<img src='images/gfirstp.gif' border='0' align='absmiddle'>&nbsp;"
	Response.Write "<img src='images/gprep.gif' border='0' align='absmiddle'>&nbsp;"
else
	bottom_bar = "<a href=""listsesmail.asp?page=" & 0 & addsortstr & "&" & getGRSN() & """><img src='images/firstp.gif' border='0' align='absmiddle'></a>&nbsp;"
	bottom_bar = bottom_bar & "<a href=""listsesmail.asp?page=" & page - 1 & addsortstr & "&" & getGRSN() & """><img src='images/prep.gif' border='0' align='absmiddle'></a>&nbsp;"
	Response.Write "<a href=""listsesmail.asp?page=" & 0 & addsortstr & "&" & getGRSN() & """><img src='images/firstp.gif' border='0' align='absmiddle'></a>&nbsp;"
	Response.Write "<a href=""listsesmail.asp?page=" & page - 1 & addsortstr & "&" & getGRSN() & """><img src='images/prep.gif' border='0' align='absmiddle'></a>&nbsp;"
end if

if ((page+1) * pageline) => show_allnum then
	bottom_bar = bottom_bar & "<img src='images/gnextp.gif' border='0' align='absmiddle'>&nbsp;"
	Response.Write "<img src='images/gnextp.gif' border='0' align='absmiddle'>&nbsp;"
else
	bottom_bar = bottom_bar & "<a href=""listsesmail.asp?page=" & page + 1 & addsortstr & "&" & getGRSN() & """><img src='images/nextp.gif' border='0' align='absmiddle'></a>&nbsp;"
	Response.Write "<a href=""listsesmail.asp?page=" & page + 1 & addsortstr & "&" & getGRSN() & """><img src='images/nextp.gif' border='0' align='absmiddle'></a>&nbsp;"
end if

if page + 1 >= allpage then
	bottom_bar = bottom_bar & "<img src='images/gendp.gif' border='0' align='absmiddle'>"
	Response.Write "<img src='images/gendp.gif' border='0' align='absmiddle'>"
else
	bottom_bar = bottom_bar & "<a href=""listsesmail.asp?page=" & allpage - 1 & addsortstr & "&" & getGRSN() & """><img src='images/endp.gif' border='0' align='absmiddle'></a>"
	Response.Write "<a href=""listsesmail.asp?page=" & allpage - 1 & addsortstr & "&" & getGRSN() & """><img src='images/endp.gif' border='0' align='absmiddle'></a>"
end if
%></span>
    </td>
  </tr>
    <tr class="title_tr">
	<td class="st_1"> 
		<a href="javascript:setsort('Priority')"><img src='images/high.gif' border='0' align='absmiddle'></a><%
if sortstr = "Priority" then
 	if sortmode = true then
		response.write "<a href=""javascript:setsort('Priority')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		response.write "<a href=""javascript:setsort('Priority')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%>
	</td>
	<td class="st_2">
		<a href="javascript:setsort('Read')"><%=s_lang_0126 %></a><%
if sortstr = "Read" then
 	if sortmode = true then
		response.write "<a href=""javascript:setsort('Read')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		response.write "<a href=""javascript:setsort('Read')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%>
	</td>
	<td class="st_3">
		<input type="checkbox" id="allcheck" onclick="return allcheck_onclick()">
	</td>
	<td class="st_4">
		<a href="javascript:setsort('Sender')"><%=s_lang_0147 %></a><%
if sortstr = "Sender" then
 	if sortmode = true then
		response.write "&nbsp;<a href=""javascript:setsort('Sender')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		response.write "&nbsp;<a href=""javascript:setsort('Sender')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%>
	</td>
	<td class="st_5">
		<a href="javascript:setsort('Subject')"><%=s_lang_0127 %></a><%
if sortstr = "Subject" then
 	if sortmode = true then
		response.write "&nbsp;<a href=""javascript:setsort('Subject')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		response.write "&nbsp;<a href=""javascript:setsort('Subject')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%>
	</td>
	<td class="st_6">
		<a href="javascript:setsort('Date')"><%=s_lang_0128 %></a><%
if sortstr = "" or sortstr = "Date" then
 	if sortmode = true then
		response.write "&nbsp;<a href=""javascript:setsort('Date')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		response.write "&nbsp;<a href=""javascript:setsort('Date')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%>
	</td>
	<td class="st_7">
		<a href="javascript:setsort('Size')"><%=s_lang_0179 %></a><%
if sortstr = "Size" then
 	if sortmode = true then
		response.write "&nbsp;<a href=""javascript:setsort('Size')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		response.write "&nbsp;<a href=""javascript:setsort('Size')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%>
	</td>
    </tr>
<%
i = page * pageline
li = 0

show_page_ex_head = true

if Show_EC_Date_Style = true and page > 0 then
	ei.GetDateEXMode exstr, allnum - i - 1, exmode, isshow
	i = i + getBeforeLines(exmode)

	exmode = NULL
	isshow = NULL

	ei.GetDateEXMode exstr, allnum - i, bf_exmode, bf_isshow
	ei.GetDateEXMode exstr, allnum - i - 1, exmode, isshow

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

	ei.getMailSessionInfo allnum - i - 1, idname, isread, priority, sendMail, sendName, subject, size, etime, mstate, msnum, msid

if Show_EC_Date_Style = false or (Show_EC_Date_Style = true and writeDateEC(allnum - i - 1) = true) then
	issign = false
	isenc = false

	if ei.MailIsSignature(allnum - i - 1) = true then
		issign = true
	end if

	if ei.MailIsEncrypted(allnum - i - 1) = true then
		isenc = true
	end if

	if subject = "" then
		subject = s_lang_0410
	end if


	xmsp = priority

	if xmsp = "High" then
		xmsp = "<img src='images/high.gif' border='0' title='" & s_lang_0130 & "'>"
	elseif xmsp = "Low" then
		xmsp = "<img src='images/low.gif' border='0' title='" & s_lang_0131 & "'>"
	else
		xmsp = "&nbsp;"
	end if

	if Len(msid) = 16 then
		ei.Get_Session_IsRead_IsStar_Labels Session("wem"), msid, ret_IsRead, ret_IsStar
	end if

	show_mail_function = " onclick=""showsession('msid=" & msid & "&" & getGRSN() & "&gourl=" & Server.URLEncode(gourl & addsortstr) & "');"""

	is_already_show = true
	Response.Write "<script>push_msid('" & msid & "');</script>" & Chr(13)
%>
    <tr id="tr_<%=li %>" class="cont_tr<% if ret_IsRead = false then Response.Write "_b" %>" onmouseover='m_over(this);' onmouseout='m_out(this);'>
	<td class="cont_td_1"><%=xmsp %></td>
	<td class="cont_td_2"><%
if mstate = 0 then
	Response.Write "<img id='icn_b_" & li & "' src='reply.gif' title='" & s_lang_0427 & "' border='0'"
elseif mstate = 1 then
	Response.Write "<img id='icn_b_" & li & "' src='forward.gif' title='" & s_lang_0428 & "' border='0'"
elseif mstate = 2 or mstate = 3 then
	if ret_IsRead = true then
		Response.Write "<img id='icn_b_" & li & "' src='rsysmail.gif' title='" & s_lang_0181 & "' border='0'"
	else
		Response.Write "<img id='icn_b_" & li & "' src='nsysmail.gif' title='" & s_lang_0429 & "' border='0'"
	end if
else
	if ret_IsRead = true then
		Response.Write "<img id='icn_b_" & li & "' src='mail.gif' title='" & s_lang_0420 & "' border='0'"
	else
		Response.Write "<img id='icn_b_" & li & "' src='newmail.gif' title='" & s_lang_0421 & "' border='0'"
	end if
end if

if issign = true then
	if ret_IsRead = true then
		Response.Write "><img id='icn_e_" & li & "' src='images/s0-1.gif' title='" & s_lang_0183 & "' border='0'"
	else
		Response.Write "><img id='icn_e_" & li & "' src='images/s0.gif' title='" & s_lang_0183 & "' border='0'"
	end if
elseif isenc = true then
	if ret_IsRead = true then
		Response.Write "><img id='icn_e_" & li & "' src='images/e0-1.gif' title='" & s_lang_0184 & "' border='0'"
	else
		Response.Write "><img id='icn_e_" & li & "' src='images/e0.gif' title='" & s_lang_0184 & "' border='0'"
	end if
end if
%>><%
exMailInfo.Load Session("wem"), idname

if exMailInfo.Have_Attachment = true then
	Response.Write "<img src='images/atta.gif' border='0'>"
end if
%></td>
	<td class="cont_td_3"><input type="checkbox" id="ck_<%=li %>" name="ck_<%=li %>" value="<%=idname %>" onclick="ck_select(this);"></td>
	<td class="cont_td_4"<%=show_mail_function %>><%=server.htmlencode(sendName) %>&nbsp;</td>
	<td id="td_subject_<%=li %>" class="cont_td_5"<%=show_mail_function %>><span class="cs_subject"><%
if msnum > 1 then
	Response.Write server.htmlencode(subject) & "(" & msnum & ")" & "</span>"
else
	Response.Write server.htmlencode(subject) & "</span>"
end if

Response.Write "<span id='sp_lb_" & li & "' style='float:right; display:inline-block;'>"
if ei.LabelCount > 0 then
	lball = ei.LabelCount
	lbi = 0
	do while lbi < lball
		mlb.GetByID ei.GetLabel(lbi), ret_id, ret_title, ret_color

		if lbi > 0 then
			Response.Write "<span id='sp_lb_bk_" & li & "_" & ret_id & "' style='float:right; width:3px;'><font style='font-size:1px;'>&nbsp;</font></span>"
		end if

		Response.Write create_label_str(ret_id, li, ret_title, ret_color, exMailInfo.FindLabel(ret_id)) & Chr(13)
		ret_id = NULL
		ret_title = NULL
		ret_color = NULL

		lbi = lbi + 1
	loop
end if
Response.Write "</span>"
%></td>
	<td class="cont_td_6"<%=show_mail_function %>><%=etime %></td>
	<td class="cont_td_7"><span class="cs_star"><%
Response.Write getShowSize(size)

if ret_IsStar = true then
	Response.Write "</span><img id='icn_star_" & li & "' src='images/star_yes.gif' border='0' style='cursor:pointer;' onclick=""set_one_star(" & li & ", '0');""></a>"
else
	Response.Write "</span><img id='icn_star_" & li & "' src='images/star_no.gif' border='0' style='cursor:pointer;' onclick=""set_one_star(" & li & ", '1');""></a>"
end if
%></td>
    </tr>
<%	
    li = li + 1
end if

	idname = NULL
	isread = NULL
	priority = NULL
	sendMail = NULL
	sendName = NULL
	subject = NULL
	size = NULL
	etime = NULL
	mstate = NULL
	msnum = NULL
	msid = NULL
	ret_IsRead = NULL
	ret_IsStar = NULL

	i = i + 1
loop


if Show_EC_Date_Style = true then
	last_exmode = -1
	do while i < allnum
		ei.GetDateEXMode exstr, allnum - i - 1, bf_exmode, bf_isshow
		tmp_bf_exmode = bf_exmode

		if bf_isshow = true or canshowlast(tmp_bf_exmode) = false then
			bf_exmode = NULL
			bf_isshow = NULL

			exit do
		end if

		if last_exmode <> bf_exmode then
			writeDateEC(allnum - i - 1)
		end if

		bf_exmode = NULL
		bf_isshow = NULL

		i = i + 1
	loop
end if
%>
<%
if is_already_show = true then
%>
<tr><td class="block_td" colspan="7"></td></tr>
<tr><td  colspan="7" class="tool_td" style="height:22px; text-align:center;">
<%=bottom_bar %>
</td></tr>
<%
end if
%>
</table>
<INPUT NAME="isremove" TYPE="hidden" value="0">
<br>
</FORM>
<div id="pmc_moveto" class="qmpanel_shadow" style="display:none; position:absolute;">
	<div class="menu_base">
	<div id="md_moveto" class="menu_bd bd">
	<table width="100%">
		<div name="mi" onclick="move('del');" class="menu_item"><%=s_lang_0334 %></div>
		<div name="mi" onclick="move('in');" class="menu_item"><%=s_lang_0327 %></div>
		<div name="mi" onclick="move('out');" class="menu_item"><%=s_lang_0332 %></div>
		<div name="mi" onclick="move('sed');" class="menu_item"><%=s_lang_0430 %></div>
<%
pfNumber = pf.FolderCount

if pfNumber > 0 then
	Response.Write "<div class='menu_item_nofun'><div style='background:#ccc; padding-top:1px; margin-top: 5px;'></div></div>"
end if

dim moveto_set_max
moveto_set_max = false

if pfNumber > 6 then
	moveto_set_max = true
end if

dim moveto_max_len
moveto_max_len = 0
i = 0
do while i < pfNumber
	spfname = pf.GetFolderName(i)

	t_len = getLength(spfname)
	if t_len > moveto_max_len then
		moveto_max_len = t_len
	end if

	Response.Write "<div name='mi' onclick=""move('" & pf.GetFolderID(spfname) & "');"" class='menu_item'>" & server.htmlencode(spfname) & "</div>"
	spfname = NULL

	i = i + 1
loop
%>
	</table>
	</div>
	</div>
</div>

<div id="pmc_bj" class="qmpanel_shadow" style="display:none; position:absolute;">
	<div class="menu_base">
	<div id="lb_bj" class="menu_bd bd">
	<table width="100%">
		<div name="mi" onclick="set_read('1');" class="menu_item"><%=s_lang_0431 %></div>
		<div name="mi" onclick="set_read('0');" class="menu_item"><%=s_lang_0432 %></div>
		<div class="menu_item_nofun"><div style="background:#ccc; padding-top:1px; margin-top: 5px;"></div></div>
		<div name="mi" onclick="set_star('1');" class="menu_item"><%=s_lang_0433 %></div>
		<div name="mi" onclick="set_star('0');" class="menu_item"><%=s_lang_0434 %></div>
<%
dim bj_set_max
dim bj_lb_max_len_title
bj_set_max = false
bj_lb_max_len_title = 0

allnum = mlb.Count

if allnum > 6 then
	bj_set_max = true
end if

if allnum > 0 then
	Response.Write "		<div class='menu_item_nofun'><div style='background:#ccc; padding-top:1px; margin-top: 5px;'></div></div>"

	i = 0
	do while i < allnum
		mlb.GetByIndex i, ret_id, ret_title, ret_color

		t_len = getLength(ret_title)
		if t_len > bj_lb_max_len_title then
			bj_lb_max_len_title = t_len
		end if
		Response.Write "<div onclick=""set_lb('" & ret_id & "', '" & ret_color & "', '" & server.htmlencode(ret_title) & "');"" name='mi' class='menu_item'><span class='wwm_color_in_line' style='background:#" & ret_color & ";'>&nbsp;</span> " & server.htmlencode(ret_title) & "</div>" & Chr(13)

		ret_id = NULL
		ret_title = NULL
		ret_color = NULL

		i = i + 1
	loop
end if
%>
	</table>
	</div>
	</div>
</div>

<div id="pmc_more" class="qmpanel_shadow" style="display:none; position:absolute;">
	<div class="menu_base">
	<div class="menu_bd bd">
	<table width="100%">
		<div name="mi" onclick="location.href='listmail.asp?set=1&mode=in&<%=getGRSN() %>';" class="menu_item"><%=s_lang_0444 %></div>
		<div class="menu_item_nofun"><div style="background:#ccc; padding-top:1px; margin-top: 5px;"></div></div>
		<div name="mi" onclick="killspam();" class="menu_item"><%=s_lang_0435 %></div>
		<div class="menu_item_nofun"><div style="background:#ccc; padding-top:1px; margin-top: 5px;"></div></div>
		<div name="mi" onclick="location.href='labels.asp?needrt=1&<%=getGRSN() %>&gourl=<%=Server.URLEncode(gourl & addsortstr) %>';" class="menu_item"><%=s_lang_0300 %></div>
	</table>
	</div>
	</div>
</div>

<div id="pmc_ck" class="qmpanel_shadow" style="display:none; position:absolute;">
	<div class="menu_base">
	<div class="menu_bd bd">
	<table width="100%">
		<div name="mi" onclick="set_ck(1);" class="menu_item"><%=s_lang_0415 %></div>
		<div name="mi" onclick="set_ck(0);" class="menu_item"><%=s_lang_0416 %></div>
		<div name="mi" onclick="set_ck(2);" class="menu_item"><%=s_lang_0417 %></div>
		<div class="menu_item_nofun" style="padding-left:5px; padding-right:5px;"><div style="background:#ccc; padding-top:1px; margin-top: 5px;"></div></div>
		<div name="mi" onclick="set_ck(3);" class="menu_item"><%=s_lang_0418 %></div>
		<div name="mi" onclick="set_ck(4);" class="menu_item"><%=s_lang_0419 %></div>
	</table>
	</div>
	</div>
</div>

<script type="text/javascript">
var mil = document.getElementsByTagName("div");
for (var i=0; i<mil.length; i++) 
{
	if (mil[i].name == "mi")
	{
		mil[i].onmouseover = function(){this.style.backgroundColor='#3470cc';this.style.color='#ffffff';}
		mil[i].onmouseout = function(){this.style.backgroundColor='#ffffff';this.style.color='#000000';}
	}
}

var is_in_menu_moveto = false;
var is_menu_show_moveto = false;
var my_menu_time_moveto;
var is_in_menu_bj = false;
var is_menu_show_bj = false;
var my_menu_time_bj;
var is_in_menu_more = false;
var is_menu_show_more = false;
var my_menu_time_more;
var is_in_menu_ck = false;
var is_menu_show_ck = false;
var my_menu_time_ck;

$(function() {
<%
if moveto_set_max = true then
	Response.Write "	$(""#md_moveto"").height(226);"
end if
%>
	$("#pm_moveto").powerFloat({
<%
if moveto_max_len > 10 then
	Response.Write "		width: " & (moveto_max_len * 6) + 38 & ","
else
	Response.Write "		width: 110,"
end if
%>
		eventType: "click",
		target: "#pmc_moveto",
		showCall: function() {
			if (is_menu_show_moveto == true)
				$.powerFloat.hide();
			else
			{
				is_menu_show_bj = false;
				is_menu_show_more = false;
				is_menu_show_ck = false;

				is_menu_show_moveto = true;
				clearTimeout(my_menu_time_moveto);
			}

			$("#pmc_moveto").mouseover(function() {
				is_in_menu_moveto = true;
				clearTimeout(my_menu_time_moveto);
			});

			$("#pmc_moveto").mouseout(function() {
				is_in_menu_moveto = false;
				my_menu_time_moveto = setTimeout("setTimeClose(1)", 1000);
			});

			$(".menu_item").click(function() {
				$.powerFloat.hide();
			});
		},
		hideCall: function() {
			setTimeout("set_menu_close(1)", 300);
		}
	});
});

$(function() {
<%
if bj_set_max = true then
	Response.Write "	$(""#lb_bj"").height(232);"
end if
%>
	$("#pm_bj").powerFloat({
<%
if bj_lb_max_len_title > 10 then
	Response.Write "		width: " & (bj_lb_max_len_title * 6) + 60 & ","
else
	Response.Write "		width: 110,"
end if
%>
		eventType: "click",
		target: "#pmc_bj",
		showCall: function() {
			if (is_menu_show_bj == true)
				$.powerFloat.hide();
			else
			{
				is_menu_show_moveto = false;
				is_menu_show_more = false;
				is_menu_show_ck = false;

				is_menu_show_bj = true;
				clearTimeout(my_menu_time_bj);
			}

			$("#pmc_bj").mouseover(function() {
				is_in_menu_bj = true;
				clearTimeout(my_menu_time_bj);
			});

			$("#pmc_bj").mouseout(function() {
				is_in_menu_bj = false;
				my_menu_time_bj = setTimeout("setTimeClose(2)", 1000);
			});

			$(".menu_item").click(function() {
				$.powerFloat.hide();
			});
		},
		hideCall: function() {
			setTimeout("set_menu_close(2)", 300);
		}
	});
});

$(function() {
	$("#pm_more").powerFloat({
		width: 130,
		eventType: "click",
		target: "#pmc_more",
		showCall: function() {
			if (is_menu_show_more == true)
				$.powerFloat.hide();
			else
			{
				is_menu_show_moveto = false;
				is_menu_show_bj = false;
				is_menu_show_ck = false;

				is_menu_show_more = true;
				clearTimeout(my_menu_time_more);
			}

			$("#pmc_more").mouseover(function() {
				is_in_menu_more = true;
				clearTimeout(my_menu_time_more);
			});

			$("#pmc_more").mouseout(function() {
				is_in_menu_more = false;
				my_menu_time_more = setTimeout("setTimeClose(3)", 1000);
			});

			$(".menu_item").click(function() {
				$.powerFloat.hide();
			});
		},
		hideCall: function() {
			setTimeout("set_menu_close(3)", 300);
		}
	});
});

$(function() {
	$("#allcheck").powerFloat({
		width: 60,
		eventType: "click",
		target: "#pmc_ck",
		showCall: function() {
			if (is_menu_show_ck == true)
				$.powerFloat.hide();
			else
			{
				is_menu_show_moveto = false;
				is_menu_show_bj = false;
				is_menu_show_more = false;

				is_menu_show_ck = true;
				clearTimeout(my_menu_time_ck);
			}

			$("#pmc_ck").mouseover(function() {
				is_in_menu_ck = true;
				clearTimeout(my_menu_time_ck);
			});

			$("#pmc_ck").mouseout(function() {
				is_in_menu_ck = false;
				my_menu_time_ck = setTimeout("setTimeClose(4)", 1000);
			});

			$(".menu_item").click(function() {
				$.powerFloat.hide();
			});
		},
		hideCall: function() {
			setTimeout("set_menu_close(4)", 300);
		}
	});
});

function set_menu_close(tgv)
{
	if (tgv == 1)
		is_menu_show_moveto = false;
	else if (tgv == 2)
		is_menu_show_bj = false;
	else if (tgv == 3)
		is_menu_show_more = false;
	else if (tgv == 4)
		is_menu_show_ck = false;
}

function setTimeClose(tgv)
{
	if (is_menu_show_moveto == true && is_in_menu_moveto == false && tgv == 1)
		$.powerFloat.hide();

	if (is_menu_show_bj == true && is_in_menu_bj == false && tgv == 2)
		$.powerFloat.hide();

	if (is_menu_show_more == true && is_in_menu_more == false && tgv == 3)
		$.powerFloat.hide();

	if (is_menu_show_ck == true && is_in_menu_ck == false && tgv == 4)
		$.powerFloat.hide();
}

function get_sel_check(with_msid)
{
	var ret_val = "maxck=<%=li %>";
	var theObj;
	var is_check = false;
	ck_array = [];

	for(var i = 0; i < <%=li %>; i++)
	{
		theObj = document.getElementById("ck_" + i);
		if (theObj != null)
		{
			if (theObj.checked == true)
			{
				ret_val += "&ck_" + i + "=" + theObj.value;
				is_check = true;
				ck_array.push(i);

				if (with_msid == true)
					ret_val += "&msid_" + i + "=" + ms_array[i];
			}
		}
	}

	if (is_check == false)
		ret_val = "";

	return ret_val;
}

function change_sel_check()
{
	var theObj;
	for(var i = 0; i < <%=li %>; i++)
	{
		theObj = document.getElementById("ck_" + i);
		if (theObj != null)
		{
			if (theObj.checked == true)
				theObj.value = ms_array[i];
		}
	}
}

function set_ck(tgmode)
{
	var theObj;
	for(var i = 0; i < <%=li %>; i++)
	{
		theObj = document.getElementById("ck_" + i);
		if (theObj != null)
		{
			if (tgmode == 0)
				theObj.checked = false;
			else if (tgmode == 1)
				theObj.checked = true;
			else if (tgmode == 2)
			{
				if (theObj.checked == true)
					theObj.checked = false;
				else
					theObj.checked = true;
			}
			else if (tgmode == 3)
			{
				if (document.getElementById("tr_" + i).className == "cont_tr_b")
					theObj.checked = true;
				else
					theObj.checked = false;
			}
			else if (tgmode == 4)
			{
				if (document.getElementById("tr_" + i).className == "cont_tr")
					theObj.checked = true;
				else
					theObj.checked = false;
			}

			ck_select(theObj);
		}
	}
}

function ischeck() {
	var theObj;
	for(var i = 0; i < <%=li %>; i++)
	{
		theObj = document.getElementById("ck_" + i);

		if (theObj != null)
		{
			if (theObj.checked == true)
				return true;
		}
	}

	return false;
}
</script>

</BODY>
</HTML>


<%
set pf = nothing

function getShowSize(bytesize)
	if bytesize < 1000 then
		getShowSize = "1K"
	else
		if bytesize < 1000000 then
			tmpSize = CDbl(bytesize/1000)
			tmpindex = InStr(1, tmpSize, ".")
			if tmpindex = 0 then
				getShowSize = tmpSize & "K"
			else
				getShowSize = CDbl(Left(tmpSize, tmpindex + 1)) & "K"
			end if
		else
			tmpSize = CStr(CDbl(bytesize/1000000))
			tmpindex = InStr(1, tmpSize, ".")
			if tmpindex = 0 then
				getShowSize = tmpSize & "M"
			else
				getShowSize = CDbl(Left(tmpSize, tmpindex + 2)) & "M"
			end if
		end if
	end if
end function


function writeDateEC(mail_index)
	writeDateEC = true

	if Show_EC_Date_Style = true then
		ei.GetDateEXMode exstr, mail_index, exmode, isshow

		if exmode = 0 then
			if show_EC_mode <> 0 then 
				show_EC_mode = 0

				if show_page_ex_head = true then
					Response.Write "<tr><td colspan=7 class='EX_TITLE'><a href='javascript:click_EC(0)'><img src='images/"
					if isshow = true then
						Response.Write "listclose.gif' title='" & s_lang_0436 & "'"
					else
						Response.Write "listopen.gif' title='" & s_lang_0437 & "'"
						writeDateEC = false
					end if

					Response.Write " border='0' align='absmiddle'></a> <font class='EX_TITLE_FONT'>" & s_lang_0134 & "</font></td></tr>"
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
					Response.Write "<tr><td colspan=7 class='EX_TITLE'><a href='javascript:click_EC(1)'><img src='images/"
					if isshow = true then
						Response.Write "listclose.gif' title='" & s_lang_0436 & "'"
					else
						Response.Write "listopen.gif' title='" & s_lang_0437 & "'"
						writeDateEC = false
					end if

					Response.Write " border='0' align='absmiddle'></a> <font class='EX_TITLE_FONT'>" & s_lang_0135 & "</font></td></tr>"
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
					Response.Write "<tr><td colspan=7 class='EX_TITLE'><a href='javascript:click_EC(2)'><img src='images/"
					if isshow = true then
						Response.Write "listclose.gif' title='" & s_lang_0436 & "'"
					else
						Response.Write "listopen.gif' title='" & s_lang_0437 & "'"
						writeDateEC = false
					end if

					Response.Write " border='0' align='absmiddle'></a> <font class='EX_TITLE_FONT'>" & s_lang_0136 & "</font></td></tr>"
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
					Response.Write "<tr><td colspan=7 class='EX_TITLE'><a href='javascript:click_EC(3)'><img src='images/"
					if isshow = true then
						Response.Write "listclose.gif' title='" & s_lang_0436 & "'"
					else
						Response.Write "listopen.gif' title='" & s_lang_0437 & "'"
						writeDateEC = false
					end if

					Response.Write " border='0' align='absmiddle'></a> <font class='EX_TITLE_FONT'>" & s_lang_0137 & "</font></td></tr>"
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
					Response.Write "<tr><td colspan=7 class='EX_TITLE'><a href='javascript:click_EC(4)'><img src='images/"
					if isshow = true then
						Response.Write "listclose.gif' title='" & s_lang_0436 & "'"
					else
						Response.Write "listopen.gif' title='" & s_lang_0437 & "'"
						writeDateEC = false
					end if

					Response.Write " border='0' align='absmiddle'></a> <font class='EX_TITLE_FONT'>" & s_lang_0138 & "</font></td></tr>"
				else
					show_page_ex_head = true
				end if

				is_show_EC = writeDateEC
			else
				writeDateEC = is_show_EC
			end if
		elseif exmode = 5 then
			if show_EC_mode <> 5 then 
				show_EC_mode = 5

				if show_page_ex_head = true then
					Response.Write "<tr><td colspan=7 class='EX_TITLE'><a href='javascript:click_EC(5)'><img src='images/"
					if isshow = true then
						Response.Write "listclose.gif' title='" & s_lang_0436 & "'"
					else
						Response.Write "listopen.gif' title='" & s_lang_0437 & "'"
						writeDateEC = false
					end if

					Response.Write " border='0' align='absmiddle'></a> <font class='EX_TITLE_FONT'>" & s_lang_0438 & "</font></td></tr>"
				else
					show_page_ex_head = true
				end if

				is_show_EC = writeDateEC
			else
				writeDateEC = is_show_EC
			end if
		elseif exmode = 6 then
			if show_EC_mode <> 6 then 
				show_EC_mode = 6

				if show_page_ex_head = true then
					Response.Write "<tr><td colspan=7 class='EX_TITLE'><a href='javascript:click_EC(6)'><img src='images/"
					if isshow = true then
						Response.Write "listclose.gif' title='" & s_lang_0436 & "'"
					else
						Response.Write "listopen.gif' title='" & s_lang_0437 & "'"
						writeDateEC = false
					end if

					Response.Write " border='0' align='absmiddle'></a> <font class='EX_TITLE_FONT'>" & s_lang_0439 & "</font></td></tr>"
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

	if sortmode = false then
		do while tmp_mode >= 0
			if Mid(exstr, tmp_mode + 1, 1) = "1" then
				if tmp_mode = 6 then
					before_num = before_num + ei.DateEX_6
				elseif tmp_mode = 5 then
					before_num = before_num + ei.DateEX_5
				elseif tmp_mode = 4 then
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
				do while tmp_mode < 7
					if Mid(exstr, tmp_mode + 1, 1) = "1" then
						if tmp_mode = 6 then
							before_num = before_num + ei.DateEX_6
						elseif tmp_mode = 5 then
							before_num = before_num + ei.DateEX_5
						elseif tmp_mode = 4 then
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
	else
		do while tmp_mode < 7
			if Mid(exstr, tmp_mode + 1, 1) = "1" then
				if tmp_mode = 6 then
					before_num = before_num + ei.DateEX_6
				elseif tmp_mode = 5 then
					before_num = before_num + ei.DateEX_5
				elseif tmp_mode = 4 then
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
			tmp_mode = tmp_mode + 1
		loop

		if Mid(exstr, tmp_tmp_mode + 1, 1) = "1" then
			tmp_mode = tmp_tmp_mode - 1
			do while tmp_mode >= 0
				if Mid(exstr, tmp_mode + 1, 1) = "1" then
					if tmp_mode = 6 then
						before_num = before_num + ei.DateEX_6
					elseif tmp_mode = 5 then
						before_num = before_num + ei.DateEX_5
					elseif tmp_mode = 4 then
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
				tmp_mode = tmp_mode - 1
			loop
		end if
	end if

	getBeforeLines = before_num
end function


function canshowlast(tmp_last_exmode)
	canshowlast = true

	if sortmode = false then
		do while tmp_last_exmode < 7
			if Mid(exstr, tmp_last_exmode + 1, 1) = "0" and getDateMails(tmp_last_exmode) > 0 then
				canshowlast = false
				exit do
			end if
			tmp_last_exmode = tmp_last_exmode + 1
		loop
	else
		do while tmp_last_exmode >= 0
			if Mid(exstr, tmp_last_exmode + 1, 1) = "0" and getDateMails(tmp_last_exmode) > 0 then
				canshowlast = false
				exit do
			end if
			tmp_last_exmode = tmp_last_exmode - 1
		loop
	end if
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
	elseif sea_ex_mode = 5 then
		getDateMails = ei.DateEX_5
	elseif sea_ex_mode = 6 then
		getDateMails = ei.DateEX_6
	end if
end function


function writeBeforeHiddenTitle(start_exmode)
	tmp_start_exmode = -1
	tmp_end_exmode = start_exmode
	if sortmode = false then
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
	else
		do while start_exmode < 7
			if Mid(exstr, start_exmode + 1, 1) = "0" then
				exit do
			else
				tmp_start_exmode = start_exmode
			end if
			start_exmode = start_exmode + 1
		loop

		if tmp_start_exmode > 0 and tmp_start_exmode >= tmp_end_exmode then
			do while tmp_start_exmode >= tmp_end_exmode
				writeHiddenTitle(tmp_start_exmode)
				tmp_start_exmode = tmp_start_exmode - 1
			loop
		end if
	end if
end function


function writeHiddenTitle(write_exmode)
	Response.Write "<tr><td colspan=7 class='EX_TITLE'><a href='javascript:click_EC(" & write_exmode & ")'><img src='images/"
	Response.Write "listopen.gif' title='" & s_lang_0437 & "' border='0' align='absmiddle'></a> <font class='EX_TITLE_FONT'>"

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
	elseif write_exmode = 5 then
		Response.Write s_lang_0438
	elseif write_exmode = 6 then
		Response.Write s_lang_0439
	end if

	Response.Write "</font></td></tr>"
end function


function create_label_str(nid, mailid, ret_title, ret_color, is_own_lb)
	isolb = "1"
	if is_own_lb = false then
		isolb = "0"
	end if
	create_label_str = "<span id=""lba_" & mailid & "_" & nid & """><span id=""lb_" & mailid & "_" & nid & """ class=""wwm_lb_box"" style=""background:#" & ret_color & ";"" onclick=""showlabel('" & nid & "', event)"" onmouseover=""up_label('" & mailid & "', '" & nid & "', " & isolb & ")"" onmouseout=""out_label('" & mailid & "', '" & nid & "', " & isolb & ")""><span class=""wwm_lb_text"">" & server.htmlencode(ret_title) & "</span><span id=""lbc_" & mailid & "_" & nid & """ class=""wwm_lb_close"" onclick=""lbclose('" & mailid & "', '" & nid & "', event, " & isolb & ")"">&nbsp;</span></span></span>"
end function


function getLength(txt)
	txt=trim(txt)
	x = len(txt)
	y = 0
	for ii = 1 to x
		if asc(mid(txt,ii,1))<0 or asc(mid(txt,ii,1))>255 then
			y = y + 2
		else
			y = y + 1
		end if
	next
	getLength= y
end function


set mlb = nothing
set ei = nothing
set exMailInfo = nothing
%>
