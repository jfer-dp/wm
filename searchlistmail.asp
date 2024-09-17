<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
dim exMailInfo
set exMailInfo = server.createobject("easymail.ExMailInfo")

dim ei
set ei = server.createobject("easymail.InfoList")

dim addsortstr

sortstr = request("sortstr")
sortmode = request("sortmode")

if sortmode = 1 then
	addsortstr = "&sortstr=" & sortstr & "&sortmode=1"
	sortmode = true
else
	addsortstr = "&sortstr=" & sortstr & "&sortmode=0"
	sortmode = false
end if

if sortstr <> "" then
	ei.SetSort sortstr, sortmode
else
	sortstr = "Date"
end if

dim mlb
set mlb = server.createobject("easymail.Labels")
mlb.Load Session("wem")

'-----------------------------------------
dim pf
set pf = server.createobject("easymail.PerFolders")
pf.Load Session("wem")

if trim(request("mailsearch")) <> "" then
	Session("SearchStr") = trim(request("mailsearch"))
end if

ei.searchstring = Session("SearchStr")

ei.LoadMailBox Session("wem"), "empty"

allnum = ei.getMailsCount

if request("page") = "" then
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

dim bottom_bar

dim is_already_show
is_already_show = false

gourl = "searchlistmail.asp?page=" & page & "&" & getGRSN()
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
<link rel="stylesheet" type="text/css" href="images/searchlistmail.css">

<STYLE type=text/css>
<!--
body {padding-top:4px;}
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

<script type="text/javascript">
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
		location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=0";
	else
<% if sortmode = false then %>
		location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=1";
<% else %>
		location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=0";
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

function showmail(s_url) {
	parent.f1.document.leftval.purl.value = "<%=gourl & addsortstr %>";
	location.href = "showmail.asp?" + s_url;
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
<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then
%>
	ck_array = [];
	ck_array.push(tgli);

	var post_date = "setmode=4&<%=getGRSN() %>&bj=" + bj + "&maxck=" + (tgli + 1) + "&ck_" + tgli + "=" + document.getElementById("ck_" + tgli).value;
	send_star(post_date);
<%
end if
%>
}

function set_star(bj)
{
	var ck_date = get_sel_check();
	if (ck_date.length < 1)
		return ;

	var post_date = "setmode=4&<%=getGRSN() %>&bj=" + bj + "&" + ck_date;
	send_star(post_date);
}

function send_star(post_date)
{
$.ajax({
	type:"POST",
	url:"listmail.asp",
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
	var ck_date = get_sel_check();
	if (ck_date.length < 1)
		return ;

	var post_date = "setmode=3&<%=getGRSN() %>&bj=" + bj + "&" + ck_date;

$.ajax({
	type:"POST",
	url:"listmail.asp",
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

function lbclose(tgli, tglbid, e, isolb)
{
	if (isolb == "0")
		return ;

	if (!e) var e = window.event;
	e.cancelBubble = true;
	if (e.stopPropagation)
	e.stopPropagation();

	var post_date = "setmode=2&<%=getGRSN() %>&lbid=" + tglbid + "&mailid=" + document.getElementById("ck_" + tgli).value;

$.ajax({
	type:"POST",
	url:"listmail.asp",
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
	var ck_date = get_sel_check();
	if (ck_date.length < 1)
		return ;

	var post_date = "setmode=1&<%=getGRSN() %>&lbid=" + tgid + "&" + ck_date;
	var theObj;

$.ajax({
	type:"POST",
	url:"listmail.asp",
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
		document.f1.gourl.value = "<%=gourl & addsortstr %>";
		document.f1.mode.value = "del";
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

function move2arc() {
	if (ischeck() == true)
	{
		if (confirm("<%=s_lang_0572 %>") == false)
			return ;

		document.f1.gourl.value = "<%=gourl & addsortstr %>";
		document.f1.mode.value = "m2arc";
		document.f1.submit();
	}
}

function realdel() {
	if (ischeck() == true)
	{
		if (confirm("<%=s_lang_0399 %>") == false)
			return ;

		document.f1.gourl.value = "<%=gourl & addsortstr %>";
		document.f1.mode.value = "del";
		document.f1.isremove.value = "1";
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
	document.f1.mode.value = "move";
	document.f1.submit();
}

function allcheck_onclick() {
	document.body.focus();
	return false;
}

function selectpage_onchange() {
	location.href = "searchlistmail.asp?<%="mode=" & Server.URLEncode(mb) & addsortstr & "&" & getGRSN() %>&page=" + document.f1.page.value;
}
//-->
</SCRIPT>

<BODY>
<FORM ACTION="mulmail.asp" METHOD="POST" name="f1" id="f1">
<INPUT NAME="mode" TYPE="hidden">
<INPUT NAME="mto" TYPE="hidden">
<INPUT NAME="gourl" TYPE="hidden">
<table id="table_main" class="table_main" align="center" cellspacing="0" cellpadding="0">
  <tr>
	<td colspan="7" class="box_title_td"><%=s_lang_0445 %><font color="#901111"><%=allnum %></font><%=s_lang_0446 %>
	</td></tr>
	<tr><td class="block_top_td" colspan="7"><div class="table_min_width"></div></td></tr>
	<tr><td colspan="7" class="tool_top_td">

<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then
%>
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
<%
end if
%>

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
	bottom_bar = "<a href=""searchlistmail.asp?mode=" & Server.URLEncode(mb) & "&page=" & 0 & addsortstr & "&" & getGRSN() & """><img src='images/firstp.gif' border='0' align='absmiddle'></a>&nbsp;"
	bottom_bar = bottom_bar & "<a href=""searchlistmail.asp?mode=" & Server.URLEncode(mb) & "&page=" & page - 1 & addsortstr & "&" & getGRSN() & """><img src='images/prep.gif' border='0' align='absmiddle'></a>&nbsp;"
	Response.Write "<a href=""searchlistmail.asp?mode=" & Server.URLEncode(mb) & "&page=" & 0 & addsortstr & "&" & getGRSN() & """><img src='images/firstp.gif' border='0' align='absmiddle'></a>&nbsp;"
	Response.Write "<a href=""searchlistmail.asp?mode=" & Server.URLEncode(mb) & "&page=" & page - 1 & addsortstr & "&" & getGRSN() & """><img src='images/prep.gif' border='0' align='absmiddle'></a>&nbsp;"
end if

if ((page+1) * pageline) => allnum then
	bottom_bar = bottom_bar & "<img src='images/gnextp.gif' border='0' align='absmiddle'>&nbsp;"
	Response.Write "<img src='images/gnextp.gif' border='0' align='absmiddle'>&nbsp;"
else
	bottom_bar = bottom_bar & "<a href=""searchlistmail.asp?mode=" & Server.URLEncode(mb) & "&page=" & page + 1 & addsortstr & "&" & getGRSN() & """><img src='images/nextp.gif' border='0' align='absmiddle'></a>&nbsp;"
	Response.Write "<a href=""searchlistmail.asp?mode=" & Server.URLEncode(mb) & "&page=" & page + 1 & addsortstr & "&" & getGRSN() & """><img src='images/nextp.gif' border='0' align='absmiddle'></a>&nbsp;"
end if

if page + 1 >= allpage then
	bottom_bar = bottom_bar & "<img src='images/gendp.gif' border='0' align='absmiddle'>"
	Response.Write "<img src='images/gendp.gif' border='0' align='absmiddle'>"
else
	bottom_bar = bottom_bar & "<a href=""searchlistmail.asp?mode=" & Server.URLEncode(mb) & "&page=" & allpage - 1 & addsortstr & "&" & getGRSN() & """><img src='images/endp.gif' border='0' align='absmiddle'></a>"
	Response.Write "<a href=""searchlistmail.asp?mode=" & Server.URLEncode(mb) & "&page=" & allpage - 1 & addsortstr & "&" & getGRSN() & """><img src='images/endp.gif' border='0' align='absmiddle'></a>"
end if
%></span>
	</td></tr>
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
		<a href="javascript:setsort('Sender')"><%=s_lang_0147 %></a>&nbsp;<%
if sortstr = "Sender" then
 	if sortmode = true then
		response.write "<a href=""javascript:setsort('Sender')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		response.write "<a href=""javascript:setsort('Sender')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%>
	</td>
	<td class="st_5">
		<a href="javascript:setsort('Subject')"><%=s_lang_0127 %></a>&nbsp;<%
if sortstr = "Subject" then
 	if sortmode = true then
		response.write "<a href=""javascript:setsort('Subject')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		response.write "<a href=""javascript:setsort('Subject')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%>
	</td>
	<td class="st_6">
		<a href="javascript:setsort('Date')"><%=s_lang_0128 %></a>&nbsp;<%
if sortstr = "" or sortstr = "Date" then
 	if sortmode = true then
		response.write "<a href=""javascript:setsort('Date')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		response.write "<a href=""javascript:setsort('Date')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%>
	</td>
	<td class="st_7">
		<a href="javascript:setsort('Size')"><%=s_lang_0179 %></a>&nbsp;<%
if sortstr = "Size" then
 	if sortmode = true then
		response.write "<a href=""javascript:setsort('Size')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		response.write "<a href=""javascript:setsort('Size')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%>
	</td>
	</tr>
<%
i = page * pageline
li = 0

do while i < allnum and li < pageline
	ei.getMailInfoEx allnum - i - 1, idname, isread, priority, sendMail, sendName, subject, size, etime, mstate

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

	show_mail_function = " onclick=""showmail('filename=" & idname & "&" & getGRSN() & "&gourl=" & Server.URLEncode(gourl & addsortstr) & "')"""

	is_already_show = true
%>
	<tr id="tr_<%=li %>" class="cont_tr<% if isread = false then Response.Write "_b" %>" onmouseover='m_over(this);' onmouseout='m_out(this);'>
	<td class="cont_td_1"><%=xmsp %></td>
	<td class="cont_td_2"><%
if mstate = 0 then
	Response.Write "<img id='icn_b_" & li & "' src='reply.gif' title='" & s_lang_0427 & "' border='0'"
elseif mstate = 1 then
	Response.Write "<img id='icn_b_" & li & "' src='forward.gif' title='" & s_lang_0428 & "' border='0'"
elseif mstate = 2 or mstate = 3 then
	if isread = true then
		Response.Write "<img id='icn_b_" & li & "' src='rsysmail.gif' title='" & s_lang_0181 & "' border='0'"
	else
		Response.Write "<img id='icn_b_" & li & "' src='nsysmail.gif' title='" & s_lang_0429 & "' border='0'"
	end if
else
	if isread = true then
		Response.Write "<img id='icn_b_" & li & "' src='mail.gif' title='" & s_lang_0420 & "' border='0'"
	else
		Response.Write "<img id='icn_b_" & li & "' src='newmail.gif' title='" & s_lang_0421 & "' border='0'"
	end if
end if

if issign = true then
	if isread = true then
		Response.Write "><img id='icn_e_" & li & "' src='images/s0-1.gif' title='" & s_lang_0183 & "' border='0'"
	else
		Response.Write "><img id='icn_e_" & li & "' src='images/s0.gif' title='" & s_lang_0183 & "' border='0'"
	end if
elseif isenc = true then
	if isread = true then
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
Response.Write server.htmlencode(subject) & "</span>"

Response.Write "<span id='sp_lb_" & li & "' style='float:right; display:inline-block;'>"

if exMailInfo.LabelCount > 0 then
	lball = exMailInfo.LabelCount
	lbi = 0
	do while lbi < lball
		mlb.GetByID exMailInfo.GetLabel(lbi), ret_id, ret_title, ret_color

		if lbi > 0 then
			Response.Write "<span id='sp_lb_bk_" & li & "_" & ret_id & "' style='float:right; width:3px;'><font style='font-size:1px;'>&nbsp;</font></span>"
		end if

		Response.Write create_label_str(ret_id, li, ret_title, ret_color, true) & Chr(13)
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

if exMailInfo.Have_Star = true then
	Response.Write "</span><img id='icn_star_" & li & "' src='images/star_yes.gif' border='0' style='cursor:pointer;' onclick=""set_one_star(" & li & ", '0');""></a>"
else
	Response.Write "</span><img id='icn_star_" & li & "' src='images/star_no.gif' border='0' style='cursor:pointer;' onclick=""set_one_star(" & li & ", '1');""></a>"
end if
%></td>
	</tr>
<%
	idname = NULL
	isread = NULL
	priority = NULL
	sendMail = NULL
	sendName = NULL
	subject = NULL
	size = NULL
	etime = NULL
	mstate = NULL

	i = i + 1
	li = li + 1
loop

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
</FORM>

<div id="pmc_moveto" class="qmpanel_shadow" style="display:none; position:absolute;">
	<div class="menu_base">
	<div id="md_moveto" class="menu_bd bd">
	<table width="100%">
		<div name="mi" onclick="move('del');" class="menu_item"><%=s_lang_0334 %></div>
		<div name="mi" onclick="move('in');" class="menu_item"><%=s_lang_0327 %></div>
		<div name="mi" onclick="move('out');" class="menu_item"><%=s_lang_0332 %></div>
		<div name="mi" onclick="move('sed');" class="menu_item"><%=s_lang_0430 %></div>
		<div class="menu_item_nofun"><div style="background:#ccc; padding-top:1px; margin-top: 5px;"></div></div>
		<div name="mi" onclick="move2arc();" class="menu_item"><%=s_lang_0571 %></div>
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

	Response.Write "<div name='mi' onclick=""move('" & server.htmlencode(spfname) & "');"" class='menu_item'>" & server.htmlencode(spfname) & "</div>"
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

function get_sel_check()
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
			}
		}
	}

	if (is_check == false)
		ret_val = "";

	return ret_val;
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
set mlb = nothing
set ei = nothing
set exMailInfo = nothing


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

function create_label_str(nid, mailid, ret_title, ret_color, is_own_lb)
	isolb = "1"
	if is_own_lb = false then
		isolb = "0"
	end if

	if isadmin() = true or Session("ReadOnlyUser") <> 1 then
		create_label_str = "<span id=""lba_" & mailid & "_" & nid & """><span id=""lb_" & mailid & "_" & nid & """ class=""wwm_lb_box"" style=""background:#" & ret_color & ";"" onclick=""showlabel('" & nid & "', event)"" onmouseover=""up_label('" & mailid & "', '" & nid & "', " & isolb & ")"" onmouseout=""out_label('" & mailid & "', '" & nid & "', " & isolb & ")""><span class=""wwm_lb_text"">" & server.htmlencode(ret_title) & "</span><span id=""lbc_" & mailid & "_" & nid & """ class=""wwm_lb_close"" onclick=""lbclose('" & mailid & "', '" & nid & "', event, " & isolb & ")"">&nbsp;</span></span></span>"
	else
		create_label_str = "<span id=""lba_" & mailid & "_" & nid & """><span id=""lb_" & mailid & "_" & nid & """ class=""wwm_lb_box"" style=""background:#" & ret_color & ";""><span class=""wwm_lb_text"">" & server.htmlencode(ret_title) & "</span><span id=""lbc_" & mailid & "_" & nid & """ class=""wwm_lb_close"">&nbsp;</span></span></span>"
	end if
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
%>
