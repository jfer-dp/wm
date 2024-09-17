<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
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

'-----------------------------------------
dim pf
set pf = server.createobject("easymail.PerFolders")
pf.Load Session("wem")

t_mailsearch = trim(request("mailsearch"))

if Len(t_mailsearch) > 0 then
	if LCase(t_mailsearch) = "null" then
		Session("SearchStr") = ""
	else
		Session("SearchStr") = t_mailsearch
		ei.searchstring = Session("SearchStr")
	end if
else
	ei.searchstring = Session("SearchStr")
end if

str_date = trim(request("date"))
ei.LoadArchiveBox Session("wem"), str_date

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

gourl = "listarc.asp?date=" & str_date & "&page=" & page & "&" & getGRSN()
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
<link rel="stylesheet" type="text/css" href="images/searchlistmail.css">
<link rel="stylesheet" type="text/css" href="images/popwin.css">

<STYLE type=text/css>
<!--
body {padding-top:4px;}
.st_1 {width:3%;}
.st_2 {width:5%;}
.st_3 {width:3%;}
.st_4 {width:17%;}
.st_5 {width:42%;}
.st_6 {width:21%;}
.st_7 {width:9%; border-right:1px solid #c1c8d2;}
-->
</STYLE>
</HEAD>

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
	location.href = "showarc.asp?" + s_url;
}

function del() {
	if (ischeck() == true)
	{
		if (confirm("<%=s_lang_0399 %>") == false)
			return ;

		document.f1.gourl.value = "<%=gourl & addsortstr %>";
		document.f1.mode.value = "arcdel";
		document.f1.submit();
	}
}

function m_return() {
	location.href = "showarchive.asp?<%=getGRSN() %>";
}

function sfind() {
	pop_show();
}

function findit() {
	if (document.getElementById('stext').value.length < 1)
		document.f1.mailsearch.value = "null";
	else
		document.f1.mailsearch.value = "\t\t" + document.getElementById('shead').value + "\t" + document.getElementById('smode').value + "\t" + document.getElementById('stext').value + "\t\tRecDate\t1\t" + get_date() + "\t\tSize\t1\t\t\tRead\t1\t1\t\tFolders\t\t";

	document.f1.action = "listarc.asp?<%=getGRSN() %>";
	document.f1.submit();
}

function get_date() {
	var mydate = new Date();
	var m_str;
	var d_str;

	if (mydate.getMonth() > 8)
		m_str = (mydate.getMonth() + 1).toString();
	else
		m_str = "0" + (mydate.getMonth() + 1).toString();

	if (mydate.getDate() > 9)
		d_str = mydate.getDate().toString();
	else
		d_str = "0" + mydate.getDate().toString();

	var fyear = mydate.getFullYear() + 1;
	return fyear.toString() + m_str + d_str;
}

function move(tgname) {
	if (ischeck() == false)
	{
		document.f1.mto.value = "";
		return ;
	}

	document.f1.mto.value = tgname;
	document.f1.gourl.value = "<%=gourl & addsortstr %>";
	document.f1.mode.value = "arc2m";
	document.f1.submit();
}

function allcheck_onclick() {
	document.body.focus();
	return false;
}

function selectpage_onchange() {
	location.href = "listarc.asp?date=<%=str_date & "&mode=" & Server.URLEncode(mb) & addsortstr & "&" & getGRSN() %>&page=" + document.f1.page.value;
}
//-->
</SCRIPT>

<BODY>
<FORM ACTION="mulmail.asp" METHOD="POST" name="f1" id="f1">
<INPUT NAME="mode" TYPE="hidden">
<INPUT NAME="mto" TYPE="hidden">
<INPUT NAME="gourl" TYPE="hidden">
<INPUT NAME="mailsearch" TYPE="hidden">
<INPUT NAME="date" TYPE="hidden" value="<%=str_date %>">
<table id="table_main" class="table_main" align="center" cellspacing="0" cellpadding="0">
  <tr>
	<td colspan="7" class="box_title_td"><font class="font_top_title"><%
temp_y = Left(str_date , 4)
temp_m = get_month(str_date)
temp_date_showstr = temp_y & s_lang_0581 & temp_m & s_lang_0582

Response.Write temp_date_showstr
%></font>&nbsp;
<%
if Len(Session("SearchStr")) > 1 then
%>
	<%=s_lang_0445 %><font color="#901111"><%=allnum %></font><%=s_lang_0446 %>
<%
else
	Response.Write s_lang_0401 & " <font color='#901111'>" & allnum & "</font>" & s_lang_0423
end if
%>
	</td></tr>
	<tr><td class="block_top_td" colspan="7"><div class="table_min_width"></div></td></tr>
	<tr><td colspan="7" class="tool_top_td">

<span class="st_span">
<a class='wwm_btnDownload btn_gray' href='javascript:m_return()'><< <%=s_lang_return %></a>
</span>
<span style='float:left; width:3px;'>&nbsp;</span>

<span class="st_span">
<a class='wwm_btnDownload btn_gray' href='javascript:del()'><%=s_lang_del %></a>
</span>
<span style='float:left; width:3px;'>&nbsp;</span>

<span class="st_span">
<a class='wwm_btnDownload btn_gray' href='javascript:sfind()'><%=s_lang_find %></a>
</span>
<span style='float:left; width:3px;'>&nbsp;</span>

<span class="st_span"><span id="pm_moveto" class="menu_pop">
<div class='attbg'><img style='margin: 6px 0pt 0pt;' src='images/popshow.gif'></div>
<div class='menu_pop_text'><%=s_lang_0404 %>...</div>
</span></span>
<span style='float:left; width:3px;'>&nbsp;</span>

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
	bottom_bar = "<a href=""listarc.asp?date=" & str_date & "&mode=" & Server.URLEncode(mb) & "&page=" & 0 & addsortstr & "&" & getGRSN() & """><img src='images/firstp.gif' border='0' align='absmiddle'></a>&nbsp;"
	bottom_bar = bottom_bar & "<a href=""listarc.asp?date=" & str_date & "&mode=" & Server.URLEncode(mb) & "&page=" & page - 1 & addsortstr & "&" & getGRSN() & """><img src='images/prep.gif' border='0' align='absmiddle'></a>&nbsp;"
	Response.Write "<a href=""listarc.asp?date=" & str_date & "&mode=" & Server.URLEncode(mb) & "&page=" & 0 & addsortstr & "&" & getGRSN() & """><img src='images/firstp.gif' border='0' align='absmiddle'></a>&nbsp;"
	Response.Write "<a href=""listarc.asp?date=" & str_date & "&mode=" & Server.URLEncode(mb) & "&page=" & page - 1 & addsortstr & "&" & getGRSN() & """><img src='images/prep.gif' border='0' align='absmiddle'></a>&nbsp;"
end if

if ((page+1) * pageline) => allnum then
	bottom_bar = bottom_bar & "<img src='images/gnextp.gif' border='0' align='absmiddle'>&nbsp;"
	Response.Write "<img src='images/gnextp.gif' border='0' align='absmiddle'>&nbsp;"
else
	bottom_bar = bottom_bar & "<a href=""listarc.asp?date=" & str_date & "&mode=" & Server.URLEncode(mb) & "&page=" & page + 1 & addsortstr & "&" & getGRSN() & """><img src='images/nextp.gif' border='0' align='absmiddle'></a>&nbsp;"
	Response.Write "<a href=""listarc.asp?date=" & str_date & "&mode=" & Server.URLEncode(mb) & "&page=" & page + 1 & addsortstr & "&" & getGRSN() & """><img src='images/nextp.gif' border='0' align='absmiddle'></a>&nbsp;"
end if

if page + 1 >= allpage then
	bottom_bar = bottom_bar & "<img src='images/gendp.gif' border='0' align='absmiddle'>"
	Response.Write "<img src='images/gendp.gif' border='0' align='absmiddle'>"
else
	bottom_bar = bottom_bar & "<a href=""listarc.asp?date=" & str_date & "&mode=" & Server.URLEncode(mb) & "&page=" & allpage - 1 & addsortstr & "&" & getGRSN() & """><img src='images/endp.gif' border='0' align='absmiddle'></a>"
	Response.Write "<a href=""listarc.asp?date=" & str_date & "&mode=" & Server.URLEncode(mb) & "&page=" & allpage - 1 & addsortstr & "&" & getGRSN() & """><img src='images/endp.gif' border='0' align='absmiddle'></a>"
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
%>></td>
	<td class="cont_td_3"><input type="checkbox" id="ck_<%=li %>" name="ck_<%=li %>" value="<%=idname %>" onclick="ck_select(this);"></td>
	<td class="cont_td_4"<%=show_mail_function %>><%=server.htmlencode(sendName) %>&nbsp;</td>
	<td id="td_subject_<%=li %>" class="cont_td_5"<%=show_mail_function %>><span class="cs_subject"><%
Response.Write server.htmlencode(subject) & "</span>"
%></td>
	<td class="cont_td_6"<%=show_mail_function %>><%=etime %></td>
	<td class="cont_td_7"><span class="cs_star"><%
Response.Write getShowSize(size)
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

	Response.Write "<div name='mi' onclick=""move('" & server.htmlencode(spfname) & "');"" class='menu_item'>" & server.htmlencode(spfname) & "</div>"
	spfname = NULL

	i = i + 1
loop
%>
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

<div id="pop_overlay">
</div>

<div id="pop_win" style="display:none; position:absolute;" class="mydiv">
	<div class="pop_base"><div class="pop_bd bd"><div class="title">
		<div class="title_left"><%=s_lang_find & " - " & temp_date_showstr %></div>
		<div class="title_right" title="<%=s_lang_close %>" id="pop_close_wind"><span>&nbsp;</span></div>
	</div>
	<div class="pop_content" style="padding-top:16px; padding-bottom:16px;">

<table width="90%" border="0" align="center" cellspacing="0">
	<tr><td nowrap>
	<select id="shead" size="1">
		<option value="Subject" selected><%=s_lang_0573 %></option>
		<option value="FromMail"><%=s_lang_0574 %></option>
		<option value="FromName"><%=s_lang_0575 %></option>
	</select>&nbsp;&nbsp;
	<select id="smode" size="1">
		<option value="1" selected><%=s_lang_0576 %></option>
		<option value="2"><%=s_lang_0577 %></option>
		<option value="3"><%=s_lang_0578 %></option>
	</select>
	</td></tr>
	<tr><td>
	<input type="text" id="stext" size="36" maxlength="60" class='b_input'>
	</td></tr>
</table>

	</div>
	<div class="title_bottom">
	<div class="title_ok_cancel_div">
	<a id="pop_ok" class="wwm_btnDownload btn_gray" href="#" onclick="javascript:findit();"><%=s_lang_find %></a>&nbsp;
	<a class="wwm_btnDownload btn_gray" href="#" onclick="javascript:pop_close()"><%=s_lang_cancel %></a>
	</div></div></div></div>
</div>

<script type="text/javascript">
var doc = document.documentElement;
var body = document.body;
var oWin;
var oLay;
var oClose;

oWin = document.getElementById("pop_win");
oLay = document.getElementById("pop_overlay");
oClose = document.getElementById("pop_close_wind");

oClose.onclick = function ()
{
	pop_close();
}

function pop_show()
{
	oLay.style.height = document.documentElement.scrollHeight + "px";
	oLay.style.width = document.documentElement.scrollWidth + "px";

	var ie_h = doc && doc.clientHeight  || body && body.clientHeight  || 0;
	var ie_w = doc && doc.clientWidth  || body && body.clientWidth  || 0;

	if (ie_h > document.documentElement.scrollHeight)
		oLay.style.height = ie_h + "px";

	if (ie_w > document.documentElement.scrollWidth)
		oLay.style.width = ie_w + "px";

	oLay.style.display = "block";
	oWin.style.display = "block"	

	document.getElementById('stext').focus();
}

function pop_close()
{
	oLay.style.display = "none";
	oWin.style.display = "none"	
}

var g_newname;

function set_name()
{
	g_newname = document.getElementById("sender_name").value;
	g_newname = g_newname.replace(/\'/g,"");
	g_newname = g_newname.replace(/\"/g,"");

	if (g_newname.length < 1)
	{
		alert("<%=b_lang_193 %>.");
		document.getElementById("sender_name").value = "";
		document.getElementById('sender_name').focus();
		return ;
	}
}


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
set ei = nothing


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

function get_month(date_str)
	if Len(date_str) = 6 then
		tmp_month = Mid(date_str, 5, 2)
		if Mid(tmp_month, 1, 1) = "0" then
			tmp_month = Mid(tmp_month, 2, 1)
		else
			tmp_month = Mid(tmp_month, 1, 2)
		end if

		get_month = tmp_month
	else
		get_month = ""
	end if
end function
%>
