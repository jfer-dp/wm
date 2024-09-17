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

dim ei
set ei = server.createobject("easymail.InfoList")

mb = trim(request("mb"))

if mb = "" then
	mb = "att"
end if

dim isfts
isfts = false

if mb <> "att" then
	ei.IsAttFolder = true
else
	if trim(request("delfts")) <> "1" then
		if trim(request("isfts")) = "1" then
			ei.isLoadZatt = true
			isfts = true
		end if
	end if
end if

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

if isfts = true then
	addsortstr = addsortstr & "&isfts=1"
end if

if sortstr <> "" then
	ei.SetSort sortstr, sortmode
else
	sortstr = "Date"
end if

'-----------------------------------------
username = Session("wem")
ei.LoadMailBox username, mb


''''''''''''''''''''''''''''''''
sname = trim(request("sname"))
sfname = trim(request("sfname"))

if sname <> "" and sfname <> "" then
	if Application("em_Enable_ShareFolder") = true then
		openresult = ei.LoadFriendMailBox(username, sname, sfname, ei.IsAttFolder)
	else
		openresult = -1
	end if

	if openresult = -1 then
		set ei = nothing
		Response.Redirect "err.asp?errstr=" & s_lang_0396
	elseif  openresult = 1 then
		set ei = nothing
		Response.Redirect "err.asp?errstr=" & s_lang_0397
	elseif  openresult = 2 then
		set ei = nothing
		Response.Redirect "err.asp?errstr=" & s_lang_0398
	end if
end if

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

gourl = "listatt.asp?page=" & page & "&" & getGRSN() & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&mb=" & Server.URLEncode(mb)

Response.Cookies("name") = Session("wem")
Response.Cookies("attfoldername") = mb

if sname = "" or sfname = "" then
	dim pf
	set pf = server.createobject("easymail.PerAttFolders")
	pf.Load Session("wem")
end if

dim is_already_show
is_already_show = false
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
<link rel="stylesheet" type="text/css" href="images/listatt.css">

<STYLE type=text/css>
<!--
body {padding-top:2px;}
.Bsbttn {font-family:<%=s_lang_font %>; font-size:10pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5; color:#000066;text-decoration:none;cursor:pointer}
.sbttn {font-family:<%=s_lang_font %>; font-size:9pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5;color:#000066;height:19px;text-decoration:none;cursor:pointer}
.st_1 {width:5%;}
.st_2 {width:3%;}
.st_3 {width:29%;}
.st_4 {width:30%;}
.st_5 {width:20%;}
.st_6 {width:5%;}
.st_7 {width:8%;}
-->
</STYLE>
</HEAD>

<script type="text/javascript" src="images/mglobal.js"></script>
<script type="text/javascript" src="images/jquery.min.js"></script>
<script type="text/javascript" src="images/jquery-powerFloat-min.js"></script>

<SCRIPT LANGUAGE=javascript>
<!--
if (ie == 6)
	document.execCommand("BackgroundImageCache", false, true); 

if (parent.f1.document.getElementById("leftval") != null)
{
<%
if Len(sname) > 0 then
%>
	parent.f1.select_one("/ff_showall.asp", "");
<%
else
%>
	parent.f1.select_one(document.location.href, "<%=mb %>");
<%
end if
%>
}

var before_changed = -1;
var before_fname = "";
var before_emsg = "";
var before_isnew = false;

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

function m_over(tag_obj) {
	var theObj = document.getElementById("ck_" + tag_obj.id.substr(3));
	if (theObj != null)
	{
		if (theObj.checked == false)
			tag_obj.style.backgroundColor = "#ecf9ff";
	}
	else
		tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj) {
	var theObj = document.getElementById("ck_" + tag_obj.id.substr(3));
	if (theObj != null)
	{
		if (theObj.checked == false)
			tag_obj.style.backgroundColor = "white";
	}
	else
		tag_obj.style.backgroundColor = "white";
}

function showatt(s_url) {
	window.open(s_url, "");
}

function showzatt() {
	location.href = "listatt.asp?mb=<%=Server.URLEncode(mb) %>&<%="sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & addsortstr & "&" & getGRSN() %>&page=" + fsa.page.value + "<%
if isfts = false then
	Response.Write "&isfts=1"
else
	Response.Write "&delfts=1"
end if
%>";
}

function allcheck_onclick() {
	document.body.focus();
	return false;
}

function stop_click(e) {
	if (!e) var e = window.event;
	e.cancelBubble = true;
	if (e.stopPropagation)
	e.stopPropagation();
}

function only_check(tag_obj, e) {
	if (!e) var e = window.event;
	e.cancelBubble = true;
	if (e.stopPropagation)
	e.stopPropagation();

	ck_select(tag_obj)
}

function del() {
	if (ischeck() == true)
	{
		if (confirm("<%=s_lang_0399 %>") == false)
			return ;

		document.f1.action = "mulmail.asp";
		document.f1.gourl.value = "<%=gourl & addsortstr %>";
		document.f1.isremove.value = "1";
		document.f1.mode.value = "del";
		document.f1.submit();
	}
}
<%
if mb = "att" then
%>
function delallatt() {
<%
	if allnum > 0 then
%>
	if (confirm("<%=s_lang_0400 %>") == false)
		return ;

	document.f1.action = "mulmail.asp?mode=cleanAtt";
	document.f1.gourl.value = "<%=gourl & addsortstr %>";
	document.f1.submit();
<%
	end if
%>
}
<%
end if
%>

function move(tg_apf_name) {
	if (ischeck() == true)
	{
		document.f1.action = "mulmail.asp";
		document.f1.gourl.value = "<%=gourl & addsortstr %>";
		document.f1.mto.value = tg_apf_name;
		document.f1.isatt.value = "1";
		document.f1.mode.value = "move";
		document.f1.submit();
	}
}

function selectpage_onchange()
{
	location.href = "listatt.asp?mb=<%=Server.URLEncode(mb) %>&<%="sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & addsortstr & "&" & getGRSN() %>&page=" + fsa.page.value;
}

function uf_changed() {
	if (document.getElementById("upfile").value.length > 1)
	{
		document.getElementById("saving").style.display = "inline";
		document.fsa.submit();
	}
}

function save_com()
{
	if (ie == false)
	{
		document.f1.appendChild(emsg);
		document.f1.appendChild(efname);
	}

	document.f1.action = "savecom.asp";
	document.f1.gourl.value = "<%=gourl & addsortstr %>";
	document.f1.submit();
}

function editcom(isnew, tid, fname, emsg)
{
<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then
%>
	var bfObj;
	if (before_changed > -1)
	{
		bfObj = document.getElementById("id_" + before_changed);
		bfObj.innerHTML = "<a href=\"JavaScript:editcom(" + before_isnew + "," + before_changed + ", '" + before_fname + "', '" + before_emsg + "')\"><img src='images/pedit.gif' border=0 title='<%=s_lang_modify %>'></a>";

		if (before_isnew == false)
		{
			if (before_emsg.length > 0)
				bfObj.innerHTML = bfObj.innerHTML + "&nbsp;" + before_emsg;
		}
		else
		{
			if (before_emsg.length > 0)
				bfObj.innerHTML = bfObj.innerHTML + "&nbsp;" + "<b>" + before_emsg + "</b>";
		}
	}

	theObj = document.getElementById("id_" + tid);
	theObj.innerHTML = "<input type='hidden' id='efname' name='efname' value='" + fname + "'><input type='text' id='emsg' name='emsg' maxlength='128' class='textbox' value='" + emsg + "'>";
	theObj.innerHTML = theObj.innerHTML + "&nbsp;<input type='button' value='<%=s_lang_save %>' class='sbttn' LANGUAGE=javascript onclick='save_com()'>";

	document.body.focus();
	document.getElementById("emsg").focus();

	before_changed = tid;
	before_fname = fname;
	before_emsg = emsg;
	before_isnew = isnew;
<%
end if
%>
}

function saveatt(filename)
{
	if (filename.length > 0)
	{
		document.f1.action = "att2att.asp";
		document.f1.gourl.value = "<%=gourl & addsortstr %>";
		document.f1.filename.value = filename;
		document.f1.submit();
	}
}
//-->
</SCRIPT>

<BODY>
<table class="table_main" align="center" cellspacing="0" cellpadding="0">
<FORM ENCTYPE="multipart/form-data" ACTION="savenetatt.asp" METHOD=POST NAME="fsa">
	<tr>
	<td class="box_title_td">
<font class="font_top_title"><%
if sname <> "" and sfname <> "" then
	Response.Write sname & "\"
end if

if mb = "att" then
	Response.Write s_lang_0296
	if isfts = true then
		Response.Write "-" & s_lang_0292
	end if
else
	Response.Write server.htmlencode(mb)
end if
%></font><%
if allnum > 0 then
	Response.Write " " & s_lang_0401 & " <font color='#901111'>" & allnum & "</font>" & s_lang_0402
end if
%>
	</td></tr>
	<tr><td class="block_top_td"><div class="table_min_width"></div></td></tr>
	<tr>
    <td class="tool_top_td">

<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then
%>
<%
if sname = "" or sfname = "" then
%>
<span class="st_span">
<a class='wwm_btnDownload btn_gray' href='javascript:del()'><%=s_lang_del %></a>
</span>
<span style='float:left; width:3px;'>&nbsp;</span>

<%
	if mb = "att" then
%>
<span class="st_span">
<a class='wwm_btnDownload btn_gray' href='javascript:delallatt()'><%=s_lang_0403 %></a>
</span>
<span style='float:left; width:3px;'>&nbsp;</span>

<span class="st_span">
<a class='wwm_btnDownload btn_gray' href='javascript:showzatt()'><%
if isfts = false then
	Response.Write s_lang_0290
else
	Response.Write s_lang_0291
end if
%></a>
</span>
<span style='float:left; width:3px;'>&nbsp;</span>
<%
	end if
%>

<span class="st_span"><span id="pm_moveto" class="menu_pop">
<div class='attbg'><img style='margin: 6px 0pt 0pt;' src='images/popshow.gif'></div>
<div class='menu_pop_text'><%=s_lang_0404 %>...</div>
</span></span>
<span style='float:left; width:3px;'>&nbsp;</span>

<span class="st_span"<%
if isMobile = true then
	Response.Write " style='display:none;'"
end if
%>>
<a class="btn_addPic" href="javascript:void(0);"><span><em>+</em><%=s_lang_0359 %></span> <input class="filePrew" tabindex="3" id="upfile" name="upfile" size="3" type="file" onchange="javascript:uf_changed()"></a>
</span>

<%
else
%>
<span class="st_span">
<a class='wwm_btnDownload btn_gray' href="javascript:location.href='ff_showall.asp?<%=getGRSN() %>'"><%=s_lang_return %></a>
</span>
<%
end if
%>
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
	bottom_bar = "<a href=""listatt.asp?mb=" & Server.URLEncode(mb) & "&page=" & 0 & addsortstr & "&" & getGRSN() & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & """><img src='images/firstp.gif' border='0' align='absmiddle'></a>&nbsp;"
	bottom_bar = bottom_bar & "<a href=""listatt.asp?mb=" & Server.URLEncode(mb) & "&page=" & page - 1 & addsortstr & "&" & getGRSN() & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & """><img src='images/prep.gif' border='0' align='absmiddle'></a>&nbsp;"
	Response.Write "<a href=""listatt.asp?mb=" & Server.URLEncode(mb) & "&page=" & 0 & addsortstr & "&" & getGRSN() & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & """><img src='images/firstp.gif' border='0' align='absmiddle'></a>&nbsp;"
	Response.Write "<a href=""listatt.asp?mb=" & Server.URLEncode(mb) & "&page=" & page - 1 & addsortstr & "&" & getGRSN() & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & """><img src='images/prep.gif' border='0' align='absmiddle'></a>&nbsp;"
end if

if ((page+1) * pageline) => allnum then
	bottom_bar = bottom_bar & "<img src='images/gnextp.gif' border='0' align='absmiddle'>&nbsp;"
	Response.Write "<img src='images/gnextp.gif' border='0' align='absmiddle'>&nbsp;"
else
	bottom_bar = bottom_bar & "<a href=""listatt.asp?mb=" & Server.URLEncode(mb) & "&page=" & page + 1 & addsortstr & "&" & getGRSN() & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & """><img src='images/nextp.gif' border='0' align='absmiddle'></a>&nbsp;"
	Response.Write "<a href=""listatt.asp?mb=" & Server.URLEncode(mb) & "&page=" & page + 1 & addsortstr & "&" & getGRSN() & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & """><img src='images/nextp.gif' border='0' align='absmiddle'></a>&nbsp;"
end if

if page + 1 >= allpage then
	bottom_bar = bottom_bar & "<img src='images/gendp.gif' border='0' align='absmiddle'>"
	Response.Write "<img src='images/gendp.gif' border='0' align='absmiddle'>"
else
	bottom_bar = bottom_bar & "<a href=""listatt.asp?mb=" & Server.URLEncode(mb) & "&page=" & allpage - 1 & addsortstr & "&" & getGRSN() & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & """><img src='images/endp.gif' border='0' align='absmiddle'></a>"
	Response.Write "<a href=""listatt.asp?mb=" & Server.URLEncode(mb) & "&page=" & allpage - 1 & addsortstr & "&" & getGRSN() & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & """><img src='images/endp.gif' border='0' align='absmiddle'></a>"
end if
%></span>
	</td>
	</tr>
</FORM>
</table>

<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<FORM METHOD="POST" name="f1">
	<tr class="title_tr">
	<td class="st_1">
		<a href="javascript:setsort('Read')"><%=s_lang_0405 %></a><%
if sortstr = "Read" then
 	if sortmode = true then
		Response.Write "<a href=""javascript:setsort('Read')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "<a href=""javascript:setsort('Read')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%>
	</td>
<%
if sname = "" or sfname = "" then
%>
	<td class="st_2">
		<input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()">
	</td>
<%
end if
%>
	<td class="st_3">
		<a href="javascript:setsort('Subject')"><%=s_lang_0406 %></a>&nbsp;<%
if sortstr = "Subject" then
 	if sortmode = true then
		Response.Write "<a href=""javascript:setsort('Subject')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "<a href=""javascript:setsort('Subject')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%>
	</td>
	<td class="st_4">
		<a href="javascript:setsort('Sender')"><%=s_lang_0407 %></a>&nbsp;<%
if sortstr = "Sender" then
 	if sortmode = true then
		Response.Write "<a href=""javascript:setsort('Sender')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "<a href=""javascript:setsort('Sender')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%>
	</td>
	<td class="st_5">
		<a href="javascript:setsort('Date')"><%=s_lang_0408 %></a>&nbsp;<%
if sortstr = "" or sortstr = "Date" then
 	if sortmode = true then
		Response.Write "<a href=""javascript:setsort('Date')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "<a href=""javascript:setsort('Date')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%>
	</td>
<%
if sname <> "" and sfname <> "" then
%>
	<td class="st_7">
		<a href="javascript:setsort('Size')"><%=s_lang_0409 %></a><%
if sortstr = "Size" then
 	if sortmode = true then
		Response.Write "&nbsp;<a href=""javascript:setsort('Size')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "&nbsp;<a href=""javascript:setsort('Size')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%>
	</td>
	<td class="st_6" style="border-right:1px solid #c1c8d2;"><%=s_lang_save %>
<%
else
%>
	<td class="st_7" style="border-right:1px solid #c1c8d2;">
		<a href="javascript:setsort('Size')"><%=s_lang_0409 %></a><%
	if sortstr = "Size" then
	 	if sortmode = true then
			Response.Write "&nbsp;<a href=""javascript:setsort('Size')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
		else
			Response.Write "&nbsp;<a href=""javascript:setsort('Size')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
		end if
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

	if subject = "" then
		subject = s_lang_0410
	end if

	show_mail_function = " onclick=""showatt('showatt.asp?isattfolder=att&filename=" & Server.URLEncode(idname) & "&count=0&" & getGRSN() & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "');"""
	is_already_show = true
%>
	<tr id="tr_<%=li %>" class="cont_tr<% if isread = false then Response.Write "_b" %>" onmouseover='m_over(this);' onmouseout='m_out(this);'<%=show_mail_function %>>
	<td class="cont_td_1">
<%
if mstate = 5 then
%>
	<img src="newmail.gif" title="<%=s_lang_0411 %>" border=0>
<%
else
%>
	<img src="mail.gif" title="<%=s_lang_0412 %>" border=0>
<%
end if
%>
	</td>
<%
if sname = "" or sfname = "" then
%>
	<td class="cont_td_2" onclick="stop_click(event);"><input type="checkbox" id="ck_<%=li %>" name="ck_<%=li %>" value="<%=server.htmlencode(idname) %>" onclick="only_check(this, event);"></td>
<%
end if
%>
	<td class="cont_td_3"><%=server.htmlencode(subject) %>&nbsp;</td>
<%
if sname = "" or sfname = "" then
	if sendName = "" then
%>
	<td id="id_<%=allnum - i - 1 %>" class="cont_td_4" onclick="stop_click(event);"><a href="JavaScript:editcom(false, <%=allnum - i - 1 %>, '<%=server.htmlencode(idname) %>', '')"><img src='images/pedit.gif' border=0 title='<%=s_lang_modify %>'></a></td>
<%
	else
%>
	<td id="id_<%=allnum - i - 1 %>" class="cont_td_4" onclick="stop_click(event);"><a href="JavaScript:editcom(false, <%=allnum - i - 1 %>, '<%=server.htmlencode(idname) %>', '<%=server.htmlencode(sendName) %>')"><img src='images/pedit.gif' border=0 title='<%=s_lang_modify %>'></a>&nbsp;<%=server.htmlencode(sendName) %></td>
<%
	end if
else
	if sendName = "" then
%>
	<td id="id_<%=allnum - i - 1 %>" class="cont_td_5">&nbsp;</td>
<%
	else
%>
	<td id="id_<%=allnum - i - 1 %>" class="cont_td_5"><%=server.htmlencode(sendName) %></td>
<%
	end if
end if
%>
	<td class="cont_td_5"><%=etime %></td>
	<td class="cont_td_7"><%=getShowSize(size) %></td>
<%
if sname <> "" and sfname <> "" then
%>
	<td class="cont_td_6" onclick="stop_click(event);"><a href="JavaScript:saveatt('<%=server.htmlencode(idname) %>')")><img src='images/saveatt.gif' border='0' align='absmiddle' title='<%=s_lang_0413 %>'></a></td>
<%
end if
%>
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
<tr><td class="block_td" colspan="8"><div class="table_min_width"></div></td></tr>
<tr><td colspan="8" class="tool_td" style="height:22px; text-align:center;">
<%=bottom_bar %>
</td></tr>
<%
end if
%>
<INPUT NAME="isremove" TYPE="hidden" value="0">
<INPUT NAME="isatt" TYPE="hidden" value="1">
<INPUT NAME="mto" TYPE="hidden">
<INPUT NAME="mode" TYPE="hidden">
<INPUT NAME="gourl" TYPE="hidden">
<INPUT NAME="filename" TYPE="hidden">
<INPUT NAME="sname" TYPE="hidden" value="<%=server.htmlencode(sname) %>">
<INPUT NAME="sfname" TYPE="hidden" value="<%=server.htmlencode(sfname) %>">
</FORM>
</table>
<div id="saving" class="wwm_msg" style="position:absolute; top:35%; left:50%; margin:0 0 0 -80px; z-index:100; display:none;"><%=s_lang_0414 %></div>

<div id="pmc_moveto" class="qmpanel_shadow" style="display:none; position:absolute;">
	<div class="menu_base">
	<div id="md_moveto" class="menu_bd bd">
	<table width="100%">
		<div name="mi" onclick="move('att');" class="menu_item"><%=s_lang_0296 %></div>
<%
dim moveto_set_max
moveto_set_max = false

if sname = "" or sfname = "" then
	pfNumber = pf.FolderCount

	if pfNumber > 0 then
		Response.Write "<div class='menu_item_nofun'><div style='background:#ccc; padding-top:1px; margin-top: 5px;'></div></div>"
	end if

	if pfNumber > 9 then
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
end if
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
	else if (tgv == 4)
		is_menu_show_ck = false;
}

function setTimeClose(tgv)
{
	if (is_menu_show_moveto == true && is_in_menu_moveto == false && tgv == 1)
		$.powerFloat.hide();

	if (is_menu_show_ck == true && is_in_menu_ck == false && tgv == 4)
		$.powerFloat.hide();
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
if sname = "" or sfname = "" then
	set pf = nothing
end if

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
%>
