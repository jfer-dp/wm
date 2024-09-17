<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
if isadmin() = false then
	Response.Redirect "noadmin.asp"
end if
%>

<%
sortstr = trim(request("sortstr"))
sortmode = trim(request("sortmode"))
user = trim(request("user"))
inout = trim(request("inout"))
purl = trim(request("purl"))

dim sysinfo
if trim(request("sam")) = "1" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	set sysinfo = server.createobject("easymail.sysinfo")
	sysinfo.Load

	if sysinfo.Enable_SendOut_Auto_Monitor = false then
		sysinfo.Enable_SendOut_Auto_Monitor = true
	end if

	sysinfo.Save
	set sysinfo = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(purl)
end if

if sortstr = "" then
	sortstr = "SendNumber"

	if inout = "in" then
		sortstr = "Date"
	end if
end if

if sortmode = "" then
	sortmode = "1"

	if inout = "in" then
		sortmode = "0"
	end if
end if

dim ei
set ei = server.createobject("easymail.ListUserMonitorMails")

dim addsortstr

if sortmode = "1" then
	addsortstr = "&sortstr=" & sortstr & "&sortmode=1"
else
	addsortstr = "&sortstr=" & sortstr & "&sortmode=0"
	sortmode = "0"
end if

if sortstr <> "" then
	if sortmode = "0" then
		ei.SetSort sortstr, false
	else
		ei.SetSort sortstr, true
	end if
end if

if inout = "in" then
	ei.Load_InMails(user)
else
	ei.Load_OutMails(user)
	inout = "out"
end if

allnum = ei.Count

if trim(request("page")) = "" then
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

gourl = "listmon.asp?user=" & Server.URLEncode(user) & "&inout=" & Server.URLEncode(inout) & "&page=" & page & "&" & getGRSN() & "&purl=" & Server.URLEncode(purl)
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<%=s_lang_meta %>
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/hwem.css">

<STYLE type=text/css>
<!--
body	{margin-bottom:20px;}
a:hover {color:black; text-decoration:underline}
a		{color:black; text-decoration:none}
.title_tr {white-space:nowrap; background:#f2f4f6; height:26px;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.table_min_width {width:660px; font-size:0pt; height:0px; width:0px; border:0px;}
.st_1,.st_2,.st_3,.st_4,.st_5,.st_6,.st_7 {text-align:center; white-space:nowrap; border-left:1px solid #A5B6C8; border-top:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8; padding-left:4px; padding-right:4px;}
.st_7 {width:10%; border-right:1px solid #A5B6C8;}
.cont_tr {background:white; height:26px; cursor:pointer;}
.cont_td {border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px;}
-->
</STYLE>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
function setsort(addsortstr){
	if ("<%=sortstr %>" != addsortstr)
	{
		if ("SendNumber" != addsortstr)
			location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=0";
		else
			location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=1";
	}
	else
<% if sortmode = "1" then %>
		location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=0";
<% else %>
		location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=1";
<% end if %>
}

function checkall(check) {
	var i = 0;
	var theObj;

	for(; i<<%
if allnum > pageline then
	if page > 0 then
		Response.Write allnum - page * pageline
	else
		Response.Write pageline
	end if
else
	Response.Write allnum
end if
%>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			theObj.checked = check;
	}
}

function del() {
	if (ischeck() == true)
	{
		document.f1.gourl.value = "<%=gourl & addsortstr %>";
		document.f1.submit();
	}
}

function ischeck() {
	var i = 0;
	var theObj;

	for(; i<<%
if allnum > pageline then
	if page > 0 then
		Response.Write allnum - page * pageline
	else
		Response.Write pageline
	end if
else
	Response.Write allnum
end if
%>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}

function allcheck_onclick() {
	if (document.f1.allcheck.checked == true)
		checkall(true);
	else
		checkall(false);
}

function selectpage_onchange() {
	location.href = "listmon.asp?user=<%=Server.URLEncode(user) %>&inout=<%=Server.URLEncode(inout) & addsortstr & "&" & getGRSN() & "&purl=" & Server.URLEncode(purl) %>&page=" + document.f1.page.value;
}

function m_over(tag_obj) {
	tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj) {
	tag_obj.style.backgroundColor = "white";
}

function only_check(e) {
	if (!e) var e = window.event;
	e.cancelBubble = true;
	if (e.stopPropagation)
	e.stopPropagation();
}
//-->
</SCRIPT>

<BODY>
<FORM ACTION="mulmonmail.asp" METHOD="POST" name="f1">
<INPUT NAME="user" TYPE="hidden" value="<%=user %>">
<INPUT NAME="inout" TYPE="hidden" value="<%=inout %>">
<INPUT NAME="gourl" TYPE="hidden">
<br>
<table width="98%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5;">
	<tr>
    <td nowrap width="43%" style="color:#444;">&nbsp;&nbsp;&nbsp;<%
if inout = "in" then
	Response.Write s_lang_0175 & " [" & server.htmlencode(user) & "] " & s_lang_0176
else
	Response.Write s_lang_0175 & " [" & server.htmlencode(user) & "] " & s_lang_0177
end if
%><font color="#901111"><%=allnum %></font><%=s_lang_0178 %></td><td width="27%" nowrap><%
if page > 0 then
	Response.Write "<a href=""listmon.asp?user=" & Server.URLEncode(user) & "&inout=" & Server.URLEncode(inout) & "&page=" & page - 1 & addsortstr & "&" & getGRSN() & "&purl=" & Server.URLEncode(purl) & """><img src='images/prep.gif' border='0' align='absmiddle'></a>&nbsp;&nbsp;"
else
	Response.Write "<img src='images/gprep.gif' border='0' align='absmiddle'>&nbsp;&nbsp;"
end if
%><select name="page" class="drpdwn" size="1" LANGUAGE=javascript onchange="selectpage_onchange()">
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
	Response.Write "<a href=""listmon.asp?user=" & Server.URLEncode(user) & "&inout=" & Server.URLEncode(inout) & "&page=" & page + 1 & addsortstr & "&" & getGRSN() & "&purl=" & Server.URLEncode(purl) & """><img src='images/nextp.gif' border='0' align='absmiddle'></a>&nbsp;"
else
	Response.Write "<img src='images/gnextp.gif' border='0' align='absmiddle'>&nbsp;"
end if
%></td>
<td width="30%" nowrap><a href="javascript:del()"><%=s_lang_del %></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<a href="javascript:location.href='<%=purl & "&" & getGRSN() %>'"><%=s_lang_return %></a>
    </td>
  </tr>
</table>
<br>
  <table width="98%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF">
	<tr><td class="block_top_td" colspan="7"><div class="table_min_width"></div></td></tr>
    <tr class="title_tr">
	<td width="3%" height="24" class="st_1">
	<div align="center"><input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()"></div>
	</td>
      <td width="6%" class="st_2">
        <div align="center"><a href="javascript:setsort('Read')"><%=s_lang_0126 %></a><%
if sortstr = "Read" then
	if sortmode = "1" then
		Response.Write "<a href=""javascript:setsort('Read')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "<a href=""javascript:setsort('Read')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%></div>
      </td>
      <td width="20%" class="st_3">
        <div align="center"><a href="javascript:setsort('Sender')"><%=s_lang_0147 %></a><%
if sortstr = "Sender" then
 	if sortmode = "1" then
		Response.Write "&nbsp;<a href=""javascript:setsort('Sender')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "&nbsp;<a href=""javascript:setsort('Sender')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%></div>
      </td>
      <td width="<%
if inout = "out" then
	Response.Write "36%"
else
	Response.Write "45%"
end if
%>" class="st_4">
        <div align="center"><a href="javascript:setsort('Subject')"><%=s_lang_0127 %></a><%
if sortstr = "Subject" then
 	if sortmode = "1" then
		Response.Write "&nbsp;<a href=""javascript:setsort('Subject')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "&nbsp;<a href=""javascript:setsort('Subject')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%></div>
      </td>
<%
if inout = "out" then
%>
      <td width="9%" align="center" class="st_5">
        <a href="javascript:setsort('SendNumber')"><%=s_lang_0180 %></a><%
if sortstr = "" or sortstr = "SendNumber" then
 	if sortmode = "1" then
		Response.Write "&nbsp;<a href=""javascript:setsort('SendNumber')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "&nbsp;<a href=""javascript:setsort('SendNumber')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	end if
end if
%></td>
<%
end if
%>
      <td width="18%" class="st_6">
        <div align="center"><a href="javascript:setsort('Date')"><%=s_lang_0128 %></a><%
if sortstr = "Date" then
 	if sortmode = "1" then
		Response.Write "&nbsp;<a href=""javascript:setsort('Date')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "&nbsp;<a href=""javascript:setsort('Date')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%></div>
      </td>
      <td width="8%" class="st_7">
        <div align="center"><a href="javascript:setsort('Size')"><%=s_lang_0179 %></a><%
if sortstr = "Size" then
 	if sortmode = "1" then
		Response.Write "&nbsp;<a href=""javascript:setsort('Size')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "&nbsp;<a href=""javascript:setsort('Size')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
	end if
end if
%></div>
      </td>
    </tr>
<%
i = page * pageline
li = 0

do while i < allnum and li < pageline
	ei.getMailInfo allnum - i - 1, idname, isread, priority, sendMail, sendName, subject, size, etime, mstate

	issign = false
	isenc = false

	if ei.MailIsSignature(allnum - i - 1) = true then
		issign = true
	end if

	if ei.MailIsEncrypted(allnum - i - 1) = true then
		isenc = true
	end if

	if subject = "" then
		subject = s_lang_0129
	end if

    Response.Write "<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);' onclick=""location.href='showmiomail.asp?filename=" & Server.URLEncode(idname) & "&user=" & Server.URLEncode(user) & "&inout=" & Server.URLEncode(inout) & "&" & getGRSN() & "&gourl=" & Server.URLEncode(gourl & addsortstr) & "'"">" & Chr(13)
%>
	<td align="center" height="26" class='cont_td' style="cursor:default;" onclick="only_check(event);"><input type="checkbox" name="check<%=li %>" value="<%=idname %>"></td>
	<td align="center" nowrap class='cont_td'>
<%
if mstate = 2 then
%>
	<img src="nsysmail.gif" title="<%=s_lang_0181 %>" border=0><%
else
%>	<img src="newmail.gif" title="<%=s_lang_0182 %>" border=0><%
end if

if issign = true then
%><img src="images/s0-1.gif" title="<%=s_lang_0183 %>" border=0>
<%
elseif isenc = true then
%><img src="images/e0-1.gif" title="<%=s_lang_0184 %>" border=0>
<%
end if
%></td>
      <td class='cont_td' style='word-break:break-all; word-wrap:break-word;'><%=server.htmlencode(sendName) %>&nbsp;</td>
      <td class='cont_td' style='word-break:break-all; word-wrap:break-word;'><%=server.htmlencode(subject) %>&nbsp;</td>
<%
if inout = "out" then
%>
      <td align="center" class='cont_td'><%=ei.MailSendNumber(allnum - i - 1) %></td>
<%
end if
%>
      <td class='cont_td'><%=etime %></td>
      <td align="right" class='cont_td'><%=getShowSize(size) %></td>
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

    li = li + 1
	i = i + 1
loop

if inout = "out" and allnum < 1 then
	dim uwt
	set uwt = server.createobject("easymail.UserWorkTimer")
	uwt.Load_User user

	if uwt.Enable_Out_Monitor = false then
		Response.Write "<tr bgcolor='#ffffff'><td height='80' align='center' colspan=10>" & s_lang_0185 & " [" & user & "] " & s_lang_0186
		Response.Write "<a href='uwtuser.asp?user=" & Server.URLEncode(user) & "&" & getGRSN() & "&gourl=" & Server.URLEncode(purl) & "'>" & s_lang_0187 & "</a></td></tr>"

		set sysinfo = server.createobject("easymail.sysinfo")
		sysinfo.Load

		if sysinfo.Enable_SendOut_Auto_Monitor = false then
			Response.Write "<tr bgcolor='#ffffff'><td height='20' align='center' colspan=10>" & s_lang_0188
			Response.Write "<a href='listmon.asp?purl=" & Server.URLEncode(purl) & "&" & getGRSN() & "&sam=1" & "'>" & s_lang_0189 & "</a></td></tr>"
		end if

		set sysinfo = nothing
	end if

	set uwt = nothing
end if
%>
</table>
</FORM>
</BODY>
</HTML>

<%
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

set ei = nothing
%>
