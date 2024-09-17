<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" -->

<%
if Application("em_EnableBBS") = false then
	response.redirect "noadmin.asp?errstr=" & Server.URLEncode(b_lang_087) & "&" & getGRSN()
end if

fileid = trim(request("fileid"))

if isadmin() = false then
	dim pfvl
	set pfvl = server.createobject("easymail.PubFolderViewLimit")
	pfvl.Load fileid

	if pfvl.IsShow(Session("mail")) = false then
		set pfvl = nothing
		Response.Redirect "noadmin.asp"
	end if

	set pfvl = nothing
end if


if pageline > 50 then
	pagelines = 50
else
	pagelines = pageline
end if

backurl = trim(request("backurl"))

page = trim(request("page"))
sortmode = request("sortmode")
sortstr = request("sortstr")

if sortmode = 0 then
	sortmode = true
else
	sortmode = false
end if

if sortstr = "" or IsNumeric(sortstr) = false then
	sortstr = "0"
end if


if IsNumeric(page) = false then
	page = "0"
end if

page = CInt(page)


dim pf
set pf = server.createobject("easymail.PubFolderManager")

pf.Order = sortmode
pf.SortMode = CInt(sortstr)

if pf.load(fileid) = false then
	set pf = nothing
	Response.Redirect "err.asp?errstr=" & Server.URLEncode(b_lang_101) & "&" & getGRSN() & "&gourl=viewmailbox.asp"
end if


allpage = CInt((pf.TopItemCount - (pf.TopItemCount mod pagelines))/ pagelines)
if pf.TopItemCount mod pagelines <> 0 then
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

dim min_show_index
min_show_index = -1

dim pfadmin

dim filename
dim ownID
dim step
dim nextstep
dim postuser
dim subject
dim time
dim length
dim state
dim searchkey
dim readcount

pf.GetFolderInfo filename, admin, permission, name, createtime, count, maxid, maxitem, itemmaxsize

dim ei
set ei = Application("em")
pfadmin = ei.GetUserMail(admin)
set ei = nothing

gourl = "showpf.asp?fileid=" & fileid & "&page=" & page & "&" & getGRSN()
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
.st_span {float:left;}
.font_top_title {font-size:11pt; color:#104A7B; font-weight:bold;}

.t_1 {border-left:1px solid #8CA5B5; border-top:1px solid #8CA5B5; border-bottom:1px solid #8CA5B5; padding-left:4px; color:#444;}
.t_2 {border-top:1px solid #8CA5B5; border-bottom:1px solid #8CA5B5;}
.t_3 {border-right:1px solid #8CA5B5; border-top:1px solid #8CA5B5; border-bottom:1px solid #8CA5B5; padding-right:10px;}

.title_tr {white-space:nowrap; background:#f2f4f6; height:24px;}
.st_left {text-align:center; white-space:nowrap; border-left:1px solid #A5B6C8; border-top:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8;}
.st_right {text-align:center; white-space:nowrap; border:1px solid #A5B6C8;}
.ct_left {text-align:center; white-space:nowrap; border-bottom:1px solid #A5B6C8;}
.ct_mid {text-align:center; white-space:nowrap; border-left:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8;}
.ct_right {text-align:left; border-left:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8;}

.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.table_min_width {width:660px; font-size:0pt; height:0px; width:0px; border:0px;}
-->
</STYLE>
</HEAD>

<script LANGUAGE=javascript>
<!--
function selectpage_onchange()
{
<%
if sortmode = true then
	smode = 0
else
	smode = 1
end if
%>
	location.href = "showpf.asp?fileid=<%=fileid %>&sortstr=<%=sortstr %>&sortmode=<%=smode %>&<%=getGRSN() %>&page=" + document.f1.page.value;
}

function setsort(addsortstr){
<% if sortmode = false then %>
	location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=0";
<% else %>
	location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=1";
<% end if %>
}

function change_sort(addsortstr){
<% if sortmode = false then %>
	location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=1";
<% else %>
	location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=0";
<% end if %>
}

<%
if isadmin() = true or LCase(pf.admin) = LCase(Session("wem")) then
%>
function godel(){
	if (ischeck() == true)
	{
		if (confirm("<%=b_lang_036 %>") == false)
			return ;

		document.f1.gourl.value = "<%=gourl & "&sortstr=" & sortstr & "&sortmode=" & smode %>";
		document.f1.submit();
	}
}

function allcheck_onclick() {
	if (document.f1.allcheck.checked == true)
		checkall(true);
	else
		checkall(false);
}

function checkall(check){
	var i = parseInt(document.f1.min_show_index.value);
	var theObj;

	for(; i < parseInt(document.f1.max_show_index.value) + 1; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			theObj.checked = check;
	}
}

function ischeck(){
	var i = parseInt(document.f1.min_show_index.value);
	var theObj;

	for(; i < parseInt(document.f1.max_show_index.value) + 1; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}
<%
end if
%>
//-->
</script>

<body>
<FORM ACTION="mdelpfmail.asp" METHOD="POST" name="f1">
<table width="98%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="margin-top:4px;">
	<tr><td class="block_top_td" colspan="4"><div class="table_min_width"></div></td></tr>
	<tr>
	<td align="left" height="28" width="40%" nowrap class="t_1">

<span class="st_span">
<%
if backurl = "" then
%>
<a class='wwm_btnDownload btn_blue' href="showallpf.asp?<%=getGRSN() %>"><%=s_lang_return %></a>
<%
else
	if InStr(1, backurl, "?") = 0 Then
		backurl = backurl & "?" & getGRSN()
	else
		backurl = backurl & "&" & getGRSN()
	end if
%>
<a class='wwm_btnDownload btn_blue' href="<%=backurl %>"><%=s_lang_return %></a>
<%
end if
%>
</span>
<span style='float:left; width:3px;'>&nbsp;</span>

<span class="st_span">
<a class='wwm_btnDownload btn_blue' href="wframe.asp?<%=getGRSN() %>&mode=post&pid=0&iniid=<%=fileid %>&gourl=<%=Server.URLEncode("showpf.asp?fileid=" & fileid & "&sortstr=" & sortstr & "&sortmode=" & smode & "&" & getGRSN() & "&page=" & page) %>"><%=b_lang_057 %></a>
</span>
<span style='float:left; width:3px;'>&nbsp;</span>

<%
if isadmin() = true or LCase(pf.admin) = LCase(Session("wem")) then
%>
<span class="st_span">
<a class='wwm_btnDownload btn_blue' href="javascript:godel()"><%=s_lang_del %></a>
</span>
<span style='float:left; width:3px;'>&nbsp;</span>
<%
end if
%>
<span class="st_span">
<a class='wwm_btnDownload btn_blue' href="findpfmail.asp?fileid=<%=fileid %>&<%=getGRSN() %>"><%=b_lang_102 %></a>
</span>
<span style='float:left; width:3px;'>&nbsp;</span>
	</td>
	<td width="20%" nowrap class="t_2">
<%
if page > 0 then
	Response.Write "<a href=""showpf.asp?fileid=" & fileid & "&sortstr=" & sortstr & "&sortmode=" & smode & "&" & getGRSN() & "&page=" & page - 1 & """><img src='images/prep.gif' border='0' align='absmiddle'></a>&nbsp;"
else
	Response.Write "<img src='images/gprep.gif' border='0' align='absmiddle'>&nbsp;"
end if
%>
<select name="page" class="drpdwn" size="1" LANGUAGE=javascript onchange="selectpage_onchange()">
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
	Response.Write "&nbsp;<a href=""showpf.asp?fileid=" & fileid & "&sortstr=" & sortstr & "&sortmode=" & smode & "&" & getGRSN() & "&page=" & page + 1 & """><img src='images/nextp.gif' border='0' align='absmiddle'></a>"
else
	Response.Write "&nbsp;<img src='images/gnextp.gif' border='0' align='absmiddle'>"
end if
%>
	</td>
	<td width="40%" align="right" nowrap class="t_3">[<font class="font_top_title"><%=server.htmlencode(name) & "</font>]&nbsp;(<font color='#901111'>" & page+1 & "</font>/" & allpage & ")" %></td>
	</tr>
</table>
<br>
<table width="98%" border="0" bgcolo="white" align="center" cellspacing="0">
	<tr><td class="block_top_td" colspan="4"><div class="table_min_width"></div></td></tr>
	<tr class="title_tr">
<%
if isadmin() = true or LCase(pf.admin) = LCase(Session("wem")) then
%>
	<td width="5%" nowrap height='28' class="st_left"><input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()"></td>
	<td width="7%" nowrap class="st_left"><%=b_lang_040 %></td>
<%
else
%>
	<td width="7%" nowrap height='28' class="st_left"><%=b_lang_040 %></td>
<%
end if
%>
	<td width="93%" align="center" nowrap class="st_right"><%=b_lang_058 %>: <%
Response.Write "<a href=""javascript:setsort('" & sortstr & "')"">" & getSortStr(sortstr) & "</a>&nbsp;"

if sortmode = true then
	Response.Write "<a href=""javascript:setsort('" & sortstr & "')""><img src='images/arrow_down.gif' border='0' align='absmiddle'></a>"
else
	Response.Write "<a href=""javascript:setsort('" & sortstr & "')""><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>"
end if

Response.Write "&nbsp;&nbsp;&nbsp;("
i = 0
do while i < 5
	if i <> CInt(sortstr) then
		Response.Write "<a href=""javascript:change_sort('" & i & "')"">" & getSortStr(i) & "</a>"

		if i <> 4 and (i = 3 and CInt(sortstr) = 4) = false then
			Response.Write "&nbsp;&nbsp;"
		end if
	end if

	i = i + 1
loop
Response.Write ")"
%>
	</td></tr>
<%
filename = NULL
admin = NULL
permission = NULL
name = NULL
createtime = NULL
count = NULL
maxid = NULL
maxitem = NULL
itemmaxsize = NULL



allnum = pf.ItemCount

dim showstep
dim nextnextstep

i = 0
dim showi
showi = 0

do while i < allnum

pf.GetItemInfoByIndex i+1, filename, ownID, nextstep, postuser, subject, time, length, state, searchkey, readcount, face, istop

filename = NULL
ownID = NULL
postuser = NULL
subject = NULL
time = NULL
length = NULL
state = NULL
searchkey = NULL
readcount = NULL
face = NULL
istop = NULL

pf.GetItemInfoByIndex i, filename, ownID, step, postuser, subject, time, length, state, searchkey, readcount, face, istop

if subject = "" then
	subject = b_lang_059
end if

showstep = step

if showstep = 0 then
	showi = showi + 1
end if

if showi > pagelines*page and showi <= pagelines*(page+1) then
	if showstep = 0 then
		if min_show_index = -1 then
			min_show_index = showi
		end if

		if isadmin() = true or LCase(pf.admin) = LCase(Session("wem")) then
			Response.Write "<tr><td height='24' class='ct_left'><input type='checkbox' name='check" & showi & "' value='" & filename & "'></td><td class='ct_mid'>" & showi & "</td>"
		else
			Response.Write "<tr><td height='24' class='ct_left'>" & showi & "</td>"
		end if

		Response.Write "<td class='ct_right'>"
		Response.Write "<ul>"

		if istop = false then
			if face > 0 then
				facestr = "<img src='images/face/" & face & ".gif' align='absmiddle' border='0'>&nbsp;"
			else
				facestr = ""
			end if
		else
			facestr = "<img src='images/face/top.gif' align='absmiddle' border='0'>&nbsp;"
		end if

		Response.Write "<li>" & facestr & "<a href='showpfmail.asp?" & getGRSN() & "&filename=" & filename & "&pid=" & ownID & "&iniid=" & fileid & "&" & getGRSN() & "&gourl=" & Server.URLEncode("showpf.asp?fileid=" & fileid & "&sortstr=" & sortstr & "&sortmode=" & smode & "&" & getGRSN() & "&page=" & page) & "'><b>"

		if postuser <> pfadmin then
			Response.Write server.htmlencode(subject)
		else
			Response.Write "<font color='#FF3333'>" & server.htmlencode(subject) & "</font>"
		end if

		Response.Write "</b></li>" & "&nbsp;[" & b_lang_060 & ":" & postuser & "&nbsp;&nbsp;" & getShowSize(length) & "&nbsp;&nbsp;" & getTimeStr(time) & "&nbsp;(" & b_lang_061 & ":" & readcount & b_lang_062 & ")]</a>" & "<ul>" & chr(13)
	else
		if face > 0 then
			facestr = "<img src='images/face/" & face & ".gif' align='absmiddle' border='0'>&nbsp;"
		else
			facestr = ""
		end if

		Response.Write "<li>" & facestr & "<a href='showpfmail.asp?" & getGRSN() & "&filename=" & filename & "&pid=" & ownID & "&iniid=" & fileid & "&" & getGRSN() & "&gourl=" & Server.URLEncode("showpf.asp?fileid=" & fileid & "&sortstr=" & sortstr & "&sortmode=" & smode & "&" & getGRSN() & "&page=" & page) & "'><b>"

		if postuser <> pfadmin then
			Response.Write server.htmlencode(subject)
		else
			Response.Write "<font color='#FF3333'>" & server.htmlencode(subject) & "</font>"
		end if

		Response.Write "</b></li>" & "&nbsp;[" & b_lang_064 & ":" & postuser & "&nbsp;&nbsp;" & getShowSize(length) & "&nbsp;&nbsp;" & getTimeStr(time) & "&nbsp;(" & b_lang_061 & ":" & readcount & b_lang_062 & ")]</a>" & chr(13)

		if nextstep <> step then
			tempstep = step

			do while nextstep > tempstep
				Response.Write "<ul>"
				tempstep = tempstep + 1
			loop

			do while nextstep < tempstep
				Response.Write "</ul>"
				tempstep = tempstep - 1
			loop
		end if
	end if


	if IsNull(nextstep) or nextstep = 0 then
		Response.Write "</td></tr>"
	end if
end if


filename = NULL
ownID = NULL
step = NULL
postuser = NULL
subject = NULL
time = NULL
length = NULL
state = NULL
searchkey = NULL
readcount = NULL
face = NULL
istop = NULL

nextstep = NULL


	i = i + 1
loop
%>
</table>
<input type="hidden" name="min_show_index" value="<%=min_show_index %>">
<input type="hidden" name="max_show_index" value="<%=showi %>">
<input type="hidden" name="gourl">
<input type="hidden" name="iniid" value="<%=fileid %>">
</FORM>
</BODY>
</HTML>

<%
pfadmin = NULL

set pf = nothing

function getTimeStr(otime)
	getTimeStr = mid(otime, 1, 4) & "-"
	getTimeStr = getTimeStr & mid(otime, 5, 2) & "-"
	getTimeStr = getTimeStr & mid(otime, 7, 2) & "&nbsp;"
	getTimeStr = getTimeStr & mid(otime, 9, 2) & ":"
	getTimeStr = getTimeStr & mid(otime, 11, 2) & ":"
	getTimeStr = getTimeStr & mid(otime, 13, 2)
end function

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

function getSortStr(sortnum)
	if sortnum = 0 then
		getSortStr = b_lang_063
	elseif sortnum = 1 then
		getSortStr = b_lang_064
	elseif sortnum = 2 then
		getSortStr = b_lang_065
	elseif sortnum = 3 then
		getSortStr = b_lang_066
	elseif sortnum = 4 then
		getSortStr = b_lang_067
	end if
end function
%>
