<!--#include file="passinc.asp" --> 
<!--#include file="language-1.asp" --> 

<%
dim nb
set nb = server.createobject("easymail.NoteBooksManager")

sortstr = request("sortstr")
o_sortmode = request("sortmode")
issort = false

searchmode = 0
if Len(trim(request("searchmode"))) > 0 and IsNumeric(trim(request("searchmode"))) = true then
	searchmode = CLng(trim(request("searchmode")))
end if

searchstr = trim(request("searchstr"))

nb.SearchMode = searchmode
nb.SearchString = searchstr

if sortstr <> "" then
	if o_sortmode = "1" then
		addsortstr = "&sortstr=" & sortstr & "&sortmode=1" & "&searchmode=" & searchmode & "&searchstr=" & server.htmlencode(searchstr)
		sortmode = true

		nb.SetSort sortstr, sortmode
		issort = true
	elseif o_sortmode = "0" then
		addsortstr = "&sortstr=" & sortstr & "&sortmode=0" & "&searchmode=" & searchmode & "&searchstr=" & server.htmlencode(searchstr)
		sortmode = false

		nb.SetSort sortstr, sortmode
		issort = true
	end if
end if

if issort = false then
	addsortstr = "&sortstr=date&sortmode=0" & "&searchmode=" & searchmode & "&searchstr=" & server.htmlencode(searchstr)
	sortstr = "date"
	o_sortmode = "0"
	sortmode = false

	nb.SetSort sortstr, sortmode
	issort = true
end if


nb.Load Session("wem")

isSearching = false
if nb.SearchMode > 0 and Len(nb.SearchString) > 0 then
	isSearching = true
	allnum = nb.MatchCount
else
	allnum = nb.Count
	searchmode = 0
end if

allnb = nb.Count

if trim(request("page")) = "" then
	page = 0
else
	page = CInt(request("page"))
end if

if page < 0 then
	page = 0
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

gourl = "nb_brow.asp?page=" & page
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
.title_tr {white-space:nowrap; background:#f2f4f6; height:24px;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.st_l, .st_r {height:24px; text-align:center; white-space:nowrap; border-left:1px solid #A5B6C8; border-top:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8;}
.st_r {border-right:1px solid #A5B6C8;}
.cont_tr {background:white; height:24px;}
.cont_td {height:24px; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px;}
.cont_td_word {height:24px; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px; word-break:break-all; word-wrap:break-word;}
.urf {color:black;}
.urf:hover {color:black;}
-->
</STYLE>
</HEAD>

<script type="text/javascript" src="images/sc_left.js"></script>

<script type="text/javascript">
<!--
function ischeck() {
	var i = 0;
	var theObj;

	for(; i<<%=allnb %>; i++)
	{
		theObj = document.getElementById("check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}

function delone(delid)
{
	if (confirm("<%=a_lang_196 %>") == false)
		return ;

	location.href = "nb_delete.asp?id=" + delid + "&thispage=<%=page %>&<%=getGRSN() & addsortstr %>&addsortstr=<%=Server.URLEncode(addsortstr) %>";
}

function mdel()
{
	if (ischeck() == true)
	{
		if (confirm("<%=a_lang_196 %>") == false)
			return ;

		document.getElementById("form1").mdel.value = "1";
		document.getElementById("form1").action = "nb_delete.asp";
		document.getElementById("form1").submit();
	}
}

function selectpage_onchange()
{
	location.href = "nb_brow.asp?<%=getGRSN() & addsortstr %>&page=" + document.getElementById("form1").page.value;
}

function setsort(addsortstr){
	if ("<%=sortstr %>" != addsortstr)
		location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=0&searchmode=<%=searchmode %>&searchstr=<%=server.htmlencode(searchstr) %>&<%=getGRSN() %>";
	else
<% if sortmode = false then %>
		location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=1&searchmode=<%=searchmode %>&searchstr=<%=server.htmlencode(searchstr) %>&<%=getGRSN() %>";
<% else %>
		location.href = "<%=gourl %>" + "&sortstr=" + addsortstr + "&sortmode=0&searchmode=<%=searchmode %>&searchstr=<%=server.htmlencode(searchstr) %>&<%=getGRSN() %>";
<% end if %>
}

function checkall(check) {
	var i = 0;
	var theObj;

	for(; i<<%=allnb %>; i++)
	{
		theObj = document.getElementById("check" + i);

		if (theObj != null)
			theObj.checked = check;
	}
}

function allcheck_onclick() {
	if (document.getElementById("form1").allcheck.checked == true)
		checkall(true);
	else
		checkall(false);
}

function msearch() {
	document.getElementById("form1").submit();
}

function window_onload() {
	var smv = <%=searchmode %>;
	if (smv < 1)
		smv = 1;

	document.getElementById("form1").searchmode.value = smv;
}

function m_over(tag_obj) {
	tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj) {
	tag_obj.style.backgroundColor = "white";
}
//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload()">
<form action="nb_brow.asp" method=post id=form1 name=form1>
<input type="hidden" name="mdel">
<input type="hidden" name="addsortstr" value="<%=server.htmlencode(addsortstr) %>">
<input type="hidden" name="sortstr" value="<%=sortstr %>">
<input type="hidden" name="sortmode" value="<%=o_sortmode %>">
<input type="hidden" name="thispage" value="<%=page %>">
<table width="90%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr>
	<td align="left" height="28" width="43%" nowrap style="padding-left:4px; color:#444444;">
<a class='wwm_btnDownload btn_blue' href="nb_add.asp?<%=getGRSN() %>&addsortstr=<%=Server.URLEncode("&page=" & page & addsortstr) %>"><%=s_lang_add %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:mdel();"><%=s_lang_del %></a>
	</td>
	<td width="30%">
<%
if page > 0 then
	Response.Write "<a href=""nb_brow.asp?" & getGRSN() & addsortstr & "&page=" & page - 1 & """><img src='images/prep.gif' border='0' align='absmiddle'></a>&nbsp;"
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
	Response.Write "&nbsp;<a href=""nb_brow.asp?" & getGRSN() & addsortstr & "&page=" & page + 1 & """><img src='images/nextp.gif' border='0' align='absmiddle'></a>"
else
	Response.Write "&nbsp;<img src='images/gnextp.gif' border='0' align='absmiddle'>"
end if
%>
	<td align='right' style="padding-right:8px; color:#444444;"><%=s_lang_301 %></td>
	</tr>
</table>
<br>

<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr class="title_tr">
    <td width="4%" class="st_l"><input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()"></td>
    <td width="5%" class="st_l"><%=a_lang_071 %></td>
<%
if issort = false then
%>
	<td width="46%" class="st_l"><a class='urf' href="javascript:setsort('title')"><%=a_lang_298 %></a></td>
<%
else
	if sortstr = "title" then
 		if sortmode = true then
%>
	<td width="46%" class="st_l"><a class='urf' href="javascript:setsort('title')"><%=a_lang_298 %></a>&nbsp;<a href="javascript:setsort('title')"><img src='images/arrow_up.gif' border='0' align='absmiddle'></a></td>
<%
		else
%>
	<td width="46%" class="st_l"><a class='urf' href="javascript:setsort('title')"><%=a_lang_298 %></a>&nbsp;<a href="javascript:setsort('title')"><img src='images/arrow_down.gif' border='0' align='absmiddle'></a></td>
<%
		end if
	else
%>
	<td width="46%" class="st_l"><a class='urf' href="javascript:setsort('title')"><%=a_lang_298 %></a></td>
<%
	end if
end if

if issort = false then
%>
	<td width="22%" class="st_l"><a class='urf' href="javascript:setsort('date')"><%=a_lang_243 %></a></td>
<%
else
	if sortstr = "date" then
		if sortmode = true then
%>
	<td width="22%" class="st_l"><a class='urf' href="javascript:setsort('date')"><%=a_lang_243 %></a>&nbsp;<a href="javascript:setsort('date')"><img src='images/arrow_up.gif' border='0' align='absmiddle'></a></td>
<%
		else
%>
	<td width="22%" class="st_l"><a class='urf' href="javascript:setsort('date')"><%=a_lang_243 %></a>&nbsp;<a href="javascript:setsort('date')"><img src='images/arrow_down.gif' border='0' align='absmiddle'></a></td>
<%
		end if
	else
%>
	<td width="22%" class="st_l"><a class='urf' href="javascript:setsort('date')"><%=a_lang_243 %></a></td>
<%
	end if
end if
%>
	<td width="11%" class="st_l"><%=s_lang_302 %></td>
	<td width="6%" class="st_l"><%=s_lang_303 %></td>
	<td width="6%" class="st_r"><%=s_lang_del %></td>
	</tr>
<%
minshowi = page * pageline
showi = 0
i = 0
li = 0

do while i < allnb and li < pageline
	si = allnb - i - 1
	nb.GetEx si, nb_date, title, text, ismatch

	isShow = true
	if isSearching = true and ismatch = false then
		isShow = false
	end if

	if isShow = true then
		showi = showi + 1
	end if

if isShow = true and showi > minshowi then
	Response.Write "<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);'>"
	Response.Write "	<td align='center' class='cont_td'><input type='checkbox' id='check" & si & "' name='check" & si & "' value='" & si & "'></td>"
	Response.Write "	<td align='center' class='cont_td'>" & minshowi + li + 1 & "</td>"

	Response.Write "	<td align='left' class='cont_td_word'><a href='nb_show.asp?id=" & si & "&page=" & page & "&" & getGRSN() & addsortstr & "'>" & server.htmlencode(title) & "</a></td>"
	Response.Write "	<td align='right' class='cont_td'>" & getShowTime(nb_date) & "&nbsp;</td>"

	Response.Write "	<td align='right' class='cont_td'>" & getShowSize(Len(text)) & "</td>"
	Response.Write "	<td align='center' class='cont_td'><a href='nb_edit.asp?id=" & si & "&page=" & page & "&" & getGRSN() & addsortstr & "&addsortstr=" & Server.URLEncode(addsortstr) & "'><img src='images/edit.gif' border='0' title='" & s_lang_303 & "'></a></td>"
	Response.Write "	<td align='center' class='cont_td'><a href='javascript:delone(" & si & ")'><img src='images/del.gif' border='0' title='" & s_lang_del & "'></a></td>"
	Response.Write "</tr>" & Chr(13)

    li = li + 1
end if

	nb_date = NULL
	title = NULL
	text = NULL
	ismatch = NULL

    i = i + 1
loop
%>
</table>
<br><br>
<table width="90%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border:1px solid #A5B6C8;'>
<td width="32%" nowrap align='left' style="padding-left:4px;">
<select name="searchmode" class="drpdwn" size="1">
<option value="1" selected><%=s_lang_304 %></option>
<option value="2"><%=s_lang_305 %></option>
<option value="3"><%=s_lang_306 %></option>
</select>
<input type="text" name="searchstr" class="n_textbox" size="14" value="<%=searchstr %>">
</td>
<td width="10%" nowrap align='left' style="padding-left:4px;">
<a class='wwm_btnDownload btn_gray' href="javascript:msearch();"><%=s_lang_307 %></a>
</td>
<td align='left' style="padding-left:16px; color:#444444;">
<%
if isSearching = true then
	Response.Write s_lang_308 & s_lang_mh & " <font color='#901111'>" & nb.MatchCount & "</font>" & s_lang_309
end if
%>
</td>
</table>
</FORM>
</BODY>
</HTML>

<%
set nb = nothing

function getShowTime(exday)
	getShowTime = ""

	if Len(exday) = 14 then
		getShowTime = Mid(Cstr(exday), 1, 4) & "-" & Mid(Cstr(exday), 5, 2) & "-" & Mid(Cstr(exday), 7, 2) & " " & Mid(Cstr(exday), 9, 2) & ":" & Mid(Cstr(exday), 11, 2) & ":" & Mid(Cstr(exday), 13, 2)
	end if
end function

function getShowSize(bytesize)
	if bytesize < 1000 then
		getShowSize = bytesize & a_lang_310
	else
		getShowSize = CLng(bytesize/1000) & a_lang_311
	end if
end function
%>
