<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" --> 

<%
if isadmin() = false and Application("em_SpamAdmin") <> LCase(Session("wem")) then
	Response.Redirect "noadmin.asp"
end if
%>

<%
dim nb
set nb = server.createobject("easymail.SpamSampleManager")
nb.Load Session("wem"), Session("tid")

if Request.ServerVariables("REQUEST_METHOD") = "POST" and trim(request("mdel")) = "1" then
	dim msg
	msg = trim(request("allmsgs"))

	if Len(msg) > 0 then
		dim item
		dim ss
		dim se
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(9))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				nb.DelByName item
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	set nb = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("ss_brow.asp?page=" & trim(request("page")) & "&searchstr=" & trim(request("searchstr")))
end if


show_page_lines = pageline
if pageline > 100 then
	show_page_lines = 100
end if

if trim(request("page")) = "" then
	page = 0
else
	page = CInt(trim(request("page")))
end if

if page < 0 then
	page = 0
end if

searchstr = trim(request("searchstr"))

dim max_lines

if Len(searchstr) < 1 then
	max_lines = nb.List(page * show_page_lines, (page + 1) * show_page_lines + 1)
else
	max_lines = nb.Find(searchstr, page * show_page_lines, (page + 1) * show_page_lines + 1)
end if

if max_lines < 1 and page > 0 then
	page = page - 1

	if page < 0 then
		page = 0
	end if

	if Len(searchstr) < 1 then
		max_lines = nb.List(page * show_page_lines, (page + 1) * show_page_lines + 1)
	else
		max_lines = nb.Find(searchstr, page * show_page_lines, (page + 1) * show_page_lines + 1)
	end if
end if

gourl = "ss_brow.asp?" & getGRSN() & "&searchstr=" & Server.URLEncode(searchstr)
%>

<HTML>
<HEAD>
<%=s_lang_meta %>
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
function ischeck() {
	var i = 0;
	var theObj;

	for(; i<<%=show_page_lines %>; i++)
	{
		theObj = eval("document.form1.check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}

function selectpage_onchange()
{
	location.href = "<%=gourl %>&page=" + document.form1.page.value;
}

function checkall(check) {
	var i = 0;
	var theObj;

	for(; i<<%=show_page_lines %>; i++)
	{
		theObj = eval("document.form1.check" + i);

		if (theObj != null)
			theObj.checked = check;
	}
}

function allcheck_onclick() {
	if (document.form1.allcheck.checked == true)
		checkall(true);
	else
		checkall(false);
}

function mdel()
{
	if (ischeck() == true)
	{
		if (confirm("<%=s_lang_0115 %>") == false)
			return ;

		var theObj;
		var tempstr = "";
		var i = 0;

		for(; i<<%=show_page_lines %>; i++)
		{
			theObj = eval("document.form1.check" + i);

			if (theObj != null && theObj.checked == true)
				tempstr = tempstr + theObj.value + "\t";
		}

		document.form1.allmsgs.value = tempstr;
		document.form1.mdel.value = "1";
		document.form1.submit();
	}
}

function msearch() {
	document.form1.submit();
}
//-->
</SCRIPT>

<BODY>
<br>
<form action="ss_brow.asp" method=post id=form1 name=form1>
<input type="hidden" name="allmsgs">
<input type="hidden" name="mdel">
<input type="hidden" name="page" value="<%=page %>">
  <table width="90%" border="0" align="center">
	<tr>
	<td width="3%">&nbsp;</td>
	<td width="10%"><a href="javascript:mdel()"><b><%=s_lang_del %></b></a></td>
	<td width="31%">
<input type="text" name="searchstr" class="textbox" size="14" value="<%=searchstr %>">
<input type="button" value="<%=s_lang_find %>" onclick="javascript:msearch();" class="sbttn">
</td>
	<td width="11%">
<%
if page > 0 then
	Response.Write "<a href=""" & gourl & "&page=" & page - 1 & """><img src='images\prep.gif' border='0' align='absmiddle'></a>"
else
	Response.Write "<img src='images\gprep.gif' border='0' align='absmiddle'>"
end if

if max_lines > show_page_lines then
	Response.Write "&nbsp;<a href=""" & gourl & "&page=" & page + 1 & """><img src='images\nextp.gif' border='0' align='absmiddle'></a>"
else
	Response.Write "&nbsp;<img src='images\gnextp.gif' border='0' align='absmiddle'>"
end if
%>
	<td width="10%"><a href="ss_brow.asp?<%=getGRSN() %>"><b><%=s_lang_all %></b></a></td>
	<td width="10%"><a href="right.asp?<%=getGRSN() %>"><b><%=s_lang_return %></b></a></td>
	<td width="25%"><font class="s"><b><%=s_lang_0111 %></b></font></td>
	</tr>
  </table>
</td></tr>
</table>
<table width="96%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
  <tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
    <td width="7%" align="center" height="25" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()"></td>
	<td width="93%" align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_msg %></b></font></td>
  </tr>
<%
nb.MoveTop
li = 0

show_lines = max_lines
if show_lines > show_page_lines then
	show_lines = max_lines - 1
end if

do while li < show_page_lines and li < show_lines
	fg = nb.GetFStr
	samstr = nb.Get

	if Len(samstr) > 0 then
		Response.Write "  <tr>"
		Response.Write "	<td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & li & "' value='" & fg & "'></td>"
		Response.Write "    <td align='left' style='border-bottom:1px " & MY_COLOR_1 & " solid;' style='word-break: break-all'>" & server.htmlencode(samstr) & "</a>&nbsp;</td>"
		Response.Write "  </tr>" & chr(13)
	end if

	nb.MoveDown

	samstr = NULL
	fg = NULL

	li = li + 1
loop
%>
</table>
  </FORM>
<br>
</BODY>
</HTML>

<%
set nb = nothing
%>
