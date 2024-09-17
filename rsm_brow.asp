<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" --> 

<%
if isadmin() = false and Application("em_SpamAdmin") <> LCase(Session("wem")) then
	Response.Redirect "noadmin.asp"
end if
%>

<%
dim nb
set nb = server.createobject("easymail.ReportSpamManager")
nb.Load

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim ischange
	ischange = false

	dim item
	dim ss
	dim se
	dim msg

	msg = trim(request("badfg"))
	if Len(msg) > 0 then
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(9))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				nb.SetSpam item, true
				ischange = true
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	msg = trim(request("okfg"))
	if Len(msg) > 0 then
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(9))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				nb.RemoveByName item
				ischange = true
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	if ischange = true then
		nb.Save
	end if

	set nb = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=rsm_brow.asp"
end if


allnum = nb.Count

show_page_lines = pageline
if pageline > 100 then
	show_page_lines = 100
end if

if trim(request("page")) = "" then
	page = 0
else
	page = CInt(request("page"))
end if

if page < 0 then
	page = 0
end if


allpage = CInt((allnum - (allnum mod show_page_lines))/ show_page_lines)


if allnum mod show_page_lines <> 0 then
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

ok_fg = ""
%>

<HTML>
<HEAD>
<%=s_lang_meta %>
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
function isHave() {
	var i = 0;
	var theObj;

	for(; i<<%=show_page_lines %>; i++)
	{
		theObj = eval("document.form1.check" + i);

		if (theObj != null)
			return true;
	}

	return false;
}

function selectpage_onchange()
{
	location.href = "rsm_brow.asp?<%=getGRSN() %>&page=" + document.form1.page.value;
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

function setspam(isspam) {
	if (document.form1.okfg.value.length < 1 && isHave() == false)
		return ;

	var i = 0;
	var theObj;
	var is_bad = false;

	for (i; i < <%=show_page_lines %>; i++)
	{
		is_bad = false;
		theObj = eval("document.form1.check" + i);

		if (theObj != null)
		{
			if (theObj.checked == true && isspam == 1)
				is_bad = true;
			else if (theObj.checked == false && isspam == 0)
				is_bad = true;

			if (is_bad == false)
				document.form1.okfg.value = document.form1.okfg.value + theObj.value + "\t";
			else
				document.form1.badfg.value = document.form1.badfg.value + theObj.value + "\t";
		}
	}

	document.form1.submit();
}
//-->
</SCRIPT>

<BODY>
<br>
<form action="rsm_brow.asp" method=post id=form1 name=form1>
  <table width="96%" border="0" align="center">
	<tr>
	<td width="3%">&nbsp;</td>
	<td width="46%">
<a href="javascript:setspam(0)"><%=s_lang_0113 %></a>
<br>
<a href="javascript:setspam(1)"><%=s_lang_0114 %></a>
	</td>
	<td width="20%">
<%
if page > 0 then
	response.write "<a href=""rsm_brow.asp?" & getGRSN() & "&page=" & page - 1 & """><img src='images\prep.gif' border='0' align='absmiddle'></a>&nbsp;"
else
	response.write "<img src='images\gprep.gif' border='0' align='absmiddle'>&nbsp;"
end if
%>
<select name="page" class="drpdwn" size="1" LANGUAGE=javascript onchange="selectpage_onchange()">
<%
i = 0

do while i < allpage
	if i <> page then
		response.write "<option value=""" & i & """>" & i + 1 & "</option>"
	else
		response.write "<option value=""" & i & """ selected>" & i + 1 & "</option>"
	end if
	i = i + 1
loop
%></select>
<%
if page < allpage - 1 then
	response.write "&nbsp;<a href=""rsm_brow.asp?" & getGRSN() & "&page=" & page + 1 & """><img src='images\nextp.gif' border='0' align='absmiddle'></a>"
else
	response.write "&nbsp;<img src='images\gnextp.gif' border='0' align='absmiddle'>"
end if
%>
	<td width="10%"><a href="right.asp?<%=getGRSN() %>"><%=s_lang_return %></a></td>
	<td width="21%"><font class="s"><b><%=s_lang_0109 %></b></font></td>
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
i = page * show_page_lines
li = 0

do while i < allnum and li < show_page_lines
	fg = nb.Get(i)
	samstr = nb.GetSample(fg)

	if Len(samstr) > 0 then
		Response.Write "  <tr>"
		Response.Write "	<td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & li & "' value='" & fg & "'></td>"
		Response.Write "    <td align='left' style='border-bottom:1px " & MY_COLOR_1 & " solid;' style='word-break: break-all'>" & server.htmlencode(samstr) & "</a>&nbsp;</td>"
		Response.Write "  </tr>" & chr(13)

		li = li + 1
	else
		ok_fg = ok_fg + fg + Chr(9)
	end if

	samstr = NULL
	fg = NULL
    i = i + 1
loop
%>
</table>
<input type="hidden" name="okfg" value="<%=ok_fg %>">
<input type="hidden" name="badfg" value="">
  </FORM>
<br>
</BODY>
</HTML>

<%
set nb = nothing
%>
