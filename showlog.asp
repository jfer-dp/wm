<%
Response.CacheControl = "no-cache" 
%>

<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if

page = trim(request("page"))

if page <> "" then
	page = CInt(page)
else
	page = 0
end if


dim a
set a = server.createobject("easymail.logs")
'-----------------------------------------
a.load
a.PageKSize = Application("em_LogPageKSize")

readnum = CInt(trim(request("index")))

allpage = a.getLogPageNumber(readnum)
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
function selectpage_onchange()
{
	location.href = "showlog.asp?<%=getGRSN() %>&index=<%=readnum %>&page=" + document.getElementById("page").value;
}
//-->
</SCRIPT>


<BODY>

<br>
<div align="center"><br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="20%" height="28">&nbsp;</td><td colspan="30"><font class="s" color="<%=MY_COLOR_4 %>"><b>日志显示 (<font color="#FF3333"><%=(page + 1) & "</font>/" & allpage %>)</b></font></td>
      <td width="41%"><%
if page - 1 < 0 then
	response.write "<img src='images\gfirstp.gif' border='0' align='absmiddle'>&nbsp;"
	response.write "<img src='images\gprep.gif' border='0' align='absmiddle'>&nbsp;&nbsp;"
else
	response.write "<a href=""showlog.asp?page=" & 0 & "&index=" & readnum & "&" & getGRSN() & """><img src='images\firstp.gif' border='0' align='absmiddle'></a>&nbsp;"
	response.write "<a href=""showlog.asp?page=" & page - 1 & "&index=" & readnum & "&" & getGRSN() & """><img src='images\prep.gif' border='0' align='absmiddle'></a>&nbsp;&nbsp;"
end if
%><select id="page" name="page" class="drpdwn" size="1" LANGUAGE=javascript onchange="selectpage_onchange()"><%
i = 0

do while i < allpage
	if i <> page then
		response.write "<option value=""" & i & """>" & i + 1 & "页</option>"
	else
		response.write "<option value=""" & i & """ selected>" & i + 1 & "页</option>"
	end if
	i = i + 1
loop
%></select>
<%
if (page + 1) => allpage then
	response.write "<img src='images\gnextp.gif' border='0' align='absmiddle'>&nbsp;"
else
	response.write "<a href=""showlog.asp?page=" & page + 1 & "&index=" & readnum & "&" & getGRSN() & """><img src='images\nextp.gif' border='0' align='absmiddle'></a>&nbsp;"
end if

if page + 1 >= allpage then
	response.write "<img src='images\gendp.gif' border='0' align='absmiddle'>&nbsp;"
else
	response.write "<a href=""showlog.asp?page=" & allpage - 1 & "&index=" & readnum & "&" & getGRSN() & """><img src='images\endp.gif' border='0' align='absmiddle'></a>&nbsp;"
end if%></td>
    </tr>
  </table>
</div>

<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr>
      <td colspan="7" align="right" bgcolor="#ffffff"><br>
		<input type="button" value=" 返回 " onclick="location.href='logs.asp?<%=getGRSN() %>'" class="Bsbttn">
		<br>&nbsp;
      </td>
    </tr>
    <tr><td>
<%
dim str

str = a.getPageLogMessage(readnum, page)

str = server.htmlencode(str)
str = replace(str, Chr(10), "<br>")
str = replace(str, Chr(32), "&nbsp;")

response.write str
%>
</td></tr>
	<tr>
      <td colspan="7" align="right" bgcolor="#ffffff"><br>
		<input type="button" value=" 返回 " onclick="location.href='logs.asp?<%=getGRSN() %>'" class="Bsbttn">
		<br>&nbsp;
      </td>
    </tr>
</table>

  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="20%" height="28">&nbsp;</td><td><font class="s" color="<%=MY_COLOR_4 %>"><b>日志显示 (<font color="#FF3333"><%=(page + 1) & "</font>/" & allpage %>)</b></font></td>
      <td width="35%"><%
if page - 1 < 0 then
	response.write "<img src='images\gfirstp.gif' border='0' align='absmiddle'>&nbsp;"
	response.write "<img src='images\gprep.gif' border='0' align='absmiddle'>&nbsp;&nbsp;"
else
	response.write "<a href=""showlog.asp?page=" & 0 & "&index=" & readnum & "&" & getGRSN() & """><img src='images\firstp.gif' border='0' align='absmiddle'></a>&nbsp;"
	response.write "<a href=""showlog.asp?page=" & page - 1 & "&index=" & readnum & "&" & getGRSN() & """><img src='images\prep.gif' border='0' align='absmiddle'></a>&nbsp;&nbsp;"
end if

if (page + 1) => allpage then
	response.write "<img src='images\gnextp.gif' border='0' align='absmiddle'>&nbsp;"
else
	response.write "<a href=""showlog.asp?page=" & page + 1 & "&index=" & readnum & "&" & getGRSN() & """><img src='images\nextp.gif' border='0' align='absmiddle'></a>&nbsp;"
end if

if page + 1 >= allpage then
	response.write "<img src='images\gendp.gif' border='0' align='absmiddle'>&nbsp;"
else
	response.write "<a href=""showlog.asp?page=" & allpage - 1 & "&index=" & readnum & "&" & getGRSN() & """><img src='images\endp.gif' border='0' align='absmiddle'></a>&nbsp;"
end if%></td>
    </tr>
  </table>
</div>
<br><br>
</BODY>
</HTML>


<%
set a = nothing
%>
