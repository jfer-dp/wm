<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" --> 

<%
dim dm
set dm = server.createobject("easymail.Domain")
dm.Load

if dm.GetUserManagerDomainCount(Session("wem")) < 1 then
	if isadmin() = false then
		set dm = nothing
		response.redirect "noadmin.asp"
	end if
end if

dim ManagerDomainString
i = 0
if isadmin() = false then
	ManagerDomainString = Chr(9)
	allnum = dm.GetUserManagerDomainCount(Session("wem"))

	do while i < allnum
		domain = dm.GetUserManagerDomain(Session("wem"), i)
		ManagerDomainString = ManagerDomainString & LCase(domain) & Chr(9)
		domain = NULL

		i = i + 1
	loop
end if


dim ei
set ei = server.createobject("easymail.mailinglist")
'-----------------------------------------

ei.LoadLists
%>

<HTML>
<HEAD>
<%=s_lang_meta %>
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<script LANGUAGE=javascript>
<!--
function mdel()
{
	document.form1.mode.value = "mdel";
	document.form1.submit();
}
//-->
</script>

<BODY>
<div align="center"><br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="5%" height="25">&nbsp;</td>
      <td width="16%"><a href="addmailinglist.asp?<%=getGRSN() %>"><%=s_lang_0085 %></a></td>
      <td width="9%"><a href="browmailinglist.asp?<%=getGRSN() %>"><%=s_lang_0086 %></a></td>
      <td width="13%"><%
if isadmin() = true then
%><a href="showsysinfo.asp?<%=getGRSN() %>#browmailinglist"><%=s_lang_enable %></a><%
end if
%></td>
      <td width="10%"><a href="<%
if isadmin() = true then
	Response.Write "right.asp"
else
	Response.Write "domainright.asp"
end if
      %>?<%=getGRSN() %>"><%=s_lang_return %></a></td>
      <td width="10%"><b><%=s_lang_0087 %></b></td>
    </tr>
  </table>
</div>
<br>
<form action="savemailinglist.asp" method=post id=form1 name=form1>
<input type="hidden" name="mode">
<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
    <td width="6%" align="center" height="25" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid;"><a href="javascript:mdel()"><img src='images\del.gif' border='0'></td>
    <td width="6%" align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid;"><%=s_lang_0088 %></td>
<%
if isadmin() = true then
%>
    <td width="39%" align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid;"><%=s_lang_0089 %></td>
<%
else
%>
    <td width="56%" align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid;"><%=s_lang_0089 %></td>
<%
end if
%>
    <td width="7%" nowrap align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid;"><%=s_lang_0174 %></td>
    <td width="7%" nowrap align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid;"><%=s_lang_0090 %></td>
    <td width="6%" align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid;"><%=s_lang_0091 %></td>
<%
if isadmin() = true then
%>
    <td width="17%" align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid;"><%=s_lang_0092 %></td>
<%
end if
%>
    <td width="6%" align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid;"><%=s_lang_0093 %></td>
    <td width="6%" align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"><%=s_lang_del %></td>
  </tr>
<%
i = 0
do while i < ei.MailingListCount
	ei.GetEx1 i, name, isSendWithMailingList, isPrivate, dManagerDomain, isShowToCc, isDisabled

	showline = true
	if isadmin() = false then
		showline = false
		if Len(dManagerDomain) > 0 then
			findp = InStr(1, dManagerDomain, "@")
			if findp > 0 then
				if InStr(1, ManagerDomainString, Chr(9) & LCase(Mid(dManagerDomain, findp + 1)) & Chr(9)) > 0 then
					showline = true
				end if
			end if
		end if
	end if

	if showline = true then
		Response.Write "  <tr>"
		Response.Write "	<td align='center' height='25' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "' value='" & name & "'></td>"
		Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & i+1 & "</td>"
		Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><a href='addmailinglist.asp?" & getGRSN() & "&rid=" & Server.URLEncode(name) & "'>" & server.htmlencode(name) & "</a></td>"

		if isDisabled = true then
			Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>Yes</td>"
		else
			Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>&nbsp;</td>"
		end if

		if isSendWithMailingList = true then
			Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>Yes</td>"
		else
			Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>No</td>"
		end if

		if isPrivate = true then
			Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>Yes</td>"
		else
			Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>No</td>"
		end if

		if isadmin() = true then
			if Len(dManagerDomain) < 1 then
				Response.Write "    <td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & s_lang_0094 & "</td>"
			else
				Response.Write "    <td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & server.htmlencode(dManagerDomain) & "</td>"
			end if
		end if

		Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><a href='addmailinglist.asp?" & getGRSN() & "&rid=" & Server.URLEncode(name) & "'><img src='images\edit.gif' border='0'></a></td>"
		Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><a href='savemailinglist.asp?" & getGRSN() & "&del=" & Server.URLEncode(name) & "'><img src='images\del.gif' border='0'></a></td>"
		Response.Write "  </tr>" & chr(10)
	end if

	name = NULL
	isSendWithMailingList = NULL
	isPrivate = NULL
	dManagerDomain = NULL
	isShowToCc = NULL
	isDisabled = NULL

    i = i + 1
loop
%>
</table>
  </FORM>
<br><br><br>
  <div align="center">
    <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr> 
        <td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
		<td width="94%">邮件列表的使用:
		<br>
		1. 首先需要启用"邮件列表"功能
		<br>
		2. 选用一个邮件列表发送人
		<br>
		3. 为此发送人选定多个接收用户
		<br>
		4. 然后再向此邮件列表发送人写信时, 邮件将会被发送给所有指定的接收用户
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
<br>
</BODY>
</HTML>

<%
set dm = nothing
set ei = nothing
%>
