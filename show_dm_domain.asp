<!--#include file="passinc.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.domain")
ei.DM_Load

'-----------------------------------------
dim eu
set eu = Application("em")
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
function domainname_onchange() {
	location.href = "show_dm_domain.asp?<%=getGRSN() %>&selectdomain=" + document.f1.domainname.value;
}
//-->
</SCRIPT>


<BODY>
<br>
<FORM ACTION="save_dm_domain.asp" METHOD="POST" NAME="f1">
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="10%" height="25">&nbsp;</td>
	<td width="37%"><a href="showsysinfo.asp?<%=getGRSN() %>#domainMonitor"><b>启动项设置</b></a></td>
	<td width="30%"><a href="right.asp?<%=getGRSN() %>"><b>返回</b></a></td>
	<td width="23%"><b>域邮件监控</b></td>
    </tr>
  </table>
<br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td width="50%" height="25" style="border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center">域名</div>
      </td>
      <td width="50%" style="border-top:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
	<div align='center'>域监控邮件接收帐号</div>
      </td>
    </tr>
	<tr><td align="center" height="25" style="border-bottom:1px <%=MY_COLOR_1 %> solid;"><select name="domainname" class="drpdwn" LANGUAGE=javascript onchange="return domainname_onchange()">
<%
i = 0
allnum = ei.GetCount()

do while i < allnum
	domain = ei.GetDomain(i)

	if domain <> trim(request("selectdomain")) then
		response.write "<option value='" & server.htmlencode(domain) & "'>" & server.htmlencode(domain) & "</option>" & Chr(13)
	else
		curdomain = domain
		response.write "<option value='" & server.htmlencode(domain) & "' selected>" & server.htmlencode(domain) & "</option>" & Chr(13)
	end if

	domain = NULL

	i = i + 1
loop


if curdomain = "" then
	curdomain = ei.GetDomain(0)
end if
%>
</select></td>
	<td align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid;"><select name="username" class="drpdwn"><option value="">[无]</option>
<%
curdomain = LCase(curdomain)
tdname = LCase(ei.DM_GetUser(curdomain))

i = 0
allnum = eu.GetUsersCount

do while i < allnum
	eu.GetUserByIndex i, name, domain, comment

	if LCase(domain) = curdomain then
		if LCase(name) = tdname then
			Response.Write "<option value='" & server.htmlencode(name) & "' selected>" & server.htmlencode(name) & "</option>" & Chr(13)
		else
			Response.Write "<option value='" & server.htmlencode(name) & "'>" & server.htmlencode(name) & "</option>" & Chr(13)
		end if
	end if

	name = NULL
	domain = NULL
	comment = NULL

	i = i + 1
loop

tdname = NULL
%>
</select></td></tr>
    <tr> 
      <td height="50" colspan="2" align="right" bgcolor="#ffffff">
	<br>
	<input type="submit" value=" 保存 " class="Bsbttn">&nbsp;
	<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
      </td>
    </tr>
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
		<td width="94%">此功能可以对用户管理中本域内设置为被监控的帐号进行监控.&nbsp;&nbsp;(注意: 启用本功能后, 域管理员也将同时启用此功能)
		<br><br>如: 在 "用户管理 | 修改" 界面中设置某一域名(如: mydomain.com)下的帐号 user 为被监控帐号后, 所有 user 帐号接收以及发送的邮件都会在指定的域监控邮件接收帐号内保留一份拷贝.
		<br>
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
<br><br>
</BODY>
</HTML>

<%
set eu = nothing
set ei = nothing
%>
