<!--#include file="passinc.asp" --> 

<%
dim esi
set esi = server.createobject("easymail.sysinfo")
esi.Load

if esi.enableDomainMonitor = false then
	set esi = nothing
	response.redirect "noadmin.asp"
end if

set esi = nothing



dim ei
set ei = server.createobject("easymail.Domain")
ei.Load

if ei.GetUserManagerDomainCount(Session("wem")) < 1 then
	set ei = nothing
	response.redirect "noadmin.asp"
end if
%>

<%
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
	location.href = "dshow_dm_domain.asp?<%=getGRSN() %>&selectdomain=" + document.f1.domainname.value;
}
//-->
</SCRIPT>


<BODY>
<br>
<br>
<FORM ACTION="dsave_dm_domain.asp" METHOD="POST" NAME="f1">
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>���ʼ����</b></font></div>
      </td>
    </tr>
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td width="50%" height="25" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center">����</div>
      </td>
      <td width="50%" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
	<div align='center'>�����ʼ������ʺ�</div>
      </td>
    </tr>
    <tr><td align="center" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid;">
<select name="domainname" class="drpdwn" LANGUAGE=javascript onchange="return domainname_onchange()">
<%
i = 0
allnum = ei.GetUserManagerDomainCount(Session("wem"))

do while i < allnum
	domain = ei.GetUserManagerDomain(Session("wem"), i)

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
	curdomain = ei.GetUserManagerDomain(Session("wem"), 0)
end if
%>
</select>
</td><td align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid;">
<select name="seluser" class="drpdwn">
<option value="">[��]</option>
<%
i = 0
allnum = eu.GetUsersCount

curdomain = LCase(curdomain)
dmuser = LCase(ei.DM_GetUser(curdomain))

do while i < allnum
	eu.GetUserByIndex i, name, domain, comment

	if LCase(domain) = curdomain then
		if dmuser = LCase(name) then
			response.write "<option value='" & name & "' selected>" & name & "</option>"
		else
			response.write "<option value='" & name & "'>" & name & "</option>"
		end if
	end if

	name = NULL
	domain = NULL
	comment = NULL

	i = i + 1
loop

dmuser = NULL
%>
</select>
    </td></tr>
    <tr> 
      <td height="50" colspan="2" align="right" bgcolor="#ffffff">
	<br>
	<input type="submit" value=" ���� " class="Bsbttn">&nbsp;
	<input type="button" value=" ȡ�� " onclick="javascript:location.href='domainright.asp?<%=getGRSN() %>';" class="Bsbttn">
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
		<td width="94%">�˹��ܿ��Զ����û������б���������Ϊ����ص��ʺŽ��м��.
		<br><br>��: �� "���û����� | �޸�" ����������ĳһ����(��: mydomain.com)�µ��ʺ� user Ϊ������ʺź�, ���� user �ʺŽ����Լ����͵��ʼ�������ָ���������ʼ������ʺ��ڱ���һ�ݿ���.
		<br>
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
