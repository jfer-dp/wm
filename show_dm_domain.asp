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
	<td width="37%"><a href="showsysinfo.asp?<%=getGRSN() %>#domainMonitor"><b>����������</b></a></td>
	<td width="30%"><a href="right.asp?<%=getGRSN() %>"><b>����</b></a></td>
	<td width="23%"><b>���ʼ����</b></td>
    </tr>
  </table>
<br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td width="50%" height="25" style="border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center">����</div>
      </td>
      <td width="50%" style="border-top:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
	<div align='center'>�����ʼ������ʺ�</div>
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
	<td align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid;"><select name="username" class="drpdwn"><option value="">[��]</option>
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
	<input type="submit" value=" ���� " class="Bsbttn">&nbsp;
	<input type="button" value=" ȡ�� " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
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
		<td width="94%">�˹��ܿ��Զ��û������б���������Ϊ����ص��ʺŽ��м��.&nbsp;&nbsp;(ע��: ���ñ����ܺ�, �����ԱҲ��ͬʱ���ô˹���)
		<br><br>��: �� "�û����� | �޸�" ����������ĳһ����(��: mydomain.com)�µ��ʺ� user Ϊ������ʺź�, ���� user �ʺŽ����Լ����͵��ʼ�������ָ���������ʼ������ʺ��ڱ���һ�ݿ���.
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
