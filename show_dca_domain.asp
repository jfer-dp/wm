<!--#include file="passinc.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim esinfo
set esinfo = server.createobject("easymail.sysinfo")

esinfo.Load


dim ei
set ei = server.createobject("easymail.domain")
'-----------------------------------------
ei.DCA_Load

allnum = ei.getcount
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>


<BODY>
<br>
<br>
<FORM ACTION="save_dca_domain.asp" METHOD="POST" NAME="f1">
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>���ʼ�Catch All</b></font></div>
      </td>
    </tr>
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td width="55%" height="25" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center">����</div>
      </td>
      <td width="45%" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
<%
if esinfo.enableCatchToOut = false then
	response.write "<div align='center'>�����ʺ�(<font color='#FF3333'>����ϵͳ���ʺ�</font>)</div>"
else
	response.write "<div align='center'>�����ʺŻ��ⲿ�ʼ���ַ</div>"
end if
%>
      </td>
    </tr>
<%
i = 0
dim tdname

do while i < allnum
	tdname = ei.GetDomain(i)

	response.write "<tr><td align='center' height='25' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='text' name='idomain" & i & "' size='40' maxlength='64' class='textbox' value='" & tdname & "' readonly>"
	response.write "</td><td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='text' name='user" & i & "' size='30' maxlength='64' class='textbox' value='" & ei.DCA_GetUser(tdname) & "'></td></tr>"
	i = i + 1
loop

tdname = NULL
%>
  </table>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr> 
      <td colspan="2" align="right" bgcolor="#ffffff">
	<br>
	<input type="submit" value=" ���� " class="Bsbttn">&nbsp;
	<input type="button" value=" ȡ�� " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
      </td>
    </tr>
  </table>
  </FORM>
<br>
  <div align="center">
    <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr> 
        <td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
		<td width="94%">��Ҳ����Ϊ���ޱ�������.&nbsp;&nbsp;(ע��: ���ñ����ܺ�, �����ԱҲ��ͬʱ���ô˹���)
		<br><br>����Ա����Ϊĳ������(��: mydomain.com)ָ��һ�������ʺ�, ����ʺžͿ��������������з����������(��: mydomain.com)�²��������ʺŵ��ʼ�.
		<br><br>����: �� mydomain.com ����, ֻ��һ�� master �ʺ�, ��: master@mydomain.com. ��δʹ�ñ�����ǰ���з�������: info@mydomain.com, poster@mydomain.com ���ʼ������޷��յ�.
		<br>�����ʹ�ñ����ܲ���ָ�������ʺ�Ϊ master ʱ, ��ô��ʹ������ info �� poster �ʺ�, master �ʺ�Ҳ�����յ����� info@mydomain.com, poster@mydomain.com ��������ʼ�.
		<br>
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
<br>
<div style="position: absolute; left: 35; top: 10;">
<table><tr bgcolor="<%=MY_COLOR_2 %>"><td nowrap style="border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
&nbsp;<a href="showsysinfo.asp?<%=getGRSN() %>#catchall" style="text-transform: none; text-decoration: none;"><font class="s" color="<%=MY_COLOR_4 %>"><b>����������</b></font>&nbsp;<img src="images\ugo.gif" border="0" align="absbottom"></a>&nbsp;
</td></tr></table>
</div>
</BODY>
</HTML>

<%
set ei = nothing
set esinfo = nothing
%>
