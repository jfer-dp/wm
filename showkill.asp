<!--#include file="passinc.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.kill")
'-----------------------------------------
ei.Load

allnum = ei.getcount
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>


<SCRIPT LANGUAGE=javascript>
<!--
function del() {
	if (ischeck() == true)
	{
		document.f1.mode.value = "del";
		document.f1.submit();
	}
}

function add() {
	document.f1.mode.value = "add";
	document.f1.submit();
}

function save() {
	document.f1.mode.value = "save";
	document.f1.submit();
}

function ischeck() {
	var i = 0;
	var theObj;

	for(; i<<%=allnum %>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}
//-->
</SCRIPT>


<BODY>
<br><br>
<FORM ACTION="savekill.asp" METHOD="POST" NAME="f1">
   <input type="hidden" name="mode">
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr>
      <td colspan="2" align="right" bgcolor="#ffffff">
	<input type="button" value=" ��� " onclick="javascript:add()" class="Bsbttn">&nbsp;
	<input type="button" value=" ɾ�� " onclick="javascript:del()" class="Bsbttn">&nbsp;
	<input type="button" value=" ���� " onclick="javascript:save()" class="Bsbttn">&nbsp;
	<input type="button" value=" ȡ�� " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	<br>&nbsp;</td>
    </tr>
    <tr bgcolor="<%=MY_COLOR_2 %>">
      <td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>���������� (������IP��ַ���ʼ�βַ���ʼ���ַ)</b></font></div>
      </td>
    </tr>
<%
i = 0

do while i < allnum
	response.write "<tr><td width='9%' height='24' align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "'></td><td width='91%' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input name='kill" & i & "' type='text' value='" & ei.GetKill(i) & "' size='70' maxlength='64' class='textbox'></td></tr>"
	i = i + 1
loop

if request("mode") <> "" then
	response.write "<tr><td width='9%' height='24' align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "'></td><td width='91%' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input name='kill" & i & "' type='text' size='70' maxlength='64' class='textbox'></td></tr>"
end if
%>
    <tr>
      <td colspan="2" align="right" bgcolor="#ffffff">
	<br>
	<input type="button" value=" ��� " onclick="javascript:add()" class="Bsbttn">&nbsp;
	<input type="button" value=" ɾ�� " onclick="javascript:del()" class="Bsbttn">&nbsp;
	<input type="button" value=" ���� " onclick="javascript:save()" class="Bsbttn">&nbsp;
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
        <td width="94%">
		ϵͳ���ܾ�����ָ��IP��ַ���ʼ�βַ���ʼ���ַ�ķ������� (ע��: ������"ϵͳ����"�е�"ϵͳ�ʼ�����"����).
		<br>�������ݿ�����IP��ַ, ��: 192.168.0.56
		<br>���ʼ���ַ @ ���βַ, ��: sex.com
		<br>���ʼ���ַ, ��: bad@sex.com
		<br><br>֧��ͨ�����ʽ. (*: ���ⳤ�ȵ��κ�����.&nbsp;&nbsp;?: һ���ַ����κ�����.)
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
<br><br>
<div style="position: absolute; left: 50; top: 20;">
<table><tr bgcolor="<%=MY_COLOR_2 %>"><td nowrap style="border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
&nbsp;<a href="showsysinfo.asp?<%=getGRSN() %>#showkill" style="text-transform: none; text-decoration: none;"><font class="s" color="<%=MY_COLOR_4 %>"><b>����������</b></font>&nbsp;<img src="images\ugo.gif" border="0" align="absbottom"></a>&nbsp;
</td></tr></table>
</div>
</BODY>
</HTML>

<%
set ei = nothing
%>
