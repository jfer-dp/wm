<!--#include file="passinc.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim eka
set eka = server.createobject("easymail.KillAttack")

eka.Load

allnum = eka.Count
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
function writerate(index, irate)
{
	var writemsg = "<select name=\"rate" + index + "\" class=\"drpdwn\">";

	var i = 0;
	for (i; i <= 100; i++)
	{
		if (irate != i)
			writemsg = writemsg + "<option value=\"" + i + "\">" + i + "%</option>"
		else
			writemsg = writemsg + "<option value=\"" + i + "\" selected>" + i + "%</option>"
	}

	document.write(writemsg);
}

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
<br>
<br>
<FORM ACTION="savekillattack.asp" METHOD="POST" NAME="f1">
	<input type="hidden" name="mode">
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td colspan="3" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>�����������</b></font></div>
      </td>
    </tr>
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td width="6%" height="25" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center">&nbsp;</div>
      </td>
      <td width="67%" height="25" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center">IP��ַ��IP��</div>
      </td>
      <td width="27%" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
	<div align='center'>�������</div>
      </td>
    </tr>
<%
i = 0
dim tdname
dim tdrate

do while i < allnum
	eka.Get i, tdname, tdrate

	response.write "<tr><td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "'></td>"
	response.write "<td align='center' height='25' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='text' name='ip" & i & "' size='50' maxlength='30' class='textbox' value='" & tdname & "'>"
	response.write "</td><td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><Script>writerate(" & i & ", " & tdrate & ")</Script></td></tr>" & Chr(13)
	i = i + 1

	tdname = NULL
	tdrate = NULL
loop

if request("mode") <> "" then
	response.write "<tr><td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "'></td><td align='center' height='25' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input name='ip" & i & "' type='text' size='50' maxlength='30' class='textbox'></td>"
	response.write "<td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><Script>writerate(" & i & ", 20)</Script></td></tr>"
end if
%>
  </table>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
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
<br><br>
  <div align="center">
    <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr> 
		<td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
		<td width="94%">����������ƹ���, ����ͨ������ָ��IP��ַ(IP��)�Ľ�����ʴӶ�����ĳЩIP��ַ(IP��)ռ�ù���Ľ�����Դ.
		<br><br>��: ����ͨ���鿴��־��������ĳ��IP��ַ(IP��)�ķ����ر�Ƶ��, ���ֲ�ȷ���Ƿ��ڹ���������ʱ, ���Ϳ���ʹ�ô˹����������������IP��ַ(IP��)�Ľ������, ��������������趨Ϊ 20%ʱ, ��ô���Դ�IP��ַ��ÿ100����������, ��ֻ��20�α�����, ������80�ν����ܾ�.
		<br><br>����Ϊ�򵥵ĸ��ʶ�Ӧ��:
		<br>100%: ����������
		<br>80%: ÿ100������, ֻ����80�ν���
		<br>60%: ÿ100������, ֻ����60�ν���
		<br>40%: ÿ100������, ֻ����40�ν���
		<br>20%: ÿ100������, ֻ����20�ν���
		<br>5%: ÿ100������, ֻ����5�ν���
		<br>0%: �����������
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
&nbsp;<a href="showsysinfo.asp?<%=getGRSN() %>#killattack" style="text-transform: none; text-decoration: none;"><font class="s" color="<%=MY_COLOR_4 %>"><b>����������</b></font>&nbsp;<img src="images\ugo.gif" border="0" align="absbottom"></a>&nbsp;
</td></tr></table>
</div>
</BODY>
</HTML>

<%
set eka = nothing
%>
