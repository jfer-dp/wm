<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.logs")
'-----------------------------------------

ei.load

allnum = ei.LogCount
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
function removeall() {
	if (confirm("确实要删除吗?") == false)
		return ;

	document.f1.mode.value = "removeall";
	document.f1.submit();
}

function logdel() {
	if (ischeck() == true)
	{
		if (confirm("确实要删除吗?") == false)
			return ;

		document.f1.submit();
	}
}

function allcheck_onclick() {
	if (document.f1.allcheck.checked == true)
		checkall(true);
	else
		checkall(false);
}

function checkall(check) {
	var i = 0;
	var theObj;

	for(; i<<%=allnum %>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			theObj.checked = check;
	}
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
<FORM ACTION="dellog.asp" METHOD="POST" NAME="f1">
<input type="hidden" name="mode">
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr>
      <td colspan="7" align="right" bgcolor="#ffffff"><br>
		<input type="button" value=" 删除 " onclick="javascript:logdel()" class="Bsbttn">&nbsp;
		<input type="button" value="删除全部" onclick="javascript:removeall()" class="Bsbttn">&nbsp;
		<input type="button" value=" 返回 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
		<br>&nbsp;
      </td>
    </tr>
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="6" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>日志管理 (<%=ei.LogCount %>)</b></font></div>
	</td>
    </tr>
    <tr bgcolor="<%=MY_COLOR_2 %>">
      <td width="5%" height="25" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()"></div>
      </td>
      <td width="58%" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center">日志名</div>
      </td>
      <td colspan="4" width="37%" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center">长度 (K)</div>
        </td>
    </tr>
    <%
i = allnum - 1

do while i >= 0
	ei.getLogInfo i, name, size

	response.write "<tr><td align='center' height='25' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "' value='" & name & "'>"
	response.write "</td><td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><a href='showlog.asp?" & getGRSN() & "&index="& i & "'>" & name
	response.write "</a></td><td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & int(size/1000)
	response.write "K</td></tr>"

	name = NULL
	size = NULL

	i = i - 1
loop
%> 
    <tr>
      <td colspan="7" align="right" bgcolor="#ffffff"><br>
		<input type="button" value=" 删除 " onclick="javascript:logdel()" class="Bsbttn">&nbsp;
		<input type="button" value="删除全部" onclick="javascript:removeall()" class="Bsbttn">&nbsp;
		<input type="button" value=" 返回 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
      </td>
    </tr>
  </table>
  </FORM>
<br>
</BODY>
</HTML>

<%
set ei = nothing
%>
