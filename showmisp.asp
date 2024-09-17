<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.IspManager")
'-----------------------------------------
ei.Load

allnum = ei.count
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
	if (document.f1.uname.value != "" && document.f1.userver.value != "" && document.f1.uport.value != "" && document.f1.uusername.value != "" && document.f1.upassword.value != "")
	{
		document.f1.mode.value = "add";
		document.f1.submit();
	}
	else
		alert("输入不完整.");
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
<FORM ACTION="savemisp.asp" METHOD="POST" NAME="f1">
<input type="hidden" name="mode">
  <table width="95%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td colspan="7" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>多ISP接收添加</b></font></div>
      </td>
    </tr>
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td width="22%" height="25" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center">名称</div>
      </td>
      <td width="25%" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">服务器地址</div>
      </td>
      <td width="8%" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">端口</div>
      </td>
      <td width="30%" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">用户名</div>
      </td>
      <td width="15%" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">密码</div>
      </td>
    </tr>
    <tr> 
      <td align="center" height="25"> 
        <input type="text" name="uname" size="15" maxlength="64" class="textbox">
      </td>
      <td align="center"> 
        <input type="text" name="userver" size="20" maxlength="64" class="textbox">
      </td>
      <td align="center"> 
        <input type="text" name="uport" size="5" maxlength="5" value="110" class="textbox">
      </td>
      <td align="center"> 
        <input type="text" name="uusername" size="22" maxlength="64" class="textbox">
      </td>
      <td align="center"> 
        <input type="password" name="upassword" size="10" maxlength="64" class="textbox">
      </td>
    </tr>
    <tr>
      <td colspan="7" align="right"><br>
		<input type="button" value=" 添加 " onClick="javascript:add()" name="button" class="Bsbttn">&nbsp;&nbsp;
		<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
      </td>
    </tr>
  </table>
<br><br>
<table width="95%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td colspan="7" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>多ISP接收设置</b></font></div>
      </td>
    </tr>
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td width="5%" height="25" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center">&nbsp;</div>
      </td>
      <td width="19%" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center">名称</div>
      </td>
      <td width="38%" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center">服务器地址</div>
      </td>
      <td width="10%" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">端口</div>
      </td>
      <td width="30%" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">用户名</div>
      </td>
    </tr>
    <%
i = 0

do while i < allnum
	ei.Get i, iname, isev, iport, iusername

	response.write "<tr><td align='center' height='25' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "'>"
	response.write "</td><td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & server.htmlencode(iname)
	response.write "&nbsp;</td><td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & server.htmlencode(isev)
	response.write "&nbsp;</td><td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & iport
	response.write "&nbsp;</td><td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & server.htmlencode(iusername)
	response.write "&nbsp;</td></tr>"

	iname = NULL
	isev = NULL
	iport = NULL
	iusername = NULL

	i = i + 1
loop
%> 
    <tr> 
      <td colspan="7" align="right"><br>
		<input type="button" value=" 删除 " onclick="javascript:del()" class="Bsbttn">&nbsp;&nbsp;
		<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
      </td>
    </tr>
  </table>
  </FORM>
<br>
<div style="position: absolute; left: 35; top: 10;">
<table><tr bgcolor="<%=MY_COLOR_2 %>"><td nowrap style="border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
&nbsp;<a href="showsysinfo.asp?<%=getGRSN() %>#showmisp" style="text-transform: none; text-decoration: none;"><font class="s" color="<%=MY_COLOR_4 %>"><b>启动项设置</b></font>&nbsp;<img src="images\ugo.gif" border="0" align="absbottom"></a>&nbsp;
</td></tr></table>
</div>
</BODY>
</HTML>

<%
set ei = nothing
%>
