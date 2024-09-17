<!--#include file="passinc.asp" --> 

<%
if isadmin() = false and isAccountsAdmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.AdminManager")
'-----------------------------------------
ei.LoadTop100

allnum = ei.TopCount

mode = trim(request("mode"))
tgname = trim(request("tgname"))

if mode <> "" and tgname <> "" then
	if mode = "clean" then
		ei.CleanMailBox(tgname)
		ei.UpdateTop100

		set ei = nothing
		response.redirect "ok.asp?" & getGRSN() & "&gourl=topsize.asp"
	end if

	if mode = "del" then
		dim emuser
		set emuser = Application("em")
		emuser.DelUserByName tgname
		set emuser = nothing

		ei.UpdateTop100

		set ei = nothing
		response.redirect "ok.asp?" & getGRSN() & "&gourl=topsize.asp"
	end if
end if

if mode = "update" then
	ei.UpdateTop100
	set ei = nothing
	response.redirect "ok.asp?" & getGRSN() & "&gourl=topsize.asp"
end if
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<script language="JavaScript">
<!--
var DOM = (document.getElementById) ? 1 : 0;
var NS4 = (document.layers) ? 1 : 0;
var IE4 = 0;
if (document.all)
{
	IE4 = 1;
	DOM = 0;
}

var win = window;   
var n   = 0;

function findIt() {
	if (document.f1.searchstr.value != "")
		findInPage(document.f1.searchstr.value);
}


function findInPage(str) {
var txt, i, found;

if (str == "")
	return false;

if (DOM)
{
	win.find(str, false, true);
	return true;
}

if (NS4) {
	if (!win.find(str))
		while(win.find(str, false, true))
			n++;
	else
		n++;

	if (n == 0)
		alert("未找到指定内容.");
}

if (IE4) {
	txt = win.document.body.createTextRange();

	for (i = 0; i <= n && (found = txt.findText(str)) != false; i++) {
		txt.moveStart("character", 1);
		txt.moveEnd("textedit");
	}

if (found) {
	txt.moveStart("character", -1);
	txt.findText(str);
	txt.select();
	txt.scrollIntoView();
	n++;
}
else {
	if (n > 0) {
		n = 0;
		findInPage(str);
	}
	else
		alert("未找到指定内容.");
	}
}

return false;
}

function del(dname)
{
	if (dname.length > 0)
	{
		if (confirm("确实要删除吗?") == false)
			return ;

		location.href = "topsize.asp?tgname=" + dname + "&mode=del&<%=getGRSN() %>";
	}
}
// -->
</script>

<BODY>
<br>
<FORM ACTION="topsize.asp" METHOD="POST" NAME="f1">
<input type="hidden" name="mode">
  <table width="90%" height="25" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="45%"><font class="s"><b>空间占用列表</b>&nbsp;&nbsp;截止:<%
ei.GetTopTime y, m, d, h, min
response.write y & "-" & m & "-" & d & " "

if h > 9 then
	response.write h
else
	response.write "0" & h
end if

response.write ":"

if min > 9 then
	response.write min
else
	response.write "0" & min
end if
%></font></td>
      <td width="14%"><a href="topsize.asp?mode=update&<%=getGRSN() %>">数据更新</a></td>
<td width="30%" nowrap><input type="text" name="searchstr" class="textbox" size="10">
<input type="button" value="页内查找" onclick="javascript:findIt();" class="sbttn">
</td>
<%
if isadmin() = true then
%>
      <td width="8%"><a href="javascript:location.href='right.asp?<%=getGRSN() %>';">返回</a></td>
<%
else
%>
      <td width="8%"><a href="javascript:location.href='showuser.asp?<%=getGRSN() %>';">返回</a></td>
<%
end if
%>
    </tr>
  </table><br>
</td></tr>
</table>
<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
  <tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
    <td width="8%" align="center" height="25" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>序号</b></font></td>
    <td width="48%" align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>帐号名称</b></font></td>
    <td width="20%" align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>占用空间 (K)</b></font></td>
    <td width="12%" align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>清空</b></font></td>
    <td width="12%" align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>删除帐号</b></font></td>
  </tr>
<%
i = 0

do while i < allnum
	ei.TopGetInfo i, name, msize

	Response.Write "  <tr height='25'>"
	Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & i + 1 & "</td>"

	Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & server.htmlencode(name) & "</a></td>"
	Response.Write "    <td align='right' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & msize & "</td>"
	Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><a href='topsize.asp?tgname=" & server.htmlencode(name) & "&mode=clean&" & getGRSN() & "'><img src='images\del.gif' border='0' alt='清空'></a></td>"
	Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><a href=""javascript:del('" & server.htmlencode(name) & "')""><img src='images\remove.gif' border='0' alt='删除帐号'></a></td>"
	Response.Write "  </tr>" & chr(13)

	name = NULL
	msize = NULL

    i = i + 1
loop
%>
</table>
<br><br><br>
  <div align="center">
    <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr> 
        <td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
        <td width="94%">空间占用管理功能可以将系统中占用空间超过1K的前100个用户列表显示出来(按照空间占用从大到小的顺序), 您可以对相关用户进行:<br>1. 邮箱空间清空(保留帐号, 但删除其所有信件及网络存储数据)<br>2. 删除帐号
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
  </FORM>
<br>
</BODY>
</HTML>

<%
set ei = nothing
%>
