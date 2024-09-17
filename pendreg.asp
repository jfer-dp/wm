<!--#include file="passinc.asp" -->

<%
if isadmin() = false and isAccountsAdmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

dim pr
set pr = server.createobject("easymail.PendRegister")
pr.Load Application("em_SignWaitDays")

allnum = pr.count


if trim(request("page")) = "" then
	page = 0
else
	page = CInt(request("page"))
end if

if page < 0 then
	page = 0
end if


allpage = CInt((allnum - (allnum mod pageline))/ pageline)


if allnum mod pageline <> 0 then
	allpage = allpage + 1
end if

if page >= allpage then
	page = allpage - 1
end if

if page < 0 then
	page = 0
end if


if allpage = 0 then
	allpage = 1
end if
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
function mdel()
{
	document.form1.mdel.value = "1";
	document.form1.action = "deletepr.asp";
	document.form1.submit();
}

function selectpage_onchange()
{
	location.href = "pendreg.asp?<%=getGRSN() %>&page=" + document.form1.page.value;
}


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
	if (document.form1.searchstr.value != "")
		findInPage(document.form1.searchstr.value);
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
//-->
</SCRIPT>

<BODY>
<br>
<form action="pendreg.asp" method=post id=form1 name=form1>
<input type="hidden" name="mdel">
<input type="hidden" name="thispage" value="<%=page %>">
  <table width="98%" height="25" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr> 
	<td width="5%">&nbsp;</td>
	<td width="26%"><font class="s"><b>邮箱申请审批: <%=name & "&nbsp;(" & page+1 & "/" & allpage & ")" %></b></font></td>
	<td width="23%">
<%
if page > 0 then
	response.write "<a href=""pendreg.asp?" & getGRSN() & "&page=" & page - 1 & """><img src='images\prep.gif' border='0' align='absmiddle'></a>&nbsp;"
else
	response.write "<img src='images\gprep.gif' border='0' align='absmiddle'>&nbsp;"
end if
%>
页数: <select name="page" class="drpdwn" size="1" LANGUAGE=javascript onchange="selectpage_onchange()">
<%
i = 0

do while i < allpage
	if i <> page then
		response.write "<option value=""" & i & """>" & i + 1 & "</option>"
	else
		response.write "<option value=""" & i & """ selected>" & i + 1 & "</option>"
	end if
	i = i + 1
loop
%></select>
<%
if page < allpage - 1 then
	response.write "&nbsp;<a href=""pendreg.asp?" & getGRSN() & "&page=" & page + 1 & """><img src='images\nextp.gif' border='0' align='absmiddle'></a>"
else
	response.write "&nbsp;<img src='images\gnextp.gif' border='0' align='absmiddle'>"
end if
%></td>
<td width="27%" nowrap><input type="text" name="searchstr" class="textbox" size="10">
<input type="button" value="页内查找" onclick="javascript:findIt();" class="sbttn">
</td>
<%
if isadmin() = true then
%>
	<td width="9%"><a href="right.asp?<%=getGRSN() %>">返回</a></td>
<%
else
%>
	<td width="9%"><a href="showuser.asp?<%=getGRSN() %>">返回</a></td>
<%
end if
%>
    </tr>
  </table>
</td></tr>
</table>
<br>
<table width="98%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
  <tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
	<td width="5%" nowrap align="center" height="25" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><a href="javascript:mdel()"><img src='images\del.gif' border='0' alt="删除所选项"></a></td>
	<td width="6%" nowrap align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>序号</b></font></td>
	<td width="14%" nowrap align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>申请注册用户名</b></font></td>
	<td width="18%" nowrap align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>域名</b></font></td>
	<td width="12%" nowrap align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>IP地址</b></font></td>
	<td width="11%" nowrap align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>申请时间</b></font></td>
	<td width="12%" nowrap align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>外部Email</b></font></td>
	<td width="6%" nowrap align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>&nbsp;</b></font></td>
	<td width="6%" nowrap align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>&nbsp;</b></font></td>
  </tr>
<%
i = page * pageline
li = 0

do while i < allnum and li < pageline
	si = allnum - i - 1
	pr.GetInfoEx si, reg_username, reg_domain, reg_ip, reg_email, reg_time, accode, viewcode

	Response.Write "  <tr>"
	Response.Write "	<td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & si & "' value='" & accode & "'></td>"
	Response.Write "    <td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & i + 1 & "</td>"

	Response.Write "    <td nowrap align='left' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>&nbsp;<a href='viewreginfo.asp?vid=" & viewcode & "&nm=" & Server.URLEncode(reg_username) & "&" & getGRSN() & "'>" & server.htmlencode(reg_username) & "</a>&nbsp;</td>"
	Response.Write "    <td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>&nbsp;" & reg_domain & "&nbsp;</td>"
	Response.Write "    <td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>&nbsp;" & reg_ip & "&nbsp;</td>"
	Response.Write "    <td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>&nbsp;" & convTimeStr(reg_time) & "&nbsp;</td>"
	Response.Write "    <td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>&nbsp;" & server.htmlencode(reg_email) & "&nbsp;</td>"
	Response.Write "    <td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>&nbsp;<a href='deletepr.asp?id=" & accode & "&thispage=" & page & "&" & getGRSN() & "'>拒绝</a>&nbsp;</td>"
	Response.Write "    <td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>&nbsp;<a href='actionit.asp?accode=" & accode & "&gourl=pendreg.asp&thispage=" & page & "&" & getGRSN() & "'>批准</a>&nbsp;</td>"
	Response.Write "  </tr>" & chr(13)

	reg_username = NULL
	reg_domain = NULL
	reg_ip = NULL
	reg_email = NULL
	reg_time = NULL
	accode = NULL
	viewcode = NULL

    i = i + 1
    li = li + 1
loop

%>
</table>
  </FORM>
<br><br><br>
  <div align="center">
    <table width="98%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr>
		<td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
		<td width="94%">当您在 "系统设置 | 资源使用设置 | 邮箱开通方式选择" 中选中了 "管理员审批后开通" 或 "管理员审批或用户邮件激活后开通" 时, 用户的邮箱申请就会被列表显示在此, 您将可以进行是否允许启用邮箱的审批.
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
set pr = nothing

function convTimeStr(ostr)
	if ostr <> "" then
		convTimeStr = Mid(ostr, 1, 4) & "-" & Mid(ostr, 5, 2) & "-" & Mid(ostr, 7, 2) & " " & Mid(ostr, 9, 2) & ":" & Mid(ostr, 11, 2)
	end if
end function
%>
