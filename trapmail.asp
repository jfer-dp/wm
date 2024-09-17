<!--#include file="passinc.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.domain")
ei.Load

set sysinfo = server.createobject("easymail.sysinfo")
sysinfo.Load

'-----------------------------------------
dim eu
set eu = Application("em")

if trim(request("wmode")) = "Save" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	TrapMail = trim(request("TrapMail"))

	se = InStr(1, TrapMail, "@")
	isok = false
	if se > 0 Then
		m_domain = Mid(TrapMail, se + 1)

		if LCase(m_domain) <> "system.mail" and ei.IsDomain(m_domain) = true then
			isok = true
		end if
	end if

	if se > 0 Then
		if eu.isUser(TrapMail) = true or eu.isUser(Left(TrapMail, se - 1)) = true then
			isok = false
		end if
	end if

	if isok = false and Len(TrapMail) > 0 then
		set eu = nothing
		set sysinfo = nothing
		set ei = nothing

		Response.Redirect "err.asp?" & getGRSN() & "&gourl=trapmail.asp"
	end if


	sysinfo.TrapMail = TrapMail

	if trim(request("EnableTrap")) <> "" then
		sysinfo.EnableTrap = true
	else
		sysinfo.EnableTrap = false
	end if

	sysinfo.Save

	Application("em_EnableTrap") = sysinfo.EnableTrap

	if sysinfo.EnableTrap = true then
		Application("em_TrapMail") = sysinfo.TrapMail
	else
		Application("em_TrapMail") = ""
	end if

	set eu = nothing
	set sysinfo = nothing
	set ei = nothing

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=trapmail.asp"
end if


domainname = trim(request("domainname"))
newTrapMail = ""

if Len(domainname) > 3 and trim(request("wmode")) = "rndCreate" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	newTrapMail = CreateRandString() & "@" & domainname
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
function isCharsInBag (s, bag)
{
	var i,c;
	for (i = 0; i < s.length; i++)
	{
		c = s.charAt(i);

		if (bag.indexOf(c) == -1)
			return false;
	}

	return true;
}

function ischinese(s)
{
	if (s.charAt(s.length - 1) == '.')
		return true;

	var badChar = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_.@";

	return !isCharsInBag(s, badChar);
}

function rndCreate() {
	document.f1.wmode.value = "rndCreate";
	document.f1.submit();
}

function save() {
	if (ischinese(document.f1.TrapMail.value) == true)
	{
		alert("邮件地址输入字符非法");
		document.f1.TrapMail.focus();
		return ;
	}
	else
	{
		document.f1.wmode.value = "Save";
		document.f1.submit();
	}
}
//-->
</SCRIPT>


<BODY>
<br>
<FORM ACTION="trapmail.asp" METHOD="POST" NAME="f1">
<input type="hidden" name="wmode">
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="2%" height="25">&nbsp;</td>
	<td width="47%"><input type="checkbox" name="EnableTrap" value="checkbox"<% if sysinfo.EnableTrap = true then Response.Write " checked"%>>启用垃圾邮件陷阱功能</td>
	<td width="29%"><a href="right.asp?<%=getGRSN() %>"><b>返回</b></a></td>
	<td width="22%" nowrap><b>垃圾邮件陷阱</b></td>
    </tr>
  </table>
<br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td height="32" style="border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center">虚拟诱饵邮件地址</div>
      </td>
    </tr>
	<tr><td align="center" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<input type="text" name="TrapMail" class="textbox" size="35" maxlength="64" value="<%
if Len(newTrapMail) > 13 then
	Response.Write newTrapMail
else
	Response.Write sysinfo.TrapMail
end if
%>">
<%
if Len(sysinfo.TrapMail) = 0 then
%>
&nbsp;&nbsp;<select name="domainname" class="drpdwn">
<%
i = 0
allnum = ei.GetCount()

do while i < allnum
	domain = ei.GetDomain(i)

	if LCase(domain) <> "system.mail" then
		if domain <> domainname then
			response.write "<option value='" & server.htmlencode(domain) & "'>" & server.htmlencode(domain) & "</option>" & Chr(13)
		else
			response.write "<option value='" & server.htmlencode(domain) & "' selected>" & server.htmlencode(domain) & "</option>" & Chr(13)
		end if
	end if

	domain = NULL

	i = i + 1
loop
%>
</select>
<%
	Response.Write "<input type=""button"" value=""随机生成"" class=""sbttn"" language=javascript onClick=""rndCreate()"">"
end if
%>
      </td>
    </tr>
    <tr> 
      <td height="50" colspan="2" align="right" bgcolor="#ffffff">
	<br>
	<input type="button" value=" 保存 " class="Bsbttn" language=javascript onClick="save()">&nbsp;
	<input type="button" value=" 返回 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
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
		<td width="94%">此功能可以虚拟一个诱饵邮件地址做为陷阱, 将所有发往这个邮址的邮件都判为垃圾邮件.
		<br><br>注意: 必须选定非 system.mail 的其他有效域名做为此诱饵邮件地址的域名.
		<br><br>为了提高此功能防垃圾邮件的效果, 请在 Internet 上广为传播此诱饵邮件地址.
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
function CreateRandString()
	do while 1
		i = 0
		temp_str = ""

		do while i < 10
			temp_str = temp_str & GetChar()
			i = i + 1
		loop

		if eu.isUser(temp_str) = false then
			exit do
		end if
	loop

	CreateRandString = temp_str
end function

function GetChar()
	Randomize
	r_Array = "abcdefghijklmnopqrstuvwxyz1234567890"
	r_k = Int((36 * Rnd) + 1)
	GetChar = Mid(r_Array, r_k, 1)
end function


set eu = nothing
set ei = nothing
set sysinfo = nothing
%>
