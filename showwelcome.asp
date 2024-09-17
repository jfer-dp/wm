<!--#include file="passinc.asp" --> 

<%
if isadmin() = false then
	dim mam
	set mam = server.createobject("easymail.AdminManager")
	mam.Load

	if mam.Enable_DomainAdmin_SetWelcomeMsg = false then
		set mam = nothing
		response.redirect "noadmin.asp"
	end if

	set mam = nothing
end if


isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

dim dm
set dm = server.createobject("easymail.Domain")
dm.Load

if dm.GetUserManagerDomainCount(Session("wem")) < 1 then
	if isadmin() = false then
		set dm = nothing
		response.redirect "noadmin.asp"
	end if
end if

'-----------------------------------------
dim ei
set ei = server.createobject("easymail.Domain_Welcome_Msg")
ei.Load
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
	location.href = "showwelcome.asp?<%=getGRSN() %>&selectdomain=" + document.f1.domainname.value;
}


function changeSystemWelcome_onclick() {
	if (document.f1.wsubject.disabled == true)
	{
		document.f1.wsubject.disabled = false;
		document.f1.wtext.disabled = false;
	}
	else
	{
		document.f1.wsubject.disabled = true;
		document.f1.wtext.disabled = true;
	}
}


function all2system()
{
	if (confirm("是否清除所有域的欢迎邮件内容, 而使用系统欢迎邮件内容? "))
	{
		document.f1.cleanall.value = "yes";
		document.f1.submit();
	}
}
//-->
</SCRIPT>


<BODY>
<br><br>
<FORM ACTION="savewelcome.asp" METHOD=POST NAME="f1">
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="5%" height="25">&nbsp;</td>
      <td width="36%"><b>选择域名</b>:&nbsp;<select name="domainname" class="drpdwn" LANGUAGE=javascript onchange="return domainname_onchange()">
<%
i = 0

if isadmin() = false then
	allnum = dm.GetUserManagerDomainCount(Session("wem"))

	do while i < allnum
		domain = dm.GetUserManagerDomain(Session("wem"), i)

		if domain <> trim(request("selectdomain")) then
			response.write "<option value='" & server.htmlencode(domain) & "'>" & server.htmlencode(domain) & "</option>" & Chr(13)
		else
			curdomain = domain
			response.write "<option value='" & server.htmlencode(domain) & "' selected>" & server.htmlencode(domain) & "</option>" & Chr(13)
		end if

		domain = NULL

		i = i + 1
	loop
else
	allnum = dm.GetCount()

	do while i < allnum
		domain = dm.GetDomain(i)

		if domain <> trim(request("selectdomain")) then
			response.write "<option value='" & server.htmlencode(domain) & "'>" & server.htmlencode(domain) & "</option>" & Chr(13)
		else
			curdomain = domain
			response.write "<option value='" & server.htmlencode(domain) & "' selected>" & server.htmlencode(domain) & "</option>" & Chr(13)
		end if

		domain = NULL

		i = i + 1
	loop
end if


if curdomain = "" then
	if isadmin() = false then
		curdomain = dm.GetUserManagerDomain(Session("wem"), 0)
	else
		curdomain = dm.GetDomain(0)
	end if
end if

haveitdm = ei.haveit(curdomain)

ei.Get curdomain, subject, text
%>
</select>
</td>
      <td width="22%"><a href="javascript:all2system()">全部使用系统欢迎邮件内容</a></td>
    </tr>
  </table>
<br>
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="30" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b><%=server.htmlencode(curdomain) %> 域欢迎邮件</b></font>
		</div>
      </td>
    </tr>
    <tr><td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" name="changeSystemWelcome" LANGUAGE=javascript onclick="return changeSystemWelcome_onclick()"<%if haveitdm = false then response.write " checked" end if %>>此域使用系统欢迎邮件内容
	</td></tr>
    <tr>
      <td width="26%" height="30">
        <div align="center">主 题:</div>
      </td>
      <td width="74%">
        <input name="wsubject" type="text" value="<%=subject %>" size="50" maxlength="250" class='textbox'<%if haveitdm = false then response.write " disabled" end if %>>
        </td>
    </tr>
    <tr>
      <td colspan="2"> 
        <div align="center">  
          <textarea name="wtext" cols="<%
if isMSIE = true then
	Response.Write "72"
else
	Response.Write "62"
end if
%>" rows="8" class='textarea'<%if haveitdm = false then response.write " disabled" end if %>><%=text %></textarea>
          </div>
          <br>
      </td>
    </tr>
    <tr> 
	<td colspan="2" align="right" bgcolor="#ffffff">
	<br><input type="submit" value=" 保存 " class="Bsbttn">&nbsp;&nbsp;
<%
if isadmin() = false then
%>
	<input type="button" value=" 取消 " onclick="javascript:location.href='domainright.asp?<%=getGRSN() %>';" class="Bsbttn">
<%
else
%>
	<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
<%
end if
%>
	</td>
	</tr>
  </table>
<input name="cleanall" type="hidden" value="">
<input name="curdomain" type="hidden" value="<%=curdomain %>">
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
		<td width="94%">您可以为每个域创建不同的新用户欢迎邮件内容, 也可以使用系统欢迎邮件内容.
		<br><br>此邮件将会被投递到本域中每一个新建用户的邮箱中.
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
curdomain = NULL
subject = NULL
text = NULL

set ei = nothing
set dm = nothing
%>
