<!--#include file="passinc.asp" --> 

<%
dim esinfo
set esinfo = server.createobject("easymail.sysinfo")
esinfo.Load

if esinfo.enableCatchAll = false then
	set esinfo = nothing
	response.redirect "noadmin.asp"
end if



dim ei
set ei = server.createobject("easymail.domain")
'-----------------------------------------
ei.Load

if ei.GetUserManagerDomainCount(Session("wem")) < 1 then
	set esinfo = nothing
	set ei = nothing
	response.redirect "noadmin.asp"
end if


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
function domainname_onchange() {
	location.href = "dshow_dca_domain.asp?<%=getGRSN() %>&selectdomain=" + document.f1.domainname.value;
}
//-->
</SCRIPT>


<BODY>
<br>
<br>
<FORM ACTION="dsave_dca_domain.asp" METHOD="POST" NAME="f1">
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>域邮件Catch All</b></font></div>
      </td>
    </tr>
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td width="50%" height="25" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center">域名</div>
      </td>
      <td width="50%" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
<%
if esinfo.enableCatchToOut = false then
	response.write "<div align='center'>接收帐号(<font color='#FF3333'>仅限系统内帐号</font>)</div>"
else
	response.write "<div align='center'>接收帐号或外部邮件地址</div>"
end if
%>
      </td>
    </tr>
    <tr><td align="center" height="28">
<select name="domainname" class="drpdwn" LANGUAGE=javascript onchange="return domainname_onchange()">
<%
i = 0

allnum = ei.GetUserManagerDomainCount(Session("wem"))

do while i < allnum
	domain = ei.GetUserManagerDomain(Session("wem"), i)

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
	curdomain = ei.GetUserManagerDomain(Session("wem"), 0)
end if
%>
</select>
</td><td align="center">
<%
	response.write "<input type='text' name='user' size='30' maxlength='64' class='textbox' value='" & ei.DCA_GetUser(curdomain) & "'>"
%>
</td></tr>
  </table>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr> 
      <td colspan="2" align="right" bgcolor="#ffffff">
	<br>
	<input type="submit" value=" 保存 " class="Bsbttn">&nbsp;
	<input type="button" value=" 取消 " onclick="javascript:location.href='domainright.asp?<%=getGRSN() %>';" class="Bsbttn">
      </td>
    </tr>
  </table>
  </FORM>
<br>
</BODY>
</HTML>

<%
set ei = nothing
set esinfo = nothing
%>
