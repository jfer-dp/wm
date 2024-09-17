<!--#include file="passinc.asp" --> 

<%
dim dm
set dm = server.createobject("easymail.Domain")
dm.Load

if dm.GetUserManagerDomainCount(Session("wem")) < 1 then
	set dm = nothing
	response.redirect "noadmin.asp"
end if

set dm = nothing


dim ei
set ei = server.createobject("easymail.sysinfo")
ei.Load

dim mam
set mam = server.createobject("easymail.AdminManager")
mam.Load
%>

<html>
<head>
<TITLE>WinWebMail</TITLE>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<LINK href="images\hwem.css" rel=stylesheet>
</head>

<script type="text/javascript" src="images/sc_left.js"></script>

<script language="JavaScript">
<!--
function theright(rurl) {
	var mrstr = String(Math.random());

	location.href = rurl + "?GRSN=" + mrstr.substring(2, 10);
}

function therightfol(rurl) {
	var mrstr = String(Math.random());

	location.href = rurl + "&GRSN=" + mrstr.substring(2, 10);
}
//-->
</script>


<body leftmargin="10" rightmargin="2" topmargin="1">
<div align="center">
<br><br><br><font class="s" color="<%=MY_COLOR_4 %>"><b>
  <table width="70%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr> 
      <td height="30" style='border-bottom:1px <%=MY_COLOR_1 %> solid;border-left:1px <%=MY_COLOR_1 %> solid;border-right:1px <%=MY_COLOR_1 %> solid;'>
        <div align="center"><font class="s"><b>域管理员设置项</b></div>
      </td>
    </tr>
  </table>
<table width="70%" border="0" align="center" cellspacing="0" bgcolor="#ffffff">
    <tr> 
      <td height="60" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b><a href="JavaScript:theright('feast.asp')">域节日设置</a></b></font><br><br>设置可以在效率手册中显示的自订制节日信息</div>
      </td>
    </tr>
<%
if ei.enableCatchAll = true then
%>
    <tr> 
      <td height="60" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b><a href="JavaScript:theright('dshow_dca_domain.asp')">域邮件Catch All</a></b></font><br><br>捕获所有发往本域的邮件(无限别名功能)</div>
      </td>
    </tr>
<%
end if

if ei.enableDomainMonitor = true then
%>
    <tr> 
      <td height="60" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b><a href="JavaScript:theright('dshow_dm_domain.asp')">域邮件监控</a></b></font><br><br>监控域用户收、发的邮件</div>
      </td>
    </tr>
<%
end if

if mam.Enable_DomainAdmin_SetWelcomeMsg = true then
%>
    <tr> 
      <td height="60" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b><a href="JavaScript:theright('showwelcome.asp')">域欢迎邮件</a></b></font><br><br>管理用户申请帐号后发给用户的域欢迎邮件内容</div>
      </td>
    </tr>
<%
end if

if mam.Enable_DomainAdmin_SetAdvertisingMsg = true then
%>
    <tr> 
      <td height="60" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b><a href="JavaScript:theright('showadv.asp')">域广告</a></b></font><br><br>管理每个域通过浏览器所发邮件尾部追加的信息</div>
      </td>
    </tr>
<%
end if

if mam.Enable_DomainAdmin_SendDomainListMail = true then
%>
    <tr> 
      <td height="60" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b><a href="JavaScript:therightfol('wframe.asp?mode=domainlistmail')">域邮件群发</a></b></font><br><br>允许向所辖域发送群发邮件</div>
      </td>
    </tr>
<%
end if
%>
    <tr> 
      <td height="60" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b><a href="JavaScript:theright('browmailinglist.asp')">邮件列表管理</a></b></font><br><br>管理系统邮件列表信息</div>
      </td>
    </tr>
  </table>
</div>
<br><br>
</body>
</html>

<%
set ei = nothing
set mam = nothing
%>
