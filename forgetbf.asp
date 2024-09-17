<%
dim webkill
set webkill = server.createobject("easymail.WebKill")
webkill.Load

rip = Request.ServerVariables("REMOTE_ADDR")

if webkill.IsKill(rip) = true then
	set webkill = nothing
	response.redirect "outerr.asp?gourl=default.asp&errstr=" & Server.URLEncode("拒绝IP地址 " & rip & " 访问") & "&" & getGRSN()
end if

set webkill = nothing
errstr = trim(request("errstr"))
%>

<!DOCTYPE html>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<STYLE type=text/css>
<!--
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function window_onload() {
<%
if errstr <> "" then
%>
	alert("<%=server.htmlencode(errstr) %>");
<%
end if
%>
	document.fm1.username.focus();
}
//-->
</script>

<body LANGUAGE=javascript onload="return window_onload()">
<br>
<form method="post" action="forget.asp" name="fm1">
<table width="82%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
忘记密码
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="left" style="padding-left:16px;">
用户名：&nbsp;<input type="text" name="username" maxlength="64" size="30" class="n_textbox">
</td></tr>

<tr><td class="block_top_td" style="height:12px;"></td></tr>

<tr><td align="left" style="background-color:white; padding-top:18px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="default.asp?<%=getGRSN() %>">取消</a>
<a class='wwm_btnDownload btn_blue' href="javascript:document.fm1.submit();">提交</a>
</td></tr>
</table>
</form>
<%
if Application("em_EnableTrap") = true then
%>
<div style="position:absolute; top:0; left:0; z-index:0; display:none;">
<a href="mailto:<%=Application("em_TrapMail") %>"><%=Application("em_TrapMail") %></a>
</div>
<%
end if
%>
</BODY>
</HTML>

<%
function getGRSN()
	dim theGRSN
	Randomize
	theGRSN = Int((9999999 * Rnd) + 1)

	getGRSN = "GRSN=" & CStr(theGRSN)
end function
%>
