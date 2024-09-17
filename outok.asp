<%
errstr = trim(request("errstr"))
gourl = trim(request("gourl"))
oreturl = trim(request("returl"))

if gourl <> "" and InStr(1, gourl, "GRSN=") = 0 then
	if InStr(1, gourl, "?") = 0 Then
		gourl = gourl & "?" & getGRSN()
	else
		gourl = gourl & "&" & getGRSN()
	end if
end if

if gourl <> "" and Len(oreturl) > 0 then
	gourl = gourl & "&returl=" & Server.URLEncode(oreturl)
end if
%>

<!DOCTYPE html>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
</HEAD>

<BODY>
<table width="90%" align="center" border="0" cellspacing="0" cellpadding="0" style="margin-top:40px;">
	<tr style="background:#EFF7FF; color:#104A7B;">
<%
if errstr = "" then
%>
	<td height="30" style="border:1px solid #8CA5B5;">&nbsp;&nbsp;您的操作<b>成功</b>。</td>
<%
else
%>
	<td height="30" style="border:1px solid #8CA5B5;">&nbsp;&nbsp;<%=server.htmlencode(errstr) %>。</td>
<%
end if
%>
	</tr>
	<tr><td style="border-bottom:1px solid #8CA5B5; height:24px;">&nbsp;</td></tr>
	<tr><td style="height:24px;">&nbsp;</td></tr>
	<tr><td align="right" style="padding-right:30px;">
<%
if gourl = "" then
%>
<a class="wwm_btnDownload btn_blue" id="okbt" href="javascript:history.back();">返回</a>
<%
else
	if Left(gourl, 11) <> "welcome.asp" then
%>
<a class="wwm_btnDownload btn_blue" id="okbt" href="javascript:location.href='<%=gourl%>';">返回</a>
<%
	else
%>
<a class="wwm_btnDownload btn_blue" id="okbt" href="javascript:parent.location.href='<%=gourl%>';">返回</a>
<%
	end if
end if
%>
	</td></tr></table>
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
