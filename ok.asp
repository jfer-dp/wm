<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
errstr = trim(request("errstr"))
gourl = trim(request("gourl"))
oreturl = trim(request("returl"))
sws = trim(request("sws"))
pgourl = trim(request("pgourl"))
lurl = trim(request("lurl"))

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

if pgourl <> "" and Len(pgourl) > 0 then
	gourl = gourl & "&gourl=" & Server.URLEncode(pgourl)
end if
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
</HEAD>

<script type="text/javascript">
var seconds = 3;
var t_id;

function window_onload()
{
	t_id = window.setInterval(redirection, 1000);
	document.getElementById("okbt").innerHTML = "<%=s_lang_return %> (" + seconds + ")";
}

function redirection()
{
	seconds--;

	if (seconds > 0)
		document.getElementById("okbt").innerHTML = "<%=s_lang_return %> (" + seconds + ")";
	else
		document.getElementById("okbt").innerHTML = "<%=s_lang_return %>";

	if (seconds < 1)
	{
		window.clearInterval(t_id);
		time_ck();
	}
}

function location_href(url, is_parent) {
	if (is_parent == 1)
		parent.location.href = url;
	else
<%
if lurl <> "1" then
%>
		location.href = url;
<%
else
%>
		location.href = parent.f1.document.leftval.purl.value;
<%
end if
%>
}
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
<table width="90%" align="center" border="0" cellspacing="0" cellpadding="0" style="margin-top:40px;">
	<tr style="background:#EFF7FF; color:#104A7B;">
<%
if errstr = "" then
%>
	<td height="30" style="border:1px solid #8CA5B5;">&nbsp;&nbsp;<%=s_lang_0314 %><%=s_lang_jh %></td>
<%
else
%>
	<td height="30" style="border:1px solid #8CA5B5;">&nbsp;&nbsp;<%
if sws = "1" then
	Response.Write errstr
else
	Response.Write "<b>" & server.htmlencode(errstr) & "</b>"
end if
%><%=s_lang_jh %>
	</td>
<%
end if
%>
	</tr>
	<tr><td style="border-bottom:1px solid #8CA5B5; height:24px;">&nbsp;</td></tr>
	<tr><td style="height:24px;">&nbsp;</td></tr>
	<tr><td align="right" style="padding-right:30px;">
<a class="wwm_btnDownload btn_blue" id="okbt" href="javascript:time_ck();"><%=s_lang_return %></a>
<%
if gourl = "" and lurl <> "1" then
%>
<script type="text/javascript">
function time_ck() {
	history.back();
}
</script>
<%
else
	if Left(gourl, 11) <> "welcome.asp" then
%>
<script type="text/javascript">
function time_ck() {
	location_href('<%=gourl%>', 0);
}
</script>
<%
	else
%>
<script type="text/javascript">
function time_ck() {
	location_href('<%=gourl%>', 1);
}
</script>
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
