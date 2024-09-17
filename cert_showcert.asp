<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
gourl = trim(request("gourl"))
pub_email = trim(request("email"))
page = trim(request("page"))

if pub_email = "" then
	Response.Redirect "err.asp?" & getGRSN() & "&gourl=cert_myothpub.asp?page=" & page
end if

dim wemcert
set wemcert = server.createobject("easymail.WebEasyMailCert")
wemcert.Load Session("wem"), Session("mail")
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<STYLE type=text/css>
<!--
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.td_l {white-space:nowrap; background-color:#EFF7FF; height:26px; text-align:right; border-bottom:1px #A5B6C8 solid; padding-left:12px; color:#444444;}
.td_r {white-space:nowrap; background-color:white; height:26px; text-align:left; border-bottom:1px #A5B6C8 solid; padding-left:6px;}
-->
</STYLE>
</head>

<script type="text/javascript">
<!--
function delpub(){
	if (confirm("<%=a_lang_131 %> <%=pub_email %><%=a_lang_132 %>") == false)
		return ;
<%
if gourl = "" then
%>
	location.href = "cert_del.asp?delmode=pub&pub_email=<%=Server.URLEncode(pub_email) %>&retstr=<%=Server.URLEncode("cert_myothpub.asp?page=" & page) %>&<%=getGRSN() %>";
<%
else
%>
	location.href = "cert_del.asp?delmode=pub&pub_email=<%=Server.URLEncode(pub_email) %>&<%=getGRSN() %>&retstr=<%=Server.URLEncode(gourl) %>";
<%
end if
%>
}
//-->
</script>

<body>
<%
wemcert.GetCertProperty pub_email, c_user_id, c_id, c_fingerprint, c_type, c_size, c_creation_time, c_expiration_time, c_validity, c_is_secret, c_is_expired
%>
<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white" style="margin-top:6px;">
<tr><td colspan="2" style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_133 %>
</td></tr>
<tr><td colspan=2 class="block_top_td" style="height:16px;"></td></tr>
<tr><td>


<table width="94%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr><td width="20%" class="td_l" style="border-top:1px #A5B6C8 solid;">
<%=a_lang_072 %><%=s_lang_mh %>
	</td>
	<td class="td_r" style="border-top:1px #A5B6C8 solid;">
<%=server.htmlencode(c_user_id) %>
	</td>
</tr>

<tr>
	<td class="td_l">
ID<%=s_lang_mh %>
	</td>
	<td class="td_r" style="color:#901111;">
<%=c_id %>
	</td>
</tr>

<tr>
	<td class="td_l">
<%=a_lang_077 %><%=s_lang_mh %>
	</td>
	<td class="td_r" style="color:#901111;">
<%=c_fingerprint %>
	</td>
</tr>

<tr>
	<td class="td_l">
<%=a_lang_078 %><%=s_lang_mh %>
	</td>
	<td class="td_r">
<%=c_type & " (" & c_size & " Bits)" %>
	</td>
</tr>

<tr>
	<td class="td_l">
<%=a_lang_079 %><%=s_lang_mh %>
	</td>
	<td class="td_r">
<%=c_creation_time %>
	</td>
</tr>

<tr>
	<td class="td_l">
<%=a_lang_080 %><%=s_lang_mh %>
	</td>
	<td class="td_r">
<%
if c_is_expired = false then
	Response.Write c_expiration_time
else
	Response.Write "<font color='#901111'>" & c_expiration_time & "</font>"
end if
%>
	</td>
</tr>

<tr>
	<td class="td_l">
<%=a_lang_081 %><%=s_lang_mh %>
	</td>
	<td class="td_r">
<%
if c_validity = true then
	Response.Write a_lang_082
else
	Response.Write "<font color='#901111'>" & a_lang_083 & "</font>"
end if
%>
	</td>
</tr>

<tr>
	<td class="td_l">
<%=a_lang_084 %><%=s_lang_mh %>
	</td>
	<td class="td_r">
<%
if c_is_secret = true then
	Response.Write a_lang_085
else
	Response.Write a_lang_086
end if
%>
	</td>
</tr>

<tr>
	<td class="td_l">
<%=a_lang_087 %><%=s_lang_mh %>
	</td>
	<td class="td_r">
<%
if c_is_expired = false then
	Response.Write a_lang_088
else
	Response.Write "<font color='#901111'>" & a_lang_089 & "</font>"
end if
%>
	</td>
</tr>

<tr>
	<td class="td_l">
<%=a_lang_090 %><%=s_lang_mh %>
	</td>
	<td class="td_r">
<a href="cert_exp.asp?<%=getGRSN() %>&mode=pub&pub_email=<%=Server.URLEncode(pub_email)%>" target='_blank'><%=a_lang_134 %></a>
	</td>
</tr>

<tr>
	<td class="td_l">
<%=a_lang_099 %><%=s_lang_mh %>
	</td>
	<td class="td_r">
<a href="Javascript:delpub();"><%=a_lang_135 %></a>
	</td>
</tr>

<tr><td colspan="2" class="block_top_td" style="height:12px;"></td></tr>

<tr>
<td nowrap colspan="2" align="left">
<%
if gourl = "" then
%>
	<a class='wwm_btnDownload btn_blue' href="cert_myothpub.asp?<%=getGRSN() %>&page=<%=page %>"><< <%=s_lang_return %></a>
<%
else
	if InStr(1, gourl, "?") = 0 Then
		gourl = gourl & "?" & getGRSN()
	else
		gourl = gourl & "&" & getGRSN()
	end if
%>
	<a class='wwm_btnDownload btn_blue' href="<%=gourl %>"><< <%=s_lang_return %></a>
<%
end if
%>
</td></tr>
</table>
</td></tr>
</table>
</body>
</html>

<%
c_user_id = NULL
c_id = NULL
c_fingerprint = NULL
c_type = NULL
c_size = NULL
c_creation_time = NULL
c_expiration_time = NULL
c_validity = NULL
c_is_secret = NULL
c_is_expired = NULL

set wemcert = nothing
%>
