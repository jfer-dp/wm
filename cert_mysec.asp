<%
Response.CacheControl = "no-cache" 
%>

<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
dim wemcert
set wemcert = server.createobject("easymail.WebEasyMailCert")

if wemcert.LightHasSecCert(Session("wem")) = false then
	set wemcert = nothing
	Response.Redirect "err.asp?" & getGRSN() & "&errstr=" & Server.URLEncode(a_lang_075) & "&gourl=cert_index.asp"
end if

wemcert.Load Session("wem"), Session("mail")

if trim(request("cleanpw")) = "1" then
	if wemcert.IsNeedSecCertPassword() = false then
		wemcert.CleanSecCertPassword()
	end if

	set wemcert = nothing

	Response.Redirect "ok.asp?gourl=cert_mysec.asp&" & getGRSN()
end if

dim mam
set mam = server.createobject("easymail.AdminManager")
mam.Load
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
function delsec(){
	if (confirm("<%=a_lang_096 %>") == false)
		return ;

	location.href = "cert_del.asp?<%=getGRSN() %>&delmode=allsec";
}

function delsecpass(){
	if (confirm("<%=a_lang_097 %>") == false)
		return ;

	location.href = "cert_mysec.asp?<%=getGRSN() %>&cleanpw=1";
}
//-->
</script>

<body>
<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white" style="margin-top:6px;">
<tr><td colspan="2" style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_057 %>
<%
wemcert.GetCertProperty Session("mail"), c_user_id, c_id, c_fingerprint, c_type, c_size, c_creation_time, c_expiration_time, c_validity, c_is_secret, c_is_expired

if c_user_id = "" or c_id = "" then
%>
&nbsp;<font style="color:#901111; font-weight:normal; font-size:12px;">[<%=a_lang_098 %>]</font>
<%
end if
%>
</td></tr>
<tr><td colspan=2 class="block_top_td" style="height:16px;"></td></tr>
<tr><td>
<table width="94%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr><td width="20%" class="td_l" style="border-top:1px #A5B6C8 solid;">

<%
if c_user_id = "" or c_id = "" then
%>
<%=a_lang_099 %><%=s_lang_mh %>
	</td>
	<td class="td_r" style="border-top:1px #A5B6C8 solid;">
<a href="Javascript:delsec();"><%=a_lang_100 %></a>
	</td>
</tr>
<%
else
%>
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
<%=a_lang_033 %><%=s_lang_mh %>
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
<%=a_lang_101 %><%=s_lang_mh %>
	</td>
	<td class="td_r">
<a href="cert_changepw.asp?<%=getGRSN() %>"><%=a_lang_020 %></a>
	</td>
</tr>

<tr>
	<td class="td_l">
<%=a_lang_102 %><%=s_lang_mh %>
	</td>
	<td class="td_r">
<%
if wemcert.IsNeedSecCertPassword() = false then
%>
<a href="Javascript:delsecpass();"><%=a_lang_103 %></a>
<%
else
%>
<%=a_lang_104 %>
<%
end if
%>
	</td>
</tr>
<%
if mam.Enable_User_Download_Private_Cert = true then
%>
<tr>
	<td class="td_l">
<%=a_lang_090 %><%=s_lang_mh %>
	</td>
	<td class="td_r">
<a href="cert_exp.asp?<%=getGRSN() %>&mode=sec" target='_blank'><%=a_lang_105 %></a>&nbsp;&nbsp;(<%=a_lang_106 %>)
	</td>
</tr>
<%
end if
%>
<tr>
	<td class="td_l">
<%=a_lang_099 %><%=s_lang_mh %>
	</td>
	<td class="td_r">
<a href="Javascript:delsec();"><%=a_lang_100 %></a>
	</td>
</tr>
<%
end if
%>
<tr>
<td colspan="2" align="left" nowrap style="padding-top:16px;">
	<a class='wwm_btnDownload btn_blue' href="cert_index.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
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
set mam = nothing
%>
