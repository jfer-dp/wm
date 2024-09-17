<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
page = trim(request("page"))

if trim(request("pub_email")) <> "" and trim(request("message")) <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim wemcert
	set wemcert = server.createobject("easymail.WebEasyMailCert")
	wemcert.Load Session("wem"), Session("mail")

	isok = wemcert.Import_Buffer_Pub_Cert(trim(request("message")), trim(request("pub_email")))
	set wemcert = nothing

	if isok <> 0 then
		Response.Redirect "err.asp?" & getGRSN() & "&errstr=" & Server.URLEncode(a_lang_107 & geterror(isok)) & "&gourl=cert_myothpub.asp?page=" & page
	else
		Response.Redirect "ok.asp?" & getGRSN() & "&errstr=" & Server.URLEncode(a_lang_108) & "&gourl=cert_myothpub.asp?page=" & page
	end if
end if
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
-->
</STYLE>
</head>

<script type="text/javascript">
<!--
function save_onclick() {
	if (document.f1.pub_email.value == "")
	{
		alert("<%=a_lang_109 %>");
		document.f1.pub_email.focus();
	}
	else
	{
		if (document.f1.message.value.length < 20)
		{
			alert("<%=a_lang_110 %>");
			document.f1.message.focus();
		}
		else
		{
			if (document.f1.message.value.length > 10000)
				alert("<%=a_lang_111 %>");
			else
				document.f1.submit();
		}
	}
}

function back_onclick() {
	location.href = "cert_myothpub.asp?<%=getGRSN() %>&page=<%=page %>";
}
//-->
</SCRIPT>

<body>
<form method="post" action="cert_pubstradd.asp" name="f1">
<input type="hidden" name="page" value="<%=page %>">

<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white" style="margin-top:6px;">
<tr><td colspan="2" style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_112 %>
</td></tr>
<tr><td colspan=2 class="block_top_td" style="height:12px; _height:14px;"></td></tr>
<tr>
	<td height="23" nowrap align="left" width="20%" style="padding-left:8px; padding-top:4px;">
<%=a_lang_113 %><%=s_lang_mh %>
	</td>
	<td align="left"><input name="pub_email" size="50" maxlength=64 class="n_textbox"></td>
</tr>

<tr>
	<td align="left" nowrap style="padding-top:6px; padding-left:8px;"><%=a_lang_114 %><%=s_lang_mh %></td>
	<td align="left" style="padding-top:10px;">
<textarea cols="70" rows="11" wrap="soft" name="message" class="n_textarea"></textarea>
	</td>
</tr>

<tr><td colspan=2 class="block_top_td" style="height:8px;"></td></tr>

<tr><td colspan="2" align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="javascript:back_onclick();"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:save_onclick();"><%=s_lang_ok %></a>
</td></tr>
</table>
</FORM>
</body>
</html>

<%
function geterror(ecode)
	if ecode = 1 then
		geterror = a_lang_115
	elseif ecode = 2 then
		geterror = a_lang_116
	elseif ecode = 3 then
		geterror = a_lang_117
	elseif ecode = 4 then
		geterror = a_lang_118
	elseif ecode = 5 then
		geterror = a_lang_119
	elseif ecode = 6 then
		geterror = a_lang_120
	elseif ecode = 7 then
		geterror = a_lang_121
	elseif ecode = 8 then
		geterror = a_lang_122
	elseif ecode = 9 then
		geterror = a_lang_123
	elseif ecode = 10 then
		geterror = a_lang_124
	end if
end function
%>
