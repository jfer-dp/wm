<%
Response.CacheControl = "no-cache" 
%>

<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
pw = trim(request("sc_pw"))
pub_email = trim(request("pub_email"))

ispub = false
if trim(request("im")) = "pub" then
	ispub = true
end if

if Request.ServerVariables("REQUEST_METHOD") = "GET" then
	if trim(request("im")) = "sec" then
		dim wemcert
		set wemcert = server.createobject("easymail.WebEasyMailCert")

		if wemcert.LightHasSecCert(Session("wem")) = true then
			set wemcert = nothing
			Response.Redirect "err.asp?" & getGRSN() & "&errstr=" & Server.URLEncode(a_lang_043) & "&gourl=cert_index.asp"
		end if

		set wemcert = nothing

		Session("cert_imp_type") = "sec"
	elseif trim(request("im")) = "pub" then
		ispub = true

		Session("cert_imp_type") = "pub"
	end if
end if

if (pw <> "" or pub_email <> "") and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if pw <> "" then
		Session("cert_imp_pw") = pw
		Session("cert_imp_save_day") = trim(request("save_day"))
	elseif pub_email <> "" then
		ispub = true

		Session("cert_imp_pw") = pub_email
	end if
end if

if trim(Session("cert_imp_pw")) = "" then
	imode = 1
else
	imode = 2
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
<!--
function checkpw(){
<%
if imode = 1 then
	if trim(request("im")) = "sec" then
%>
	if (document.fsa.sc_pw.value == "")
	{
		alert("<%=a_lang_044 %>");
		document.fsa.sc_pw.focus();
	}
	else
		document.fsa.submit();
<%
	else
%>
	if (document.fsa.pub_email.value == "")
	{
		alert("<%=a_lang_045 %>");
		document.fsa.pub_email.focus();
	}
	else
		document.fsa.submit();
<%
	end if
else
%>
	if (document.fsa.upfile.value == "")
	{
		alert("<%=a_lang_046 %>");
		document.fsa.upfile.focus();
	}
	else
		document.fsa.submit();
<%
end if
%>
}
//-->
</script>

<body>
<%
if imode = 2 then
%>
<FORM ENCTYPE="multipart/form-data" ACTION="cert_imp_end.asp" METHOD=POST NAME="fsa">
<%
else
%>
<form name="fsa" METHOD="POST" action="cert_imp.asp">
<%
end if
%>
<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white" style="margin-top:6px;">
<tr><td colspan="2" style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_047 %><%=imode %><%=a_lang_048 %>
</td></tr>

<%
if imode = 2 then
%>
	<tr>
	<td align="right" nowrap width="50%" height="50" style="padding-left:12px; padding-top:10px; padding-bottom:10px; border-bottom:1px #a7c5e2 solid;">
	<input name="upfile" type="file" class='up_textbox' size="30">
	</td><td align="left" nowrap style="padding-left:12px; padding-top:10px; padding-bottom:10px; border-bottom:1px #a7c5e2 solid;">
	<a class='wwm_btnDownload btn_gray' href="javascript:checkpw();"><%=a_lang_049 %></a>
	<input type="hidden" name="impmode">
	</td>
	</tr>
<%
else
	if trim(request("im")) = "sec" then
%>
	<tr>
	<td align="left" height="34" style="padding-left:12px; padding-top:16px;"><%=a_lang_050 %><%=s_lang_mh %>&nbsp;<input type="password" name="sc_pw" class='n_textbox'></td>
	</tr>
	<tr><td align="left" height="34" style="padding-left:12px; padding-bottom:16px; border-bottom:1px #a7c5e2 solid;"><%=a_lang_051 %><%=s_lang_mh %>
<select name="save_day" class="drpdwn">
<option value="-1" selected><%=a_lang_025 %></option>
<option value="0"><%=a_lang_026 %></option>
<%
	now_temp = 10

	do while now_temp < 999
		Response.Write "<option value='" & now_temp & "'>" & now_temp & a_lang_027 & "</option>" & Chr(13)

		now_temp = now_temp + 10
	loop
%>
</select>
	</td></tr>
<%
	elseif trim(request("im")) = "pub" then
%>
	<tr>
	<td align="left" nowrap colspan="2" height="50" style="padding-left:12px; padding-top:10px; padding-bottom:10px; border-bottom:1px #a7c5e2 solid;">
<%=a_lang_052 %><%=s_lang_mh %>&nbsp;<input type="text" name="pub_email" class='n_textbox' size="35">
	</td>
	</tr>
<%
	end if
end if
%>
	<tr> 
	<td colspan="2" align="right" style="background-color:white; padding-right:16px;"><br>
<%
if imode <> 2 then
%>
	<a class='wwm_btnDownload btn_blue' href="javascript:checkpw();"><%=s_lang_ok %></a>&nbsp;<%
end if

if ispub = true then
%><a class='wwm_btnDownload btn_blue' href="cert_myothpub.asp?<%=getGRSN() %>"><%=s_lang_cancel %></a><%
else
%><a class='wwm_btnDownload btn_blue' href="cert_index.asp?<%=getGRSN() %>"><%=s_lang_cancel %></a>
<%
end if
%>
	</td></tr>
</table>
</FORM>
</body>
</html>
