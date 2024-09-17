<%
Response.CacheControl = "no-cache" 
%>

<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
dim wemcert
set wemcert = server.createobject("easymail.WebEasyMailCert")

if wemcert.LightHasSecCert(Session("wem")) = true then
	set wemcert = nothing
	Response.Redirect "err.asp?" & getGRSN() & "&errstr=" & Server.URLEncode(a_lang_028) & "&gourl=cert_index.asp"
end if

set wemcert = nothing

if trim(Session("cert_ca")) = "1" then
	Response.Redirect "err.asp?" & getGRSN() & "&errstr=" & Server.URLEncode(a_lang_029) & "&gourl=cert_index.asp"
end if

pw = trim(request("pw"))

if pw <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	t_year = trim(request("t_year"))
	if t_year = "" then
		t_year = "0000"
	end if

	t_month = trim(request("t_month"))
	if t_month = "" then
		t_month = "00"
	end if

	t_day = trim(request("t_day"))
	if t_day = "" then
		t_day = "00"
	end if


	dim em
	set em = server.createobject("easymail.emmail")

	if em.CreateCert(getusername(Session("wem")), Session("mail"), pw, CInt(trim(request("c_type"))), CInt(trim(request("c_keysize"))), CInt(trim(request("c_subkeysize"))), t_year & t_month & t_day) = false then
		response.write a_lang_030
	else
		Session("cert_ca") = "1"
	end if

	set em = nothing

	Response.End
end if
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
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
function checkpw(){
	if (document.fc.pw.value.length < 8)
	{
		alert("<%=a_lang_018 %>");
		document.fc.pw.focus();
		return ;
	}

	if (document.fc.pw.value != document.fc.pw1.value)
		alert("<%=a_lang_019 %>");
	else
		document.fc.submit();
}

function date1_onclick() {
	document.fc.t_year.value = "0000";
	document.fc.t_month.value = "00";
	document.fc.t_day.value = "00";
	document.fc.t_year.disabled = true;
	document.fc.t_month.disabled = true;
	document.fc.t_day.disabled = true;
}

function date2_onclick() {
	document.fc.t_year.disabled = false;
	document.fc.t_month.disabled = false;
	document.fc.t_day.disabled = false;
}

function type_onchange() {
	if (document.fc.typeselect.value == "1")
	{
		document.fc.c_type.value = "0";
		document.fc.c_keysize.value = "1024";
		document.fc.c_subkeysize.value = "1024";
	}
	else if (document.fc.typeselect.value == "2")
	{
		document.fc.c_type.value = "0";
		document.fc.c_keysize.value = "1024";
		document.fc.c_subkeysize.value = "2048";
	}
	else if (document.fc.typeselect.value == "3")
	{
		document.fc.c_type.value = "0";
		document.fc.c_keysize.value = "1024";
		document.fc.c_subkeysize.value = "4096";
	}
	else if (document.fc.typeselect.value == "4")
	{
		document.fc.c_type.value = "1";
		document.fc.c_keysize.value = "1024";
		document.fc.c_subkeysize.value = "1024";
	}
	else if (document.fc.typeselect.value == "5")
	{
		document.fc.c_type.value = "1";
		document.fc.c_keysize.value = "2048";
		document.fc.c_subkeysize.value = "2048";
	}
}
//-->
</script>

<body>
<form name="fc" METHOD="POST" action="cert_createca.asp">
<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white" style="margin-top:6px;">
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_031 %>
</td></tr>
<tr><td style="padding-left:12px;">
<table width="100%" border="0" align="left" cellspacing="0" bgcolor="white" height="210">
	<tr><td colspan=2 class="block_top_td" style="height:1px; _height:8px;"></td></tr>

	<tr><td align="left" nowrap height="20" width="8%"><%=a_lang_032 %><%=s_lang_mh %></td>
	<td align="left">
	<input type="text" name="username" value="<%=getusername(Session("wem")) & " <" & Session("mail") & ">" %>" class="n_textbox" size="26" readonly>
	</td></tr>

	<tr><td align="left" height="20" nowrap style="padding-top:4px;"><%=a_lang_033 %><%=s_lang_mh %></td>
	<td align="left">
		<select name="typeselect" class="drpdwn" size="1" LANGUAGE=javascript onchange="return type_onchange()">
		<option value="1">1024 Bits DH/DSS</option>
		<option value="2" selected>2048 Bits DH/DSS</option>
		<option value="3">4096 Bits DH/DSS</option>
		<option value="4">1024 Bits RSA</option>
		<option value="5">2048 Bits RSA</option>
		</select>
	</td></tr>

	<tr><td align="left" height="20" nowrap style="padding-top:4px;"><%=a_lang_034 %><%=s_lang_mh %></td>
	<td align="left" style="color:#444444;">
	<input type="password" name="pw" maxlength="64" class="n_textbox">&nbsp;&nbsp;[<%=a_lang_035 %>]
	</td></tr>

	<tr><td align="left" height="20" nowrap style="padding-top:4px;"><%=a_lang_036 %><%=s_lang_mh %></td>
	<td align="left">
	<input type="password" name="pw1" maxlength="64" class="n_textbox">
	</td></tr>

	<tr><td align="left" height="20" colspan="2" nowrap>
		<input type=radio checked value="0" LANGUAGE=javascript onclick="return date1_onclick()" name="crmode"> <%=a_lang_037 %>&nbsp;
		<input type=radio value="1" LANGUAGE=javascript onclick="return date2_onclick()" name="crmode"> <%=a_lang_038 %><br>
	</td></tr>

	<tr><td align="left" height="20" colspan="2" nowrap style="padding-bottom:8px;">
<select name="t_year" class="drpdwn" disabled>
<option value="0000">------</option>
<%
	now_temp = Year(Now())
	end_year = CInt(now_temp) + 11

	do while now_temp < end_year
		response.write "<option value='" & now_temp & "'>" & now_temp & a_lang_039 & "</option>" & Chr(13)

		now_temp = now_temp + 1
	loop
%>
</select>&nbsp;
<select name="t_month" class="drpdwn" disabled>
<option value="00">----</option>
<%
	i = 1
	do while i < 13
		if i < 10 then
			response.write "<option value='0" & i & "'>" & i & a_lang_040 & "</option>"
		else
			response.write "<option value='" & i & "'>" & i & a_lang_040 & "</option>"
		end if

		i = i + 1
	loop
%>
</select>&nbsp;
<select name="t_day" class="drpdwn" disabled>
<option value="00">----</option>
<%
	i = 1
	do while i < 32
		if i < 10 then
			response.write "<option value='0" & i & "'>" & i & a_lang_041 & "</option>"
		else
			response.write "<option value='" & i & "'>" & i & a_lang_041 & "</option>"
		end if

		i = i + 1
	loop
%>
</select>
	</td></tr>
</table>
</td></tr>

<tr>
<td nowrap style="border-top:1px #a7c5e2 solid; color:#093665; padding-left:6px; padding-top:12px;">
	<a class='wwm_btnDownload btn_blue' href="cert_index.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
	<a class='wwm_btnDownload btn_blue' href="javascript:checkpw();"><%=a_lang_042 %></a>
</td></tr>
</table>

<input type="hidden" name="c_type" value="0">
<input type="hidden" name="c_keysize" value="1024">
<input type="hidden" name="c_subkeysize" value="2048">
</form>
</body>
</html>

<%
function getusername(temp_domain_account)
	se = InStr(1, temp_domain_account, "@")

	if se <> 0 then
		getusername = Mid(temp_domain_account, 1, se - 1)
	else
		getusername = temp_domain_account
	end if
end function
%>
