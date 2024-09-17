<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
dim id
id = trim(request("id"))

dim issave
issave = trim(request("issave"))
gourl = trim(Request("gourl"))

dim ads
set ads = server.createobject("easymail.Addresses")
ads.Load Session("wem")

if issave = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ads.CreateNewEmail

	ads.nickname = trim(request("CName"))
	ads.email = trim(request("Email"))
	ads.first_name = trim(request("First_Name"))
	ads.last_name = trim(request("Last_Name"))
	ads.company = trim(request("Company"))
	ads.other_email = trim(request("Other_Email"))

	ads.pm_home = trim(request("PM_Home"))
	ads.pm_work = trim(request("PM_Work"))
	ads.mobile = trim(request("PM_Mobile"))

	ads.wi_zip = trim(request("WI_ZIP"))
	ads.wi_country = trim(request("WI_Country"))
	ads.wi_state = trim(request("WI_State"))
	ads.wi_city = trim(request("WI_City"))
	ads.wi_address = trim(request("WI_Address"))

	ads.birthday = trim(request("Birthday"))
	ads.homepage = trim(request("HomePage"))

if IsNumeric(id) = true then
	ads.ModifyEmail CInt(id)
	ads.Save
end if

	set ads = nothing

	if gourl = "" then
		Response.Redirect "ads_brow.asp?" & getGRSN()
	else
		Response.Redirect gourl & "&" & getGRSN()
	end if
end if

if IsNumeric(id) = true then
	ads.MoveTo CInt(id)
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
body {padding-left:8px; padding-right:8px;}
.title_td {padding:3px 2px 2px 6px; _padding-top:4px;}
-->
</STYLE>
</head>

<script type="text/javascript">
<!-- 
function gosub() {
	if (document.f1.CName.value == "")
	{
		alert("<%=s_lang_0376 %>");
		document.f1.CName.focus();
	}
	else if (document.f1.Email.value == "")
	{
		alert("<%=s_lang_0352 %>");
		document.f1.Email.focus();
	}
	else
		document.f1.submit();
}
// -->
</script>

<body>
<form name="f1" method=post action="ads_edit.asp">
<input type="hidden" name="gourl" value="<%=gourl %>">
<input type="hidden" name="issave" value="1">
<input type="hidden" name="id" value="<%=id %>">
<table width="90%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr><td align="left" height="28" style="padding-left:4px;">
	<a class='wwm_btnDownload btn_blue' href="javascript:history.back();"><< <%=s_lang_return %></a>&nbsp;
	<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
	</td></tr>
</table>
<br>
<table width="90%" border="0" align="center" bgcolor="white" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr><td height="30">
		<table border=0 cellspacing="0" cellpadding=4 width="100%">
			<tr>
			<td nowrap width="6%"><%=s_lang_0360 %><%=s_lang_mh %></td>
			<td nowrap width="44%"><input type=text size="20" name="CName" value="<%=ads.nickname %>" class="n_textbox" maxlength="64"> <font color="#901111">*</font></td>
			<td nowrap width="13%"><%=s_lang_0377 %><%=s_lang_mh %></td>
			<td nowrap width="37%"><input type=text size="20" name="Email" value="<%=ads.email %>" class="n_textbox" maxlength="128"> <font color="#901111">*</font></td>
			</tr>

			<tr> 
			<td nowrap><%=s_lang_0378 %><%=s_lang_mh %></td>
			<td><input type=text size="20" name="First_Name" value="<%=ads.first_name %>" class="n_textbox" maxlength="32"></td>
			<td nowrap><%=s_lang_0379 %><%=s_lang_mh %></td>
			<td><input type=text size="20" name="Last_Name" value="<%=ads.last_name %>" class="n_textbox" maxlength="32"></td>
			</tr>

			<tr> 
			<td nowrap><%=s_lang_0381 %><%=s_lang_mh %></td>
			<td><input type=text size="20" name="Company" value="<%=ads.company %>" class="n_textbox" maxlength="64"></td>
			<td nowrap><%=s_lang_0380 %><%=s_lang_mh %></td>
			<td><input type=text size="20" name="Other_Email" value="<%=ads.other_email %>" class="n_textbox" maxlength="128"></td>
			</tr>
		</table>
	</td></tr>
</table>

<br>
<table width="90%" border="0" align="center" bgcolor="white" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr bgcolor="#104A7B">
	<td align=left class="title_td"><font color="white"><%=s_lang_0382 %></font></td>
	</tr>

	<tr><td>
	<table border=0 cellspacing=0 cellpadding=4 width="100%">
		<tr>
		<td nowrap width="10%"><%=s_lang_0383 %><%=s_lang_mh %></td>
		<td width="40%"><input type=text size="20" name="PM_Home" value="<%=ads.pm_home %>" class="n_textbox" maxlength="32"></td>
		<td nowrap width="10%"><%=s_lang_0384 %><%=s_lang_mh %></td>
		<td width="40%"><input type=text size="20" name="PM_Work" value="<%=ads.pm_work %>" class="n_textbox" maxlength="32"></td>
		</tr>

		<tr> 
		<td nowrap><%=s_lang_0385 %><%=s_lang_mh %></td>
		<td><input type=text size="20" name="PM_Mobile" value="<%=ads.mobile %>" class="n_textbox" maxlength="16"></td>
		</tr>
	</table>
	</td></tr>
</table>

<br>
<table width="90%" border="0" align="center" bgcolor="white" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr bgcolor="#104A7B">
	<td align=left class="title_td"><font color="white"><%=s_lang_0386 %></font></td>
	</tr>
	<tr><td>
	<table border=0 cellspacing=0 cellpadding=4 width="100%">
		<tr>
		<td nowrap width="10%"><%=s_lang_0387 %><%=s_lang_mh %></td>
		<td width="40%"><input type=text size="20" name="WI_ZIP" value="<%=ads.wi_zip %>" class="n_textbox" maxlength="16"></td>
		<td nowrap width="6%"><%=s_lang_0388 %><%=s_lang_mh %></td>
		<td width="44%"><input type=text size="20" name="WI_Address" class="n_textbox" value="<%=ads.wi_address %>" maxlength="64"></td>
		</tr>

		<tr> 
		<td nowrap><%=s_lang_0389 %><%=s_lang_mh %></td>
		<td><input type=text size="20" name="WI_City" value="<%=ads.wi_city %>" class="n_textbox" maxlength="64"></td>
		<td><%=s_lang_0390 %><%=s_lang_mh %></td>
		<td><input type=text size="20" name="WI_State" value="<%=ads.wi_state %>" class="n_textbox" maxlength="64"></td>
		</tr>

		<tr> 
		<td nowrap><%=s_lang_0391 %><%=s_lang_mh %></td>
		<td><input type=text size="20" name="WI_Country" value="<%=ads.wi_country %>" class="n_textbox" maxlength="64"></td>
		</tr>
	</table>
	</td></tr>
</table>

<br>
<table width="90%" border="0" align="center" bgcolor="white" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr bgcolor="#104A7B">
	<td align=left class="title_td"><font color="white"><%=s_lang_0392 %></font></td>
	</tr>
	<tr><td>
	<table border=0 cellspacing=0 cellpadding=4 width="100%">
		<tr>
		<td nowrap width="6%"><%=s_lang_0393 %><%=s_lang_mh %></td>
		<td width="44%"><input type=text maxlength="16" name="Birthday" value="<%=ads.birthday %>" class="n_textbox"></td>
		<td nowrap width="6%"><%=s_lang_0394 %><%=s_lang_mh %></td>
		<td width="44%"><input type=text size="20" name="HomePage" value="<%=ads.homepage %>" class="n_textbox" maxlength="128"></td>
		</tr>
	</table>
	</td></tr>
</table>
</form>
</body>
</html>

<%
set ads = nothing

function mReplace(tempstr)
	dim cstr
	cstr = replace(tempstr , """", "'")
	mReplace = replace(cstr , "'", "''")
end function

function mmReplace(tempstr)
	mmReplace = replace(tempstr , "'","''")
end function
%>
