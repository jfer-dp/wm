<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
dim is_domain_admin
is_domain_admin = false

if isadmin() = false then
	dim dm
	set dm = server.createobject("easymail.Domain")
	dm.Load

	allnum = dm.GetUserManagerDomainCount(Session("wem"))

	curdomain = Mid(Session("mail"), InStr(Session("mail"), "@") + 1)

	i = 0
	do while i < allnum
		domain = dm.GetUserManagerDomain(Session("wem"), i)

		if LCase(curdomain) = LCase(domain) then
			is_domain_admin = true
		end if

		domain = NULL

		i = i + 1
	loop

	set dm = nothing
else
	is_domain_admin = true
end if
%>

<%
dim id
id = trim(request("id"))

dim issave
issave = trim(request("issave"))
gourl = trim(Request("gourl"))

dim ads
set ads = server.createobject("easymail.DomainPubAddresses")
ads.Load Session("wem")

if issave = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if is_domain_admin = false then
		set ads = nothing
		response.redirect "noadmin.asp"
	end if

	ads.CreateNewEmail

	ads.nickname = trim(request("CName"))
	ads.email = trim(request("Email"))
	ads.first_name = trim(request("First_Name"))
	ads.last_name = trim(request("Last_Name"))
	ads.company = trim(request("Company"))

	ads.pm_home = trim(request("PM_Home"))
	ads.pm_work = trim(request("PM_Work"))
	ads.mobile = trim(request("PM_Mobile"))

if IsNumeric(id) = true then
	ads.ModifyEmail CInt(id)
	ads.Save
end if

	set ads = nothing

	if gourl = "" then
		Response.Redirect "ads_dm_pubbrow.asp?" & getGRSN()
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
//-->
</script>

<body>
<form name="f1" method=post action="ads_dm_pubedit.asp">
<input type="hidden" name="gourl" value="<%=gourl %>">
<input type="hidden" name="issave" value="1">
<input type="hidden" name="id" value="<%=trim(request("id")) %>">
<table width="90%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr><td align="left" height="28" style="padding-left:4px;">
	<a class='wwm_btnDownload btn_blue' href="javascript:history.back();"><< <%=s_lang_return %></a>
<% if is_domain_admin = true then %>
	&nbsp;<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
<% end if %>
	</td></tr>
</table>

<br>
<table width="90%" border="0" align="center" bgcolor="white" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr><td height="30">
		<table border=0 cellspacing="0" cellpadding=4 width="100%">
			<tr>
			<td nowrap width="6%"><%=s_lang_0360 %><%=s_lang_mh %></td>
			<td nowrap width="44%"><input type=text size="20" name="CName" value="<%=ads.nickname %>" class="n_textbox" maxlength="64"> <font color="#901111">*</font></td>
			<td nowrap width="10%"><%=s_lang_0377 %><%=s_lang_mh %></td>
			<td nowrap width="40%"><input type=text size="20" name="Email" value="<%=ads.email %>" class="n_textbox" maxlength="128"> <font color="#901111">*</font></td>
			</tr>

			<tr> 
			<td><%=s_lang_0378 %><%=s_lang_mh %></td>
			<td><input type=text size="20" name="First_Name" value="<%=ads.first_name %>" class="n_textbox" maxlength="32"></td>
			<td><%=s_lang_0379 %><%=s_lang_mh %></td>
			<td><input type=text size="20" name="Last_Name" value="<%=ads.last_name %>" class="n_textbox" maxlength="32"></td>
			</tr>

			<tr> 
			<td nowrap><%=s_lang_0381 %><%=s_lang_mh %></td>
			<td><input type=text size="20" name="Company" value="<%=ads.company %>" class="n_textbox" maxlength="64"></td>
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
		<td nowrap width="9%"><%=s_lang_0383 %><%=s_lang_mh %></td>
		<td width="41%"><input type=text size="20" name="PM_Home" value="<%=ads.pm_home %>" class="n_textbox" maxlength="32"></td>
		<td nowrap width="9%"><%=s_lang_0384 %><%=s_lang_mh %></td>
		<td width="41%"><input type=text size="20" name="PM_Work" value="<%=ads.pm_work %>" class="n_textbox" maxlength="32"></td>
		</tr>

		<tr> 
		<td nowrap><%=s_lang_0385 %><%=s_lang_mh %></td>
		<td><input type=text size="20" name="PM_Mobile" value="<%=ads.mobile %>" class="n_textbox" maxlength="16"></td>
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
