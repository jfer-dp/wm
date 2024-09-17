<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
issave = trim(request("issave"))
mode = trim(Request("mode"))
gourl = trim(Request("gourl"))

dim ads
set ads = server.createobject("easymail.Addresses")
ads.Load Session("wem")

if issave = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	isok = ads.AddGroupInfo(trim(request("CName")), trim(request("Email")))

	ads.Save

	set ads = nothing

	if isok = true then
		if mode = "saveadd" then
			response.redirect "adg_add.asp?gourl=" & Server.URLEncode(gourl) & "&" & getGRSN()
		else
			response.redirect gourl & "&" & getGRSN()
		end if
	else
		response.redirect "err.asp?" & getGRSN()
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
body {padding-left:8px; padding-right:8px;}
.title_td {padding:3px 2px 2px 6px; _padding-top:4px;}
-->
</STYLE>
</head>

<script type="text/javascript">
<!-- 
function back_onclick() {
	location.href = "<%=gourl & "&" & getGRSN() %>";
}

function gosub() {
	if (document.f1.CName.value == "")
	{
		alert("<%=s_lang_0351 %>");
		document.f1.CName.focus();
	}
	else if (document.f1.Email.value == "")
	{
		alert("<%=s_lang_0352 %>");
		document.f1.Email.focus();
	}
	else
	{
		document.f1.mode.value = "save"

		if (document.f1.Email.value.length > 10000)
			document.f1.Email.value = document.f1.Email.value.substring(0, 10000);

		document.f1.submit();
	}
}

function saveadd_onclick() {
	if (document.f1.CName.value == "")
	{
		alert("<%=s_lang_0351 %>");
		document.f1.CName.focus();
	}
	else if (document.f1.Email.value == "")
	{
		alert("<%=s_lang_0352 %>");
		document.f1.Email.focus();
	}
	else
	{
		document.f1.mode.value = "saveadd"

		if (document.f1.Email.value.length > 10000)
			document.f1.Email.value = document.f1.Email.value.substring(0, 10000);

		document.f1.submit();
	}
}

function select_sel_ads_add() {
	var emails_str = "," + document.f1.Email.value.toLowerCase() + ",";
	var search_str = "," + document.f1.sel_ads_add.value.toLowerCase() + ",";

	if (emails_str.indexOf(search_str) == -1)
	{
		if (document.f1.Email.value.length == 0)
			document.f1.Email.value = document.f1.sel_ads_add.value;
		else if (document.f1.Email.value.length > 0)
			document.f1.Email.value = document.f1.Email.value + "," + document.f1.sel_ads_add.value;
	}

	document.f1.sel_ads_add.selectedIndex = 0;
}
// -->
</script>

<body>
<form name="f1" method=post action="adg_add.asp">
<input type="hidden" name="issave" value="1">
<input type="hidden" name="mode">
<input type="hidden" name="gourl" value="<%=gourl %>">
<input type="hidden" name="id" value="<%=id %>">
<table width="90%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr><td align="left" height="28" style="padding-left:4px;">
	<a class='wwm_btnDownload btn_blue' href="javascript:back_onclick();"><< <%=s_lang_return %></a>&nbsp;
	<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>&nbsp;
	<a class='wwm_btnDownload btn_blue' href="javascript:saveadd_onclick();"><%=s_lang_0375 %></a>
	</td></tr>
</table>
<br>

<table width="90%" border="0" align="center" bgcolor="white" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr><td height="30">
		<table border=0 cellspacing="0" cellpadding=4 width="100%">
			<tr style="padding-bottom:3px;">
			<td align="left" nowrap width="10%" style="padding-top:8px;"><%=s_lang_0319 %><%=s_lang_mh %></td>
			<td align="left"><input type=text size="20" name="CName" class="n_textbox" maxlength="64"></td>
			</tr>

			<tr>
			<td align="left" nowrap valign="top" style="padding-top:10px;"><%=s_lang_0353 %><%=s_lang_mh %></td>
			<td align="left">
			<textarea rows="7" cols="60" wrap="virtual" name="Email" class="n_textarea"></textarea>
			<br><font color="#444444"><%=s_lang_0354 %></font>
			</td>
			</tr>

			<tr> 
			<td>&nbsp;</td>
			<td align="left">
<select name="sel_ads_add" class="drpdwn" size="1" LANGUAGE=javascript onchange="select_sel_ads_add()">
<option value='' selected>---<%=s_lang_0355 %>---</option>
<%
i = 0
all_ads_num = ads.EmailCount

do while i < all_ads_num
	ads.MoveTo i

	Response.Write "<option value='" & ads.email & "'>" & ads.nickname & " | " & ads.email & "</option>"

    i = i + 1
loop
%>
</select>
			</td>
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
