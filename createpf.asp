<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
if trim(request("PFPermission")) <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim pf
	set pf = server.createobject("easymail.PubFolderManager")

	isok = pf.CreatePubFolder(trim(request("Admin")), CInt(trim(request("PFPermission"))), trim(request("NewPFName")))
	set pf = nothing

	if isok = true then
		Response.Redirect "ok.asp?gourl=showallpf.asp&" & getGRSN()
	else
		Response.Redirect "err.asp?gourl=showallpf.asp&" & getGRSN()
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
.st_l, .st_r {height:28px; white-space:nowrap;}
.st_l {text-align:right;}
.st_r {text-align:left;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function gosub()
{
	if (document.f1.NewPFName.value != "" && document.f1.Admin.value != "")
		document.f1.submit();
}
//-->
</script>

<body>
<FORM ACTION="createpf.asp" METHOD="POST" NAME="f1">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_148 %>
</td></tr>
<tr><td class="block_top_td" style="height:12px; _height:14px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr>
	<td width="10%" nowrap class="st_l"><%=a_lang_149 %><%=s_lang_mh %></td>
	<td class="st_r"><input type="text" name="NewPFName" size="40" class='n_textbox'></td>
	</tr>

	<tr>
	<td nowrap class="st_l"><%=a_lang_150 %><%=s_lang_mh %></td>
	<td class="st_r">
<select name="PFPermission" class="drpdwn" size="1">
<%
i = 0
ichoose = 1

do while i < 5
	if ichoose <> i then
		response.write "<option value='" & i & "'>" & getPermissionStr(i) & "</option>"
	else
		response.write "<option value='" & i & "' selected>" & getPermissionStr(i) & "</option>"
	end if

	i = i + 1
loop
%>
</select>
	</td>
	</tr>

	<tr>
	<td class="st_l"><%=a_lang_151 %><%=s_lang_mh %></td>
	<td class="st_r">
<%
dim eu
set eu = Application("em")
%>
<select name="Admin" class="drpdwn" size="1">
<%
i = 0
allnum = eu.GetUsersCount

do while i < allnum
	eu.GetUserByIndex i, name, domain, comment

	response.write "<option value='" & name & "'>" & name & "</option>"

	name = NULL
	domain = NULL
	comment = NULL

	i = i + 1
loop

set eu = nothing
%>
</select>
	</td>
	</tr>
</table>
</td></tr>
<tr><td class="block_top_td" style="height:8px;"></td></tr>

<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="javascript:history.back();"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=a_lang_152 %></a>
</td></tr>
</table>
</Form>
</BODY>
</HTML>

<%
function getPermissionStr(pm)
	if pm = 0 then
		getPermissionStr = a_lang_153
	elseif pm = 1 then
		getPermissionStr = a_lang_154
	elseif pm = 2 then
		getPermissionStr = a_lang_155
	elseif pm = 3 then
		getPermissionStr = a_lang_156
	elseif pm = 4 then
		getPermissionStr = a_lang_157
	end if
end function
%>
