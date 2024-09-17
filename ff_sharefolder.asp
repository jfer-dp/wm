<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
foldername = trim(request("foldername"))

if trim(request("mode")) = "" then
	isatt = false
else
	isatt = true
end if

dim shareMode
dim showInList
dim perfolders

if isatt = false then
	set perfolders = server.createobject("easymail.PerFolders")
else
	set perfolders = server.createobject("easymail.PerAttFolders")
end if

perfolders.Load Session("wem")

if trim(request("save")) = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	pw1 = trim(request("pw1"))
	pw2 = trim(request("pw2"))
	fshare = trim(request("fshare"))

	if trim(request("showInList")) = "" then
		showInList = false
	else
		showInList = true
	end if

	if pw1 = pw2 and IsNumeric(fshare) = true then
		if pw1 <> "" then
			perfolders.SetPassword foldername, pw1
		end if

		perfolders.SetShareMode foldername, CInt(fshare)
		perfolders.SetShowInList foldername, showInList

		perfolders.Save


		sharetoname = trim(request("shareto"))
		if sharetoname <> "" and Session("wem") <> sharetoname and (pw1 <> "" or fshare = "2") then
			dim userweb
			set userweb = server.createobject("easymail.UserWeb")
			userweb.Load sharetoname

			if userweb.HaveThisFriendFolder(Session("wem"), foldername, isatt) = false then
				userweb.AddFriendFolder Session("wem"), foldername, pw1, isatt
			else
				userweb.SetPassword Session("wem"), foldername, pw1, isatt
			end if

			userweb.Save
			set userweb = nothing
		end if
	end if

	set perfolders = nothing

	if isatt = false and foldername <> "att" then
		Response.Redirect "viewmailbox.asp?" & getGRSN()
	else
		Response.Redirect "attfolders.asp?" & getGRSN()
	end if
end if

perfolders.GetInfoByName foldername, shareMode, showInList
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
.cont_td {height:24px; white-space:nowrap; border-bottom:1px solid #A5B6C8; padding-left:2px; padding-right:2px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function gosub()
{
	if (document.f1.shareto.value != "" && document.f1.fshare.value == 0)
	{
		alert("<%=a_lang_210 %>");
		return ;
	}

	if (document.f1.shareto.value != "" && document.f1.fshare.value == 1 && (document.f1.pw1.value == "" || document.f1.pw2.value == ""))
	{
		alert("<%=a_lang_211 %>");
		return ;
	}

	if (document.f1.pw1.value == document.f1.pw2.value)
		document.f1.submit();
	else
		alert("<%=a_lang_212 %>");
}
//-->
</script>

<body>
<form name="f1" method="post" action="ff_sharefolder.asp">
<input type="hidden" name="foldername" value="<%=foldername %>">
<input type="hidden" name="save" value="1">
<input type="hidden" name="mode" value="<%
if isatt = true then
	Response.Write "att"
end if
%>">
<%
if foldername = "att" then
	foldername = a_lang_202
end if
%>
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px; padding-right:6px; word-break:break-all; word-wrap:break-word;">
<%
if isatt = true then
	Response.Write a_lang_213
end if
%><%=a_lang_214 %>&nbsp;(<%=server.htmlencode(foldername) %>)
</td></tr>
<tr><td class="block_top_td" style="height:4px; _height:6px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr>
	<td width="20%" align="right" class="cont_td"><%=a_lang_215 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td">
<select name="fshare" class="drpdwn" size="1">
<%
i = 0

do while i < 3
	if shareMode <> i then
		Response.Write "<option value='" & i & "'>" & getShareStr(i) & "</option>"
	else
		Response.Write "<option value='" & i & "' selected>" & getShareStr(i) & "</option>"
	end if

	i = i + 1
loop
%>
</select>
	</td>
	</tr>

	<tr>
	<td align="right" class="cont_td"><%=a_lang_216 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td"><input type="checkbox" name="showInList" value="checkbox" <% if showInList = true then Response.Write "checked"%>> <%=a_lang_217 %>
	</td>
	</tr>

	<tr>
	<td align="right" class="cont_td"><%=a_lang_205 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td"><input type="password" name="pw1" maxlength="32" class="n_textbox"> <font color='#444444'>(<%=a_lang_218 %>)</font>
	</td>
	</tr>

	<tr>
	<td align="right" class="cont_td"><%=a_lang_023 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td"><input type="password" name="pw2" maxlength="32" class="n_textbox"> <font color='#444444'></font>
	</td>
	</tr>

	<tr>
	<td align="right" class="cont_td"><%=a_lang_219 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td"><input type="text" name="shareto" maxlength="64" size="40" class="n_textbox">
	</td>
	</tr>

	<tr><td align="left" colspan="2" style="background-color:white; padding-right:16px; padding-top:18px; padding-bottom:10px;">
<a class='wwm_btnDownload btn_blue' href="javascript:history.back();"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
	</td></tr>
</table>
</td></tr>
</table>

<table width="89%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px #A5B6C8 solid; margin-top:80px; padding-bottom:10px;'>
	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:6px; color:#444444;">
	<%=a_lang_220 %>
	</td>
	</tr>

	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:6px; color:#444444;">
	<%=a_lang_221 %>
	</td>
	</tr>

	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:6px; color:#444444;">
	<%=a_lang_222 %>
	</td>
	</tr>
</table>
</form>
<div style="position:absolute; left:12px; top:10px;">
<a href="help.asp#sharefolder" target="_blank"><img src="images/help.gif" border="0" title="<%=s_lang_help %>"></a></div>
</BODY>
</HTML>

<%
shareMode = NULL
showInList = NULL

set perfolders = nothing


function getShareStr(md)
	if md = 0 then
		getShareStr = a_lang_223
	elseif md = 1 then
		getShareStr = a_lang_224
	elseif md = 2 then
		getShareStr = a_lang_225
	end if
end function
%>
