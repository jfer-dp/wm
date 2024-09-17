<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
issave = trim(Request("issave"))
id = trim(Request("id"))

if id = "" then
	response.redirect "nb_brow.asp?" & getGRSN()
end if

if IsNumeric(id) = false then
	response.redirect "nb_brow.asp?" & getGRSN()
end if

id = CInt(id)

dim nb
set nb = server.createobject("easymail.NoteBooksManager")

sortstr = request("sortstr")
o_sortmode = request("sortmode")
addsortstr = trim(Request("addsortstr"))
issort = false

if sortstr <> "" then
	if o_sortmode = "1" then
		sortmode = true

		nb.SetSort sortstr, sortmode
		issort = true
	elseif o_sortmode = "0" then
		sortmode = false

		nb.SetSort sortstr, sortmode
		issort = true
	end if
end if

nb.Load Session("wem")

if issave = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if trim(Request("name")) <> "" then
		nb.Modify id, trim(Request("name")), Request("message")
		nb.Save
	end if

	set nb = nothing
	Response.Redirect "nb_brow.asp?page=" & trim(request("page")) & "&" & getGRSN() & addsortstr
end if

nb.Get id, nb_date, title, text
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
function save_onclick() {
	if (document.getElementById("f1").name.value=="")
	{
		alert("<%=a_lang_296 %>");
		document.getElementById("f1").name.focus();
	}
	else
	{
		if (document.getElementById("f1").message.value.length > 10000)
			document.getElementById("f1").message.value = document.getElementById("f1").message.value.substring(0, 10000);

		document.getElementById("f1").submit();
	}
}

function back_onclick() {
	history.back();
}
//-->
</SCRIPT>

<body>
<form method="post" action="nb_edit.asp" id="f1" name="f1">
<input type="hidden" name="issave" value="1">
<input type="hidden" name="addsortstr" value="<%=addsortstr %>">
<input type="hidden" name="sortstr" value="<%=sortstr %>">
<input type="hidden" name="sortmode" value="<%=o_sortmode %>">
<input type="hidden" name="id" value="<%=id %>">
<input type="hidden" name="page" value="<%=trim(request("page")) %>">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=s_lang_303 %>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="left">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr>
	<td width="8%" nowrap align="right" style="height:24px;"><%=a_lang_298 %><%=s_lang_mh %></td>
	<td align="left"><input name="name" size="70" maxlength=124 class="n_textbox" value="<%=title %>"></td>
	</tr>

	<tr>
	<td align="right" nowrap style="height:24px;"><%=a_lang_299 %><%=s_lang_mh %></td>
	<td>
	<textarea cols="75" rows="11" wrap="soft" name="message" class="n_textarea"><%=server.htmlencode(text) %></textarea>
	</td>
	</tr>
</table>
</td></tr>

<tr><td class="block_top_td" style="height:8px;"></td></tr>

<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="javascript:back_onclick();"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:save_onclick();"><%=s_lang_save %></a>
</td></tr>
</table>
</FORM>
</body>
</html>

<%
nb_date = NULL
title = NULL
text = NULL

set nb = nothing


function mReplace(tempstr)
	dim cstr
	cstr = replace(tempstr , """", "'")
	mReplace = replace(cstr , "'", "''")
end function

function mmReplace(tempstr)
	mmReplace = replace(tempstr , "'","''")
end function
%>
