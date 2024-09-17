<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
issave = trim(Request("issave"))
mode = trim(Request("mode"))
addsortstr = trim(Request("addsortstr"))

if issave = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim nb
	set nb = server.createobject("easymail.NoteBooksManager")
	nb.Load Session("wem")

	if trim(Request("name")) <> "" then
		nb.Add trim(Request("name")), Request("message")
		nb.Save
	end if

	set nb = nothing

	if mode = "saveadd" then
		response.redirect "nb_add.asp?" & getGRSN() & "&mode=saveadd" & trim(Request("appsortstr"))
	else
		response.redirect "nb_brow.asp?" & getGRSN() & addsortstr
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
</HEAD>

<script type="text/javascript">
<!--
function save_onclick() {
	if (document.getElementById("f1").name.value == "")
	{
		alert("<%=a_lang_296 %>");
		document.getElementById("f1").name.focus();
	}
	else
	{
		document.getElementById("f1").mode.value = "save"

		if (document.getElementById("f1").message.value.length > 10000)
			document.getElementById("f1").message.value = document.getElementById("f1").message.value.substring(0, 10000);

		document.getElementById("f1").submit();
	}
}

function saveadd_onclick() {
	if (document.getElementById("f1").name.value == "")
	{
		alert("<%=a_lang_296 %>");
		document.getElementById("f1").name.focus();
	}
	else
	{
		document.getElementById("f1").mode.value = "saveadd"

		if (document.getElementById("f1").message.value.length > 10000)
			document.getElementById("f1").message.value = document.getElementById("f1").message.value.substring(0, 10000);

		document.getElementById("f1").submit();
	}
}

function back_onclick() {
<%
if mode = "saveadd" then
%>
	location.href = "nb_brow.asp?<%=getGRSN() & addsortstr %>";
<%
else
%>
	history.back();
<%
end if
%>
}
//-->
</SCRIPT>

<body>
<form method="post" action="nb_add.asp" id="f1" name="f1">
<input type="hidden" name="addsortstr" value="<%=addsortstr %>">
<input type="hidden" name="appsortstr" value="&addsortstr=<%=Server.URLEncode(addsortstr) %>">
<input type="hidden" name="mode">
<input type="hidden" name="issave" value="1">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_297 %>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="left">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr>
	<td width="8%" nowrap align="right" style="height:24px; padding-top:4px;"><%=a_lang_298 %><%=s_lang_mh %></td>
	<td align="left"><input name="name" size="70" maxlength=124 class="n_textbox"></td>
	</tr>

	<tr>
	<td align="right" nowrap style="height:24px;"><%=a_lang_299 %><%=s_lang_mh %></td>
	<td>
	<textarea cols="75" rows="11" wrap="soft" name="message" class="n_textarea"></textarea>
	</td>
	</tr>
</table>
</td></tr>

<tr><td class="block_top_td" style="height:8px;"></td></tr>

<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="javascript:back_onclick();"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:save_onclick();"><%=s_lang_save %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:saveadd_onclick();"><%=s_lang_300 %></a>
</td></tr>
</table>
</FORM>
</body>
</html>

<%
function mReplace(tempstr)
	dim cstr
	cstr = replace(tempstr , """", "'")
	mReplace = replace(cstr , "'", "''")
end function

function mmReplace(tempstr)
	mmReplace = replace(tempstr , "'","''")
end function
%>
