<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
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
sortmode = request("sortmode")
issort = false

if sortstr <> "" then
	if sortmode = "1" then
		addsortstr = "&sortstr=" & sortstr & "&sortmode=1"
		sortmode = true

		nb.SetSort sortstr, sortmode
		issort = true
	elseif sortmode = "0" then
		addsortstr = "&sortstr=" & sortstr & "&sortmode=0"
		sortmode = false

		nb.SetSort sortstr, sortmode
		issort = true
	end if
end if

nb.Load Session("wem")

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
function back_onclick() {
	history.back();
}
//-->
</script>

<body>
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%= server.htmlencode(title) %>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="left" style="padding-left:6px;">
<%
t = server.htmlencode(text)
t = replace(t, Chr(10), "<br>")
t = replace(t, Chr(32), "&nbsp;")

Response.Write t
%>&nbsp;</td>
</tr>

<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="javascript:back_onclick();"><< <%=s_lang_return %></a>
</td></tr>
</table>
</body>
</html>

<%
nb_date = NULL
title = NULL
text = NULL

set nb = nothing
%>
