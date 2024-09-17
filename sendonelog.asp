<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
user = ""
if isadmin() = true then
	user = trim(request("user"))
end if

filename = trim(request("filename"))
gourl = trim(request("gourl"))
returl = trim(request("returl"))

dim slm
set slm = server.createobject("easymail.SendLogManager")

if Len(user) < 1 then
	slm.LoadOne Session("wem"), filename
else
	slm.LoadOne user, filename
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
td {padding-left:3px; padding-right:3px;}
-->
</STYLE>
</HEAD>

<script language="JavaScript">
<!-- 
function back() {
<% if gourl = "" then %>
	history.back();
<% else %>
	location.href = "<%=gourl %>&<%=getGRSN() %>&returl=<%=Server.URLEncode(returl) %>";
<% end if %>
}

function show_it(name) {
	var show_span = document.getElementById(name + "_span")

	if (show_span.style.display == "none")
		show_span.style.display = "inline";
	else
		show_span.style.display = "none";
}

function m_over(tag_obj) {
	tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj) {
	tag_obj.style.backgroundColor = "white";
}
//-->
</script>

<BODY>
<table width="90%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr>
	<td align="left" height="28" style="padding-left:15px; color:#444;"><%=s_lang_0276 %></td>
	<td align="right" width="20%" style="padding-right:15px;"><a class='wwm_btnDownload btn_blue' href="javascript:back()"><%=s_lang_return %></a></td>
	</tr>
</table>
<br>
<table width="90%" border="0" align="center" cellspacing="0" style='border:1px solid #336699;'>
	<tr>
	<td width="12%" bgcolor="#EFF7FF" height="25" align="right" style="color:#104A7B;"><%=s_lang_0128 %><%=s_lang_mh %></td>
	<td width="88%" style='border-bottom:1px solid #8CA5B5;'>
	<table width="100%" border="0" cellspacing="0" align="center"><tr><td><%=get_date_showstr(slm.DateStr) %></td><td align="right"><img src="images/<%
if slm.Is_End = true then
	Response.Write "rc_end.gif"" title=""" & s_lang_0273
else
	Response.Write "rc_noend.gif"" title=""" & s_lang_0274
end if
%>" align='absmiddle' border='0'></td></tr>
	</table>
	</td></tr>
	<tr>
	<td height="25" bgcolor="#EFF7FF" style='border-bottom:1px solid #8CA5B5; color:#104A7B;' align="right"><%=s_lang_0127 %><%=s_lang_mh %></td>
	<td style='border-bottom:1px solid #8CA5B5; word-break:break-all; word-wrap:break-word;'><%=server.htmlencode(slm.Subject) %>&nbsp;</td>
	</tr>
	<tr><td colspan=2 style="padding:0px;">
	<table width="100%" border="0" cellspacing="0" align="center">
	<tr style='background-color:#EFF7FF; color:#444;'>
<td width='70%' height="25" style='padding-left:8px; border-right:1px solid #8CA5B5; border-bottom:1px solid #8CA5B5;'><%=s_lang_0149 %></td>
<td width='20%' style='padding-left:8px; border-right:1px solid #8CA5B5; border-bottom:1px solid #8CA5B5;'><%=s_lang_0128 %></td>
<td align="center" width='10%' style='border-bottom:1px solid #8CA5B5;'><%=s_lang_0126 %></td>
</tr>
<%
i = 0
allnum = slm.EmailCount

do while i < allnum
	slm.GetOneInfo i, email, name, is_send_ok, send_date, cmd_str

	if i + 1 < allnum then
		bottom_line = " style='word-break:break-all; word-wrap:break-word; border-bottom:1px solid #8CA5B5;'"
	else
		bottom_line = " style='word-break:break-all; word-wrap:break-word;'"
	end if

	Response.Write "<tr onmouseover='m_over(this);' onmouseout='m_out(this);'>"
	Response.Write "<td height='24'" & bottom_line & ">" & server.htmlencode(name)

	if Len(cmd_str) > 2 then
		Response.Write "<br><span id=""" & i & "_span"" name=""" & i & "_span"" style=""display:none""><font color='#901111'>"
		ht = server.htmlencode(cmd_str)
		ht = replace(ht, Chr(13), "<br>")
		ht = replace(ht, Chr(32), "&nbsp;")
		ht = replace(ht, Chr(9), "&nbsp;&nbsp;&nbsp;&nbsp;")
		Response.Write ht
		Response.Write "</font></span>"
	end if

	Response.Write "</td>"

	if Len(send_date) > 4 then
		Response.Write "<td nowrap" & bottom_line & ">" & get_date_showstr(send_date) & "</td>"
	else
		Response.Write "<td nowrap" & bottom_line & ">" & get_date_showstr(Left(slm.DateStr, 12)) & "</td>"
	end if

	Response.Write "<td align='center' nowrap"
	if i + 1 < allnum then
		Response.Write " style='border-bottom:1px solid #8CA5B5;'"
	end if
	Response.Write ">"

	if is_send_ok = true then
		if Len(cmd_str) > 2 then
			Response.Write "<a href=""javascript:show_it('" & i & "')"">" & s_lang_0277 & "</a></td>"
		else
			Response.Write s_lang_0277 & "</td>"
		end if
	else
		if Len(cmd_str) > 2 then
			Response.Write "<a href=""javascript:show_it('" & i & "')"">" & s_lang_0278 & "</a></td>"
		else
			Response.Write s_lang_0278 & "</td>"
		end if
	end if

	Response.Write "</tr>" & Chr(13)

	email = NULL
	name = NULL
	is_send_ok = NULL
	send_date = NULL
	cmd_str = NULL

	i = i + 1
loop
%>
</table>
</td></tr>
</table>
</BODY>
</HTML>

<%
set slm = nothing

function get_date_showstr(show_date_str)
	if Len(show_date_str) = 14 or Len(show_date_str) = 12 then
		tmp_month = Mid(show_date_str, 5, 2)
		if Mid(tmp_month, 1, 1) = "0" then
			tmp_month = Mid(tmp_month, 2, 1)
		end if

		tmp_day = Mid(show_date_str, 7, 2)
		if Mid(tmp_day, 1, 1) = "0" then
			tmp_day = Mid(tmp_day, 2, 1)
		end if

		get_date_showstr = Mid(show_date_str, 1, 4) & s_lang_0139 & tmp_month & s_lang_0140 & tmp_day & s_lang_0141 & " " & Mid(show_date_str, 9, 2) & ":" & Mid(show_date_str, 11, 2)

		if Len(show_date_str) = 14 then
			get_date_showstr = get_date_showstr & ":" & Mid(show_date_str, 13, 2)
		end if
	else
		get_date_showstr = ""
	end if
end function
%>
