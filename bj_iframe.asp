<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" -->

<%
if isadmin() = false then
	Response.Redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.MailboxBanjia")
ei.Load
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<LINK href="images/hwem.css" rel=stylesheet>

<STYLE type=text/css>
html {overflow: scroll; overflow-x: hidden; overflow-y: auto !important;}
.td_null {white-space:nowrap; background-color:white; border-bottom:1px #A5B6C8 solid; height:22px; color:#202020;}
.td_yes {white-space:nowrap; background-color:#e3f6f7; border-bottom:1px #A5B6C8 solid; height:22px; color:#202020;}
.td_no {white-space:nowrap; background-color:#f8dff6; border-bottom:1px #A5B6C8 solid; height:22px; color:#202020;}
</STYLE>
</HEAD>

<BODY style="padding:0px; margin:0px;">
<table width="100%" border="0" align="center" cellspacing="0">
<%
i = 0
allnum = ei.Count
run_index = ei.Processing_Index

if ei.IsRun = false then
	run_index = -1
end if

do while i < allnum
	ei.Get i, s_name, s_state

	if s_state = 2 then
		if i = 0 then
			Response.Write "<tr><td width='30%' class='td_no' align='center' style='border-top:1px #A5B6C8 solid;'>" & b_lang_388 & "</td><td class='td_no' style='border-top:1px #A5B6C8 solid;'>"
		else
			Response.Write "<tr><td class='td_no' align='center'>" & b_lang_388 & "</td><td class='td_no'>"
		end if
	elseif s_state = 1 then
		if i = 0 then
			Response.Write "<tr><td width='30%' class='td_yes' align='center' style='border-top:1px #A5B6C8 solid;'>" & b_lang_389 & "</td><td class='td_yes' style='border-top:1px #A5B6C8 solid;'>"
		else
			Response.Write "<tr><td class='td_yes' align='center'>" & b_lang_389 & "</td><td class='td_yes'>"
		end if
	else
		if i = run_index then
			if i = 0 then
				Response.Write "<tr><td width='30%' class='td_null' align='center' style='border-top:1px #A5B6C8 solid;'><img src='images/bjing.gif' border=0 align='absmiddle' title='" & b_lang_390 & "'>&nbsp;" & b_lang_390 & "</td><td class='td_null' style='border-top:1px #A5B6C8 solid;'>"
			else
				Response.Write "<tr><td class='td_null' align='center'><img src='images/bjing.gif' border=0 align='absmiddle' title='" & b_lang_390 & "'>&nbsp;" & b_lang_390 & "</td><td class='td_null'>"
			end if
		else
			if i = 0 then
				Response.Write "<tr><td width='30%' class='td_null' style='border-top:1px #A5B6C8 solid;'>&nbsp;</td><td class='td_null' style='border-top:1px #A5B6C8 solid;'>"
			else
				Response.Write "<tr><td class='td_null'>&nbsp;</td><td class='td_null'>"
			end if
		end if
	end if

	Response.Write server.htmlencode(s_name) & "</td></tr>" & Chr(13)

	s_name = NULL
	s_state = NULL
	i = i + 1
loop
%>
</td></tr>
</table>
</BODY>
</HTML>

<%
set ei = nothing
%>
