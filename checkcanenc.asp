<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" --> 

<%
cr_to = trim(request("EasyMail_To"))
cr_cc = trim(request("EasyMail_Cc"))
cr_bcc = trim(request("EasyMail_Bcc"))

allrc = cr_to

if cr_cc <> "" then
	if allrc = "" then
		allrc = cr_cc
	else
		allrc = allrc & "," & cr_cc
	end if
end if

if cr_bcc <> "" then
	if allrc = "" then
		allrc = cr_bcc
	else
		allrc = allrc & "," & cr_bcc
	end if
end if

dim msg

if Len(allrc) > 0 then
	dim item
	dim ss
	dim se
	dim dse
	dim fse
	ss = 1
	se = 1
	dse = 1
	fse = 1

	Do While 1
		dse = InStr(ss, allrc, ",")
		fse = InStr(ss, allrc, ";")

		if dse < 1 and fse < 1 then
			se = 0
		elseif dse < 1 then
			se = fse
		elseif fse < 1 then
			se = dse
		else
			if dse < fse then
				se = dse
			else
				se = fse
			end if
		end if

		If se <> 0 Then
			item = remove_kh(Mid(allrc, ss, se - ss))
			if Len(item) > 0 then
				msg = msg & item & ","
			end if
		Else
			item = remove_kh(Mid(allrc, ss))
			if Len(item) > 0 then
				msg = msg & item
			end if

			Exit Do
		End If

		ss = se + 1
	Loop
end if


dim wemcert
set wemcert = server.createobject("easymail.WebEasyMailCert")
wemcert.Load Session("wem"), Session("mail")

wemcert.CheckEncodeEmailsHaveCert msg, badrc, rcok

msg = ""

if Len(badrc) > 0 then
	ss = 1
	se = 1

	Do While 1
		se = InStr(ss, badrc, ",")

		If se <> 0 Then
			item = Mid(badrc, ss, se - ss) & "<br>"
			msg = msg & item
		Else
			item = Mid(badrc, ss)
			msg = msg & item

			Exit Do
		End If

		ss = se + 1
	Loop
end if
%>

<!DOCTYPE html>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
<link rel="stylesheet" type="text/css" href="images/hwem.css">
</HEAD>

<BODY>
<br><br>
<table border="0" width="60%" cellpadding=2 cellspacing=0 align="center" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_6 %>">
	<td height="20" width="60%" nowrap align="center">
<%
if rcok = true then
%> 
	<font class="fw" size=2 style=font-size:9pt color="#ffffff"><b><%=s_lang_0293 %></b></font>
<%
else
%> 
	<font class="fw" size=2 style=font-size:9pt color="#ffffff"><b><%=s_lang_0294 %></b></font>
<%
end if
%> 
    </td></tr>
	<tr bgcolor="#ffffff">
    <td height="20" align="left"> 
<font size=2 style=font-size:9pt><%=msg %></font>
    </td></tr>
</table>

<table width="60%" cellpadding=0 cellspacing=0 align="center">
<tr>
<td align="right">
<br>
<hr size="1" color="<%=MY_COLOR_1 %>">
<a class="wwm_btnDownload btn_blue" style="width: 40px;" href="#" onclick="javascript:self.close();"><%=s_lang_0059 %></a>
</td>
</tr>
</table>

</BODY>
</HTML>

<%
badrc = NULL
rcok = NULL

set wemcert = nothing

function remove_kh(one_str)
	remove_kh = one_str
	dim kh_s
	dim kh_e
	kh_s = 0
	kh_e = 0
	kh_s = InStr(1, one_str, "<")

	if kh_s > 0 then
		kh_s = kh_s + 1
		kh_e = InStr(kh_s, one_str, ">")

		if kh_e > 0 then
			remove_kh = Mid(one_str, kh_s, kh_e - kh_s)
		end if
	end if
end function
%>
