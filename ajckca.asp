<%
Response.Charset="GB2312"
%>

<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" -->

<%
cr_to = UnEscape(trim(request("to")))
cr_cc = UnEscape(trim(request("cc")))
cr_bcc = UnEscape(trim(request("bcc")))

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

if rcok = true then
%>
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:12px;"></td></tr>
<tr><td align="center">
<div class="wwm_line_msg"><%=s_lang_0293 %></div>
</td></tr></table>
<%
else
%>
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="white">
	<tr bgcolor="#104A7B">
	<td height="20" nowrap align="center" style="color:white;">
	<%=s_lang_0294 %>
	</td></tr>
	<tr><td height="20" align="left" style="color:#333333; padding:2px;">
	<%=msg %>
	</td></tr>
</table>
<%
end if

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

Function UnEscape(str)
    dim i,s,c
    s=""
    For i=1 to Len(str)
        c=Mid(str,i,1)
        If Mid(str,i,2)="%u" and i<=Len(str)-5 Then
            If IsNumeric("&H" & Mid(str,i+2,4)) Then
                s = s & CHRW(CInt("&H" & Mid(str,i+2,4)))
                i = i+5
            Else
                s = s & c
            End If
        ElseIf c="%" and i<=Len(str)-2 Then
            If IsNumeric("&H" & Mid(str,i+1,2)) Then
                s = s & CHRW(CInt("&H" & Mid(str,i+1,2)))
                i = i+2
            Else
                s = s & c
            End If
        Else
            s = s & c
        End If
    Next
    UnEscape = s
End Function
%>
