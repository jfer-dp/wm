<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei

set ei = server.createobject("easymail.stakeout")
ei.Load
'-----------------------------------------

ei.RemoveAll


dim msg
msg = trim(request("addlist"))

if Len(msg) > 0 then
	dim ss
	dim se
	ss = 1
	se = 1

    Do While 1
        se = InStr(ss, msg, Chr(9))

        If se <> 0 Then
   	        item = Mid(msg, ss, se - ss)
			ei.Add item
		Else
            Exit Do
   	    End If

        ss = se + 1
    Loop

	ei.Save
else
	ei.RemoveAll
	ei.Save
end if


set ei = nothing


if err.number = 0 then
	response.redirect "ok.asp?" & getGRSN() & "&gourl=showstakeout.asp"
else
	response.redirect "err.asp?" & getGRSN() & "&gourl=showstakeout.asp"
end if
%>
