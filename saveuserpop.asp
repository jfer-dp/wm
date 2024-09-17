<!--#include file="passinc.asp" -->

<%
dim ei
set ei = server.createobject("easymail.UserMessages")
'-----------------------------------------
ei.Load Session("wem")

i = ei.GetMulPop3Count


if trim(request("checkpop")) <> "" then
	ei.POP3Support = true
else
	ei.POP3Support = false
end if

if trim(request("mode")) = "save" then
	ei.SaveMPOP3
elseif trim(request("mode")) = "add" then
	isdel = trim(request("uisdel"))
	if isdel = "" then
		isdel = false
	else
		isdel = true
	end if

	ei.AddMulPop3 trim(request("uname")), trim(request("userver")), trim(request("uport")), trim(request("uusername")), trim(request("upassword")), isdel
	ei.SaveMPOP3
elseif trim(request("mode")) = "del" then
	do while i >= 0
		if trim(request("check" & i)) <> "" then
			ei.DeleteMulPop3 i
		end if 

	    i = i - 1
	loop

	ei.SaveMPOP3
end if

set ei = nothing

response.redirect "ok.asp?gourl=showuserpop.asp"
%>
