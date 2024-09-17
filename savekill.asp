<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei

set ei = server.createobject("easymail.kill")
'-----------------------------------------
ei.Load

allnum = ei.getcount


ei.RemoveAll


i = 0
if trim(request("mode")) = "save" then
	do while i < allnum + 1
		if trim(request("kill" & i)) <> "" then
			ei.AddKill trim(request("kill" & i))
		end if 

	    i = i + 1
	loop
elseif trim(request("mode")) = "add" then
	do while i < allnum + 1
		if trim(request("kill" & i)) <> "" then
			ei.AddKill trim(request("kill" & i))
		end if 

	    i = i + 1
	loop
elseif trim(request("mode")) = "del" then
	do while i < allnum + 1
		if trim(request("check" & i)) = "" and trim(request("kill" & i)) <> "" then
			ei.AddKill trim(request("kill" & i))
		end if 

	    i = i + 1
	loop
end if

ei.Save

set ei = nothing


if trim(request("mode")) = "add" then
	response.redirect "showkill.asp?mode=add&" & getGRSN()
else
	response.redirect "ok.asp?" & getGRSN() & "&gourl=showkill.asp"
end if
%>
