<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.IspManager")
'-----------------------------------------
ei.Load

i = ei.count


if trim(request("mode")) = "add" then
	ei.Add trim(request("uname")), trim(request("userver")), trim(request("uport")), trim(request("uusername")), trim(request("upassword"))
	ei.Save
elseif trim(request("mode")) = "del" then
	do while i >= 0
		if trim(request("check" & i)) <> "" then
			ei.DeleteByIndex i
		end if

	    i = i - 1
	loop

	ei.Save
end if

set ei = nothing

response.redirect "ok.asp?" & getGRSN() & "&gourl=showmisp.asp"
%>
