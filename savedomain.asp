<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei

set ei = server.createobject("easymail.domain")
'-----------------------------------------
ei.Load

i = 0
allnum = ei.getcount


i = 0
if trim(request("mode")) = "save" then
	do while i < allnum + 1
		if trim(request("domain" & i)) <> "" and trim(request("domain" & i)) <> "system.mail" then
			ei.AddDomain trim(request("domain" & i))
		end if 

	    i = i + 1
	loop
elseif trim(request("mode")) = "add" then
	do while i < allnum + 1
		if trim(request("domain" & i)) <> "" and trim(request("domain" & i)) <> "system.mail" then
			ei.AddDomain trim(request("domain" & i))
		end if 

	    i = i + 1
	loop
elseif trim(request("mode")) = "del" then
	do while i < allnum + 1
		if trim(request("check" & i)) <> "" and trim(request("domain" & i)) <> "" and trim(request("domain" & i)) <> "system.mail" then
			ei.DelDomainByName trim(request("domain" & i))
		end if 

	    i = i + 1
	loop
end if

ei.Save

set ei = nothing


if trim(request("mode")) = "add" then
	response.redirect "showdomain.asp?mode=add&" & getGRSN()
else
	response.redirect "ok.asp?" & getGRSN() & "&gourl=showdomain.asp"
end if
%>
