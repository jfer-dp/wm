<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim ei

	set ei = server.createobject("easymail.KillAttack")
	'-----------------------------------------
	ei.Load

	i = 0
	allnum = ei.Count

	ei.RemoveAll

	if trim(request("mode")) = "save" then
		do while i < allnum + 1
			if trim(request("ip" & i)) <> "" then
				ei.Add trim(request("ip" & i)), CLng(trim(request("rate" & i)))
			end if 

	    	i = i + 1
		loop
	elseif trim(request("mode")) = "add" then
		do while i < allnum + 1
			if trim(request("ip" & i)) <> "" then
				ei.Add trim(request("ip" & i)), CLng(trim(request("rate" & i)))
			end if 

		    i = i + 1
		loop
	elseif trim(request("mode")) = "del" then
		do while i < allnum + 1
			if trim(request("check" & i)) = "" and trim(request("ip" & i)) <> "" then
				ei.Add trim(request("ip" & i)), CLng(trim(request("rate" & i)))
			end if 

		    i = i + 1
		loop
	end if

	ei.Save

	set ei = nothing
end if

if trim(request("mode")) = "add" then
	response.redirect "showkillattack.asp?mode=add&" & getGRSN()
else
	response.redirect "ok.asp?" & getGRSN() & "&gourl=showkillattack.asp"
end if
%>
