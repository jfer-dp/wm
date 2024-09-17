<!--#include file="passinc.asp" -->

<%
if isadmin() = false and isAccountsAdmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim pr
set pr = server.createobject("easymail.PendRegister")
pr.Load Application("em_SignWaitDays")

mdel = trim(Request("mdel"))

dim isedit
isedit = false

if mdel = "1" then
	themax = pr.Count

	do while themax >= 0
		if trim(request("check" & themax)) <> "" then
			md = trim(request("check" & themax))

			if md <> "" then
				pr.RemoveSign md
				isedit = true
			end if
		end if 

	    themax = themax - 1
	loop
else
	if trim(request("id")) <> "" then
		md = trim(request("id"))

		if md <> "" then
			pr.RemoveSign md
			isedit = true
		end if
	end if 
end if

if isedit = true then
	pr.Save
end if

set pr = nothing
response.redirect "pendreg.asp?" & getGRSN() & "&page=" & trim(request("thispage"))
%>
