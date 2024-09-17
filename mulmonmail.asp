<!--#include file="passinc.asp" --> 

<%
if isadmin() = false then
	Response.Redirect "noadmin.asp"
end if
%>

<%
user = trim(request("user"))
inout = trim(request("inout"))
filename = trim(request("filename"))

dim ei
set ei = server.createobject("easymail.ListUserMonitorMails")

if inout = "in" then
	ei.Load_InMails(user)
else
	ei.Load_OutMails(user)
end if

if filename <> "" then
	ei.DelMail filename
else
	allnum = ei.Count

	dim themax
	if allnum > pageline then
		themax = pageline
	else
		themax = allnum
	end if

	i = 0
	do while i <= themax
		if trim(request("check" & i)) <> "" then
			ei.DelMail trim(request("check" & i))
		end if 

	    i = i + 1
	loop
end if

set ei = nothing

if err.Number = 0 then
	if trim(request("gourl")) = "" then
		Response.Redirect "ok.asp"
	else
		Response.Redirect "ok.asp?gourl=" & Server.URLEncode(trim(request("gourl")))
	end if
else
	Response.Redirect "err.asp"
end if
%>
