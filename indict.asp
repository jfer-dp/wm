<!--#include file="passinc.asp" -->

<%
filename = trim(request("filename"))

dim ei
set ei = server.createobject("easymail.emmail")

if trim(request("mode")) = "at" then
	ei.AddTrust filename, Session("wem"), Session("tid")
	set ei = nothing

	Response.Redirect "movemail.asp?filename=" & Server.URLEncode(filename) & "&mto=in&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl")))
end if

ei.AddSpamEx filename, Session("wem"), Session("tid")

set ei = nothing

if trim(request("realdel")) = "1" then
	Response.Redirect "delmail.asp?realdel=1&filename=" & filename & "&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl")))
else
	rdel = false

	dim usg
	set usg = server.createobject("easymail.UserSpamGuard")
	usg.LightLoad Session("wem")

	if usg.SpamProcessMode = 0 then
		rdel = true
	end if

	set usg = nothing

	if rdel = true then
		Response.Redirect "delmail.asp?realdel=1&filename=" & filename & "&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl")))
	else
		Response.Redirect "delmail.asp?filename=" & filename & "&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl")))
	end if
end if
%>
