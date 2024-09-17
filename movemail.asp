<!--#include file="passinc.asp" --> 

<%
if isadmin() = false and Session("ReadOnlyUser") = 1 then
	Response.Redirect "err.asp"
end if


filename = trim(request("filename"))
mto = trim(request("mto"))

if LCase(mto) = ".arc" then
	set march = server.createobject("easymail.MailArchive")
	march.Load Session("wem")

	march.Move_mail_to_archive filename
	set march = nothing

	if trim(request("gourl")) = "" then
		Response.Redirect "ok.asp"
	else
		Response.Redirect "ok.asp?gourl=" & Server.URLEncode(trim(request("gourl")))
	end if
end if


if LCase(Right(filename, 4)) = ".arc" then
	set march = server.createobject("easymail.MailArchive")
	march.Load Session("wem")

	march.Move_archive_to_mail filename, mto
	set march = nothing

	if trim(request("gourl")) = "" then
		Response.Redirect "ok.asp"
	else
		Response.Redirect "ok.asp?gourl=" & Server.URLEncode(trim(request("gourl")))
	end if
end if


dim ei
set ei = server.createobject("easymail.InfoList")

if ei.MoveMail(Session("wem"), filename, mto) = true then
	set ei = nothing

	if trim(request("nextfile")) = "" then
		if trim(request("gourl")) = "" then
			response.redirect "ok.asp"
		else
			response.redirect "ok.asp?gourl=" & Server.URLEncode(trim(request("gourl")))
		end if
	else
		response.redirect "showmail.asp?filename=" & trim(request("nextfile")) & "&gourl=" & Server.URLEncode(trim(request("gourl")))
	end if
else
	set ei = nothing
	response.redirect "err.asp"
end if
%>
