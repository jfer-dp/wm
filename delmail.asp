<!--#include file="passinc.asp" --> 

<%
if isadmin() = false and Session("ReadOnlyUser") = 1 then
	Response.Redirect "err.asp"
end if


filename = trim(request("filename"))
realdel = trim(request("realdel"))
gourl = trim(request("gourl"))
enc_gourl = Server.URLEncode(gourl)

if LCase(Right(filename, 4)) = ".arc" then
	set march = server.createobject("easymail.MailArchive")
	march.Load Session("wem")

	march.Del filename
	set march = nothing

	if gourl = "" then
		Response.Redirect "ok.asp"
	else
		Response.Redirect "ok.asp?gourl=" & enc_gourl
	end if
end if


dim ei
set ei = server.createobject("easymail.InfoList")

dim wmeth
set wmeth = server.createobject("easymail.WMethod")

if Len(filename) > 6 then
	tmpfile = Mid(filename, Len(filename) - 2, 3)

	if realdel <> "1" then
		if tmpfile = "del" then
			wmeth.Revoke_Delete_Begin Session("wem")
			wmeth.Revoke_Delete_File Session("wem"), filename
			ei.DelMail Session("wem"), filename
			wmeth.Revoke_Delete_End Session("wem")
		else
			ei.MoveMail Session("wem"), filename, "del"
		end if
	else
		wmeth.Revoke_Delete_Begin Session("wem")
		wmeth.Revoke_Delete_File Session("wem"), filename
		ei.DelMail Session("wem"), filename
		wmeth.Revoke_Delete_End Session("wem")
	end if

	set wmeth = nothing
	set ei = nothing

	if err.Number = 0 then
		if trim(request("nextfile")) = "" then
			if gourl = "" then
				Response.Redirect "ok.asp"
			else
				Response.Redirect "ok.asp?gourl=" & enc_gourl
			end if
		else
			Response.Redirect "showmail.asp?filename=" & trim(request("nextfile")) & "&gourl=" & enc_gourl
		end if
	else
		Response.Redirect "err.asp"
	end if
else
	msid = trim(request("msid"))
	mto = trim(request("mto"))

	if Len(msid) > 6 then
		if realdel <> "1" then
			if Len(mto) > 0 then
				ei.Move_Session Session("wem"), msid, mto
			else
				ei.Move_Session Session("wem"), msid, "del"
			end if
		else
			wmeth.Revoke_Delete_Begin Session("wem")
			ei.Del_Session Session("wem"), msid
			wmeth.Revoke_Delete_End Session("wem")
		end if

		set wmeth = nothing
		set ei = nothing

		if err.Number = 0 then
			if trim(request("nextfile")) = "" then
				if gourl = "" then
					Response.Redirect "ok.asp"
				else
					Response.Redirect "ok.asp?gourl=" & enc_gourl
				end if
			else
				Response.Redirect "showsession.asp?msid=" & trim(request("nextfile")) & "&gourl=" & enc_gourl
			end if
		else
			Response.Redirect "err.asp"
		end if
	end if
end if
%>
