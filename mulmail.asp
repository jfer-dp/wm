<!--#include file="passinc.asp" --> 

<%
if isadmin() = false and Session("ReadOnlyUser") = 1 then
	Response.Redirect "err.asp"
end if


dim themax
dim tmpfile
dim ei

shmode = trim(request("mode"))
shfname = trim(request("mto"))

if LCase(Left(shmode, 2)) = "sh" then
	if Session("SH_Admin") = false then
		Response.Redirect "noadmin.asp"
	end if

	dim shm
	set shm = server.createobject("easymail.SH_Manager")

	ispass = false
	if Right(shmode, 1) = "1" then
		ispass = true
	end if

	sh_mode = LCase(Mid(shmode, 3, 3))
	if sh_mode = "one" then
		shm.SetPass Left(shfname, Len(shfname) - 3), ispass
	else
		if sh_mode = "all" then
			shm.SetAllPass ispass
		else
			if sh_mode = "mul" then
				set ei = server.createobject("easymail.InfoList")
				ei.Load_SH_Mails
				allnum = ei.getMailsCount

				if allnum > pageline then
					themax = pageline
				else
					themax = allnum
				end if

				i = 0
				do while i <= themax
					tmp_ck = trim(request("ck_" & i))
					if tmp_ck <> "" then
						shm.SetPass Left(tmp_ck, Len(tmp_ck) - 3), ispass
					end if 

				    i = i + 1
				loop

				set ei = nothing
			end if
		end if
	end if

	set shm = nothing

	if err.Number = 0 then
		if trim(request("gourl")) = "" then
			Response.Redirect "ok.asp"
		else
			Response.Redirect "ok.asp?gourl=" & Server.URLEncode(trim(request("gourl")))
		end if
	else
		Response.Redirect "err.asp"
	end if
end if


set ei = server.createobject("easymail.InfoList")

dim wmeth
set wmeth = server.createobject("easymail.WMethod")

dim march

if trim(request("isatt")) = "1" then
	ei.IsAttFolder = true
end if

if trim(request("mode")) = "cleanAtt" then
	ei.LoadMailBox Session("wem"), "att"
	allnum = ei.getMailsCount

	i = 0

	do while i < allnum
		ei.getMailInfo allnum - i - 1, idname, isread, priority, sendMail, sendName, subject, size, etime
		ei.DelMail Session("wem"), idname

		idname = NULL
		isread = NULL
		priority = NULL
		sendMail = NULL
		sendName = NULL
		subject = NULL
		size = NULL
		etime = NULL

		i = i + 1
	loop

	set wmeth = nothing
	set ei = nothing

	if err.Number = 0 then
		if trim(request("gourl")) = "" then
			response.redirect "ok.asp"
		else
			response.redirect "ok.asp?gourl=" & Server.URLEncode(trim(request("gourl")))
		end if
	else
		response.redirect "err.asp"
	end if
end if


if trim(request("mode")) = "cleanTrash" then
	ei.LoadMailBox Session("wem"), "del"
	allnum = ei.getMailsCount

	i = 0

	do while i < allnum
		ei.getMailInfo allnum - i - 1, idname, isread, priority, sendMail, sendName, subject, size, etime
		ei.DelMail Session("wem"), idname

		idname = NULL
		isread = NULL
		priority = NULL
		sendMail = NULL
		sendName = NULL
		subject = NULL
		size = NULL
		etime = NULL

		i = i + 1
	loop

	set wmeth = nothing
	set ei = nothing

	if err.Number = 0 then
		if trim(request("gourl")) = "" then
			response.redirect "ok.asp"
		else
			response.redirect "ok.asp?gourl=" & Server.URLEncode(trim(request("gourl")))
		end if
	else
		response.redirect "err.asp"
	end if
end if


if trim(request("mode")) = "mitt" then
	ei.LoadMailBox Session("wem"), "in"
	allnum = ei.getMailsCount

	i = 0

	do while i < allnum
		ei.getMailInfo allnum - i - 1, idname, isread, priority, sendMail, sendName, subject, size, etime
		ei.MoveMail Session("wem"), idname, "del"

		idname = NULL
		isread = NULL
		priority = NULL
		sendMail = NULL
		sendName = NULL
		subject = NULL
		size = NULL
		etime = NULL

		i = i + 1
	loop

	set wmeth = nothing
	set ei = nothing

	if err.Number = 0 then
		if trim(request("gourl")) = "" then
			response.redirect "ok.asp"
		else
			response.redirect "ok.asp?gourl=" & Server.URLEncode(trim(request("gourl")))
		end if
	else
		response.redirect "err.asp"
	end if
end if


ei.LoadSizeInfo Session("wem")
mto = trim(request("mto"))

allnum = ei.allMailCount

if allnum > pageline then
	themax = pageline
else
	themax = allnum
end if


if trim(request("mode")) = "arcdel" then
	set march = server.createobject("easymail.MailArchive")
	march.Load Session("wem")
	i = 0

	do while i <= themax
		tmp_ck = trim(request("ck_" & i))
		if tmp_ck <> "" then
			march.Del tmp_ck
		end if

	    i = i + 1
	loop

	set march = nothing
	set wmeth = nothing
	set ei = nothing

	if err.Number = 0 then
		if trim(request("gourl")) = "" then
			response.redirect "ok.asp"
		else
			response.redirect "ok.asp?gourl=" & Server.URLEncode(trim(request("gourl")))
		end if
	else
		response.redirect "err.asp"
	end if
end if


dim emm
set emm = server.createobject("easymail.emmail")

i = 0

if trim(request("isremove")) = "0" or trim(request("isremove")) = "" then
	if trim(request("mode")) = "move" then
		do while i <= themax
			tmp_ck = trim(request("ck_" & i))
			if tmp_ck <> "" then
				ei.MoveMail Session("wem"), tmp_ck, mto
			end if 

		    i = i + 1
		loop
	elseif trim(request("mode")) = "semove" then
		do while i <= themax
			tmp_ck = trim(request("ck_" & i))
			if tmp_ck <> "" then
				ei.Move_Session Session("wem"), tmp_ck, mto
			end if 

		    i = i + 1
		loop
	elseif trim(request("mode")) = "spam" then
		rdel = false

		dim usg
		set usg = server.createobject("easymail.UserSpamGuard")
		usg.LightLoad Session("wem")

		if usg.SpamProcessMode = 0 then
			rdel = true
		end if

		set usg = nothing


		do while i <= themax
			tmp_ck = trim(request("ck_" & i))
			if tmp_ck <> "" then
				emm.AddSpamEx tmp_ck, Session("wem"), Session("tid")

				if rdel = true then
					ei.DelMail Session("wem"), tmp_ck
				else
					ei.MoveMail Session("wem"), tmp_ck, "del"
				end if
			end if 

		    i = i + 1
		loop
	elseif trim(request("mode")) = "m2arc" then
		set march = server.createobject("easymail.MailArchive")
		march.Load Session("wem")

		do while i <= themax
			tmp_ck = trim(request("ck_" & i))
			if tmp_ck <> "" then
				march.Move_mail_to_archive tmp_ck
			end if 

		    i = i + 1
		loop

		set march = nothing
	elseif trim(request("mode")) = "arc2m" then
		set march = server.createobject("easymail.MailArchive")
		march.Load Session("wem")

		do while i <= themax
			tmp_ck = trim(request("ck_" & i))
			if tmp_ck <> "" then
				march.Move_archive_to_mail tmp_ck, mto
			end if 

		    i = i + 1
		loop

		set march = nothing
	else
		wmeth.Revoke_Delete_Begin Session("wem")
		do while i <= themax
			tmp_ck = trim(request("ck_" & i))
			if tmp_ck <> "" then
				tmpfile = Mid(tmp_ck, Len(tmp_ck) - 2, 3)

				if tmpfile = "del" then
					wmeth.Revoke_Delete_File Session("wem"), tmp_ck
					ei.DelMail Session("wem"), tmp_ck
				else
					ei.MoveMail Session("wem"), tmp_ck, "del"
				end if
			end if 

		    i = i + 1
		loop
		wmeth.Revoke_Delete_End Session("wem")
	end if
else
	if trim(request("mode")) = "del" then
		wmeth.Revoke_Delete_Begin Session("wem")
		do while i <= themax
			tmp_ck = trim(request("ck_" & i))
			if tmp_ck <> "" then
				wmeth.Revoke_Delete_File Session("wem"), tmp_ck
				ei.DelMail Session("wem"), tmp_ck
			end if

		    i = i + 1
		loop
		wmeth.Revoke_Delete_End Session("wem")
	elseif trim(request("mode")) = "sedel" then
		wmeth.Revoke_Delete_Begin Session("wem")
		do while i <= themax
			tmp_ck = trim(request("ck_" & i))
			if tmp_ck <> "" then
				ei.Del_Session Session("wem"), tmp_ck
			end if

		    i = i + 1
		loop
		wmeth.Revoke_Delete_End Session("wem")
	end if
end if

set wmeth = nothing
set ei = nothing
set emm = nothing


if err.Number = 0 then
	if trim(request("gourl")) = "" then
		response.redirect "ok.asp"
	else
		response.redirect "ok.asp?gourl=" & Server.URLEncode(trim(request("gourl")))
	end if
else
	response.redirect "err.asp"
end if
%>
