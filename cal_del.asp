<!--#include file="passinc.asp" -->

<%
if Len(Session("svcal")) > 0 then
	Response.Redirect "noadmin.asp"
end if

returl = trim(request("returl"))
calid = trim(request("calid"))
calmode = trim(request("calmode"))

dim ecal

if Len(calid) > 10 and calmode = "1" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	set ecal = server.createobject("easymail.Calendar")
	ecal.Load Session("wem")

	isok = ecal.RemoveByID(calid)

	if isok = true then
		isok = ecal.Save()
	end if

	set ecal = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	end if
end if


if Len(calid) < 1 and calmode = "2" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	set ecal = server.createobject("easymail.Calendar")
	ecal.Load Session("wem")

	allnum = ecal.Count
	if pageline < allnum then
		allnum = pageline
	end if

	isok = true
	haveok = false
	i = 0
	do while i < allnum
		temp_id = trim(request("check" & i))

		if temp_id <> "" then
			if ecal.RemoveByID(temp_id) = false then
				isok = false
			else
				haveok = true
			end if
		end if 

	    i = i + 1
	loop

	if haveok = true then
		if ecal.Save() = false then
			isok = false
		end if
	end if

	set ecal = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	end if
end if


if Len(calid) < 1 and calmode = "3" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	set ecal = server.createobject("easymail.CalTask")
	ecal.Load Session("wem")

	allnum = ecal.Count
	if pageline < allnum then
		allnum = pageline
	end if

	isok = true
	haveok = false
	i = 0
	do while i < allnum
		temp_id = trim(request("check" & i))

		if temp_id <> "" then
			if ecal.RemoveByID(temp_id) = false then
				isok = false
			else
				haveok = true
			end if
		end if 

	    i = i + 1
	loop

	if haveok = true then
		if ecal.Save() = false then
			isok = false
		end if
	end if

	set ecal = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	end if
end if


if Len(calid) < 1 and calmode = "4" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	set ecal = server.createobject("easymail.CalTask")
	ecal.Load Session("wem")

	allnum = ecal.Count
	if pageline < allnum then
		allnum = pageline
	end if

	isok = true
	haveok = false
	i = 0
	do while i < allnum
		temp_id = trim(request("check" & i))

		if temp_id <> "" then
			if ecal.SetCompleteByID(temp_id, true) = false then
				isok = false
			else
				haveok = true
			end if
		end if 

	    i = i + 1
	loop

	if haveok = true then
		if ecal.Save() = false then
			isok = false
		end if
	end if

	set ecal = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	end if
end if


if Len(calid) > 10 and calmode = "5" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	set ecal = server.createobject("easymail.CalTask")
	ecal.Load Session("wem")

	isok = ecal.RemoveByID(calid)

	if isok = true then
		isok = ecal.Save()
	end if

	set ecal = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	end if
end if


if Len(calid) > 10 and calmode = "6" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	set ecal = server.createobject("easymail.CalTask")
	ecal.Load Session("wem")

	isok = ecal.SetCompleteByID(calid, true)

	if isok = true then
		isok = ecal.Save()
	end if

	set ecal = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_tasknew.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_tasknew.asp")
		end if
	end if
end if


if Len(calid) < 1 and calmode = "7" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	set ecal = server.createobject("easymail.CalendarNotice")
	ecal.Load Session("wem")

	allnum = ecal.Count
	if pageline < allnum then
		allnum = pageline
	end if

	isok = true
	haveok = false
	i = 0
	do while i < allnum
		temp_id = trim(request("check" & i))

		if temp_id <> "" then
			if ecal.RemoveByID(temp_id) = false then
				isok = false
			else
				haveok = true
			end if
		end if 

	    i = i + 1
	loop

	if haveok = true then
		if ecal.Save() = false then
			isok = false
		end if
	end if

	set ecal = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_listinvited.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_listinvited.asp")
		end if
	end if
end if


if Len(calid) > 0 and calmode = "8" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	purl = trim(request("purl"))

	set ecal = server.createobject("easymail.CalendarExtend")
	ecal.Load Session("wem"), calid

	allnum = ecal.Count
	if pageline < allnum then
		allnum = pageline
	end if

	isok = true
	haveok = false
	i = 0
	do while i < allnum
		temp_id = trim(request("check" & i))

		if temp_id <> "" and LCase(temp_id) <> LCase(Session("mail")) then
			if ecal.RemoveByEmail(temp_id) = false then
				isok = false
			else
				haveok = true
			end if
		end if

	    i = i + 1
	loop

	if haveok = true then
		if ecal.Save() = false then
			isok = false
		end if
	end if

	set ecal = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl & "&" & purl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl & "&" & purl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	end if
end if


if Len(calid) > 0 and calmode = "9" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	email = trim(request("email"))
	purl = trim(request("purl"))

	set ecal = server.createobject("easymail.CalendarExtend")
	ecal.Load Session("wem"), calid

	isok = false
	if email <> "" and LCase(email) <> LCase(Session("mail")) then
		if ecal.RemoveByEmail(email) = true then
			if ecal.Save() = true then
				isok = true
			end if
		end if
	end if

	set ecal = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl & "&" & purl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl & "&" & purl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	end if
end if


if Len(calid) > 10 and calmode = "10" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	set ecal = server.createobject("easymail.CalendarNotice")
	ecal.Load Session("wem")

	isok = ecal.RemoveByID(calid)

	if isok = true then
		if ecal.Save() = false then
			isok = false
		end if
	end if

	set ecal = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_listinvited.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_listinvited.asp")
		end if
	end if
end if


if Len(returl) > 3 then
	Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
else
	Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
end if
%>
