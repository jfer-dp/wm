<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	dim mam
	set mam = server.createobject("easymail.AdminManager")
	mam.Load

	if mam.Enable_DomainAdmin_SetWelcomeMsg = false then
		set mam = nothing
		response.redirect "noadmin.asp"
	end if

	set mam = nothing
end if


dim dm
set dm = server.createobject("easymail.Domain")
dm.Load

if dm.GetUserManagerDomainCount(Session("wem")) < 1 then
	if isadmin() = false then
		set dm = nothing
		response.redirect "noadmin.asp"
	end if
end if


if Request.ServerVariables("REQUEST_METHOD") = "POST" and trim(request("curdomain")) <> "" then
	dim ei
	set ei = server.createobject("easymail.Domain_Welcome_Msg")

	ei.Load

	if isadmin() = true then
		if trim(request("cleanall")) = "yes" then
			ei.RemoveAll
		else
			if trim(request("changeSystemWelcome")) <> "" then
				ei.Set trim(request("curdomain")), "", ""
			else
				ei.Set trim(request("curdomain")), trim(request("wsubject")), trim(request("wtext"))
			end if
		end if
	else
		allnum = dm.GetUserManagerDomainCount(Session("wem"))
		isok = false
		i = 0

		do while i < allnum
			if trim(request("curdomain")) = dm.GetUserManagerDomain(Session("wem"), i) then
				isok = true
	            exit do
			end if

			i = i + 1
		loop

		if isok = true then
			if trim(request("cleanall")) = "yes" then
				i = 0

				do while i < allnum
					ei.Set dm.GetUserManagerDomain(Session("wem"), i), "", ""

					i = i + 1
				loop
			else
				if trim(request("changeSystemWelcome")) <> "" then
					ei.Set trim(request("curdomain")), "", ""
				else
					ei.Set trim(request("curdomain")), trim(request("wsubject")), trim(request("wtext"))
				end if
			end if
		end if
	end if

	ei.Save

	set ei = nothing
end if

set dm = nothing

response.redirect "ok.asp?" & getGRSN() & "&gourl=showwelcome.asp?selectdomain=" & trim(request("curdomain"))
%>
