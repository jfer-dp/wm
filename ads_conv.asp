<!--#include file="passinc.asp" -->

<%
gourl = trim(request("gourl"))
cfrom = trim(request("cfrom"))
cto = trim(request("cto"))

if cto = "pads" then
	if isadmin() = false then
		response.redirect "noadmin.asp"
	end if
end if

if cto = "dmpads" then
	is_domain_admin = false

	if isadmin() = false then
		dim dm
		set dm = server.createobject("easymail.Domain")
		dm.Load

		allnum = dm.GetUserManagerDomainCount(Session("wem"))

		curdomain = Mid(Session("mail"), InStr(Session("mail"), "@") + 1)

		i = 0
		do while i < allnum
			domain = dm.GetUserManagerDomain(Session("wem"), i)
	
			if LCase(curdomain) = LCase(domain) then
				is_domain_admin = true
			end if

			domain = NULL

			i = i + 1
		loop

		set dm = nothing
	else
		is_domain_admin = true
	end if

	if is_domain_admin = false then
		response.redirect "noadmin.asp"
	end if
end if



dim ads

if cfrom = "ads" then
	set ads = server.createobject("easymail.Addresses")
	ads.Load Session("wem")
elseif cfrom = "pads" then
	set ads = server.createobject("easymail.Pub_Addresses")
	ads.Load
elseif cfrom = "dmpads" then
	set ads = server.createobject("easymail.DomainPubAddresses")
	ads.Load Session("wem")
end if


dim isok
isok = false

dim msg
msg = trim(Request("upinfo"))

if Len(msg) > 0 then
	dim item
	dim ss
	dim se
	ss = 1
	se = 1

	Do While 1
		se = InStr(ss, msg, Chr(9))

		If se <> 0 Then
			item = Mid(msg, ss, se - ss)

			if item <> "" then
				if cto = "pads" then
					if ads.AddInPublicAddresses(item) = true then
						isok = true
					end if
				elseif cto = "ads" then
					if cfrom = "pads" then
						if ads.AddInPrivateAddresses(Session("wem"), item) = true then
							isok = true
						end if
					elseif cfrom = "dmpads" then
						if ads.AddInPrivateAddresses(item) = true then
							isok = true
						end if
					end if
				elseif cto = "dmpads" then
					if cfrom = "pads" then
						if ads.AddInDomainPublicAddresses(Session("wem"), item) = true then
							isok = true
						end if
					elseif cfrom = "ads" then
						if ads.AddInDomainPublicAddresses(item) = true then
							isok = true
						end if
					end if
				end if
			end if 
		Else
			Exit Do
		End If

		ss = se + 1
	Loop
end if

set ads = nothing

if isok = false then
	Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
else
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if
%>
