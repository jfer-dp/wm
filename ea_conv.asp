<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
gourl = trim(request("gourl"))
rq_pid = trim(request("pid"))
cf = trim(request("cf"))

cv_nosame = trim(request("ns"))
cv_comm = trim(request("cm"))
cv_nodisabled = trim(request("nd"))
cv_domain = LCase(trim(request("dm")))

dim ads
themax = 0
if cf = "1" then
	set ads = server.createobject("easymail.Pub_Addresses")
	ads.Load
	themax = ads.Count
elseif cf = "2" then
	set ads = server.createobject("easymail.DomainPubAddresses")
	ads.Load Session("wem")
	themax = ads.Count
elseif cf = "3" then
	set ads = server.createobject("easymail.Addresses")
	ads.Load Session("wem")
	themax = ads.EmailCount
end if

dim eads
set eads = server.createobject("easymail.EntAddress")
eads.Load

isok = false
i = 0
if cf = "1" or cf = "2" or cf = "3" then
	do while i < themax
		ads.MoveTo i
		if cv_nosame = "0" or (cv_nosame = "1" and eads.HaveAds(ads.email) = false) then
			if eads.AddAds(ads.email, ads.nickname, "", rq_pid) = true then
				isok = true
			end if
		end if

	    i = i + 1
	loop

	set ads = nothing
elseif cf = "4" then
	dim ei
	set ei = Application("em")
	themax = ei.GetUsersCount

	do while i < themax
		ei.GetUserByIndex1 i, outname, outdomain, outcomment, outforbid, outlasttime

		if cv_comm = "0" then
			outcomment = ""
		end if

		if cv_domain = "" or (Len(cv_domain) > 1 and cv_domain = LCase(outdomain)) then
			if cv_nodisabled = "0" or (cv_nodisabled = "1" and outforbid = false) then
				if InStr(outname, "@") > 0 then
					if cv_nosame = "0" or (cv_nosame = "1" and eads.HaveAds(outname) = false) then
						if eads.AddAds(outname, outname, outcomment, rq_pid) = true then
							isok = true
						end if
					end if
				else
					if cv_nosame = "0" or (cv_nosame = "1" and eads.HaveAds(outname & "@" & outdomain) = false) then
						if eads.AddAds(outname & "@" & outdomain, outname, outcomment, rq_pid) = true then
							isok = true
						end if
					end if
				end if
			end if
		end if

		outname = NULL
		outdomain = NULL
		outcomment = NULL
		outforbid = NULL
		outlasttime = NULL

	    i = i + 1
	loop

	set ei = nothing
elseif cf = "5" then
	dim ldap
	set ldap = server.createobject("easymail.LDAP")
	isok = ldap.UpdateEntAddress(trim(request("ldapsel")))
	set ldap = nothing
end if

if cf <> "5" then
	if isok = true then
		isok = eads.Save()
	end if
end if

set eads = nothing

if isok = false then
	Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
else
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if
%>
