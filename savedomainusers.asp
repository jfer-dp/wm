<!--#include file="passinc.asp" -->

<%
dim cdomain
cdomain = trim(request("cdomain"))
if cdomain = "" then
	response.redirect "noadmin.asp"
end if


dim ed
set ed = server.createobject("easymail.domain")
ed.Load

if ed.GetUserManagerDomainCount(Session("wem")) < 1 then
	set ed = nothing
	response.redirect "noadmin.asp"
end if


i = 0
allnum = ed.GetUserManagerDomainCount(Session("wem"))

dim isok
isok = false

do while i < allnum
	cdomainstr = ed.GetUserManagerDomain(Session("wem"), i)

	if cdomainstr = cdomain then
		isok = true
	end if

	cdomainstr = NULL

	i = i + 1
loop

set ed = nothing


if isok = false then
	response.redirect "noadmin.asp"
end if
%>


<%
dim ei
set ei = Application("em")
'-----------------------------------------

i = ei.GetUsersCount


if trim(request("mode")) = "del" then

	do while i >= 0
		if trim(request("check" & i)) <> "" then
			ei.GetUserByIndex1 i, name, domain, comment, forbid, lasttime

			if domain = cdomain and name <> Session("wem") and name <> Application("em_SystemAdmin") then
				ei.DelUserByIndex i
			end if
		end if 

		name = NULL
		domain = NULL
		comment = NULL
		forbid = NULL
		lasttime = NULL

	    i = i - 1
	loop

elseif trim(request("mode")) = "forbid" then

	do while i >= 0
		if trim(request("check" & i)) <> "" then
			ei.GetUserByIndex1 i, name, domain, comment, forbid, lasttime

			if domain = cdomain and name <> Session("wem") and name <> Application("em_SystemAdmin") then
				ei.ForbidUserByIndex i, TRUE
			end if
		end if 

		name = NULL
		domain = NULL
		comment = NULL
		forbid = NULL
		lasttime = NULL

	    i = i - 1
	loop

elseif trim(request("mode")) = "clear" then

	do while i >= 0
		if trim(request("check" & i)) <> "" then
			ei.GetUserByIndex1 i, name, domain, comment, forbid, lasttime

			if domain = cdomain then
				ei.ForbidUserByIndex i, FALSE
			end if
		end if 

		name = NULL
		domain = NULL
		comment = NULL
		forbid = NULL
		lasttime = NULL

	    i = i - 1
	loop

end if


set ei = nothing


searchtext = trim(request("searchtext"))
page = trim(request("page"))
sortby = trim(request("sortby"))

response.redirect "showdomainusers.asp?" & getGRSN() & "&sortby=" & sortby & "&cdomain=" & cdomain & "&page=" & page & "&searchtext=" & searchtext
%>
