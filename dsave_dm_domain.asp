<!--#include file="passinc.asp" -->

<%
dim esi
set esi = server.createobject("easymail.sysinfo")
esi.Load

if esi.enableDomainMonitor = false then
	set esi = nothing
	response.redirect "noadmin.asp"
end if

set esi = nothing


dim ei
set ei = server.createobject("easymail.Domain")
ei.Load

if ei.GetUserManagerDomainCount(Session("wem")) < 1 then
	set ei = nothing
	response.redirect "noadmin.asp"
end if
%>

<%
domainname = trim(request("domainname"))
seluser = trim(request("seluser"))

if Request.ServerVariables("REQUEST_METHOD") = "POST" and domainname <> "" then
	allnum = ei.GetUserManagerDomainCount(Session("wem"))
	isok = false
	i = 0

	do while i < allnum
		if domainname = ei.GetUserManagerDomain(Session("wem"), i) then
			isok = true
            exit do
		end if

		i = i + 1
	loop

	if isok = true then
		ei.DM_ModifyUser domainname, seluser
		ei.DM_Save
	end if 
end if

set ei = nothing

Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("dshow_dm_domain.asp?selectdomain=" & domainname)
%>
