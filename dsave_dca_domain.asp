<!--#include file="passinc.asp" -->

<%
dim esinfo
set esinfo = server.createobject("easymail.sysinfo")
esinfo.Load

if esinfo.enableCatchAll = false then
	set esinfo = nothing
	response.redirect "noadmin.asp"
end if

set esinfo = nothing



dim ei
set ei = server.createobject("easymail.Domain")
ei.Load

if ei.GetUserManagerDomainCount(Session("wem")) < 1 then
	set ei = nothing
	response.redirect "noadmin.asp"
end if
%>


<%
if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	allnum = ei.GetUserManagerDomainCount(Session("wem"))
	isok = false
	i = 0

	do while i < allnum
		if trim(request("domainname")) = ei.GetUserManagerDomain(Session("wem"), i) then
			isok = true
            exit do
		end if

		i = i + 1
	loop

	if isok = true then
		ei.DCA_ModifyUser trim(request("domainname")), trim(request("user"))
		ei.DCA_Save
	end if 
end if

set ei = nothing

response.redirect "ok.asp?" & getGRSN() & "&gourl=dshow_dca_domain.asp?selectdomain=" & trim(request("domainname"))
%>
