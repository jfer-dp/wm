<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
domainname = trim(request("domainname"))
username = trim(request("username"))

if Request.ServerVariables("REQUEST_METHOD") = "POST" and domainname <> "" then
	dim ei
	set ei = server.createobject("easymail.domain")
	ei.DM_Load

	ei.DM_ModifyUser domainname, username

	ei.DM_Save

	set ei = nothing
end if

Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("show_dm_domain.asp?selectdomain=" & domainname)
%>
