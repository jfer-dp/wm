<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
fid = trim(request("fid"))
gourl = trim(request("gourl"))

if fid <> "" then
	dim poll
	set poll = server.createobject("easymail.Poll")
	poll.LoadOne fid
	poll.PI_Remove_Poll fid

	set poll = nothing

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if

Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
%>
