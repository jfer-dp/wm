<!--#include file="passinc.asp" --> 

<%
efname = trim(request("efname"))
emsg = trim(request("emsg"))

emsg = replace(emsg, Chr(13), "")
emsg = replace(emsg, """", "")
emsg = replace(emsg, "'", "")
emsg = replace(emsg, "\", "")
emsg = replace(emsg, "<", "")
emsg = replace(emsg, ">", "")

if efname <> "" then
	dim am
	set am = server.createobject("easymail.Attachments")

	isok = am.SetAttFileComment(Session("wem"), efname, emsg)

	set am = nothing
end if

if isok = true then
	response.redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl")))
else
	response.redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl")))
end if
%>
