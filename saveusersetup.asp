<!--#include file="passinc.asp" -->

<%
dim ei
set ei = server.createobject("easymail.UserMessages")
'-----------------------------------------
ei.Load Session("wem")


if trim(request("checkauto")) <> "" then
	ei.UseAutoReply = true
else
	ei.UseAutoReply = false
end if

if trim(request("checkautoforward")) <> "" then
	ei.UseAutoForward = true
else
	ei.UseAutoForward = false
end if

if trim(request("checklocalsave")) <> "" then
	ei.LocalSave = true
else
	ei.LocalSave = false
end if

ei.AutoForwardTo = trim(request("AutoForwardTo"))


ei.AutoReplySubject = trim(request("subject"))
ei.AutoReplyText = trim(request("text"))

if trim(request("checkauto")) <> "" then
	sy = CLng(trim(request("sy")))
	sm = CLng(trim(request("sm")))
	sd = CLng(trim(request("sd")))
	ey = CLng(trim(request("ey")))
	em = CLng(trim(request("em")))
	ed = CLng(trim(request("ed")))
else
	sy = 0
	sm = 0
	sd = 0
	ey = 0
	em = 0
	ed = 0
end if

ei.SetAutoReplyDateLimit sy, sm, sd, ey, em, ed

ei.Save

set ei = nothing

Response.Redirect "ok.asp?" & getGRSN() & "&gourl=showusersetup.asp"
%>
