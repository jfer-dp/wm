<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim mam
set mam = server.createobject("easymail.AdminManager")
mam.Load

if trim(request("Enable_DAdminAllotSize")) = "" then
	mam.Enable_DAdminAllotSize = false
else
	mam.Enable_DAdminAllotSize = true
end if

mam.Save
set mam = nothing
'-----------------------------------------


dim ei

set ei = server.createobject("easymail.domain")
ei.Load
'-----------------------------------------

if IsNumeric(trim(request("maxuser"))) = true then
	mn = CLng(trim(request("maxuser")))
else
	mn = 30000
end if


if IsNumeric(trim(request("maxsize"))) = true then
	msize = CLng(trim(request("maxsize")))
else
	msize = 0
end if

if IsNumeric(trim(request("allsize"))) = true then
	asize = CLng(trim(request("allsize")))
else
	asize = 0
end if

if IsNumeric(trim(request("expire"))) = true then
	dexpire = CLng(trim(request("expire")))
else
	dexpire = 0
end if


if trim(request("checkshow")) = "" then
	ei.ModifyControlMsgEx trim(request("curdomain")), false, mn, trim(request("username")), msize, asize, dexpire
else
	ei.ModifyControlMsgEx trim(request("curdomain")), true, mn, trim(request("username")), msize, asize, dexpire
end if


ei.SaveControlMsg

set ei = nothing

if err.number = 0 then
	response.redirect "ok.asp?" & getGRSN() & "&gourl=m_showcondomain.asp?selectdomain=" & trim(request("curdomain"))
else
	response.redirect "err.asp?" & getGRSN() & "&gourl=m_showcondomain.asp?selectdomain=" & trim(request("curdomain"))
end if
%>
