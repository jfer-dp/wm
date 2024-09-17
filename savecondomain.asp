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

isok = false
i = 0
allnum = ei.getcount

ei.RemoveAllControlMsg

do while i < allnum + 1
	if IsNumeric(trim(request("maxuser" & i))) = true then
		mn = CLng(trim(request("maxuser" & i)))
	else
		mn = 30000
	end if


	if IsNumeric(trim(request("maxsize" & i))) = true then
		msize = CLng(trim(request("maxsize" & i)))
	else
		msize = 0
	end if

	if IsNumeric(trim(request("allsize" & i))) = true then
		asize = CLng(trim(request("allsize" & i)))
	else
		asize = 0
	end if

	if IsNumeric(trim(request("expire" & i))) = true then
		dexpire = CLng(trim(request("expire" & i)))
	else
		dexpire = 0
	end if

	if trim(request("check" & i)) = "" then
		ei.AddControlMsgEx trim(request("domain" & i)), false, mn, trim(request("username" & i)), msize, asize, dexpire
	else
		ei.AddControlMsgEx trim(request("domain" & i)), true, mn, trim(request("username" & i)), msize, asize, dexpire
	end if

    i = i + 1


	if i = allnum then
		isok = true
	end if
loop


if isok = true then
	ei.SaveControlMsg
end if

set ei = nothing

if trim(request("gourl")) = "s_showcondomain.asp" then
	if err.number = 0 then
		response.redirect "ok.asp?" & getGRSN() & "&gourl=s_showcondomain.asp"
	else
		response.redirect "err.asp?" & getGRSN() & "&gourl=s_showcondomain.asp"
	end if
else
	if err.number = 0 then
		response.redirect "ok.asp?" & getGRSN() & "&gourl=showcondomain.asp"
	else
		response.redirect "err.asp?" & getGRSN() & "&gourl=showcondomain.asp"
	end if
end if
%>
