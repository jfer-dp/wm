<!--#include file="passinc.asp" --> 

<%
iniid = trim(request("iniid"))
min_show_index = trim(request("min_show_index"))
max_show_index = trim(request("max_show_index"))

if isadmin() = false then
	dim pfvl
	set pfvl = server.createobject("easymail.PubFolderViewLimit")
	pfvl.Load iniid

	if pfvl.IsShow(Session("mail")) = false then
		set pfvl = nothing
		Response.Redirect "noadmin.asp"
	end if

	set pfvl = nothing
end if

if iniid <> "" and min_show_index <> "" and IsNumeric(min_show_index) = true and max_show_index <> "" and IsNumeric(max_show_index) = true and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	i = CInt(min_show_index)
	max_show_index = CInt(max_show_index)
	isok = false

	dim pf
	set pf = server.createobject("easymail.PubFolderManager")
	pf.load iniid

	if isadmin() = true or LCase(pf.admin) = LCase(Session("wem")) then
		do while i <= max_show_index
			if trim(request("check" & i)) <> "" then
				pf.Add_Remove_PID(trim(request("check" & i)))
			end if

		    i = i + 1
		loop

		isok = pf.Remove_PID_List()
	end if

	set pf = nothing
end if


if isok = true then
	if trim(request("gourl")) = "" then
		Response.Redirect "ok.asp"
	else
		Response.Redirect "ok.asp?gourl=" & Server.URLEncode(trim(request("gourl")))
	end if
else
	if trim(request("gourl")) = "" then
		Response.Redirect "err.asp"
	else
		Response.Redirect "err.asp?gourl=" & Server.URLEncode(trim(request("gourl")))
	end if
end if
%>
