<!--#include file="passinc.asp" -->

<%
Set Obj=Server.CreateObject("EasyMail.EMSend")

dim is_f4
is_f4 = false

attfoldername = trim(request.Cookies("attfoldername"))

if attfoldername = "" then
	attfoldername = "att"
else
	if attfoldername = "zatt" then
		is_f4 = true
		attfoldername = "att"
	end if
end if

dim exname
exname = ""

if attfoldername = "att" then
	exname = "att"
else
	if attfoldername <> "" then
		dim pf
		set pf = server.createobject("easymail.PerAttFolders")
		pf.Load Session("wem")

		exname = pf.GetFolderID(attfoldername)

		set pf = nothing
	end if
end if

if is_f4 = false then
	if Obj.SaveFileEx(Session("wem"), Session("tid"), exname) = 0 then
		Set Obj=nothing
		exname = NULL
		Response.Redirect "ok.asp?gourl=listatt.asp?mb=" & Server.URLEncode(attfoldername) & "&" & getGRSN()
	else
		Set Obj=nothing
		exname = NULL
		Response.Redirect "err.asp?gourl=listatt.asp?mb=" & Server.URLEncode(attfoldername) & "&" & getGRSN()
	end if
else
	Obj.SaveNetZFile Session("wem"), Session("tid"), "zatt", ret_isok, ret_id, ret_size
	Set Obj=nothing

	if ret_isok = true then
		exname = "addatt.asp?zsize=" & Server.URLEncode(getShowSize(ret_size)) & "&zid=" & Server.URLEncode(ret_id) & "&" & getGRSN()
	else
		exname = "addatt.asp?" & getGRSN()
	end if

	ret_isok = NULL
	ret_id = NULL
	ret_size = NULL

	Response.Redirect exname
end if

function getShowSize(bytesize)
	if bytesize < 1000 then
		getShowSize = "1K"
	else
		if bytesize < 1000000 then
			getShowSize = CLng(bytesize/1000) & "K"
		else
			tmpSize = CStr(CDbl(bytesize/1000000))
			tmpindex = InStr(1, tmpSize, ".")
			if tmpindex = 0 then
				getShowSize = tmpSize & "M"
			else
				getShowSize = Left(tmpSize, tmpindex + 2) & "M"
			end if
		end if
	end if
end function
%>
