<%
sname = trim(request("sname"))
sfname = trim(request("sfname"))


dim a
set a = server.createobject("easymail.emmail")
'-----------------------------------------

if Session("wem") = "" then
	myaccount = trim(request.Cookies("name"))
else
	myaccount = Session("wem")
end if

if sname <> "" and sfname <> "" then
	if Session("wem") = "" then
		openresult = a.OpenFriendFolder(trim(request.Cookies("name")), sname, sfname, false)
	else
		openresult = a.OpenFriendFolder(Session("wem"), sname, sfname, false)
	end if

	if openresult = -1 then
		set a = nothing
		Response.Redirect "err.asp?errstr=ʧ��"
	elseif  openresult = 1 then
		set a = nothing
		Response.Redirect "err.asp?errstr=�������"
	elseif  openresult = 2 then
		set a = nothing
		Response.Redirect "err.asp?errstr=�ļ��в����ڻ��������"
	end if
end if

filename = trim(request("filename"))

Response.ContentType = "text/plain"
a.ShowEmail myaccount, filename

'-----------------------------------------
set a = nothing

Response.End
%>
