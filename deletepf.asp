<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
if trim(request("fileid")) <> "" then
	dim pf
	set pf = server.createobject("easymail.PubFolderManager")

	pf.RemovePubFolder trim(request("fileid"))
	set pf = nothing

	Response.Redirect "showallpf.asp?" & getGRSN()
end if
%>
