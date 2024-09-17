<!--#include file="passinc.asp" -->

<%
index = trim(request("index"))

if index = "" then
	Response.Redirect "ff_showall.asp?" & getGRSN()
end if

if IsNumeric(index) = false then
	Response.Redirect "ff_showall.asp?" & getGRSN()
end if

dim userweb
set userweb = server.createobject("easymail.UserWeb")
userweb.Load Session("wem")

userweb.RemoveFriendFolderByIndex CInt(index)

userweb.Save

set userweb = nothing

Response.Redirect "ff_showall.asp?" & getGRSN()
%>
