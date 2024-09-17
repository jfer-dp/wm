<!--#include file="passinc.asp" --> 

<%
sname = trim(request("sname"))
sfname = trim(request("sfname"))
filename = trim(request("filename"))

if sname = "" or sfname = "" or filename = "" or Request.ServerVariables("REQUEST_METHOD") <> "POST" then
	Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl")))
end if


dim ei
set ei = server.createobject("easymail.emmail")

if sfname = "att" then
	openresult = ei.OpenFriendFolder(Session("wem"), sname, sfname, false)
else
	openresult = ei.OpenFriendFolder(Session("wem"), sname, sfname, true)
end if

if openresult = -1 then
	set ei = nothing
	Response.Redirect "err.asp?errstr=失败&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl")))
elseif  openresult = 1 then
	set ei = nothing
	Response.Redirect "err.asp?errstr=密码错误&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl")))
elseif  openresult = 2 then
	set ei = nothing
	Response.Redirect "err.asp?errstr=文件夹不存在或不允许访问&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl")))
end if


'-----------------------------------------
isok = ei.SaveFriendAttFile2Att(filename)

set ei = nothing

if isok = true then
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl")))
else
	Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl")))
end if
%>
