<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	Response.Redirect "noadmin.asp"
end if

i = trim(request("count"))
filename = trim(request("filename"))
user = trim(request("user"))
inout = trim(request("inout"))
pt = trim(request("pt"))

if pt = "" then
	pt = "0"
end if

dim ei
set ei = server.createobject("easymail.emmail")

if inout = "in" then
	ei.LoadAll_MonInMail user, filename, CDbl(pt), ""
else
	ei.LoadAll_MonOutMail user, filename, CDbl(pt), ""
end if

if trim(request("ishtml")) <> "1" then
	Response.ContentType = ei.GetContentType(cint(i))
else
	Response.Charset = ei.GetCharSet(cint(i))
end if

if trim(request("isdown")) = "1" then
	ei.ShowAttachment cint(i), true
else
	ei.ShowAttachment cint(i), false
end if

set ei = nothing
Response.End
%>
