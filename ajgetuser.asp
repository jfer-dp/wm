<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	Response.Redirect "noadmin.asp"
end if
%>

<%
s_domain = LCase(trim(request("domain")))
Response.Write "<select class=""drpdwn"" style=""WIDTH:300px"" multiple size=22 id=""sysuserlist"" name=""sysuserlist"" width=""300"">"

if Len(s_domain) < 1 then
	Response.Write "</select>"
	Response.end
end if

dim ei
set ei = Application("em")

i = 0
allnum = ei.GetUsersCount

do while i < allnum
	ei.GetUserByIndex i, name, domain, comment

	if s_domain = "all" then
		Response.Write "<option value='" & server.htmlencode(name) & "'>" & server.htmlencode(name) & "</option>" & Chr(13)
	else
		if LCase(domain) = s_domain then
			Response.Write "<option value='" & server.htmlencode(name) & "'>" & server.htmlencode(name) & "</option>" & Chr(13)
		end if
	end if

	name = NULL
	domain = NULL
	comment = NULL

	i = i + 1
loop

Response.Write "</select>"
set ei = nothing
%>
