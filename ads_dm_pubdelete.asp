<!--#include file="passinc.asp" -->

<%
dim is_domain_admin
is_domain_admin = false

if isadmin() = false then
	dim dm
	set dm = server.createobject("easymail.Domain")
	dm.Load

	allnum = dm.GetUserManagerDomainCount(Session("wem"))

	curdomain = Mid(Session("mail"), InStr(Session("mail"), "@") + 1)

	i = 0
	do while i < allnum
		domain = dm.GetUserManagerDomain(Session("wem"), i)

		if LCase(curdomain) = LCase(domain) then
			is_domain_admin = true
		end if

		domain = NULL

		i = i + 1
	loop

	set dm = nothing
else
	is_domain_admin = true
end if


if is_domain_admin = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim id
id = trim(request("id"))
gourl = trim(Request("gourl"))
mdel = trim(Request("mdel"))
mode = trim(Request("mode"))

dim ads
set ads = server.createobject("easymail.DomainPubAddresses")
ads.Load Session("wem")

if mdel = "1" then
	dim msg
	msg = trim(Request("upinfo"))

	if Len(msg) > 0 then
		dim item
		dim ss
		dim se
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(9))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				ads.RemoveEmailByNickName item
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if
else
	ads.RemoveEmailByNickName id
end if

ads.Save
set ads = nothing

if gourl = "" then
	Response.Redirect "ads_dm_pubbrow.asp?" & getGRSN()
else
	Response.Redirect gourl & "&" & getGRSN()
end if
%>
