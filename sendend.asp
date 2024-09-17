<!--#include file="passinc.asp" -->

<%
newaddnum = trim(request("newaddnum"))
gourl = trim(request("gourl"))

if Request.ServerVariables("REQUEST_METHOD") <> "POST" or newaddnum = "" then
	if gourl <> "" then
		Response.Redirect gourl & "&" & getGRSN()
	else
		Response.Redirect "viewmailbox.asp?" & getGRSN()
	end if
end if

if IsNumeric(newaddnum) = false or CInt(newaddnum) < 1 then
	if gourl <> "" then
		Response.Redirect gourl & "&" & getGRSN()
	else
		Response.Redirect "viewmailbox.asp?" & getGRSN()
	end if
end if


allnum = CInt(newaddnum)

dim ads
set ads = server.createobject("easymail.Addresses")
ads.Load Session("wem")

i = 0
do while i < allnum
	nadd = trim(request("newadd" & i))
	if nadd <> "" then
		gan_name = ""
		gan_email = ""

		gan_s = InStr(nadd, "<")
		if gan_s > 0 then
			gan_e = InStr(gan_s, nadd, ">")

			if gan_e > 0 then
				gan_name = Mid(nadd, 1, gan_s - 1)
				gan_email = Mid(nadd, gan_s + 1, gan_e - gan_s - 1)
				ads.Simple_Add_Email gan_name, gan_email
			else
				ads.Simple_Add_Email nadd, nadd
			end if
		else
			ads.Simple_Add_Email nadd, nadd
		end if
	end if

    i = i + 1
loop

ads.Save

set ads = nothing

if gourl <> "" then
	Response.Redirect gourl & "&" & getGRSN()
else
	Response.Redirect "viewmailbox.asp?" & getGRSN()
end if
%>
