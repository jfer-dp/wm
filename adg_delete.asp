<!--#include file="passinc.asp" -->

<%
dim id
id = trim(request("id"))
gourl = trim(Request("gourl"))
mdel = trim(Request("mdel"))

dim ads
set ads = server.createobject("easymail.Addresses")
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
				ads.RemoveGroupByNickName item
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if
else
	ads.RemoveGroupByNickName id
end if

ads.Save
set ads = nothing

if gourl = "" then
	Response.Redirect "adg_brow.asp?" & getGRSN()
else
	Response.Redirect gourl & "&" & getGRSN()
end if
%>
