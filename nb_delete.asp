<!--#include file="passinc.asp" -->

<%
dim nb
set nb = server.createobject("easymail.NoteBooksManager")

sortstr = request("sortstr")
sortmode = request("sortmode")
addsortstr = trim(Request("addsortstr"))
issort = false

if sortstr <> "" then
	if sortmode = "1" then
		sortmode = true

		nb.SetSort sortstr, sortmode
		issort = true
	elseif sortmode = "0" then
		sortmode = false

		nb.SetSort sortstr, sortmode
		issort = true
	end if
end if

nb.Load Session("wem")

mdel = trim(Request("mdel"))

dim isedit
isedit = false

if mdel = "1" then
	themax = nb.count - 1

	do while themax >= 0
		if trim(request("check" & themax)) <> "" then
			md = trim(request("check" & themax))

			if IsNumeric(md) = true then
				nb.DeleteByIndex CInt(md)
				isedit = true
			end if
		end if 

	    themax = themax - 1
	loop
else
	if trim(request("id")) <> "" then
		md = trim(request("id"))

		if IsNumeric(md) = true then
			nb.DeleteByIndex CInt(md)
			isedit = true
		end if
	end if 
end if

if isedit = true then
	nb.Save
end if

set nb = nothing
Response.Redirect "nb_brow.asp?" & getGRSN() & "&page=" & trim(request("thispage")) & addsortstr
%>
