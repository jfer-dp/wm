<!--#include file="passinc.asp" --> 

<%
	id = trim(request("id"))

	if id <> "" then
		dim am
		set am = server.createobject("easymail.Attachments")
		am.Load Session("wem"), Session("tid")

		am.DeleteAtt id

		set am = nothing
	end if

	response.redirect "addatt.asp?" & getGRSN()
%>
