<!--#include file="passinc.asp" -->

<%
	CName = trim(request("CName"))
	Email = trim(request("Email"))

	if CName = "" then
		CName = Email
	end if

	dim ads
	set ads = server.createobject("easymail.Addresses")
	ads.Load Session("wem")

	if ads.Simple_Add_Email(CName, Email) = false then
		set ads = nothing
		response.redirect "err.asp"
	end if

	ads.Save
	set ads = nothing

	response.redirect "ok.asp"
%>
