<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim ei

	set ei = server.createobject("easymail.domain")
	'-----------------------------------------
	ei.DCA_Load

	i = 0
	allnum = ei.getcount

	do while i < allnum + 1
		if trim(request("idomain" & i)) <> "" then
			ei.DCA_ModifyUser trim(request("idomain" & i)), trim(request("user" & i))
		end if 

	    i = i + 1
	loop

	ei.DCA_Save

	set ei = nothing
end if

response.redirect "ok.asp?" & getGRSN() & "&gourl=show_dca_domain.asp"
%>
