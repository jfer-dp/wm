<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
if Request.ServerVariables("REQUEST_METHOD") = "POST" and trim(request("curdomain")) <> "" then
	dim ei
	set ei = server.createobject("easymail.DomainDefaultMailBoxSize")

	ei.Load

	if trim(request("cleanall")) = "yes" then
		ei.RemoveAll
	else
		if trim(request("changemyselect")) <> "" then
			ei.Remove trim(request("curdomain"))
		else
			if IsNumeric(trim(request("ksize"))) = true then
				ei.Set trim(request("curdomain")), CLng(trim(request("ksize")))
			end if
		end if
	end if

	ei.Save

	set ei = nothing
end if

response.redirect "ok.asp?" & getGRSN() & "&gourl=showddsize.asp?selectdomain=" & trim(request("curdomain"))
%>
