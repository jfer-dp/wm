<!--#include file="passinc.asp" --> 

<%
dim em
set em = Application("em")

if em.Kick(Session("wem"), trim(request("ip"))) = true then
	set em = nothing
	Response.Redirect "ok.asp?gourl=" & Server.URLEncode("viewmailbox.asp")
else
	set em = nothing
	Response.Redirect "err.asp?gourl=" & Server.URLEncode("viewmailbox.asp")
end if
%>
