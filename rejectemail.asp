<!--#include file="passinc.asp" -->

<%
dim ei
set ei = server.createobject("easymail.usermessages")
ei.Load Session("wem")

ei.AddRejectEMail trim(request("kill"))

ei.SaveReject

set ei = nothing

Response.Redirect "ok.asp?" & getGRSN()
%>
