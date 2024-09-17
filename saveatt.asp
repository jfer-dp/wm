<!--#include file="passinc.asp" -->

<%
Set Obj=Server.CreateObject("EasyMail.EMSend")

if Obj.SaveAtt(Session("wem"), Session("tid")) = 0 then
	Set Obj=nothing
	Response.Redirect "addatt.asp?" & getGRSN()
else
	Set Obj=nothing
	Response.Redirect "addatt.asp?errcode=上传附件过大&" & getGRSN()
end if
%>
