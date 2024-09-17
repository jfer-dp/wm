<!--#include file="passinc.asp" -->

<%
Set Obj=Server.CreateObject("EasyMail.EMSend")

if Obj.Import_CSV(Session("wem"), Session("tid")) = true then
	Set Obj=nothing
	Response.Redirect "ok.asp?gourl=ads_brow.asp&" & getGRSN()
else
	Set Obj=nothing
	Response.Redirect "err.asp?gourl=ads_brow.asp&" & getGRSN()
end if
%>
