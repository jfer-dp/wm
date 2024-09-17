<%
if Session("wem") = "" then
	Response.Redirect "default.asp"
end if
%>

<%
dim emm
set emm = server.createobject("easymail.emmail")

emm.Export_CSV Session("wem")

set emm = nothing

Response.End
%>
