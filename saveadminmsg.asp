<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.adminmsg")
'-----------------------------------------
ei.Load

ei.errback_subject = trim(request("err_subject"))
ei.errback_text = trim(request("err_text"))
ei.welcome_subject = trim(request("welcome_subject"))
ei.welcome_text = trim(request("welcome_text"))

ei.ReadBack_Subject = trim(request("ReadBack_Subject"))
ei.ReadBack_Text = trim(request("ReadBack_Text"))
ei.Fill_Subject = trim(request("Fill_Subject"))
ei.Fill_Text = trim(request("Fill_Text"))

ei.Virus_Subject = trim(request("Virus_Subject"))
ei.Virus_Text = trim(request("Virus_Text"))

ei.NoSpam_Affirm_Subject = trim(request("NoSpam_Affirm_Subject"))
ei.NoSpam_Affirm_Text = trim(request("NoSpam_Affirm_Text"))

ei.UserExp_Affirm_Subject = trim(request("UserExp_Affirm_Subject"))
ei.UserExp_Affirm_Text = trim(request("UserExp_Affirm_Text"))

ei.TrashMsg_Subject = trim(request("TrashMsg_Subject"))
ei.TrashMsg_Text = trim(request("TrashMsg_Text"))

ei.Recall_Subject = trim(request("Recall_Subject"))
ei.Recall_Text = trim(request("Recall_Text"))

ei.Save

set ei = nothing

response.redirect "ok.asp?" & getGRSN() & "&gourl=showadminmsg.asp"
%>
