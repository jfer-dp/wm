<!--#include file="passinc.asp" -->

<%
other_account = trim(Request("other_account"))
other_mail = trim(Request("other_mail"))

dim em
set em = server.createobject("easymail.emmail")

if trim(request("mode")) = "sec" then
	dim mam
	set mam = server.createobject("easymail.AdminManager")
	mam.Load

	if mam.Enable_User_Download_Private_Cert = true then
		if em.Export_Sec_Cert(Session("wem"), Session("mail")) = false then
			Response.Write "下载数字证书失败."
		end if
	end if

	set mam = nothing
else
	if other_account <> "" and other_mail <> "" then
		if em.Export_Pub_Cert(other_account, other_mail, other_mail) = false then
			Response.Write "下载公共数字证书失败."
		end if
	else
		if em.Export_Pub_Cert(Session("wem"), Session("mail"), trim(request("pub_email"))) = false then
			Response.Write "下载公共数字证书失败."
		end if
	end if
end if

set em = nothing

Response.End
%>
