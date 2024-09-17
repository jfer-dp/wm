<%
Response.CacheControl = "no-cache" 
%>

<!--#include file="passinc.asp" -->

<%
if trim(Session("cert_imp_pw")) = "" or trim(Session("cert_imp_type")) = "" then
	Session("cert_imp_type") = ""
	Session("cert_imp_pw") = ""
	Session("cert_imp_save_day") = ""

	Response.Redirect "err.asp?" & getGRSN() & "&errstr=" & Server.URLEncode("�������") & "&gourl=cert_index.asp"
end if


save_day = -1

if trim(Session("cert_imp_save_day")) <> "" and IsNumeric(trim(Session("cert_imp_save_day"))) = true then
	save_day = CInt(trim(Session("cert_imp_save_day")))
end if


dim ispub
ispub = false

dim em
set em = server.createobject("easymail.EMsend")
'-----------------------------------------

if trim(Session("cert_imp_type")) = "sec" then
	isok = em.Import_Sec_Cert(Session("wem"), Session("tid"), Session("mail"), trim(Session("cert_imp_pw")), save_day)
elseif trim(Session("cert_imp_type")) = "pub" then
	ispub = true
	isok = em.Import_Pub_Cert(Session("wem"), Session("tid"), Session("mail"), trim(Session("cert_imp_pw")))
end if

set em = nothing


Session("cert_imp_type") = ""
Session("cert_imp_pw") = ""
Session("cert_imp_save_day") = ""

if ispub = false then
	if isok <> 0 then
		Response.Redirect "err.asp?" & getGRSN() & "&errstr=" & Server.URLEncode("����֤�鵼��ʧ��: " & geterror(isok)) & "&gourl=cert_index.asp"
	else
		Response.Redirect "ok.asp?" & getGRSN() & "&errstr=" & Server.URLEncode("����֤�鵼��ɹ�") & "&gourl=cert_index.asp"
	end if
else
	if isok <> 0 then
		Response.Redirect "err.asp?" & getGRSN() & "&errstr=" & Server.URLEncode("����֤�鵼��ʧ��: " & geterror(isok)) & "&gourl=cert_myothpub.asp"
	else
		Response.Redirect "ok.asp?" & getGRSN() & "&errstr=" & Server.URLEncode("����֤�鵼��ɹ�") & "&gourl=cert_myothpub.asp"
	end if
end if


function geterror(ecode)
	if ecode = 1 then
		geterror = "����֪�Ĵ���"
	elseif ecode = 2 then
		geterror = "�������"
	elseif ecode = 3 then
		geterror = "�ļ��򿪴���"
	elseif ecode = 4 then
		geterror = "�ļ��Ҳ���"
	elseif ecode = 5 then
		geterror = "�������"
	elseif ecode = 6 then
		geterror = "�����֤�����ʹ���"
	elseif ecode = 7 then
		geterror = "�����˽��֤��͵�ǰ�ʺŲ���"
	elseif ecode = 8 then
		geterror = "�Ѵ�����ͬ�ʼ���ַ��֤��"
	elseif ecode = 9 then
		geterror = "һ�ε������֤��"
	elseif ecode = 10 then
		geterror = "����Ĺ���֤���ʼ���ַ���������ʼ���ַ����"
	end if
end function
%>
