<%
Response.CacheControl = "no-cache" 
%>

<!--#include file="passinc.asp" -->

<%
if trim(Session("cert_imp_pw")) = "" or trim(Session("cert_imp_type")) = "" then
	Session("cert_imp_type") = ""
	Session("cert_imp_pw") = ""
	Session("cert_imp_save_day") = ""

	Response.Redirect "err.asp?" & getGRSN() & "&errstr=" & Server.URLEncode("输入错误") & "&gourl=cert_index.asp"
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
		Response.Redirect "err.asp?" & getGRSN() & "&errstr=" & Server.URLEncode("数字证书导入失败: " & geterror(isok)) & "&gourl=cert_index.asp"
	else
		Response.Redirect "ok.asp?" & getGRSN() & "&errstr=" & Server.URLEncode("数字证书导入成功") & "&gourl=cert_index.asp"
	end if
else
	if isok <> 0 then
		Response.Redirect "err.asp?" & getGRSN() & "&errstr=" & Server.URLEncode("数字证书导入失败: " & geterror(isok)) & "&gourl=cert_myothpub.asp"
	else
		Response.Redirect "ok.asp?" & getGRSN() & "&errstr=" & Server.URLEncode("数字证书导入成功") & "&gourl=cert_myothpub.asp"
	end if
end if


function geterror(ecode)
	if ecode = 1 then
		geterror = "不可知的错误"
	elseif ecode = 2 then
		geterror = "输入错误"
	elseif ecode = 3 then
		geterror = "文件打开错误"
	elseif ecode = 4 then
		geterror = "文件找不到"
	elseif ecode = 5 then
		geterror = "密码错误"
	elseif ecode = 6 then
		geterror = "导入的证书类型错误"
	elseif ecode = 7 then
		geterror = "导入的私人证书和当前帐号不符"
	elseif ecode = 8 then
		geterror = "已存在相同邮件地址的证书"
	elseif ecode = 9 then
		geterror = "一次导入过多证书"
	elseif ecode = 10 then
		geterror = "导入的公共证书邮件地址和所输入邮件地址不符"
	end if
end function
%>
