<!--#include file="passinc.asp" -->

<%
gourl = trim(Request("gourl"))
other_account = trim(Request("other_account"))
other_mail = trim(Request("other_mail"))

if other_account = "" or other_mail = "" then
	Response.Redirect "err.asp?" & getGRSN() & "&gourl=cert_share.asp?page=" & page
end if

dim wemcert
set wemcert = server.createobject("easymail.WebEasyMailCert")
wemcert.Load Session("wem"), Session("mail")

isok = wemcert.Import_Other_Account_Pub_Cert(other_account, other_mail)

set wemcert = nothing


if isok = false then
	Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
else
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if
%>
