<!--#include file="passinc.asp" -->

<%
delmode = trim(request("delmode"))

dim wemcert
set wemcert = server.createobject("easymail.WebEasyMailCert")
wemcert.Load Session("wem"), Session("mail")

dim retstr
retstr = trim(request("retstr"))

if delmode = "all" then
	isok = wemcert.RemoveAllCert()
elseif delmode = "allpub" then
	isok = wemcert.RemoveAllPubCert()
elseif delmode = "allsec" then
	isok = wemcert.RemoveAllSecCert()
elseif delmode = "pub" then
	isok = wemcert.RemovePubCert(trim(request("pub_email")))
end if

set wemcert = nothing

if retstr = "" then
	if isok = true then
		Response.Redirect "ok.asp?gourl=cert_index.asp&" & getGRSN()
	else
		Response.Redirect "err.asp?gourl=cert_index.asp&" & getGRSN()
	end if
else
	if isok = true then
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(retstr)
	else
		Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(retstr)
	end if
end if
%>
