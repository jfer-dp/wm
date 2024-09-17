<!--#include file="passinc.asp" -->

<%
dim wemcert
set wemcert = server.createobject("easymail.WebEasyMailCert")
wemcert.Load Session("wem"), Session("mail")

themax = wemcert.PubCertCount

do while themax >= 0
	if trim(request("check" & themax)) <> "" then
		md = trim(request("check" & themax))

		if md <> "" then
			wemcert.RemovePubCert md
		end if
	end if 

    themax = themax - 1
loop

set wemcert = nothing

Response.Redirect "cert_myothpub.asp?" & getGRSN() & "&page=" & trim(request("page"))
%>
