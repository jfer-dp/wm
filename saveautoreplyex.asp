<!--#include file="passinc.asp" -->

<%
if Request.ServerVariables("REQUEST_METHOD") = "POST" and trim(request("arindex")) <> "" and IsNumeric(trim(request("arindex"))) = true then
	dim arex
	set arex = server.createobject("easymail.AutoReplyEx")
	arex.Load Session("wem")

	arex.Get CInt(trim(request("arindex"))), are_name, are_subject, are_text
	retindex = CInt(trim(request("arindex")))

	delok = true

	if trim(request("isdel")) = "yes" then
		if arex.Remove(are_name) = false then
			delok = false
		else
			retindex = -1
		end if
	else
		if CInt(trim(request("arindex"))) = -1 then
			arex.Set trim(request("mare_name")), trim(request("mare_subject")), trim(request("mare_text"))
		else
			if trim(request("mare_subject")) = "" and trim(request("mare_text")) = "" then
				if arex.Remove(are_name) = false then
					delok = false
				else
					retindex = -1
				end if
			else
				arex.Set are_name, trim(request("mare_subject")), trim(request("mare_text"))
			end if
		end if
	end if

	are_name = NULL
	are_subject = NULL
	are_text = NULL

	arex.Save

	set arex = nothing
end if

if delok = true then
	response.redirect "ok.asp?" & getGRSN() & "&gourl=showautoreplyex.asp?selectar=" & CStr(retindex) & "&returl=" & trim(request("returl"))
else
	response.redirect "err.asp?" & getGRSN() & "&gourl=showautoreplyex.asp?selectar=" & CStr(retindex) & "&returl=" & trim(request("returl")) & "&errstr=删除失败! 此项自动回复内容可能正被邮件过滤项所使用"
end if
%>
