<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.logs")
'-----------------------------------------
ei.load


' 删除全部日志
if trim(request("mode")) = "removeall" then
	ei.RemoveAll()
else
	allnum = ei.LogCount
	i = 0

	do while i < allnum
		if trim(request("check" & i)) <> "" then
			ei.deleteLog trim(request("check" & i))				' 根据名称删除
		end if 
	
	    i = i + 1
	loop
end if


set ei = nothing

if err.number = 0 then
	response.redirect "ok.asp?" & getGRSN() & "&gourl=logs.asp"
else
	response.redirect "err.asp?" & getGRSN() & "&gourl=logs.asp"
end if
%>