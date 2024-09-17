<!--#include file="passinc.asp" --> 

<%
mode = trim(request("mode"))

attnum = trim(request("attnum"))

if IsNumeric(attnum) = false then
	Response.Redirect "err.asp?errstr=失败"
end if


dim ei
set ei = server.createobject("easymail.emmail")

if mode = "post" then
	ei.IsInPublicFolder = true
	ei.PublicFolderName = trim(request("iniid"))
end if

sname = trim(request("sname"))
sfname = trim(request("sfname"))

if sname <> "" and sfname <> "" then
	openresult = ei.OpenFriendFolder(Session("wem"), sname, sfname, false)

	if openresult = -1 then
		set ei = nothing
		Response.Redirect "err.asp?errstr=失败"
	elseif  openresult = 1 then
		set ei = nothing
		Response.Redirect "err.asp?errstr=密码错误"
	elseif  openresult = 2 then
		set ei = nothing
		Response.Redirect "err.asp?errstr=文件夹不存在或不允许访问"
	end if
end if

'-----------------------------------------
filename = trim(request("filename"))

pt = trim(request("pt"))

if pt <> "" then
	bd = trim(request("bd"))

	if bd <> "" then
		ei.LoadAll2 Session("wem"), filename, CDbl(pt), bd
	else
		ei.LoadAll1 Session("wem"), filename, CDbl(pt)
	end if
else
	ei.LoadAll Session("wem"), filename
end if


isok = ei.SaveToAttFile(CInt(attnum), Session("wem"))

nextisok = true

if isok = false then
	nextisok = ei.SaveToAttFileWithName(CInt(attnum), Session("wem"), "Empty")
end if

set ei = nothing

if isok = true then
	Response.Redirect "ok.asp?" & getGRSN()
else
	if nextisok = true then
		Response.Redirect "ok.asp?errstr=原操作失败, 该附件被强制保存为Empty&" & getGRSN()
	else
		Response.Redirect "err.asp?" & getGRSN()
	end if
end if
%>
