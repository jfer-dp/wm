<!--#include file="passinc.asp" -->

<%
if Session("ReadOnlyUser") <> 2 then
	Session("ReadOnlyUser") = 0

	dim rou
	set rou = server.createobject("easymail.ReadOnlyUsers")
	rou.Load
	if rou.isReadOnlyUser(Session("wem")) = true then
		Session("ReadOnlyUser") = 1
	else
		Session("ReadOnlyUser") = 2
	end if
	set rou = nothing
end if

dim need_change_pw
need_change_pw = false

if trim(request("grsn")) = "" then
	dim uwt
	set uwt = server.createobject("easymail.UserWorkTimer")
	uwt.Load_User Session("wem")

	if uwt.password_strong > 0 then
		need_change_pw = true
	end if

	set uwt = nothing
end if

if need_change_pw = true then
	Response.Redirect "newpw.asp?" & getGRSN()
end if

dim ei
set ei = server.createobject("easymail.UserWeb")
'-----------------------------------------

ei.Load Session("wem")

if ei.ChooseColorIndex > 0 and ei.ChooseColorIndex < 10 then
	dim mam
	set mam = server.createobject("easymail.AdminManager")
	mam.Load

	mam.GetSysColor ei.ChooseColorIndex, sc_name, sc_color
	showcstr = sc_color

	sc_name = NULL
	sc_color = NULL

	set mam = nothing
else
	showcstr = ei.OwnColor
end if

if Len(showcstr) <> 66 then
	showcstr = csbi_color_str_default
end if

Session("my_Show_Color") = showcstr


Session("pl") = ei.pageLines
Session("addomail") = ei.orMailForReply
Session("delProc") = ei.delProc

if ei.addInSubjectForReply = 0 then
	Session("addsubject") = "> "
elseif ei.addInSubjectForReply = 1 then
	Session("addsubject") = "Re: "
elseif ei.addInSubjectForReply = 2 then
	Session("addsubject") = "»Ø¸´: "
end if

set ei = nothing

dim si
set si = server.createobject("easymail.sysinfo")
si.Load
defaultMailsNumber = si.defaultMailsNumber
Application("em_EnableVerification") = si.EnableVerification

Response.Cookies("cookie_ZATT_Is_Enable") = CStr(si.ZATT_Is_Enable)
Response.Cookies("cookie_ZATT_Is_Enable").Expires = DateAdd("d", 2, Now())

if si.ZATT_URL = "http://localhost/downatt.asp" then
	bf_url = Request.ServerVariables("HTTP_REFERER")
	bf_index = InStr(bf_url, "?")

	if bf_index > 0 then
		bf_url = Left(bf_url, bf_index)
	end if

	bf_index = InStrRev(bf_url, "/")
	if bf_index > 0 then
		bf_url = Left(bf_url, bf_index)
		if Len(bf_url) > 7 then
			si.ZATT_URL = bf_url & "downatt.asp"
			si.Save
		end if
	end if
end if

set si = nothing
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>WinWebMail - <%=Session("mail") %></TITLE>
</HEAD>

<frameset cols="150,*" frameborder="NO" border="0" framespacing="0" rows="*"> 
  <frame id=f1 name="f1" scrolling="AUTO" noresize src="left.asp?<%=getGRSN() %>" marginWidth=5 marginHeight=20>
  <frame id=f2 name="f2" src="viewmailbox.asp?rla=1&dmn=<%=defaultMailsNumber %>&noticemsg=<%=trim(request("noticemsg")) %>&<%=getGRSN() %>">
</frameset>
</HTML>
