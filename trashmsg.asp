<%
if IsEmpty(Application("em_MaxMPOP3")) and IsEmpty(Application("em_MaxSigns")) then
	dim mam
	set mam = server.createobject("easymail.AdminManager")

	tmp_num = 0
	do while tmp_num < 30
		mam.LoadExt

		if mam.IsLoadOK = true then
			Exit Do
		end if

		mam.Sleep 500
		tmp_num = tmp_num + 1
	loop

	if mam.IsLoadOK = true then
		Application("em_MaxMPOP3") = mam.MaxMPOP3
		Application("em_MaxSigns") = mam.MaxSigns
		Application("em_SystemAdmin") = LCase(mam.SystemAdmin)
		Application("em_EnableBBS") = mam.EnableBBS
		Application("em_Enable_SignHold") = mam.Enable_SignHold
		Application("em_Enable_FreeSign") = mam.Enable_FreeSign
		Application("em_Enable_SignWithDomainUser") = mam.Enable_SignWithDomainUser
		Application("em_Enable_SignNumberLimit") = mam.Enable_SignNumberLimit
		Application("em_SignNumberLimitDays") = mam.SignNumberLimitDays
		Application("em_Enable_ShareFolder") = mam.Enable_ShareFolder
		Application("em_Enable_SignEnglishName") = mam.Enable_SignEnglishName
		Application("em_LogPageKSize") = mam.LogPageKSize
		Application("em_TestAccounts") = LCase(mam.TestAccounts)
		Application("em_SignMode") = mam.SignMode
		Application("em_SignWaitDays") = mam.SignWaitDays
		Application("em_am_Name") = mam.am_Name
		Application("em_am_Accounts") = LCase(mam.am_Accounts)
		Application("em_AccountsAdmin") = LCase(mam.AccountsAdmin)
		Application("em_EnableEntAddress") = mam.Enable_Show_EntAddress
		Application("em_SpamAdmin") = LCase(mam.SpamAdmin)

		Application("em_EnableTrap") = mam.EnableTrap
		if mam.EnableTrap = true then
			Application("em_TrapMail") = mam.TrapMail
		end if

		set mam = nothing
	else
		set mam = nothing
		Response.Redirect "default.asp?errstr=" & Server.URLEncode("超时, 请重试") & "&" & getGRSN()
	end if
end if


dim isok
isok = false

mode = trim(request("mode"))
user = trim(request("user"))
allnum = trim(request("allnum"))
fid = trim(request("fid"))

if IsNumeric(allnum) = true then
	allnum = CLng(allnum)
else
	allnum = 0
end if

delnum = -1
dim tm

if user <> "" and fid <> "" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	set tm = server.createobject("easymail.TrashMsg")
	tm.Load user
	isok = tm.MoveToInbox(fid)
	set tm = nothing
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<BODY>
<br><br>
<table width="100%"><tr><td width="30"></td><td>
<table border=0 cellspacing=0 cellpadding=0 width="90%">
  <tr bgcolor="#e8e8e8"> 
<%
if isok = true then
%>
	<td height="60" style='border:1px #8B969D solid;'><font class="Wf" color="#104A7B">&nbsp;&nbsp;<b>您已成功挽回该邮件</b>。<br><br>&nbsp;&nbsp;您可以登录邮箱在收件箱中查看，或通过客户端再次收取该邮件。</font></td>
<%
else
%>
	<td height="30" style='border:1px #8B969D solid;'>&nbsp;<img src="images/error.gif" align="absmiddle" border="0"><font class="Wf" color="#104A7B">&nbsp;挽回邮件处理<b>失败</b>。</font></td>
<%
end if
%>
  </tr>
<br>
<tr><td><br><hr size="1" color="#8B969D"></td></tr>
<tr><td align="right"><br>
<input type="button" value=" 关闭 " LANGUAGE=javascript onclick="javascript:self.close();" class="Bsbttn">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</td></tr></table>
</td></tr></table>
</body>
</html>
<%
else

rd = trim(request("rd"))
if mode = "empty" and user <> "" and rd <> "" and fid = "" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	set tm = server.createobject("easymail.TrashMsg")
	isok = tm.CheckRand(user, rd)

	if isok = true then
		dim mil
		set mil = server.createobject("easymail.InfoList")
		mil.LoadMailBox user, "del"

		allnum = mil.getMailsCount
		i = 0

		do while i < allnum
			mil.getMailInfo allnum - i - 1, idname, isread, priority, sendMail, sendName, subject, size, etime
			mil.DelMail user, idname

			idname = NULL
			isread = NULL
			priority = NULL
			sendMail = NULL
			sendName = NULL
			subject = NULL
			size = NULL
			etime = NULL

			i = i + 1
		loop

		set mil = nothing
		tm.DelRand user
	end if

	set tm = nothing
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<BODY>
<br><br>
<table width="100%"><tr><td width="30"></td><td>
<table border=0 cellspacing=0 cellpadding=0 width="90%">
  <tr bgcolor="#e8e8e8"> 
<%
if isok = true then
%>
	<td height="30" style='border:1px #8B969D solid;'><font class="Wf" color="#104A7B">&nbsp;&nbsp;<b>已清空垃圾箱中邮件</b>。</td>
<%
else
%>
	<td height="30" style='border:1px #8B969D solid;'>&nbsp;<img src="images/error.gif" align="absmiddle" border="0"><font class="Wf" color="#104A7B">&nbsp;清空垃圾箱处理<b>失败</b>。</font></td>
<%
end if
%>
  </tr>
<br>
<tr><td><br><hr size="1" color="#8B969D"></td></tr>
<tr><td align="right"><br>
<input type="button" value=" 关闭 " LANGUAGE=javascript onclick="javascript:self.close();" class="Bsbttn">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</td></tr></table>
</td></tr></table>
</body>
</html>
<%
else

if mode <> "" and user <> "" and allnum > 0 and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	set tm = server.createobject("easymail.TrashMsg")
	tm.Load user

	i = 0
	if mode = "0" then
		delnum = 0
		do while i <= allnum
			if trim(request("ck" & i)) <> "" then
				oneisok = tm.DelTrashMsg(trim(request("ck" & i)))
				delnum = delnum + 1

				if isok = false and oneisok = true then
					isok = true
				end if
			end if

		    i = i + 1
		loop

		if isok = true and delnum < allnum then
			tm.MoveMulToInbox(trim(request("maxmid")))
		end if
	elseif mode = "1" then
		do while i <= allnum
			if trim(request("ck" & i)) <> "" then
				oneisok = tm.DelTrashMsg(trim(request("ck" & i)))

				if isok = false and oneisok = true then
					isok = true
				end if
			end if

		    i = i + 1
		loop
	elseif mode = "2" then
		do while i <= allnum
			if trim(request("ck" & i)) <> "" then
				oneisok = tm.MoveToInbox(trim(request("ck" & i)))

				if isok = false and oneisok = true then
					isok = true
				end if
			end if

		    i = i + 1
		loop
	end if

	set tm = nothing
end if
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<BODY>
<br><br>
<table width="100%"><tr><td width="30"></td><td>
<table border=0 cellspacing=0 cellpadding=0 width="90%">
  <tr bgcolor="#e8e8e8"> 
<%
if isok = false then
	if mode <> "" then
		if mode = "1" or mode = "2" or (mode = "0" and delnum > 0) then
%>
	<td height="30" style='border:1px #8B969D solid;'><font class="Wf" color="#104A7B">&nbsp;&nbsp;已完成对您垃圾箱内邮件的处理。</font></td>
<%
		else
%>
	<td height="30" style='border:1px #8B969D solid;'>&nbsp;<img src="images/error.gif" align="absmiddle" border="0"><font class="Wf" color="#104A7B">&nbsp;垃圾箱邮件处理<b>失败</b>。</font></td>
<%
		end if
	else
%>
	<td height="30" style='border:1px #8B969D solid;'><font class="Wf" color="#104A7B">&nbsp;&nbsp;欢迎使用垃圾箱邮件处理功能。</font></td>
<%
	end if
else
%>
	<td height="30" style='border:1px #8B969D solid;'><font class="Wf" color="#104A7B">&nbsp;&nbsp;垃圾箱邮件处理<b>成功</b>。</font></td>
<%
end if
%>
  </tr>
<br>
<tr><td><br><hr size="1" color="#8B969D"></td></tr>
<tr><td align="right"><br>
<input type="button" value=" 关闭 " LANGUAGE=javascript onclick="javascript:self.close();" class="Bsbttn">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</td></tr></table>
</td></tr></table>
</body>
</html>

<%
end if
end if

function getGRSN()
	dim theGRSN
	Randomize
	theGRSN = Int((9999999 * Rnd) + 1)

	getGRSN = "GRSN=" & CStr(theGRSN)
end function
%>
