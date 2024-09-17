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
		Response.Redirect "default.asp?errstr=" & Server.URLEncode("≥¨ ±, «Î÷ÿ ‘") & "&" & getGRSN()
	end if
end if

accode = trim(request("accode"))

dim isok
isok = false

dim pr
set pr = server.createobject("easymail.PendRegister")
pr.Load Application("em_SignWaitDays")

if pr.ActionIt(accode) = true then
	isok = true
	pr.Save
end if

set pr = nothing

if trim(request("thispage")) <> "" and trim(request("gourl")) <> "" then
	response.redirect "pendreg.asp?" & getGRSN() & "&page=" & trim(request("thispage"))
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
if accode = "" or isok = false then
%>
	<td height="30" style='border-top:1px #8B969D solid; border-left:1px #8B969D solid; border-bottom:1px #8B969D solid; border-right:1px #8B969D solid;' colspan="2">&nbsp;<img src="images/error.gif" align="absmiddle" border="0"><font class="Wf" color="#104A7B">&nbsp;” œ‰’ ∫≈º§ªÓ<b> ß∞‹</b>°£</font></td>
<%
else
%>
	<td height="30" style='border-top:1px #8B969D solid; border-left:1px #8B969D solid; border-bottom:1px #8B969D solid; border-right:1px #8B969D solid;' colspan="2"><font class="Wf" color="#104A7B">&nbsp;&nbsp;” œ‰’ ∫≈º§ªÓ<b>≥…π¶</b>°£</font></td>
<%
end if
%>
  </tr>
<br>
<tr><td><br><hr size="1" color="#8B969D"></td></tr>
<tr><td align="right" colspan="2"><br>
<input type="button" value=" »∑∂® " onclick="javascript:location.href='default.asp';" class="Bsbttn">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</td></tr></table>
</td></tr></table>
</body>

</html>

<%
function getGRSN()
	dim theGRSN
	Randomize
	theGRSN = Int((9999999 * Rnd) + 1)

	getGRSN = "GRSN=" & CStr(theGRSN)
end function
%>
