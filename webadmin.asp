<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if

Response.Buffer = TRUE
%>

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

dim mam
set mam = server.createobject("easymail.AdminManager")
mam.Load

dim pr
set pr = server.createobject("easymail.PendRegister")
pr.Load Application("em_SignWaitDays")


if trim(request("MaxMPOP3")) <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if trim(request("MaxFolders")) <> "" and IsNumeric(trim(request("MaxFolders"))) = true then
		mam.MaxFolders = CInt(trim(request("MaxFolders")))

		if mam.MaxFolders < 0 or mam.MaxFolders > 99 then
			mam.MaxFolders = 10
		end if
	end if

	if trim(request("MaxMPOP3")) <> "" and IsNumeric(trim(request("MaxMPOP3"))) = true then
		Application("em_MaxMPOP3") = CInt(trim(request("MaxMPOP3")))

		if Application("em_MaxMPOP3") < 0 or Application("em_MaxMPOP3") > 99 then
			Application("em_MaxMPOP3") = 10
		end if

		mam.MaxMPOP3 = Application("em_MaxMPOP3")
	end if

	if trim(request("MaxSigns")) <> "" and IsNumeric(trim(request("MaxSigns"))) = true then
		Application("em_MaxSigns") = CInt(trim(request("MaxSigns")))

		if Application("em_MaxSigns") < 0 or Application("em_MaxSigns") > 99 then
			Application("em_MaxSigns") = 10
		end if

		mam.MaxSigns = Application("em_MaxSigns")
	end if

	if trim(request("SystemAdmin")) <> "" then
		Application("em_SystemAdmin") = LCase(trim(request("SystemAdmin")))

		mam.SystemAdmin = Application("em_SystemAdmin")
	end if

	if trim(request("AccountsAdmin")) <> "" then
		Application("em_AccountsAdmin") = LCase(trim(request("AccountsAdmin")))

		mam.AccountsAdmin = Application("em_AccountsAdmin")
	end if

	if trim(request("SpamAdmin")) <> "" then
		Application("em_SpamAdmin") = LCase(trim(request("SpamAdmin")))

		mam.SpamAdmin = Application("em_SpamAdmin")
	end if

	if trim(request("EnableBBS")) = "" then
		Application("em_EnableBBS") = false
		mam.EnableBBS = false
	else
		Application("em_EnableBBS") = true
		mam.EnableBBS = true
	end if

	if trim(request("Enable_SignHold")) = "" then
		Application("em_Enable_SignHold") = false
		mam.Enable_SignHold = false
	else
		Application("em_Enable_SignHold") = true
		mam.Enable_SignHold = true
	end if

	if trim(request("Enable_User_Download_Private_Cert")) = "" then
		mam.Enable_User_Download_Private_Cert = false
	else
		mam.Enable_User_Download_Private_Cert = true
	end if

	if trim(request("Enable_FreeSign")) = "" then
		Application("em_Enable_FreeSign") = false
		mam.Enable_FreeSign = false
	else
		Application("em_Enable_FreeSign") = true
		mam.Enable_FreeSign = true
	end if

	if trim(request("Enable_SignWithDomainUser")) = "" then
		Application("em_Enable_SignWithDomainUser") = false
		mam.Enable_SignWithDomainUser = false
	else
		Application("em_Enable_SignWithDomainUser") = true
		mam.Enable_SignWithDomainUser = true
	end if

	if trim(request("Enable_SignNumberLimit")) = "" then
		Application("em_Enable_SignNumberLimit") = false
		mam.Enable_SignNumberLimit = false
	else
		Application("em_Enable_SignNumberLimit") = true
		mam.Enable_SignNumberLimit = true
	end if

	if trim(request("SignNumberLimitDays")) <> "" and IsNumeric(trim(request("SignNumberLimitDays"))) = true then
		Application("em_SignNumberLimitDays") = CInt(trim(request("SignNumberLimitDays")))

		if Application("em_SignNumberLimitDays") < 0 or Application("em_SignNumberLimitDays") > 999 then
			Application("em_SignNumberLimitDays") = 1
		end if

		mam.SignNumberLimitDays = Application("em_SignNumberLimitDays")
	end if

	if trim(request("Enable_ShareFolder")) = "" then
		Application("em_Enable_ShareFolder") = false
		mam.Enable_ShareFolder = false
	else
		Application("em_Enable_ShareFolder") = true
		mam.Enable_ShareFolder = true
	end if

	if trim(request("Enable_SignEnglishName")) = "" then
		Application("em_Enable_SignEnglishName") = false
		mam.Enable_SignEnglishName = false
	else
		Application("em_Enable_SignEnglishName") = true
		mam.Enable_SignEnglishName = true
	end if

	if trim(request("LogPageKSize")) <> "" and IsNumeric(trim(request("LogPageKSize"))) = true then
		Application("em_LogPageKSize") = CInt(trim(request("LogPageKSize")))

		if Application("em_LogPageKSize") < 1 or Application("em_LogPageKSize") > 999 then
			Application("em_LogPageKSize") = 50
		end if

		mam.LogPageKSize = Application("em_LogPageKSize")
	end if

	if trim(request("SignMode")) <> "" and IsNumeric(trim(request("SignMode"))) = true then
		Application("em_SignMode") = CInt(trim(request("SignMode")))

		if Application("em_SignMode") < 0 or Application("em_SignMode") > 3 then
			Application("em_SignMode") = 0
		end if

		mam.SignMode = Application("em_SignMode")
	end if

	if trim(request("SignWaitDays")) <> "" and IsNumeric(trim(request("SignWaitDays"))) = true then
		Application("em_SignWaitDays") = CInt(trim(request("SignWaitDays")))

		if Application("em_SignWaitDays") < 1 or Application("em_SignWaitDays") > 99 then
			Application("em_SignWaitDays") = 7
		end if

		mam.SignWaitDays = Application("em_SignWaitDays")
	else
		Application("em_SignWaitDays") = 7
		mam.SignWaitDays = Application("em_SignWaitDays")
	end if

	if trim(request("am_Name")) <> "" then
		Application("em_am_Name") = trim(request("am_Name"))

		mam.am_Name = Application("em_am_Name")
	end if

	if trim(request("am_Accounts")) <> "" then
		Application("em_am_Accounts") = LCase(trim(request("am_Accounts")))

		mam.am_Accounts = Application("em_am_Accounts")
	end if


	Application("em_TestAccounts") = LCase(trim(request("TestAccounts")))
	mam.TestAccounts = Application("em_TestAccounts")


	pr.acSubject = trim(request("acSubject"))
	pr.acText = trim(request("acText"))



	if trim(request("Enable_DomainAdmin_SetWelcomeMsg")) = "" then
		mam.Enable_DomainAdmin_SetWelcomeMsg = false
	else
		mam.Enable_DomainAdmin_SetWelcomeMsg = true
	end if

	if trim(request("Enable_DomainAdmin_SetAdvertisingMsg")) = "" then
		mam.Enable_DomainAdmin_SetAdvertisingMsg = false
	else
		mam.Enable_DomainAdmin_SetAdvertisingMsg = true
	end if

	if trim(request("Enable_DomainAdmin_SendDomainListMail")) = "" then
		mam.Enable_DomainAdmin_SendDomainListMail = false
	else
		mam.Enable_DomainAdmin_SendDomainListMail = true
	end if

	if trim(request("Enable_Show_User_Memo")) = "" then
		mam.Enable_Show_User_Memo = false
	else
		mam.Enable_Show_User_Memo = true
	end if

	if trim(request("Enable_Show_EntAddress")) = "" then
		Application("em_EnableEntAddress") = false
		mam.Enable_Show_EntAddress = false
	else
		Application("em_EnableEntAddress") = true
		mam.Enable_Show_EntAddress = true
	end if

	if trim(request("Enable_SignWithInputMoreInfo")) = "" then
		mam.Enable_SignWithInputMoreInfo = false
	else
		mam.Enable_SignWithInputMoreInfo = true
	end if

	if trim(request("Enable_Puny_DBCS_SignName")) = "" then
		mam.Enable_Puny_DBCS_SignName = false
	else
		mam.Enable_Puny_DBCS_SignName = true
	end if

	mam.AssumpsitString = trim(request("AssumpsitString"))


	if trim(request("Sign_AccountMinLen")) <> "" and IsNumeric(trim(request("Sign_AccountMinLen"))) = true then
		mam.Sign_AccountMinLen = CLng(trim(request("Sign_AccountMinLen")))
	end if

	if trim(request("Sign_PassWordMinLen")) <> "" and IsNumeric(trim(request("Sign_PassWordMinLen"))) = true then
		mam.Sign_PassWordMinLen = CLng(trim(request("Sign_PassWordMinLen")))
	end if

	if trim(request("Sign_AccessMode")) <> "" and IsNumeric(trim(request("Sign_AccessMode"))) = true then
		mam.Sign_AccessMode = CLng(trim(request("Sign_AccessMode")))
	end if


	mam.save
	pr.Save

	set mam = nothing
	set pr = nothing


	if err.number = 0 then
		response.redirect "ok.asp?" & getGRSN() & "&gourl=webadmin.asp"
	else
		response.redirect "err.asp?" & getGRSN() & "&gourl=webadmin.asp"
	end if
end if 


dim eu
set eu = Application("em")
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<script language="JavaScript">
<!--
var userlist = "<%

i = 0
allnum = eu.GetUsersCount

do while i < allnum
	eu.GetUserByIndex i, name, domain, comment

	Response.Write "<option value=" & Chr(92) & """" & server.htmlencode(name) & Chr(92) & """>" & server.htmlencode(name) & "</option>" & Chr(92) & Chr(13)
	Response.Flush

	name = NULL
	domain = NULL
	comment = NULL

	i = i + 1
loop

Response.Write "</select>"
%>";


function window_onload() {
	document.f1.SystemAdmin.value = "<%=Application("em_SystemAdmin") %>";
	document.f1.am_Accounts.value = "<%=mam.am_Accounts %>";
	document.f1.AccountsAdmin.value = "<%=mam.AccountsAdmin %>";
	document.f1.TestAccounts.value = "<%=mam.TestAccounts %>";
	document.f1.SpamAdmin.value = "<%=mam.SpamAdmin %>";

	document.f1.save1.disabled = false;
	document.f1.save2.disabled = false;
}

function ischangeAdmin() {
	if (document.f1.SignMode.value == "3" && document.f1.AssumpsitString.value.length < 1)
	{
		alert("验证码不可为空");
		document.f1.AssumpsitString.focus();
		return ;
	}

	if ("<%=Application("em_SystemAdmin") %>" != document.f1.SystemAdmin.value)
	{
		if (confirm("重要: 当前系统管理员帐号为: <%=Application("em_SystemAdmin") %>\r\n\r\n您是否确定要修改系统管理员帐号为: " + document.f1.SystemAdmin.value + " \r\n\r\n注意: 按\"确定\"后当前帐号将无权继续进行系统管理员操作"))
			document.f1.submit();
	}
	else
		document.f1.submit();
}

function writeselect(name, haveNULL) {
	document.write("<select name=\"" + name + "\" class=\"drpdwn\">");

	if (haveNULL == true)
		document.write("<option value=\"\">无</option>");

	document.write(userlist);
}
// -->
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
<br>
<form method="post" action="webadmin.asp" name="f1">
  <table width="90%" border="0" align="center">
    <tr> 
      <td> 
        <div align="center"> 
          <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>">
            <tr> 
              <td valign=center align=right bgcolor=#ffffff 
    height=40 colspan="2"> 
                <div align="right"> 
				<input name="save1" type="button" value=" 保存 " onclick="javascript:ischangeAdmin();" class="Bsbttn" disabled>&nbsp;&nbsp;
				<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
                </div>
              </td>
            </tr>
            <tr> 
              <td valign=center align=right width="50%" height=30 style='border-top:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <div align="right">允许创建的最大私人文件夹数&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <input type="text" name="MaxFolders" class='textbox' value="<%=mam.MaxFolders %>" size="10" maxlength="2">
				</td>
            </tr>
            <tr> 
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">允许创建的最大多POP3下载数&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <input type="text" name="MaxMPOP3" class='textbox' value="<%=Application("em_MaxMPOP3") %>" size="10" maxlength="2">
				</td>
            </tr>
            <tr> 
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">允许创建的最大签名数&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <input type="text" name="MaxSigns" class='textbox' value="<%=Application("em_MaxSigns") %>" size="10" maxlength="2">
				</td>
            </tr>
            <tr> 
              <td valign=center align=right height=30 style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <div align="right">设定WebMail系统管理员&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
<script>writeselect("SystemAdmin", false)</script>
				</td>
            </tr>
            <tr> 
              <td valign=center align=right height=30 style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <div align="right">设定帐号管理员&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
<script>writeselect("AccountsAdmin", false)</script>
				</td>
            </tr>
            <tr> 
              <td valign=center align=right height=30 style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <div align="right">设定垃圾邮件管理员&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
<script>writeselect("SpamAdmin", false)</script>
				</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">启用公共文件夹(BBS)&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="EnableBBS" value="checkbox" <% if mam.EnableBBS = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">启用共享文件夹功能&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_ShareFolder" value="checkbox" <% if mam.Enable_ShareFolder = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
    <a name="signhold"></a>
                <div align="right">启用帐号保留功能&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_SignHold" value="checkbox" <% if mam.Enable_SignHold = true then response.write "checked"%>>
				&nbsp;<a href="signhold.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">允许用户下载其私钥&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_User_Download_Private_Cert" value="checkbox" <% if mam.Enable_User_Download_Private_Cert = true then response.write "checked"%>>
			</td>
            </tr>
            <tr> 
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">查看日志文件时每页显示&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <select name="LogPageKSize" class="drpdwn">
<%
i = 10
if i <> Application("em_LogPageKSize") then
	response.write "<option value='" & i & "'>" & i & "K</option>"
else
	response.write "<option value='" & i & "' selected>" & i & "K</option>"
end if

i = 20
if i <> Application("em_LogPageKSize") then
	response.write "<option value='" & i & "'>" & i & "K</option>"
else
	response.write "<option value='" & i & "' selected>" & i & "K</option>"
end if

i = 50
do while i < 999
	if i <> Application("em_LogPageKSize") then
		response.write "<option value='" & i & "'>" & i & "K</option>" & Chr(10)
	else
		response.write "<option value='" & i & "' selected>" & i & "K</option>" & Chr(10)
	end if

	i = i + 50
loop
%>
        </select>
				</td>
            </tr>
            <tr> 
              <td valign=center align=right height=30 style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <div align="right">设定公开测试帐号为&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
<script>writeselect("TestAccounts", true)</script>
				</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">允许域管理员设置域欢迎邮件内容&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_DomainAdmin_SetWelcomeMsg" value="checkbox" <% if mam.Enable_DomainAdmin_SetWelcomeMsg = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">允许域管理员设置域广告内容&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_DomainAdmin_SetAdvertisingMsg" value="checkbox" <% if mam.Enable_DomainAdmin_SetAdvertisingMsg = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">允许域管理员发送域列表邮件&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_DomainAdmin_SendDomainListMail" value="checkbox" <% if mam.Enable_DomainAdmin_SendDomainListMail = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">允许在用户管理中显示用户的备注信息&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_Show_User_Memo" value="checkbox" <% if mam.Enable_Show_User_Memo = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">启用企业地址簿功能&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_Show_EntAddress" value="checkbox" <% if mam.Enable_Show_EntAddress = true then response.write "checked"%>>
			</td>
            </tr>
            </table><br>
			<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>">
            <tr>
              <td valign=center align=right width="50%" height=30 bgcolor="<%=MY_COLOR_2 %>" style='border-top:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;' 
    height=30> 
                <div align="right">允许公开申请邮箱&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_FreeSign" value="checkbox" <% if mam.Enable_FreeSign = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">公开申请邮箱时输入帐号的最小长度&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <input type="text" name="Sign_AccountMinLen" class='textbox' value="<%=mam.Sign_AccountMinLen %>" size="5" maxlength="1">
				</td>
            </tr>
            <tr>
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">公开申请邮箱时输入密码的最小长度&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <input type="text" name="Sign_PassWordMinLen" class='textbox' value="<%=mam.Sign_PassWordMinLen %>" size="5" maxlength="1">
				</td>
            </tr>
            <tr>
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30>
                <div align="right">公开申请邮箱的缺省接入模式&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
	<select name="Sign_AccessMode" class=drpdwn size="1">
<%
anum = 0

do while anum < 7
	if anum <> mam.Sign_AccessMode then
		Response.Write "<option value=""" & anum & """>" & getaccessmode(anum) & "</option>" & Chr(13)
	else
		Response.Write "<option value=""" & anum & """ selected>" & getaccessmode(anum) & "</option>" & Chr(13)
	end if

	anum = anum + 1
loop
%>
	</select>
				</td>
            </tr>
            <tr> 
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">邮箱开通方式选择&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <select name="SignMode" class="drpdwn">
<%
i = 0

do while i < 4
	if i <> mam.SignMode then
		response.write "<option value='" & i & "'>"
	else
		response.write "<option value='" & i & "' selected>"
	end if

	if i = 0 then
		response.write "申请邮箱后立即开通"
	elseif i = 1 then
		response.write "管理员审批后开通"
	elseif i = 2 then
		response.write "管理员审批或用户邮件激活后开通"
	elseif i = 3 then
		response.write "核对验证码内容正确后开通"
	end if

	response.write "</option>" & Chr(10)

	i = i + 1
loop
%>
        </select>
				</td>
            </tr>
            <tr> 
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30>
                <div align="right">邮箱申请验证码&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <input type="text" name="AssumpsitString" class='textbox' value="<%=mam.AssumpsitString %>" size="30" maxlength="128">
				</td>
            </tr>
            <tr>
<a name="InputMoreInfo"></a>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">公开申请时需要填写详细注册信息&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_SignWithInputMoreInfo" value="checkbox" <% if mam.Enable_SignWithInputMoreInfo = true then response.write "checked"%>>
				&nbsp;<a href="setreginfo.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">公开申请的是含域名帐号&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_SignWithDomainUser" value="checkbox" <% if mam.Enable_SignWithDomainUser = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">只允许公开申请英文邮箱&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_SignEnglishName" value="checkbox" <% if mam.Enable_SignEnglishName = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">创建非英文邮箱时使用Puny编码&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_Puny_DBCS_SignName" value="checkbox" <% if mam.Enable_Puny_DBCS_SignName = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">启用申请次数限制功能&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_SignNumberLimit" value="checkbox" <% if mam.Enable_SignNumberLimit = true then response.write "checked"%>>&nbsp;已申请的用户, <input type="text" name="SignNumberLimitDays" class='textbox' value="<%=mam.SignNumberLimitDays %>" size="3" maxlength="3"> 天内不允许再申请
			</td>
            </tr>
            <tr> 
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">帐号申请保留天数&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <input type="text" name="SignWaitDays" class='textbox' value="<%=mam.SignWaitDays %>" size="10" maxlength="2">
				</td>
            </tr>
            <tr>
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30>
                <div align="right">激活邮件的发件人名&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <input type="text" name="am_Name" class='textbox' value="<%=mam.am_Name %>" maxlength="64">
				</td>
            </tr>
            <tr>
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30>
                <div align="right">激活邮件的发件邮件帐号&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
<script>writeselect("am_Accounts", false)</script>
				</td>
            </tr><tr><td colspan="2" align="center">
	<table><tr>
		<td width="26%" height="24">
		<div align="center"><br>激活邮件主题:</div>
	</td>
	<td width="74%"><br>
		<input name="acSubject" type="text" value="<%=pr.acSubject %>" size="40" class='textbox'>
	</td>
	</tr>
	<tr>
	<td colspan="2">
	<div align="center">
	<textarea name="acText" cols="<%
if isMSIE = true then
	Response.Write "60"
else
	Response.Write "50"
end if
%>" rows="7" class='textarea'><%=pr.acText %></textarea>
	</div>
	</td>
	</tr></table><br></td></tr>
            <tr> 
              <td valign=center align=right bgcolor=#ffffff 
    height=40 colspan="2"> 
                <div align="right"> 
				<input name="save2" type="button" value=" 保存 " onclick="javascript:ischangeAdmin();" class="Bsbttn" disabled>&nbsp;&nbsp;
				<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
                </div>
              </td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
  </table><br><br><br>
  <div align="center">
    <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10">
        </td>
      </tr>
      <tr> 
        <td width="30%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'><font color="#FF3333">&nbsp;WebMail系统管理员</font></td>
        <td width="70%"> 出于安全性的考虑, 建议将系统管理员帐号设置成为非admin的其他帐号(尽量避免让其他人知道), 以防止可能针对此帐号及密码的攻击. 
        但需注意如果启用了帐号自动清理功能后, 长期不登录此非admin帐号时, 该帐号会被删除. 也可手工修改此帐号: 打开\adminmsg\webadmin.ini 将 "SystemAdmin" 项设置为指定帐号后重启IIS即可.<br>
          <br>
        </td>
      </tr>
      <tr>
		<td width="30%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'><font color="#FF3333">&nbsp;公开测试帐号</font></td>
		<td width="70%"> 当您选定一个用户为公开测试帐号时, 此用户将不可修改自己的密码.<br>
          <br>
        </td>
      </tr>
      <tr>
		<td width="30%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'><font color="#FF3333">&nbsp;帐号申请保留天数</font></td>
		<td width="70%"> 在用户申请邮箱后, 系统将保留其所申请的帐号不被其他用户所用, 直到此帐号被激活或是超过了系统保留天数而被删除.<br>
          <br>
        </td>
      </tr>
      <tr>
        <td width="30%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'><font color="#FF3333">&nbsp;激活邮件</font></td>
		<td width="70%"> 当您设置邮箱开通方式为"管理员审批或用户邮件激活后开通"后, 当用户申请邮箱时系统将会把"激活邮件"发送到用户指定的接收信箱中, 当用户点击了信中的邮箱激活链接后, 此邮箱将正式开通使用.<br>
		<br>
		%time% : 将被替换为用户申请帐号的时间
		<br>
		%signmail% : 将被替换为用户申请开通的邮件地址
		<br>
		%username% : 将被替换为登录系统的用户名
		<br>
		%ip% : 将被替换为用户申请邮箱时的IP地址
		<br>
		%accode% : 将被替换为激活用户邮箱的串号
		<br><br>
		需要特别注意的是: 您必须将 "http://localhost/actionit.asp" 更改为您服务器的有效http地址时才可以被激活.<br>
		例: 如果您的域名为: <font color="#FF3333">www.mydomain.com</font>, WebMail的虚拟目录是: <font color="#FF3333">mail</font>时, 您需要把这个链接改为:<br>
		http://<font color="#FF3333">www.mydomain.com</font>/<font color="#FF3333">mail</font>/actionit.asp?accode=%accode%
		<br><br>
        </td>
      </tr>
      <tr>
		<td width="30%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'><font color="#FF3333">&nbsp;邮箱的接入模式</font></td>
		<td width="70%"> 是用来指定公开申请的帐号可以使用邮箱的哪些服务, 其中包括:
		<br>http(webmail)、smtp、pop3(含imap4)这三种服务的组合.<br>
          <br>
        </td>
      </tr>
      <tr>
		<td width="30%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'><font color="#FF3333">&nbsp;申请验证码</font></td>
		<td width="70%"> 由管理员随意填写一个申请邮箱的验证码串, 只有知道这一验证码的用户才可以通过邮箱申请页面注册到邮箱.
		<br><br>
        </td>
      </tr>
      <tr>
		<td width="30%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'><font color="#FF3333">&nbsp;填写详细注册信息</font></td>
		<td width="70%"> 为了获得更多的注册用户信息(比如: 真实姓名, 性别等资料), 管理员可以通过生成注册信息表格的方式要求注册邮箱的用户进行填写.
		<br><br>
        </td>
      </tr>
    </table>
  </div>
</Form>
<br>
</BODY>
</HTML>

<%
set mam = nothing
set eu = nothing
set pr = nothing


function getaccessmode(amode)
	if amode = 0 then
		getaccessmode = "http/smtp/pop3,imap4"
	elseif amode = 1 then
		getaccessmode = "smtp/pop3,imap4"
	elseif amode = 2 then
		getaccessmode = "http/smtp"
	elseif amode = 3 then
		getaccessmode = "http/pop3,imap4"
	elseif amode = 4 then
		getaccessmode = "http"
	elseif amode = 5 then
		getaccessmode = "smtp"
	elseif amode = 6 then
		getaccessmode = "pop3,imap4"
	end if
end function
%>
