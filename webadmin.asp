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
		alert("��֤�벻��Ϊ��");
		document.f1.AssumpsitString.focus();
		return ;
	}

	if ("<%=Application("em_SystemAdmin") %>" != document.f1.SystemAdmin.value)
	{
		if (confirm("��Ҫ: ��ǰϵͳ����Ա�ʺ�Ϊ: <%=Application("em_SystemAdmin") %>\r\n\r\n���Ƿ�ȷ��Ҫ�޸�ϵͳ����Ա�ʺ�Ϊ: " + document.f1.SystemAdmin.value + " \r\n\r\nע��: ��\"ȷ��\"��ǰ�ʺŽ���Ȩ��������ϵͳ����Ա����"))
			document.f1.submit();
	}
	else
		document.f1.submit();
}

function writeselect(name, haveNULL) {
	document.write("<select name=\"" + name + "\" class=\"drpdwn\">");

	if (haveNULL == true)
		document.write("<option value=\"\">��</option>");

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
				<input name="save1" type="button" value=" ���� " onclick="javascript:ischangeAdmin();" class="Bsbttn" disabled>&nbsp;&nbsp;
				<input type="button" value=" ȡ�� " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
                </div>
              </td>
            </tr>
            <tr> 
              <td valign=center align=right width="50%" height=30 style='border-top:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <div align="right">�����������˽���ļ�����&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <input type="text" name="MaxFolders" class='textbox' value="<%=mam.MaxFolders %>" size="10" maxlength="2">
				</td>
            </tr>
            <tr> 
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">������������POP3������&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <input type="text" name="MaxMPOP3" class='textbox' value="<%=Application("em_MaxMPOP3") %>" size="10" maxlength="2">
				</td>
            </tr>
            <tr> 
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">�����������ǩ����&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <input type="text" name="MaxSigns" class='textbox' value="<%=Application("em_MaxSigns") %>" size="10" maxlength="2">
				</td>
            </tr>
            <tr> 
              <td valign=center align=right height=30 style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <div align="right">�趨WebMailϵͳ����Ա&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
<script>writeselect("SystemAdmin", false)</script>
				</td>
            </tr>
            <tr> 
              <td valign=center align=right height=30 style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <div align="right">�趨�ʺŹ���Ա&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
<script>writeselect("AccountsAdmin", false)</script>
				</td>
            </tr>
            <tr> 
              <td valign=center align=right height=30 style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <div align="right">�趨�����ʼ�����Ա&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
<script>writeselect("SpamAdmin", false)</script>
				</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">���ù����ļ���(BBS)&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="EnableBBS" value="checkbox" <% if mam.EnableBBS = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">���ù����ļ��й���&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_ShareFolder" value="checkbox" <% if mam.Enable_ShareFolder = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
    <a name="signhold"></a>
                <div align="right">�����ʺű�������&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_SignHold" value="checkbox" <% if mam.Enable_SignHold = true then response.write "checked"%>>
				&nbsp;<a href="signhold.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">�����û�������˽Կ&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_User_Download_Private_Cert" value="checkbox" <% if mam.Enable_User_Download_Private_Cert = true then response.write "checked"%>>
			</td>
            </tr>
            <tr> 
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">�鿴��־�ļ�ʱÿҳ��ʾ&nbsp;:&nbsp;</div>
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
                <div align="right">�趨���������ʺ�Ϊ&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
<script>writeselect("TestAccounts", true)</script>
				</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">���������Ա������ӭ�ʼ�����&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_DomainAdmin_SetWelcomeMsg" value="checkbox" <% if mam.Enable_DomainAdmin_SetWelcomeMsg = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">���������Ա������������&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_DomainAdmin_SetAdvertisingMsg" value="checkbox" <% if mam.Enable_DomainAdmin_SetAdvertisingMsg = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">���������Ա�������б��ʼ�&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_DomainAdmin_SendDomainListMail" value="checkbox" <% if mam.Enable_DomainAdmin_SendDomainListMail = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">�������û���������ʾ�û��ı�ע��Ϣ&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_Show_User_Memo" value="checkbox" <% if mam.Enable_Show_User_Memo = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">������ҵ��ַ������&nbsp;:&nbsp;</div>
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
                <div align="right">��������������&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_FreeSign" value="checkbox" <% if mam.Enable_FreeSign = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">������������ʱ�����ʺŵ���С����&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <input type="text" name="Sign_AccountMinLen" class='textbox' value="<%=mam.Sign_AccountMinLen %>" size="5" maxlength="1">
				</td>
            </tr>
            <tr>
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">������������ʱ�����������С����&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <input type="text" name="Sign_PassWordMinLen" class='textbox' value="<%=mam.Sign_PassWordMinLen %>" size="5" maxlength="1">
				</td>
            </tr>
            <tr>
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30>
                <div align="right">�������������ȱʡ����ģʽ&nbsp;:&nbsp;</div>
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
                <div align="right">���俪ͨ��ʽѡ��&nbsp;:&nbsp;</div>
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
		response.write "���������������ͨ"
	elseif i = 1 then
		response.write "����Ա������ͨ"
	elseif i = 2 then
		response.write "����Ա�������û��ʼ������ͨ"
	elseif i = 3 then
		response.write "�˶���֤��������ȷ��ͨ"
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
                <div align="right">����������֤��&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <input type="text" name="AssumpsitString" class='textbox' value="<%=mam.AssumpsitString %>" size="30" maxlength="128">
				</td>
            </tr>
            <tr>
<a name="InputMoreInfo"></a>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">��������ʱ��Ҫ��д��ϸע����Ϣ&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_SignWithInputMoreInfo" value="checkbox" <% if mam.Enable_SignWithInputMoreInfo = true then response.write "checked"%>>
				&nbsp;<a href="setreginfo.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">����������Ǻ������ʺ�&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_SignWithDomainUser" value="checkbox" <% if mam.Enable_SignWithDomainUser = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">ֻ����������Ӣ������&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_SignEnglishName" value="checkbox" <% if mam.Enable_SignEnglishName = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">������Ӣ������ʱʹ��Puny����&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_Puny_DBCS_SignName" value="checkbox" <% if mam.Enable_Puny_DBCS_SignName = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">��������������ƹ���&nbsp;:&nbsp;</div>
              </td>
			<td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<input type="checkbox" name="Enable_SignNumberLimit" value="checkbox" <% if mam.Enable_SignNumberLimit = true then response.write "checked"%>>&nbsp;��������û�, <input type="text" name="SignNumberLimitDays" class='textbox' value="<%=mam.SignNumberLimitDays %>" size="3" maxlength="3"> ���ڲ�����������
			</td>
            </tr>
            <tr> 
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30> 
                <div align="right">�ʺ����뱣������&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
                <input type="text" name="SignWaitDays" class='textbox' value="<%=mam.SignWaitDays %>" size="10" maxlength="2">
				</td>
            </tr>
            <tr>
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30>
                <div align="right">�����ʼ��ķ�������&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <input type="text" name="am_Name" class='textbox' value="<%=mam.am_Name %>" maxlength="64">
				</td>
            </tr>
            <tr>
              <td valign=center align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'
    height=30>
                <div align="right">�����ʼ��ķ����ʼ��ʺ�&nbsp;:&nbsp;</div>
              </td>
              <td align=left bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
<script>writeselect("am_Accounts", false)</script>
				</td>
            </tr><tr><td colspan="2" align="center">
	<table><tr>
		<td width="26%" height="24">
		<div align="center"><br>�����ʼ�����:</div>
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
				<input name="save2" type="button" value=" ���� " onclick="javascript:ischangeAdmin();" class="Bsbttn" disabled>&nbsp;&nbsp;
				<input type="button" value=" ȡ�� " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
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
        <td width="30%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'><font color="#FF3333">&nbsp;WebMailϵͳ����Ա</font></td>
        <td width="70%"> ���ڰ�ȫ�ԵĿ���, ���齫ϵͳ����Ա�ʺ����ó�Ϊ��admin�������ʺ�(����������������֪��), �Է�ֹ������Դ��ʺż�����Ĺ���. 
        ����ע������������ʺ��Զ������ܺ�, ���ڲ���¼�˷�admin�ʺ�ʱ, ���ʺŻᱻɾ��. Ҳ���ֹ��޸Ĵ��ʺ�: ��\adminmsg\webadmin.ini �� "SystemAdmin" ������Ϊָ���ʺź�����IIS����.<br>
          <br>
        </td>
      </tr>
      <tr>
		<td width="30%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'><font color="#FF3333">&nbsp;���������ʺ�</font></td>
		<td width="70%"> ����ѡ��һ���û�Ϊ���������ʺ�ʱ, ���û��������޸��Լ�������.<br>
          <br>
        </td>
      </tr>
      <tr>
		<td width="30%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'><font color="#FF3333">&nbsp;�ʺ����뱣������</font></td>
		<td width="70%"> ���û����������, ϵͳ����������������ʺŲ��������û�����, ֱ�����ʺű�������ǳ�����ϵͳ������������ɾ��.<br>
          <br>
        </td>
      </tr>
      <tr>
        <td width="30%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'><font color="#FF3333">&nbsp;�����ʼ�</font></td>
		<td width="70%"> �����������俪ͨ��ʽΪ"����Ա�������û��ʼ������ͨ"��, ���û���������ʱϵͳ�����"�����ʼ�"���͵��û�ָ���Ľ���������, ���û���������е����伤�����Ӻ�, �����佫��ʽ��ͨʹ��.<br>
		<br>
		%time% : �����滻Ϊ�û������ʺŵ�ʱ��
		<br>
		%signmail% : �����滻Ϊ�û����뿪ͨ���ʼ���ַ
		<br>
		%username% : �����滻Ϊ��¼ϵͳ���û���
		<br>
		%ip% : �����滻Ϊ�û���������ʱ��IP��ַ
		<br>
		%accode% : �����滻Ϊ�����û�����Ĵ���
		<br><br>
		��Ҫ�ر�ע�����: �����뽫 "http://localhost/actionit.asp" ����Ϊ������������Чhttp��ַʱ�ſ��Ա�����.<br>
		��: �����������Ϊ: <font color="#FF3333">www.mydomain.com</font>, WebMail������Ŀ¼��: <font color="#FF3333">mail</font>ʱ, ����Ҫ��������Ӹ�Ϊ:<br>
		http://<font color="#FF3333">www.mydomain.com</font>/<font color="#FF3333">mail</font>/actionit.asp?accode=%accode%
		<br><br>
        </td>
      </tr>
      <tr>
		<td width="30%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'><font color="#FF3333">&nbsp;����Ľ���ģʽ</font></td>
		<td width="70%"> ������ָ������������ʺſ���ʹ���������Щ����, ���а���:
		<br>http(webmail)��smtp��pop3(��imap4)�����ַ�������.<br>
          <br>
        </td>
      </tr>
      <tr>
		<td width="30%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'><font color="#FF3333">&nbsp;������֤��</font></td>
		<td width="70%"> �ɹ���Ա������дһ�������������֤�봮, ֻ��֪����һ��֤����û��ſ���ͨ����������ҳ��ע�ᵽ����.
		<br><br>
        </td>
      </tr>
      <tr>
		<td width="30%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'><font color="#FF3333">&nbsp;��д��ϸע����Ϣ</font></td>
		<td width="70%"> Ϊ�˻�ø����ע���û���Ϣ(����: ��ʵ����, �Ա������), ����Ա����ͨ������ע����Ϣ���ķ�ʽҪ��ע��������û�������д.
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
