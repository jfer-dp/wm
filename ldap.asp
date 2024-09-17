<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
isamg = false
amg = trim(request("amg"))
gourl = trim(request("gourl"))
userid = trim(request("id"))
name = trim(request("name"))
domain = trim(request("domain"))

if amg = "1" and Len(gourl) > 0 and Len(userid) > 0 and Len(name) > 0 and Len(domain) > 0 then
	isamg = true
end if


if isamg = true then
	if isadmin() = false and isAccountsAdmin() = false then
		dim ed
		set ed = server.createobject("easymail.domain")
		ed.Load

		if ed.GetUserManagerDomainCount(Session("wem")) < 1 then
			set ed = nothing
			Response.Redirect "noadmin.asp"
		end if

		i = 0
		allnum = ed.GetUserManagerDomainCount(Session("wem"))

		dim isok
		isok = false

		do while i < allnum
			cdomainstr = ed.GetUserManagerDomain(Session("wem"), i)

			if LCase(cdomainstr) = LCase(domain) then
				isok = true
			end if

			cdomainstr = NULL

			i = i + 1
		loop

		set ed = nothing

		if isok = false then
			Response.Redirect "noadmin.asp"
		end if
	end if
end if


issave = trim(request("issave"))
islight = trim(request("islight"))

isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

dim ldap
set ldap = server.createobject("easymail.LDAP")

if isamg = false then
	if islight = "1" then
		ldap.Load Session("wem"), false, true
	else
		ldap.Load Session("wem"), false, false
	end if
else
	if islight = "1" then
		ldap.Load name, false, true
	else
		ldap.Load name, false, false
	end if
end if

canSave = true
if ldap.Enable_LDAP_User_Edit = false and isamg = false then
	canSave = false
end if


if issave = "1" and canSave = true and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	mylog = ""
	isok = false

	need_modify = true
	need_add = false
	need_del = false

	if islight = "1" then
		need_modify = false
	end if

	is_change_LDAP_Logon = false
	if isamg = true then
		if trim(request("Enable_LDAP_User_Edit")) <> "" then
			ldap.Enable_LDAP_User_Edit = true
		else
			ldap.Enable_LDAP_User_Edit = false
		end if

		if trim(request("Enable_LDAP_Logon")) <> "" then
			if ldap.Enable_LDAP_Logon = false then
				is_change_LDAP_Logon = true
				ldap.Enable_LDAP_Logon = true
				need_modify = true
			end if
		else
			if ldap.Enable_LDAP_Logon = true then
				is_change_LDAP_Logon = true
				ldap.Enable_LDAP_Logon = false
				need_del = true
				need_add = true
			end if
		end if

		if trim(request("Enable_LDAP_Show")) <> "" then
			if ldap.Enable_LDAP_Show = false then
				ldap.Enable_LDAP_Show = true
				need_add = true
			end if
		else
			if ldap.Enable_LDAP_Show = true then
				ldap.Enable_LDAP_Show = false
				need_del = true
				need_add = false
			end if
		end if

		if trim(request("Enable_LDAP_User_Edit_OU")) <> "" then
			ldap.Enable_LDAP_User_Edit_OU = true
		else
			ldap.Enable_LDAP_User_Edit_OU = false
		end if
	end if

	ldap.mail = trim(request("mail"))
	ldap.homePhone = trim(request("homePhone"))
	ldap.pager = trim(request("pager"))
	ldap.mobile = trim(request("mobile"))
	ldap.telephoneNumber = trim(request("telephoneNumber"))
	ldap.facsimileTelephoneNumber = trim(request("facsimileTelephoneNumber"))
	ldap.title = trim(request("title"))
	ldap.ou = trim(request("ou"))
	ldap.physicalDeliveryOfficeName = trim(request("physicalDeliveryOfficeName"))
	ldap.sn = trim(request("sn"))
	ldap.givenName = trim(request("givenName"))
	ldap.cn = trim(request("cn"))

	if trim(request("Enable_Synchro_Password")) <> "" then
		ldap.Enable_Synchro_Password = true
	else
		ldap.Enable_Synchro_Password = false
	end if

	if trim(request("Enable_LDAP_EntAddress_Show")) <> "" then
		ldap.Enable_LDAP_EntAddress_Show = true
	else
		ldap.Enable_LDAP_EntAddress_Show = false
	end if

	if ldap.Save = true then
		isok = true
	end if


	ldappw = trim(request("ldappw"))
	if isamg = true and Len(ldappw) > 0 then
		ldap.Submit_Password name, ldappw
	end if

	if isok = true then
		need_log = false

		if is_change_LDAP_Logon = true then
			if isamg = false then
				ldap.Set_User_LDAPLogin Session("wem"), ldap.Enable_LDAP_Logon
			else
				ldap.Set_User_LDAPLogin name, ldap.Enable_LDAP_Logon
			end if
		end if

		if need_del = true then
			if isamg = false then
				ldap.Submit_Del Session("wem"), false, true
			else
				ldap.Submit_Del name, false, true
			end if
			need_log = true

			if need_add = true then
				if isamg = false then
					ldap.Submit_Add Session("wem"), false, true
				else
					ldap.Submit_Add name, false, true
				end if
			end if
		elseif need_add = true then
			if isamg = false then
				ldap.Submit_Add Session("wem"), false, true
			else
				ldap.Submit_Add name, false, true
			end if
			need_log = true
		elseif need_modify = true then
			if isamg = false then
				ldap.Submit_Modify Session("wem"), false, true
			else
				ldap.Submit_Modify name, false, true
			end if
			need_log = true
		end if

		if need_log = true then
			mylog = ldap.Log
			if mylog <> "" then
				isok = false
			end if
		end if
	end if

	set ldap = nothing

	if isamg = false then
		if isok = true then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("ldap.asp")
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("ldap.asp") & "&errstr=" & Server.URLEncode(mylog)
		end if
	else
		Response.Redirect "changepw.asp?" & getGRSN() & "&fo=1&id=" & userid & "&gourl=" & Server.URLEncode(gourl)
	end if
end if
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<STYLE type=text/css>
<!--
.st_l {height:24px; text-align:right; white-space:nowrap;}
.st_r {height:24px; text-align:left; white-space:nowrap;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!-- 
var old_info = "";
function gosub() {
	var new_info = document.f1.cn.value + '\f' + document.f1.mail.value + '\f' + document.f1.ou.value + '\f' + document.f1.physicalDeliveryOfficeName.value + '\f' + document.f1.title.value + '\f' + document.f1.mobile.value + '\f' + document.f1.telephoneNumber.value + '\f' + document.f1.facsimileTelephoneNumber.value + '\f' + document.f1.sn.value + '\f' + document.f1.givenName.value + '\f' + document.f1.homePhone.value + '\f' + document.f1.pager.value;

	if (old_info == new_info)
		document.f1.islight.value = "1";

	document.f1.submit();
}

function window_onload() {
	old_info = document.f1.cn.value + '\f' + document.f1.mail.value + '\f' + document.f1.ou.value + '\f' + document.f1.physicalDeliveryOfficeName.value + '\f' + document.f1.title.value + '\f' + document.f1.mobile.value + '\f' + document.f1.telephoneNumber.value + '\f' + document.f1.facsimileTelephoneNumber.value + '\f' + document.f1.sn.value + '\f' + document.f1.givenName.value + '\f' + document.f1.homePhone.value + '\f' + document.f1.pager.value;
}

function showentads() {
	var remote = null;
	remote = window.open("ldap_oupop.asp?ofm=ou&<%=getGRSN() %>", "", "top=80; left=130; height=330,width=510,scrollbars=yes,resizable=yes,status=no,toolbar=no,menubar=no,location=no");
}

function goback()
{
<%
if isamg = false then
%>
	location.href = "myreginfo.asp?<%=getGRSN() %>";
<%
else
%>
	location.href = "changepw.asp?<%=getGRSN() %>&fo=1&id=<%=userid %>&gourl=<%=Server.URLEncode(gourl) %>";
<%
end if
%>
}
// -->
</script>

<body LANGUAGE=javascript onload="return window_onload()">
<form name="f1" method=post action="ldap.asp">
<input type="hidden" name="issave" value="1">
<input type="hidden" name="islight" value="0">

<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=s_lang_0032 %>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="center">

<table align="center" cellpadding=4 cellspacing=0 border=0 width="97%" style="border:1px #A5B6C8 solid;">
	<tr>
	<td width="12%" class="st_l"><b><%=s_lang_0034 %></b><%=s_lang_mh %></td><td colspan="4" class="st_r"><input type=text size="26" name="dn" id="dn" class="n_textbox" readonly value="<%=ldap.dn %>"></td>
	</tr>
	<tr>
	<td width="12%" class="st_l"><%=s_lang_0035 %><%=s_lang_mh %></td><td width="38%" class="st_r"><input type=text size="26" name="cn" class="n_textbox" maxlength="64" value="<%=ldap.cn %>"></td>
	<td width="12%" class="st_l"><%=s_lang_0036 %><%=s_lang_mh %></td><td width="38%" class="st_r"><input type=text size="26" name="mail" class="n_textbox" maxlength="64" value="<%=ldap.mail %>"></td>
	</tr>
	<tr>
	<td class="st_l"><%=s_lang_0043 %><%=s_lang_mh %></td><td class="st_r"><input type=text name="ou" id="ou" class="n_textbox" maxlength="64" value="<%=ldap.ou %>" <%
if ldap.Enable_LDAP_User_Edit_OU = true or isamg = true then
	Response.Write "size=""20"">&nbsp;<input type=button value=""..."" LANGUAGE=javascript onclick=""showentads()"" class=""sbttn"" style=""WIDTH: 22px"">"
else
	Response.Write "size=""26"" readonly>"
end if
%></td>
	<td class="st_l"><%=s_lang_0044 %><%=s_lang_mh %></td><td class="st_r"><input type=text size="26" name="physicalDeliveryOfficeName" class="n_textbox" maxlength="64" value="<%=ldap.physicalDeliveryOfficeName %>"></td>
	</tr>
	<tr>
	<td class="st_l"><%=s_lang_0042 %><%=s_lang_mh %></td><td class="st_r"><input type=text size="26" name="title" class="n_textbox" maxlength="64" value="<%=ldap.title %>"></td>
	<td class="st_l"><%=s_lang_0039 %><%=s_lang_mh %></td><td class="st_r"><input type=text size="26" name="mobile" class="n_textbox" maxlength="64" value="<%=ldap.mobile %>"></td>
	</tr>
	<tr>
	<td class="st_l"><%=s_lang_0040 %><%=s_lang_mh %></td><td class="st_r"><input type=text size="26" name="telephoneNumber" class="n_textbox" maxlength="64" value="<%=ldap.telephoneNumber %>"></td>
	<td class="st_l"><%=s_lang_0041 %><%=s_lang_mh %></td><td class="st_r"><input type=text size="26" name="facsimileTelephoneNumber" class="n_textbox" maxlength="64" value="<%=ldap.facsimileTelephoneNumber %>"></td>
	</tr>
	<tr>
	<td class="st_l"><%=s_lang_0045 %><%=s_lang_mh %></td><td class="st_r"><input type=text size="26" name="sn" class="n_textbox" maxlength="64" value="<%=ldap.sn %>"></td>
	<td class="st_l"><%=s_lang_0046 %><%=s_lang_mh %></td><td class="st_r"><input type=text size="26" name="givenName" class="n_textbox" maxlength="64" value="<%=ldap.givenName %>"></td>
	</tr>
	<tr>
	<td class="st_l"><%=s_lang_0037 %><%=s_lang_mh %></td><td class="st_r"><input type=text size="26" name="homePhone" class="n_textbox" maxlength="64" value="<%=ldap.homePhone %>"></td>
	<td class="st_l"><%=s_lang_0038 %><%=s_lang_mh %></td><td class="st_r"><input type=text size="26" name="pager" class="n_textbox" maxlength="64" value="<%=ldap.pager %>"></td>
	</tr>
	<tr><td colspan="4" bgcolor="white" class="st_r" style="border-top:1px solid #A5B6C8;">
	<input type="checkbox" name="Enable_Synchro_Password" id="Enable_Synchro_Password" <% if ldap.Enable_Synchro_Password = true then response.write "checked"%>>
	<%=s_lang_0051 %>
	</td></tr>
	<tr><td colspan="4" bgcolor="white" class="st_r" style="border-top:1px solid #A5B6C8;">
	<input type="checkbox" name="Enable_LDAP_EntAddress_Show" id="Enable_LDAP_EntAddress_Show" <% if ldap.Enable_LDAP_EntAddress_Show = true then response.write "checked"%>>
	<%=s_lang_0052 %>
	</td></tr>
<% if isamg = true then %>
	<tr><td colspan="4" bgcolor="white" class="st_r" style="border-top:1px solid #A5B6C8;">
	<input type="checkbox" name="Enable_LDAP_User_Edit" id="Enable_LDAP_User_Edit" <% if ldap.Enable_LDAP_User_Edit = true then response.write "checked"%>>
	<%=s_lang_0053 %>
	</td></tr>
	<tr><td colspan="4" bgcolor="white" class="st_r" style="border-top:1px solid #A5B6C8;">
	<input type="checkbox" name="Enable_LDAP_Show" id="Enable_LDAP_Show" <% if ldap.Enable_LDAP_Show = true then response.write "checked"%>>
	<%=s_lang_0054 %>
	</td></tr>
	<tr><td colspan="4" bgcolor="white" class="st_r" style="border-top:1px solid #A5B6C8;">
	<input type="checkbox" name="Enable_LDAP_Logon" id="Enable_LDAP_Logon" <% if ldap.Enable_LDAP_Logon = true then response.write "checked"%>>
	<%=s_lang_0055 %>
	</td></tr>
	<tr><td colspan="4" bgcolor="white" class="st_r" style="border-top:1px solid #A5B6C8;">
	<input type="checkbox" name="Enable_LDAP_User_Edit_OU" id="Enable_LDAP_User_Edit_OU" <% if ldap.Enable_LDAP_User_Edit_OU = true then response.write "checked"%>>
	<%=s_lang_0056 %>
	</td></tr>
	<tr><td colspan="4" bgcolor="white" class="st_r" style="border-top:1px solid #A5B6C8;">
	<%=s_lang_0061 %><%=s_lang_mh %>
	<input type="password" name="ldappw" maxlength="64" size="26" class="n_textbox">&nbsp;<font color='#444444'><%=s_lang_0062 %></font>
	</td></tr>
<% end if %>
</table>
	</td></tr>
	<tr><td bgcolor="white" align="left"><br>
<a class='wwm_btnDownload btn_blue' href="javascript:goback();"><< <%=s_lang_return %></a>

<% if ldap.Enable_LDAP_User_Edit = true or isamg = true then %>
	<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
<% end if %>
	</td></tr>
</table>
	<input type="hidden" name="gourl" value="<%=gourl %>">
	<input type="hidden" name="id" value="<%=userid %>">
	<input type="hidden" name="name" value="<%=name %>">
	<input type="hidden" name="domain" value="<%=domain %>">
	<input type="hidden" name="amg" value="<%=amg %>">
</form>
</body>
</html>

<%
set ldap = nothing
%>
