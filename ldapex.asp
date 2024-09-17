<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if

name = trim(request("exnm"))
mode = trim(request("mode"))
exldap = trim(request("exldap"))
post_uid = trim(request("uid"))

isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if


if mode = "3" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim sysinfo
	set sysinfo = server.createobject("easymail.sysinfo")
	sysinfo.Load
	needSave = false

	if trim(request("EnableLDAPServer")) <> "" then
		if sysinfo.EnableLDAPServer = false then
			sysinfo.EnableLDAPServer = true
			needSave = true
		end if
	else
		if sysinfo.EnableLDAPServer = true then
			sysinfo.EnableLDAPServer = false
			needSave = true
		end if
	end if

	if needSave = true then
		sysinfo.Save
	end if

	set sysinfo = nothing

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("ldapex.asp")
end if


dim ldap
set ldap = server.createobject("easymail.LDAP")


if mode = "2" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ldap.Submit_Del exldap, true, true

	isok = ldap.Extend_Del(exldap)
	set ldap = nothing

	if isok = true then
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("ldapex.asp")
	else
		Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("ldapex.asp")
	end if
end if


isNew = false
if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if Len(exldap) < 1 then
		exldap = post_uid
		isNew = true
	end if

	ldap.Load exldap, true, false
else
	ldap.Load name, true, false
end if


if mode = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	mylog = ""
	isok = false

	if Len(ldap.uid) < 1 then
		ldap.uid = post_uid
	end if

	dim eusers
	set eusers = Application("em")

	if eusers.isUser(ldap.uid) = true then
		set eusers = nothing
		Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("ldapex.asp") & "&errstr=" & Server.URLEncode(s_lang_0068)
	end if

	set eusers = nothing

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

	if ldap.Save = true then
		isok = true
	end if

	if isok = true then
		if isNew = true then
			ldap.Submit_Add exldap, true, true
		else
			ldap.Submit_Modify exldap, true, true
		end if

		mylog = ldap.Log
		if mylog <> "" then
			isok = false
		end if
	end if

	set ldap = nothing

	if isok = true then
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("ldapex.asp?exnm=" & exldap)
	else
		Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("ldapex.asp?exnm=" & exldap) & "&errstr=" & Server.URLEncode(mylog)
	end if
end if

set sysinfo = server.createobject("easymail.sysinfo")
sysinfo.Load
%>

<html>
<head>
<%=s_lang_meta %>
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</head>

<script language="JavaScript">
<!-- 
var old_info = "";
function gosub() {
	if (document.f1.uid.value.indexOf('\"') > -1 || document.f1.uid.value.indexOf('\\') > -1
		|| document.f1.uid.value.indexOf('\/') > -1 || document.f1.uid.value.indexOf(':') > -1
		|| document.f1.uid.value.indexOf('*') > -1 || document.f1.uid.value.indexOf('?') > -1
		|| document.f1.uid.value.indexOf('<') > -1 || document.f1.uid.value.indexOf('>') > -1
		|| document.f1.uid.value.indexOf('=') > -1 || document.f1.uid.value.indexOf(',') > -1
		|| document.f1.uid.value.indexOf('|') > -1 || document.f1.uid.value.indexOf('\t') > -1 || document.f1.uid.value.indexOf('\f') > -1)
	{
		alert("<%=s_lang_0066 %>");
		return false;
	}

	var new_info = document.f1.uid.value + '\f' + document.f1.cn.value + '\f' + document.f1.mail.value + '\f' + document.f1.ou.value + '\f' + document.f1.physicalDeliveryOfficeName.value + '\f' + document.f1.title.value + '\f' + document.f1.mobile.value + '\f' + document.f1.telephoneNumber.value + '\f' + document.f1.facsimileTelephoneNumber.value + '\f' + document.f1.sn.value + '\f' + document.f1.givenName.value + '\f' + document.f1.homePhone.value + '\f' + document.f1.pager.value;

	if (old_info != new_info)
	{
		var isfind = false;
		var i = 0;
		for (i; i < document.f1.exldap.length; i++)
		{
			if (document.f1.exldap[i].value.toLowerCase() == document.f1.uid.value.toLowerCase())
			{
				isfind = true
				break;
			}
		}

		if (isfind == false)
			document.f1.submit();
		else
			alert("<%=s_lang_0068 %>");
	}
}

function window_onload() {
	old_info = document.f1.uid.value + '\f' + document.f1.cn.value + '\f' + document.f1.mail.value + '\f' + document.f1.ou.value + '\f' + document.f1.physicalDeliveryOfficeName.value + '\f' + document.f1.title.value + '\f' + document.f1.mobile.value + '\f' + document.f1.telephoneNumber.value + '\f' + document.f1.facsimileTelephoneNumber.value + '\f' + document.f1.sn.value + '\f' + document.f1.givenName.value + '\f' + document.f1.homePhone.value + '\f' + document.f1.pager.value;
}

function showentads() {
	var remote = null;
	remote = window.open("ldap_oupop.asp?ofm=ou&<%=getGRSN() %>", "", "top=80; left=130; height=330,width=510,scrollbars=yes,resizable=yes,status=no,toolbar=no,menubar=no,location=no");
}

function goback() {
	location.href = "right.asp?<%=getGRSN() %>";
}

function select_exldap() {
	location.href = "ldapex.asp?exnm=" + document.f1.exldap.value + "&<%=getGRSN() %>";
}

function godel() {
	document.f1.mode.value = "2";
	document.f1.submit();
}

function set_ldap() {
	document.f1.mode.value = "3";
	document.f1.submit();
}
// -->
</script>


<body LANGUAGE=javascript onload="return window_onload()">
<form name="f1" method=post action="ldapex.asp">
<input type="hidden" name="mode" id="mode" value="1">
<br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="2%" height="25">&nbsp;</td>
	<td width="30%"><input type="checkbox" id="EnableLDAPServer" name="EnableLDAPServer" value="checkbox"<% if sysinfo.EnableLDAPServer = true then Response.Write " checked"%>><%=s_lang_0071 %></td>
	<td width="45%" nowrap><%=s_lang_0072 %>:&nbsp;<input type=text size="20" class="textbox" value="<%=sysinfo.LDAP_Organization %>" readonly></td>
	<td width="23%"><input type="button" value="<%=s_lang_setting %>" class="sbttn" style="WIDTH: 60px" language=javascript onClick="set_ldap()"<% if IsEnterpriseVersion = false then Response.Write " disabled" %>></td>
    </tr>
  </table><br><br>
<div align="center">
<table cellpadding=0 cellspacing=0 border=0 width="90%">
	<tr bgcolor=999999>
	<td> 
	<table cellpadding=3 cellspacing=1 border=0 width="100%">
	<tr>
	<td align="center" height="26" bgcolor="<%=MY_COLOR_6 %>"><font face="Arial,Helvetica" color="#ffffff"><b><%=s_lang_0064 %></b></font>&nbsp;
	<select name="exldap" id="exldap" class="drpdwn" size="1" LANGUAGE=javascript onchange="select_exldap()">
	<option value=""><%=s_lang_0065 %></option>
<%
allnum = ldap.ExtendCount

i = 0
do while i < allnum
	exfn = ldap.GetExtendFileNameByIndex(i)

	if LCase(name) = LCase(exfn) then
		Response.Write "<option value=""" & server.htmlencode(exfn) & """ selected>" & server.htmlencode(exfn) & "</option>"
	else
		Response.Write "<option value=""" & server.htmlencode(exfn) & """>" & server.htmlencode(exfn) & "</option>"
	end if

	exfn = NULL
	i = i + 1
loop
%>
	</select>
	</td>
	</tr>
	<tr>
	<td bgcolor="<%=MY_COLOR_3 %>"> 
	<table border=0 cellspacing=0 cellpadding=4 width="100%">
	<tr>
	<td><b>UID</b></td><td><input type=text size="26" name="uid" id="uid" class="textbox" value="<%=ldap.uid %>"<% if Len(name) > 0 then Response.Write " readonly" %>></td>
	</tr>
	<tr>
	<td><%=s_lang_0035 %></td><td><input type=text size="26" name="cn" class="textbox" maxlength="64" value="<%=ldap.cn %>"></td>
	<td><%=s_lang_0036 %></td><td><input type=text size="26" name="mail" class="textbox" maxlength="64" value="<%=ldap.mail %>"></td>
	</tr>
	<tr>
	<td><%=s_lang_0043 %></td><td><input type=text name="ou" id="ou" class="textbox" maxlength="64" value="<%=ldap.ou %>" size="20">&nbsp;<input type=button value="..." LANGUAGE=javascript onclick="showentads()" class="sbttn" style="WIDTH: 22px"></td>
	<td><%=s_lang_0044 %></td><td><input type=text size="26" name="physicalDeliveryOfficeName" class="textbox" maxlength="64" value="<%=ldap.physicalDeliveryOfficeName %>"></td>
	</tr>
	<tr>
	<td><%=s_lang_0042 %></td><td><input type=text size="26" name="title" class="textbox" maxlength="64" value="<%=ldap.title %>"></td>
	<td><%=s_lang_0039 %></td><td><input type=text size="26" name="mobile" class="textbox" maxlength="64" value="<%=ldap.mobile %>"></td>
	</tr>
	<tr>
	<td><%=s_lang_0040 %></td><td><input type=text size="26" name="telephoneNumber" class="textbox" maxlength="64" value="<%=ldap.telephoneNumber %>"></td>
	<td><%=s_lang_0041 %></td><td><input type=text size="26" name="facsimileTelephoneNumber" class="textbox" maxlength="64" value="<%=ldap.facsimileTelephoneNumber %>"></td>
	</tr>
	<tr>
	<td><%=s_lang_0045 %></td><td><input type=text size="26" name="sn" class="textbox" maxlength="64" value="<%=ldap.sn %>"></td>
	<td><%=s_lang_0046 %></td><td><input type=text size="26" name="givenName" class="textbox" maxlength="64" value="<%=ldap.givenName %>"></td>
	</tr>
	<tr>
	<td><%=s_lang_0037 %></td><td><input type=text size="26" name="homePhone" class="textbox" maxlength="64" value="<%=ldap.homePhone %>"></td>
	<td><%=s_lang_0038 %></td><td><input type=text size="26" name="pager" class="textbox" maxlength="64" value="<%=ldap.pager %>"></td>
	</tr>
	</table>
	</td>
	</tr>
	</table>
	</td>
	</tr>
</table>
</div>
<br>
<div align="center">
<table width="90%" border="0" align="center" cellspacing="0">
	<tr><td height="30" align="right">
<%
if Len(name) > 0 then
%>
	<input type=button value=" <%=s_lang_del %> " LANGUAGE=javascript onclick="godel()" class="Bsbttn"<% if IsEnterpriseVersion = false then Response.Write " disabled" %>>&nbsp;&nbsp;
<%
end if
%>
	<input type=button value=" <%=s_lang_save %> " LANGUAGE=javascript onclick="gosub()" class="Bsbttn"<% if IsEnterpriseVersion = false then Response.Write " disabled" %>>&nbsp;&nbsp;
	<input type=button value=" <%=s_lang_return  %> " LANGUAGE=javascript onclick="goback()" class="Bsbttn">
	</td></tr>
</table>
</div>
</form>
<br>
</body>
</html>

<%
set sysinfo = nothing
set ldap = nothing
%>
