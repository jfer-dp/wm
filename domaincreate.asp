<!--#include file="passinc.asp" -->

<%
domain = trim(request("domain"))


dim ei
set ei = server.createobject("easymail.domain")
ei.Load

if ei.GetUserManagerDomainCount(Session("wem")) < 1 then
	set ei = nothing
	response.redirect "noadmin.asp"
end if


if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim mam
	set mam = server.createobject("easymail.AdminManager")
	mam.Load

	Enable_Puny_DBCS_SignName = mam.Enable_Puny_DBCS_SignName

	set mam = nothing
end if


i = 0
allnum = ei.GetUserManagerDomainCount(Session("wem"))

dim iscontrol
iscontrol = false

dim wmethod
set wmethod = server.createobject("easymail.WMethod")

dim showdomainselect
showdomainselect = ""

do while i < allnum
	cdomainstr = ei.GetUserManagerDomain(Session("wem"), i)

	if cdomainstr = domain then
		iscontrol = true
		showdomainselect = showdomainselect & "<option value='" & cdomainstr & "' selected>" & wmethod.Puny_To_Domain(cdomainstr) & "</option>"
	else
		showdomainselect = showdomainselect & "<option value='" & cdomainstr & "'>" & wmethod.Puny_To_Domain(cdomainstr) & "</option>"
	end if

	cdomainstr = NULL

	i = i + 1
loop

set wmethod = nothing

if domain <> "" and iscontrol = false then
	set ei = nothing
	response.redirect "noadmin.asp"
end if



'======================================================
dim errstr

if trim(request("errstr")) <> "" then
	errstr = trim(request("errstr"))
else
	errstr = "����������������û����Լ�������Ϣ����ѡ������<br>&nbsp;&nbsp;��ע�⣺ֻ�����뺬�����ʺţ�"
end if

username = trim(request("username"))
if Enable_Puny_DBCS_SignName = true and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	set wmethod = server.createobject("easymail.WMethod")

	if wmethod.isHaveDBCS(username) = true then
		username = wmethod.Str_To_Puny(username)
	end if

	set wmethod = nothing
end if

pw = trim(request("pw"))
pw1 = trim(request("pw1"))

if pw <> pw1 then
	errstr = "��������벻��ͬ"
end if

if username <> "" and domain <> "" then
	if pw = "" or pw1 = "" then
		errstr = "���벻��Ϊ��"
	end if
end if



dim comeinadd
comeinadd = false

if iscontrol = true and Session("Reg") = "next" and username <> "" and domain <> "" and pw <> "" and pw1 <> "" and pw = pw1 and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim isok
	isok = true

	dim isdomain
	isdomain = false

	ei.GetControlMsg domain, isshow, maxuser, manager
	mdn = ei.GetUserNumberInDomain(domain)
	isdomain = ei.IsDomain(domain)

	if mdn >= maxuser then
		errstr="��ǰ���е��û�������"
		isok = false
	end if

	if isdomain = false then
		errstr="��Ч����"
		isok = false
	end if


	if InStr(username, "!") or InStr(username, """") or InStr(username, "#") or InStr(username, "$") or InStr(username, "%") then
		errstr="�û����а����Ƿ��ַ�"
		isok = false
	end if

	if InStr(username, "&") or InStr(username, "`") or InStr(username, "(") or InStr(username, ")") or InStr(username, "*") then
		errstr="�û����а����Ƿ��ַ�"
		isok = false
	end if

	if InStr(username, "+") or InStr(username, ",") or InStr(username, "/") or InStr(username, ":") then
		errstr="�û����а����Ƿ��ַ�"
		isok = false
	end if

	if InStr(username, ";") or InStr(username, "<") or InStr(username, "=") or InStr(username, ">") or InStr(username, "?") then
		errstr="�û����а����Ƿ��ַ�"
		isok = false
	end if

	if InStr(username, "@") or InStr(username, "[") or InStr(username, "\") or InStr(username, "]") or InStr(username, "^") then
		errstr="�û����а����Ƿ��ַ�"
		isok = false
	end if

	if InStr(username, "'") or InStr(username, "{") or InStr(username, "|") or InStr(username, "}") or InStr(username, "~") then
		errstr="�û����а����Ƿ��ַ�"
		isok = false
	end if

	if InStr(username, " ") or InStr(username, Chr(9)) then
		errstr="�û����а����Ƿ��ַ�"
		isok = false
	end if



	'-----
	Set easymail = Application("em")
	if easymail.isUser(username & "@" & domain) = true then
		errstr="ϵͳ�����д��û�"
		isok = false
	end if    
	Set easymail = nothing


	if isok = true then
		comeinadd = true
	end if

end if

Session("Reg") = ""


'------------------------------------------------
if comeinadd = true and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	accessmode = trim(request("accessmode"))

	if IsNumeric(accessmode) = false then
		accessmode = "0"
	end if

	Set easymail = Application("em")
	easymail.adduser1 username & "@" & domain, pw, domain, "From: " & Request.ServerVariables("REMOTE_ADDR"), CInt(accessmode)
	Set easymail = nothing
%>
<html>
<head>
<title>��������</title>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<LINK href="images\hwem.css" rel=stylesheet>
</head>

<body>
<br><br>
<div align="center"> 
<form name="fc" action="domaincreate.asp">
  <table width="450"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" height="210">
    <tr> 
      <td height="28" align="center" bgcolor="<%=MY_COLOR_2 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>�����ɹ�</b></font></td>
    </tr>
    <tr bgcolor="<%=MY_COLOR_3 %>" align="center"> 
      <td><br>
        <div align="center"> 
          <table width="90%" border="0">
            <tr> 
              <td> 
                <div align="center"></div>
              </td>
            </tr>
            <tr> 
              <td height="1"> 
                <div align="center"> 
                  <hr size="1">
                </div>
              </td>
            </tr>
            <tr> 
              <td height="1">
                <div align="center" style="font-size:9pt;"><font class="s">���� [ <b><%=username & "@" & domain %></b> ] �����ɹ�.<br><br>
                ��¼�û�����: <font color="#FF3333"><b><%=username & "@" & domain %></b></font></font></div>
              </td>
            </tr>
            <tr> 
              <td height="1"> 
                <div align="center"> 
                  <hr size="1">
                </div>
              </td>
            </tr>
          </table>
          <br>
        </div>
      </td>
    </tr>
    <tr> 
      <td height="40" bgcolor="#ffffff">
        <div align="right">
          <input type="submit" value="<< ��һ��" class="Bsbttn">&nbsp;&nbsp;
          <input type="button" value=" �˳� " onClick="javascript:location.href='showdomainusers.asp?<%=getGRSN() %>'" class="Bsbttn">
        </div>
      </td>
    </tr>
  </table>
<input type="hidden" name="GRSN" value="<%=createGRSN() %>">
<input type="hidden" name="GRSN" value="<%=createGRSN() %>">
<input type="hidden" name="domain" value="<%=domain %>">
<input type="hidden" name="accessmode" value="<%=accessmode %>">
</form>
</div>
</body>
</html>
<%
else
%>
<html>
<head>
<title>��������</title>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<LINK href="images\hwem.css" rel=stylesheet>

<SCRIPT LANGUAGE=javascript>
<!--
function checkpw(){
	if (document.fc.username.value == "")
	{
		alert("�û�������Ϊ��");
		document.fc.username.focus();
		return ;
	}

	if (document.fc.pw.value == "")
	{
		alert("���벻��Ϊ��");
		document.fc.pw.focus();
		return ;
	}

	if (document.fc.pw1.value == "")
	{
		alert("���벻��Ϊ��");
		document.fc.pw1.focus();
		return ;
	}

	if (document.fc.pw.value != document.fc.pw1.value)
		alert("��������벻��ͬ");
	else
		document.fc.submit();
}
//-->
</script>
</head>

<body>
<br><br>
<div align="center"> 
<form name="fc" METHOD="POST" action="domaincreate.asp">
  <table width="450"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" height="254">
    <tr bgcolor="<%=MY_COLOR_2 %>"> 
      <td height="28" align="center" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>��������</b></font></td>
    </tr>
    <tr bgcolor="<%=MY_COLOR_3 %>" align="center">
      <td><br>
        <div align="center">
          <table width="90%" border="0">
            <tr>
              <td>
                <div align="center"></div>
              </td>
            </tr>
            <tr>
              <td height="1">
                <div align="center">
                  <hr size="1">
                </div>
              </td>
            </tr>
            <tr>
              <td height="1">
                <div style="font-size:9pt;">&nbsp;&nbsp;<%=errstr %></div>
              </td>
            </tr>
            <tr>
              <td height="1">
                <div align="center">
                  <hr size="1">
                </div>
              </td>
            </tr>
            <tr>
              <td height="1">
                <div align="center">
                  <table width="60%" border="0">
                    <tr>
                      <td height="30" style="font-size:9pt;">�û���:</td>
                      <td height="30">
                        <input type="text" name="username" value="<%=username %>" maxlength="32" class="textbox">
                      </td>
                    </tr>
                    <tr>
                      <td height="30" style="font-size:9pt;">����:</td>
                      <td height="30">
<select name="domain" class="drpdwn" size="1">
<%=showdomainselect %>
</select>
					</td>
                    </tr>
                    <tr>
                      <td height="30" style="font-size:9pt;">����:</td>
                      <td height="30">
                        <input type="password" name="pw" maxlength="32" class="textbox">
                      </td>
                    </tr>
                    <tr>
                      <td height="30" style="font-size:9pt;">ȷ������:</td>
                      <td height="30">
                        <input type="password" name="pw1" maxlength="32" class="textbox">
                      </td>
                    </tr>
                    <tr>
                      <td height="30" style="font-size:9pt;">���ʷ�ʽ:</td>
                      <td height="30">
	<select name="accessmode" class="drpdwn" size="1">
<%
amode = 0
if IsNumeric(trim(request("accessmode"))) = true then
	amode = CInt(trim(request("accessmode")))
end if


anum = 0
do while anum < 7
	if amode = anum then
		response.write "<option value=""" & anum & """ selected>" & getaccessmode(anum) & "</option>"
	else
		response.write "<option value=""" & anum & """>" & getaccessmode(anum) & "</option>"
	end if
	anum = anum + 1
loop
%>
	</select>
                      </td>
                    </tr>
                  </table>
                </div>
              </td>
            </tr>
          </table><br>
        </div>
      </td>
    </tr>
    <tr>
      <td height="40" bgcolor="#ffffff" align="right">
          <input type="button" value=" �ύ " onClick="javascript:checkpw();" class="Bsbttn">&nbsp;&nbsp;
          <input type="button" value=" ȡ�� " onClick="javascript:history.back();" class="Bsbttn">
      </td>
    </tr>
  </table>
</form>
</div>
</body>
</html>

<%
	Session("Reg") = "next"
end if
%>

<%
set ei = nothing
%>

<%
function createGRSN()
	Randomize
	createGRSN = Int((9999999 * Rnd) + 1)
end function

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
