<!--#include file="passinc.asp" -->

<%
if isadmin() = false and isAccountsAdmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
domain = LCase(trim(request("dmname")))


dim ei
set ei = server.createobject("easymail.domain")
ei.Load

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim mam
	set mam = server.createobject("easymail.AdminManager")
	mam.Load

	Enable_Puny_DBCS_SignName = mam.Enable_Puny_DBCS_SignName

	set mam = nothing
end if

i = 0
allnum = ei.GetCount()

dim showdomainselect
showdomainselect = ""

dim wmethod
set wmethod = server.createobject("easymail.WMethod")

do while i < allnum
	cdomainstr = ei.GetDomain(i)

	if cdomainstr = domain then
		showdomainselect = showdomainselect & "<option value='" & cdomainstr & "' selected>" & wmethod.Puny_To_Domain(cdomainstr) & "</option>"
	else
		showdomainselect = showdomainselect & "<option value='" & cdomainstr & "'>" & wmethod.Puny_To_Domain(cdomainstr) & "</option>"
	end if

	cdomainstr = NULL

	i = i + 1
loop

set wmethod = nothing



'======================================================
dim errstr

if trim(request("errstr")) <> "" then
	errstr = trim(request("errstr"))
else
	errstr = "请输入您想申请的用户名以及密码信息, 并选择域名"
end if

username = LCase(trim(request("username")))
if Enable_Puny_DBCS_SignName = true and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	set wmethod = server.createobject("easymail.WMethod")

	if wmethod.isHaveDBCS(username) = true then
		username = wmethod.Str_To_Puny(username)
	end if

	set wmethod = nothing
end if

pw = trim(request("pw"))
pw1 = trim(request("pw1"))
crmode = trim(request("crmode"))

if pw <> pw1 then
	errstr = "输入的密码不相同"
end if

if username <> "" and domain <> "" then
	if pw = "" or pw1 = "" then
		errstr = "密码不可为空"
	end if
end if



dim comeinadd
comeinadd = false

if Session("Reg") = "next" and username <> "" and domain <> "" and pw <> "" and pw1 <> "" and pw = pw1 and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim isok
	isok = true

	dim isdomain
	isdomain = false

	isdomain = ei.IsDomain(domain)

	if isdomain = false then
		errstr="无效域名"
		isok = false
	end if


	if InStr(username, "!") or InStr(username, """") or InStr(username, "#") or InStr(username, "$") or InStr(username, "%") then
		errstr="用户名中包含非法字符"
		isok = false
	end if

	if InStr(username, "&") or InStr(username, "`") or InStr(username, "(") or InStr(username, ")") or InStr(username, "*") then
		errstr="用户名中包含非法字符"
		isok = false
	end if

	if InStr(username, "+") or InStr(username, ",") or InStr(username, "/") or InStr(username, ":") then
		errstr="用户名中包含非法字符"
		isok = false
	end if

	if InStr(username, ";") or InStr(username, "<") or InStr(username, "=") or InStr(username, ">") or InStr(username, "?") then
		errstr="用户名中包含非法字符"
		isok = false
	end if

	if InStr(username, "@") or InStr(username, "[") or InStr(username, "\") or InStr(username, "]") or InStr(username, "^") then
		errstr="用户名中包含非法字符"
		isok = false
	end if

	if InStr(username, "'") or InStr(username, "{") or InStr(username, "|") or InStr(username, "}") or InStr(username, "~") then
		errstr="用户名中包含非法字符"
		isok = false
	end if

	if InStr(username, " ") or InStr(username, Chr(9)) then
		errstr="用户名中包含非法字符"
		isok = false
	end if



	'-----
	Set easymail = Application("em")
	if crmode = "0" then
		if easymail.isUser(username) = true then
			errstr="系统中已有此用户"
			isok = false
		end if
	else
		if easymail.isUser(username & "@" & domain) = true then
			errstr="系统中已有此用户"
			isok = false
		end if
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
	if crmode = "0" then
		easymail.adduser1 username, pw, domain, "From: " & Request.ServerVariables("REMOTE_ADDR"), CInt(accessmode)
	else
		easymail.adduser1 username & "@" & domain, pw, domain, "From: " & Request.ServerVariables("REMOTE_ADDR"), CInt(accessmode)
	end if
	Set easymail = nothing
%>
<html>
<head>
<title>创建邮箱</title>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<LINK href="images\hwem.css" rel=stylesheet>
</head>

<body>
<br><br>
<div align="center"> 
<form name="fc" action="admincreate.asp">
  <table width="450"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" height="210">
    <tr> 
      <td height="28" align="center" bgcolor="<%=MY_COLOR_2 %>" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>创建成功</b></font></td>
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
                <div align="center" style="font-size:9pt;"><font class="s">邮箱 [ <b><%=username & "@" & domain %></b> ] 创建成功.<br><br>
<%
if crmode = "0" then
%>
                登录用户名是: <font class="s" color="#FF3333"><b><%=username %></b></font></font></div>
<%
else
%>
                登录用户名是: <font color="#FF3333"><b><%=username & "@" & domain %></b></font></font></div>
<%
end if
%>
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
          <input type="submit" value="<< 上一步" class="Bsbttn">&nbsp;&nbsp;
          <input type="button" value=" 退出 " onClick="javascript:location.href='showuser.asp?<%=getGRSN() %>'" class="Bsbttn">
        </div>
      </td>
    </tr>
  </table>
<input type="hidden" name="GRSN" value="<%=createGRSN() %>">
<input type="hidden" name="dmname" value="<%=domain %>">
<input type="hidden" name="accessmode" value="<%=accessmode %>">
<input type="hidden" name="crmode" value="<%=crmode %>">
</form>
</div>
</body>
</html>
<%
else
%>
<html>
<head>
<title>创建邮箱</title>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<LINK href="images\hwem.css" rel=stylesheet>

<SCRIPT LANGUAGE=javascript>
<!--
function checkpw(){
	if (document.fc.username.value == "")
	{
		alert("用户名不可为空");
		document.fc.username.focus();
		return ;
	}

	if (document.fc.pw.value == "")
	{
		alert("密码不可为空");
		document.fc.pw.focus();
		return ;
	}

	if (document.fc.pw1.value == "")
	{
		alert("密码不可为空");
		document.fc.pw1.focus();
		return ;
	}

	if (document.fc.pw.value != document.fc.pw1.value)
	{
		alert("输入的密码不相同");
		document.fc.pw1.focus();
		return ;
	}

	var i = 0
	var mct = false;

	for(i = document.fc.domain.length - 1; i >= 0 ; i--)
	{
		if (document.fc.dmname.value.toLowerCase() == document.fc.domain.options[i].value.toLowerCase())
		{
			mct = true;
			break;
		}
	}

	if (mct == false)
	{
		alert("域名输入有误");
		document.fc.dmname.focus();
	}
	else
		document.fc.submit();
}

function selectdomain_onchange()
{
	document.fc.dmname.value = document.fc.domain.value;
}

function window_onload()
{
	selectdomain_onchange();
}
//-->
</script>
</head>

<BODY LANGUAGE=javascript onload="return window_onload()">
<br><br>
<div align="center"> 
<form name="fc" METHOD="POST" action="admincreate.asp">
  <table width="60%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" height="254">
    <tr bgcolor="<%=MY_COLOR_2 %>"> 
      <td height="28" align="center" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>创建邮箱</b></font></td>
    </tr>
    <tr align="center"> 
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
                <div style="font-size:9pt;">&nbsp;&nbsp;<%=errstr %>.</div>
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
                  <table width="80%" border="0">
                    <tr><td colspan="2" height="30" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
					<input type=radio <% if crmode = "0" then response.write "checked"%> value="0" name="crmode"> 创建普通帐号&nbsp;
					<input type=radio <% if crmode <> "0" then response.write "checked"%> value="1" name="crmode"> 创建含域名帐号<br>
					</td>
                    </tr>
                    <tr> 
                      <td height="30" width="20%" style="font-size:9pt;"><font class="s">用户名:</font></td>
                      <td height="30"> 
                        <input type="text" name="username" value="<%=username %>" maxlength="32" class="textbox">
                      </td>
                    </tr>
                    <tr> 
                      <td height="30" style="font-size:9pt;"><font class="s">域名:</font></td>
                      <td height="30"> 
<input type="text" name="dmname" id="dmname" size="20" class="textbox">
<select name="domain" class="drpdwn" size="1" LANGUAGE=javascript onchange="selectdomain_onchange()">
<%=showdomainselect %>
</select>
					</td>
                    </tr>
                    <tr> 
                      <td height="30" style="font-size:9pt;"><font class="s">密码:</font></td>
                      <td height="30"> 
                        <input type="password" name="pw" maxlength="32" class="textbox">
                      </td>
                    </tr>
                    <tr> 
                      <td height="30" style="font-size:9pt;" nowrap><font class="s">确认密码:</font></td>
                      <td height="30"> 
                        <input type="password" name="pw1" maxlength="32" class="textbox">
                      </td>
                    </tr>
                    <tr> 
                      <td height="30" style="font-size:9pt;" nowrap><font class="s">访问方式:</font></td>
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
          <input type="button" value=" 提交 " onClick="javascript:checkpw();" class="Bsbttn">&nbsp;&nbsp;
          <input type="button" value=" 取消 " onClick="javascript:history.back();" class="Bsbttn">
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
