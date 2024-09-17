<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim sysinfo
set sysinfo = server.createobject("easymail.sysinfo")
sysinfo.Load

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	Relay_ServerName = trim(request("Relay_ServerName"))
	Relay_ServerPort = trim(request("Relay_ServerPort"))
	Relay_Email = trim(request("Relay_Email"))
	Relay_Accounts = trim(request("Relay_Accounts"))
	Relay_Password = trim(request("Relay_Password"))
	Relay_Pop3Server = trim(request("Relay_Pop3Server"))
	Relay_Pop3Port = trim(request("Relay_Pop3Port"))
	Relay_Maintain = trim(request("Relay_Maintain"))


	if trim(request("Relay_Need_ErrorMsg")) = "" then
		sysinfo.Relay_Need_ErrorMsg = false
	else
		sysinfo.Relay_Need_ErrorMsg = true
	end if

	if trim(request("EnableRelayServerSend")) = "" then
		sysinfo.EnableRelayServerSend = false
	else
		sysinfo.EnableRelayServerSend = true
	end if

	sysinfo.Relay_ServerName = Relay_ServerName

	if Relay_ServerPort <> "" and IsNumeric(Relay_ServerPort) = true then
		sysinfo.Relay_ServerPort = CLng(Relay_ServerPort)
	else
		sysinfo.Relay_ServerPort = 25
	end if

	sysinfo.Relay_Email = Relay_Email

	if trim(request("Relay_NeedAuth")) = "" then
		sysinfo.Relay_NeedAuth = false
	else
		sysinfo.Relay_NeedAuth = true
	end if

	if trim(request("Relay_ReplaceMailFrom")) = "" then
		sysinfo.Relay_ReplaceMailFrom = false
	else
		sysinfo.Relay_ReplaceMailFrom = true
	end if

	if trim(request("Relay_ReplaceFrom")) = "" then
		sysinfo.Relay_ReplaceFrom = false
	else
		sysinfo.Relay_ReplaceFrom = true
	end if

	sysinfo.Relay_Accounts = Relay_Accounts
	sysinfo.Relay_Password = Relay_Password

	sysinfo.Relay_Pop3Server = Relay_Pop3Server

	if Relay_Pop3Port <> "" and IsNumeric(Relay_Pop3Port) = true then
		sysinfo.Relay_Pop3Port = CLng(Relay_Pop3Port)
	else
		sysinfo.Relay_Pop3Port = 110
	end if

	sysinfo.Relay_Maintain = CLng(Relay_Maintain)


	sysinfo.save
	set sysinfo = nothing

	response.redirect "ok.asp?" & getGRSN() & "&gourl=relayserver.asp"
end if
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
function emailisok(emailads)
{
	var mailisok = true;
	var sp = emailads.indexOf("@");
	if (sp == -1)
		mailisok = false;
	else
	{
		sp = emailads.indexOf("@", sp + 1);
		if (sp != -1)
			mailisok = false;
		else
		{
			if (emailads.charAt(0) == '@' || emailads.charAt(emailads.length - 1) == '@')
			{
				mailisok = false;
			}
		}
	}

	if (mailisok == true)
		return true;

	return false;
}

function gosub()
{
	if (document.fm1.Relay_Email.value.length > 0)
	{
		if (emailisok(document.fm1.Relay_Email.value) == false)
		{
			alert("无效的邮件地址.");
			document.fm1.Relay_Email.focus();
			return ;
		}
	}

	document.fm1.submit();
}
//-->
</SCRIPT>

<BODY>
<br><br>
<form method="post" action="relayserver.asp" name="fm1">
  <table width="79%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="1%" height="28">&nbsp;</td>
      <td width="77%"><input type="checkbox" name="EnableRelayServerSend" value="checkbox"<% if sysinfo.EnableRelayServerSend = true then Response.Write " checked"%>>允许在投递失败后, 使用中继服务器转发邮件</td>
      <td width="22%"><font class="s" color="<%=MY_COLOR_4 %>"><b>中继服务器</b></font></td>
    </tr>
  </table><br>
  <table width="80%" border="0" align="center">
    <tr> 
      <td> 
        <div align="center"> 
          <table width="100%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-top:1px <%=MY_COLOR_1 %> solid;">
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;' height="30">
                <div align="right">中继转发成功后需要发送退信&nbsp;:&nbsp;</div>
              </td>
			<td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <input type="checkbox" name="Relay_Need_ErrorMsg" value="checkbox" <% if sysinfo.Relay_Need_ErrorMsg = true then response.write "checked"%>>
			</td>
            </tr>
            <tr> 
              <td width="55%" valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;' height=30> 
                <div align="right">中继服务器地址(建议填写IP地址)&nbsp;:&nbsp;</div>
              </td>
              <td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <input type="text" name="Relay_ServerName" class='textbox' value="<%=sysinfo.Relay_ServerName %>" size="20" maxlength="32">
              </td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;' height=30> 
                <div align="right">中继服务器端口&nbsp;:&nbsp;</div>
              </td>
              <td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <input type="text" name="Relay_ServerPort" class='textbox' value="<%=sysinfo.Relay_ServerPort %>" size="7" maxlength="5">
              </td>
            </tr>
            <tr> 
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;' height=30> 
                <div align="right">邮件地址&nbsp;:&nbsp;</div>
              </td>
              <td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <input type="text" name="Relay_Email" class='textbox' value="<%=sysinfo.Relay_Email %>" size="20" maxlength="32">
              </td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;' height="30">
                <div align="right">中继投递时需要身份认证&nbsp;:&nbsp;</div>
              </td>
			<td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <input type="checkbox" name="Relay_NeedAuth" value="checkbox" <% if sysinfo.Relay_NeedAuth = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;' height="30">
                <div align="right">允许替换SMTP命令中的 Mail From 信息&nbsp;:&nbsp;</div>
              </td>
			<td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <input type="checkbox" name="Relay_ReplaceMailFrom" value="checkbox" <% if sysinfo.Relay_ReplaceMailFrom = true then response.write "checked"%>>
			</td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;' height="30">
                <div align="right">允许替换邮件中的 From 信息&nbsp;:&nbsp;</div>
              </td>
			<td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <input type="checkbox" name="Relay_ReplaceFrom" value="checkbox" <% if sysinfo.Relay_ReplaceFrom = true then response.write "checked"%>>
			</td>
            </tr>
            <tr> 
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;' height=30> 
                <div align="right">用户名&nbsp;:&nbsp;</div>
              </td>
              <td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <input type="text" name="Relay_Accounts" class='textbox' value="<%=sysinfo.Relay_Accounts %>" size="20" maxlength="32">
              </td>
            </tr>
            <tr> 
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;' height=30> 
                <div align="right">密码&nbsp;:&nbsp;</div>
              </td>
              <td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <input type="password" name="Relay_Password" class='textbox' value="<%=sysinfo.Relay_Password %>" size="8" maxlength="64">
              </td>
            </tr>
            <tr> 
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;' height=30>
                <div align="right">POP3服务器地址&nbsp;:&nbsp;</div>
              </td>
              <td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <input type="text" name="Relay_Pop3Server" class='textbox' value="<%=sysinfo.Relay_Pop3Server %>" size="20" maxlength="32">
              </td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;' height=30> 
                <div align="right">POP3服务器端口&nbsp;:&nbsp;</div>
              </td>
              <td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
                <input type="text" name="Relay_Pop3Port" class='textbox' value="<%=sysinfo.Relay_Pop3Port %>" size="7" maxlength="5">
              </td>
            </tr>
            <tr>
              <td valign=center align=right bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;' height=30> 
                <div align="right">中继帐号维护方式&nbsp;:&nbsp;</div>
              </td>
              <td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				<select name="Relay_Maintain" class="drpdwn">
<%
i = sysinfo.Relay_Maintain
%>
<option value='0'<% if i = 0 then Response.Write " selected" %>>不进行维护</option>
<option value='1'<% if i = 1 then Response.Write " selected" %>>定时登录维护</option>
<option value='2'<% if i = 2 then Response.Write " selected" %>>定时删除邮件维护</option>
				</select>
              </td>
            </tr>
            <tr>
			<td valign=center align=right bgcolor=#ffffff height="40" colspan="2"><br>
                <div align="right"> 
				<input type="button" value=" 保存 " onclick="javascript:gosub()" class="Bsbttn">&nbsp;&nbsp;
				<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
                </div>
              </td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
  </table>
</Form>
<br>
</BODY>
</HTML>

<%
set sysinfo = nothing
%>
