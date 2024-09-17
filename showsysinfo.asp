<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if

Response.Buffer = TRUE
%>

<%
dim ei
set ei = server.createobject("easymail.sysinfo")

dim eu
set eu = Application("em")

'-----------------------------------------
ei.Load

dim wms
set wms = server.createobject("easymail.WebMailSet")
wms.Load
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<SCRIPT LANGUAGE=javascript>
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

function sub() {
	document.f1.submit();
}



var DOM = (document.getElementById) ? 1 : 0;
var NS4 = (document.layers) ? 1 : 0;
var IE4 = 0;
if (document.all)
{
	IE4 = 1;
	DOM = 0;
}

var win = window;   
var n   = 0;

function findIt() {
	if (document.f1.searchstr.value != "")
		findInPage(document.f1.searchstr.value);
}


function findInPage(str) {
var txt, i, found;

if (str == "")
	return false;

if (DOM)
{
	win.find(str, false, true);
	return true;
}

if (NS4) {
	if (!win.find(str))
		while(win.find(str, false, true))
			n++;
	else
		n++;

	if (n == 0)
		alert("未找到指定内容.");
}

if (IE4) {
	txt = win.document.body.createTextRange();

	for (i = 0; i <= n && (found = txt.findText(str)) != false; i++) {
		txt.moveStart("character", 1);
		txt.moveEnd("textedit");
	}

if (found) {
	txt.moveStart("character", -1);
	txt.findText(str);
	txt.select();
	txt.scrollIntoView();
	n++;
}
else {
	if (n > 0) {
		n = 0;
		findInPage(str);
	}
	else
		alert("未找到指定内容.");
	}
}

return false;
}

function writeselect(name, haveNULL) {
	document.write("<select name=\"" + name + "\" class=\"drpdwn\">");

	if (haveNULL == true)
		document.write("<option value=\"\">无</option>");

	document.write(userlist);
}

function window_onload() {
	document.f1.ErrSender.value = "<%=ei.errsender %>";
	document.f1.listSender.value = "<%=ei.listSender %>";
	document.f1.stakeOutTo.value = "<%=ei.stakeOutTo %>";
	document.f1.mailMoveTo.value = "<%=ei.mailMoveTo %>";
	document.f1.CheckIPClass.value = "<%=ei.CheckIPClass %>";
	document.f1.EnableSmtpCheckError2Trash.value = "<%
if ei.EnableSmtpCheckError2Trash = true then
	Response.Write "0"
else
	Response.Write "1"
end if
%>";
	document.f1.DNSExpiresDays.value = "<%=ei.DNSExpiresDays %>";
	document.f1.DisableRelayEmail_Mode.value = "<%=ei.DisableRelayEmail_Mode %>";
	document.f1.Collection_SysError_Mail_Mode.value = "<%=ei.Collection_SysError_Mail_Mode %>";

	document.f1.save1.disabled = false;
	document.f1.save2.disabled = false;
}
//-->
</SCRIPT>


<BODY LANGUAGE=javascript onload="return window_onload()">
<br>
<FORM ACTION="savesysinfo.asp" METHOD="POST" NAME="f1">
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="#ffffff">
      <td colspan="3" align="center"><br>
	<table width="100%"><tr><td align="left">
<input type="text" name="searchstr" class="textbox" size="20">
<input type="button" value="页内查找" onclick="javascript:findIt();" class="sbttn">
	</td><td align="right">
		<input type="button" value=" 保存 " name="save1" onclick="javascript:sub()" class="Bsbttn" disabled>&nbsp;&nbsp;
		<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td></tr>
	</table><br>
	</td>
    </tr>
    <tr>
      <td width="55%" height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-top:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left">各类型系统邮件的发送人 (错误邮件的处理人)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-top:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;'>
<script>writeselect("ErrSender", false)</script>
        </td>
    </tr>
	</table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">系统以及Web邮件的缺省字符集</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="System_Mail_CharSet" class="textbox" size="12" maxlength="30" value="<%=ei.System_Mail_CharSet %>">
        </td>
    </tr>
	</table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">普通邮件的最大发送次数</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="MaxSendNum" class="textbox" size="10" maxlength="5" value="<%= ei.MaxSendNum %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">错误回复邮件的最大发送次数</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="ErrMaxSendNum" class="textbox" size="10" maxlength="5" value="<%= ei.ErrMaxSendNum %>">
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">允许对信任邮件地址发送退信</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="Enable_ErrBackToTrustMail" <% if ei.Enable_ErrBackToTrustMail = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">允许对外部邮址发送退信</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableErrBackToOut" <% if ei.EnableErrBackToOut = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">允许对发往本系统不存在帐号的外部邮址退信</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableErrBackToOutForLocalNoSuchUser" <% if ei.EnableErrBackToOutForLocalNoSuchUser = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">发送重试间隔时间</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="Delivery_Retry_Interval" class="textbox" size="7" maxlength="5" value="<%=ei.Delivery_Retry_Interval %>">&nbsp;秒
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">发送超时时间</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="DeliveryTimeout" class="textbox" size="5" maxlength="3" value="<%=ei.DeliveryTimeout %>">&nbsp;秒
        </td>
    </tr>
	</table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">外发邮件时Helo命令后的内容</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="HELO_STRING" class="textbox" size="20" maxlength="128" value="<%=ei.HELO_STRING %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">限定邮件发送时的最大长度</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="restrictSmtpMailSize" <% if ei.restrictSmtpMailSize = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">限定邮件接收时的最大长度</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="restrictPop3MailSize" <% if ei.restrictPop3MailSize = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">允许发送的邮件最大长度(兆)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="smtpMailMaxSize" class="textbox" size="10" maxlength="3" value="<%=ei.smtpMailMaxSize %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">允许接收的邮件最大长度(兆)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="pop3MailMaxSize" class="textbox" size="10" maxlength="3" value="<%=ei.pop3MailMaxSize %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">缺省邮箱大小(兆)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="defaultMailBoxSize" class="textbox" size="10" maxlength="5" value="<%=ei.defaultMailBoxSize %>">
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">邮箱最多邮件数(10 - 9999封)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="defaultMailsNumber" class="textbox" size="10" maxlength="4" value="<%=ei.defaultMailsNumber %>">
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">邮件抄送的最大收件人数 (1 - 999)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="MaxRecipients" class="textbox" size="10" maxlength="3" value="<%=ei.MaxRecipients %>">
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">Web下发送邮件的最大接收地址数 (1 - 999)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="Web_Max_Recipients" class="textbox" size="10" maxlength="3" value="<%=ei.Web_Max_Recipients %>">
        </td>
    </tr>
	</table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">只允许本系统用户发信 (<font color="#FF3333">建议选中</font>)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="onlyMailFromSystem" <% if ei.onlyMailFromSystem = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">只允许发信给本系统用户 (<font color="#FF3333">一般不选中</font>)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="onlyRcptToSystem" <% if ei.onlyRcptToSystem = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">只验证本地收件人名, 不验证所在域名</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="recvByName" <% if ei.recvByName = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用SMTP发信身份认证功能 (<font color="#FF3333">建议选中</font>)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="useAuth" <% if ei.useAuth = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用接收认证功能</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableIngoingAuth" <% if ei.EnableIngoingAuth = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left"><%=s_lang_0029 %> (<font color="#FF3333"><%=s_lang_recom %></font>)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableOpenRelay" <% if ei.enableOpenRelay = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left"><%=s_lang_0028 %>&nbsp;&nbsp;&nbsp;<a href="authtrustip.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableAuthTrustIP" <% if ei.EnableAuthTrustIP = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">允许 MAIL FROM 命令中出现空地址 (MAIL FROM: <>)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableMailFromNULL" <% if ei.EnableMailFromNULL = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left"><%=s_lang_0073 %></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableResponse_SMTP_NoUser" <% if ei.EnableResponse_SMTP_NoUser = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left"><%=s_lang_0074 %></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableReceiveOutMail" <% if ei.EnableReceiveOutMail = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left"><%=s_lang_0075 %></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableAutoReply" <% if ei.EnableAutoReply = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">只允许自动转发到系统内帐号</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="Enable_OnlyAutoForwardToLocal" <% if ei.Enable_OnlyAutoForwardToLocal = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
<a name="handpoint2"></a>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用外发地址定向解析功能&nbsp;&nbsp;&nbsp;<a href="handpoint2.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableHandPoint2" <% if ei.enableHandPoint2 = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">当DNS查询MX记录失败, 从DNS根服务器查询</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="DNS_can_by_ROOT" <% if ei.DNS_can_by_ROOT = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">允许使用本地TCP/IP进行DNS查询</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableLocalNetwork" <% if ei.EnableLocalNetwork = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">DNS缓冲保存时间:</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <select name="DNSExpiresDays" class="drpdwn">
<%
i = 1
do while i < 100
	Response.Write "<option value='" & i & "'>" & i & "天</option>"

	i = i + 1
loop
%>
        </select>
        </td>
    </tr>
    <tr>
<a name="killhelo"></a>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">拒绝 HELO/EHLO 主机名设置&nbsp;&nbsp;&nbsp;<a href="killhelo.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableKillHeloDomain" <% if ei.EnableKillHeloDomain = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
<a name="relayserver"></a>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">允许在投递失败后, 使用中继服务器转发邮件&nbsp;&nbsp;&nbsp;<a href="relayserver.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableRelayServerSend" <% if ei.EnableRelayServerSend = true then response.write "checked"%>>
        </td>
    </tr>
	</table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用SMTP服务</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="useSmtp" <% if ei.useSmtp = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">SMTP服务端口</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="smtpport" class="textbox" size="10" maxlength="5" value="<%=ei.smtpport %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用SMTP加密传输服务(SSL)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="useSSL_SMTP" <% if ei.useSSL_SMTP = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">SMTP加密传输服务(SSL)端口</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="sslsmtpport" class="textbox" size="10" maxlength="5" value="<%=ei.sslsmtpport %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用POP3服务</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="usePOP3" <% if ei.usePOP3 = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">POP3服务端口</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="pop3port" class="textbox" size="10" maxlength="5" value="<%=ei.pop3port %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用POP3加密传输服务(SSL)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="useSSL_POP3" <% if ei.useSSL_POP3 = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">POP3加密传输服务(SSL)端口</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="sslpop3port" class="textbox" size="10" maxlength="5" value="<%=ei.sslpop3port %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用IMAP4服务</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="use_IMAP4" <% if ei.use_IMAP4 = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">IMAP4服务端口</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="imap4port" class="textbox" size="10" maxlength="5" value="<%=ei.imap4port %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用IMAP4加密传输服务(SSL)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="useSSL_IMAP4" <% if ei.useSSL_IMAP4 = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">IMAP4加密传输服务(SSL)端口</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="sslimap4port" class="textbox" size="10" maxlength="5" value="<%=ei.sslimap4port %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用DayTime服务</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="useDayTime" <% if ei.useDayTime = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">DayTime服务端口</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="daytimeport" class="textbox" size="10" maxlength="5" value="<%=ei.daytimeport %>">
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">加密通讯协议 (默认为: SSL 2.0 或 3.0)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <select name="SSL_Mode" class="drpdwn">
<%
if ei.SSL_Mode = 1 then
	response.write "<option value='1' selected>SSL 2.0</option>"
	response.write "<option value='2'>SSL 3.0</option>"
	response.write "<option value='3'>SSL 2.0 或 3.0</option>"
	response.write "<option value='4'>TLS 1.0</option>"
elseif ei.SSL_Mode = 2 then
	response.write "<option value='1'>SSL 2.0</option>"
	response.write "<option value='2' selected>SSL 3.0</option>"
	response.write "<option value='3'>SSL 2.0 或 3.0</option>"
	response.write "<option value='4'>TLS 1.0</option>"
elseif ei.SSL_Mode = 3 then
	response.write "<option value='1'>SSL 2.0</option>"
	response.write "<option value='2'>SSL 3.0</option>"
	response.write "<option value='3' selected>SSL 2.0 或 3.0</option>"
	response.write "<option value='4'>TLS 1.0</option>"
elseif ei.SSL_Mode = 4 then
	response.write "<option value='1'>SSL 2.0</option>"
	response.write "<option value='2'>SSL 3.0</option>"
	response.write "<option value='3'>SSL 2.0 或 3.0</option>"
	response.write "<option value='4' selected>TLS 1.0</option>"
end if
%>
        </select>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">最大工作进程数</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="MaxConnNum" class="textbox" size="5" maxlength="4" value="<%=ei.MaxConnNum %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">限制SMTP服务范围 [慎用] (<font color="#FF3333">一般不选中</font>)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableLimitSmtp" <% if ei.EnableLimitSmtp = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">限制POP3服务范围 [慎用] (<font color="#FF3333">一般不选中</font>)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableLimitPop3" <% if ei.EnableLimitPop3 = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
<a name="accessip"></a>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用信任IP(不受限IP地址或IP段)功能&nbsp;&nbsp;&nbsp;<a href="accessip.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableAlwaryCanAccessIP" <% if ei.EnableAlwaryCanAccessIP = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
<a name="trustuser"></a>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用信任帐号功能&nbsp;&nbsp;&nbsp;<a href="trustuser.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableTrustUser" <% if ei.EnableTrustUser = true then response.write "checked"%>>
        </td>
    </tr>
	</table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">POP3代理下载的休眠时间(分钟)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="POP3DownSleepTime" class="textbox" size="10" maxlength="3" value="<%=ei.POP3DownSleepTime %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">ISP代理下载的休眠时间(分钟)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="IspDownSleepTime" class="textbox" size="10" maxlength="3" value="<%=ei.IspDownSleepTime %>">
        </td>
    </tr>
	</table><br>
	<a name="showmisp"></a>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用ISP邮件接收功能&nbsp;&nbsp;&nbsp;<a href="showmisp.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="canUseIsp" <% if ei.canUseIsp = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">ISP接收的服务器名</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="ispServerName" class="textbox" size="20" maxlength="64" value="<%=ei.ispServerName %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">ISP接收的端口号</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="ispPop3Port" class="textbox" size="10" maxlength="5" value="<%=ei.ispPop3Port %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">ISP接收的用户名</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="ispUserName" class="textbox" size="20" maxlength="64" value="<%=ei.ispUserName %>">
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">ISP接收的密码</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="password" name="ispPassword" class="textbox" size="8" maxlength="64" value="<%=ei.ispPassword %>">
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">ISP邮件的接收处理方式</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <select name="manageisp" class="drpdwn">
<%
if ei.manageisp = 0 then
	response.write "<option value='0' selected>接收并交由管理员处理</option>"
	response.write "<option value='1'>不下载此邮件</option>"
else
	response.write "<option value='0'>接收并交由管理员处理</option>"
	response.write "<option value='1' selected>不下载此邮件</option>"
end if
%>
        </select>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用ISP邮件自动分发功能</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="useDistributeISPMail" <% if ei.useDistributeISPMail = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">选择ISP邮件的自动分发处理列表</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <select name="distributeISPMailinglist" class="drpdwn">
        <option value=''>分发给所有人</option>
<%
dim ml
set ml = server.createobject("easymail.mailinglist")

ml.Load ""

i = 0
allnum = ml.MailingListCount 

do while i < allnum
	mlname = ml.GetMailingListNameByIndex(i)

	if mlname = ei.distributeISPMailinglist then
		response.write "<option value='" & mlname & "' selected>" & server.htmlencode(mlname) & "</option>"
	else
		response.write "<option value='" & mlname & "'>" & server.htmlencode(mlname) & "</option>"
	end if

	mlname = NULL

	i = i + 1
loop

set ml = nothing
%>
        </select>&nbsp;(需启用邮件列表功能)
        </td>
    </tr>
	</table><br>
	<a name="browmailinglist"></a>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用邮件列表功能&nbsp;&nbsp;&nbsp;<a href="browmailinglist.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="useMailList" <% if ei.useMailList = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">DNS地址 (修改后, 需要重启邮件服务程序才能生效)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="DNS" class="textbox" size="30" maxlength="64" value="<%=ei.DNS %>">
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">客户端登录时需要验证admin密码</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="useLogonPass" <% if ei.useLogonPass = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">Web下允许发送的附件总长度(1 - 9,999,999K)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="webMailMaxLen" class="textbox" size="10" maxlength="7" value="<%=ei.webMailMaxLen %>">&nbsp;K
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">时区设置</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="timeZone" class="textbox" size="5" maxlength="5" value="<%=ei.timeZone %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用短消息数据存储功能 (一般不选中)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableSMS" <% if ei.EnableSMS = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
	<a name="keywords"></a> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用主题过滤功能&nbsp;&nbsp;&nbsp;<a href="keywords.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableKeywordFilter" <% if ei.EnableKeywordFilter = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">申请邮箱时填写的注册信息文件的保留天数</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="TempRegInfoKeepDays" class="textbox" size="10" maxlength="3" value="<%=ei.TempRegInfoKeepDays %>">
        </td>
    </tr>
	</table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr> 
<a name="groupmail"></a>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用邮件群发功能 (一般不选中)&nbsp;&nbsp;&nbsp;<a href="groupmail.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="useList" <% if ei.useList = true then response.write "checked"%>>&nbsp;&nbsp;(在发送群发邮件后请立即禁用此项)
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">群发邮件发送人</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
<script>writeselect("listSender", false)</script>
        </td>
    </tr>
	</table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr> 
<a name="showkill"></a>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用黑名单功能&nbsp;&nbsp;&nbsp;<a href="showkill.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="startKill" <% if ei.startKill = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">相关事件保存到日志中</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="Enable_KillSaveLog" <% if ei.Enable_KillSaveLog = true then response.write "checked"%>>
        </td>
    </tr>
	</table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">备份方式 (一般设为"增量式备份")</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <select name="backupmode" class="drpdwn">
<%
if ei.backupmode <> 2 then
	response.write "<option value='1' selected>增量式备份</option>"
	response.write "<option value='2'>不备份</option>"
else
	response.write "<option value='1'>增量式备份</option>"
	response.write "<option value='2' selected>不备份</option>"
end if
%>
        </select>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">备份间隔天数</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="backupdate" class="textbox" size="10" maxlength="3" value="<%=ei.backupdate %>">
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">系统每天自动备份的时间</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <select name="sysBackupTime" class="drpdwn">
<%
i = 0
do while i < 24
	if i < 10 then
		showhour = "0" & i
	else
		showhour = i
	end if

	if i = ei.sysBackupTime then
		response.write "<option value='" & i & "' selected>" & showhour & ":00</option>"
	else
		response.write "<option value='" & i & "'>" & showhour & ":00</option>"
	end if

	i = i + 1
loop
%>
        </select>
        </td>
    </tr>
	</table><br>
<a name="showstakeout"></a>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用邮件监控功能&nbsp;&nbsp;&nbsp;<a href="showstakeout.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="useStakeOut" <% if ei.useStakeOut = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">监控邮件处理人(用来接收所有监控邮件的帐号)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
<script>writeselect("stakeOutTo", false)</script>
        </td>
    </tr>
    <tr> 
<a name="domainMonitor"></a>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用域邮件监控功能&nbsp;&nbsp;&nbsp;<a href="show_dm_domain.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableDomainMonitor" <% if ei.enableDomainMonitor = true then response.write "checked"%>>
        </td>
    </tr>
	</table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用邮件自动清理功能</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="useAutoMailClean" <% if ei.useAutoMailClean = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">超过多少天的邮件将被自动转移(10 - 9999天)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="mailMoveDays" class="textbox" size="10" maxlength="4" value="<%=ei.mailMoveDays %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">此邮件将被转移至</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
<script>writeselect("mailMoveTo", false)</script>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">超过多少天的邮件将被自动删除(10 - 9999天)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="mailDeleteDays" class="textbox" size="10" maxlength="4" value="<%=ei.mailDeleteDays %>">
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">邮件自动清理时包含定期帐号</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="cleanMailIncludeExpiresUser" <% if ei.cleanMailIncludeExpiresUser = true then response.write "checked"%>>
        </td>
    </tr>
	</table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用用户自动清理功能</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="useAutoUserClean" <% if ei.useAutoUserClean = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">超过多少天未登录系统的用户将被禁用帐号(1 - 9999天)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="forbidUserDays" class="textbox" size="10" maxlength="4" value="<%=ei.forbidUserDays %>">
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">帐号禁用多少天后用户将被删除(1 - 9999天)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="deleteUserDays" class="textbox" size="10" maxlength="4" value="<%=ei.deleteUserDays %>">
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">用户自动清理时包含定期帐号</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="cleanAccoutIncludeExpiresUser" <% if ei.cleanAccoutIncludeExpiresUser = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用帐号到期警告功能</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableUserExpAffirm" <% if ei.enableUserExpAffirm = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">帐号到期前几天开始警告(1 - 999天)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="daysUserExpAffirm" class="textbox" size="5" maxlength="3" value="<%=ei.daysUserExpAffirm %>">
        </td>
    </tr>
	</table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用密码保护功能</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnablePreHacker" <% if ei.EnablePreHacker = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">进入密码保护后休眠的时间(1 - 999分钟)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="bWaitMinute" class="textbox" size="5" maxlength="3" value="<%=ei.bWaitMinute %>">&nbsp;分钟
        </td>
    </tr>
    </table><br>
<a name="catchall"></a>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用邮件 Catch All 功能&nbsp;&nbsp;&nbsp;<a href="show_dca_domain.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableCatchAll" <% if ei.enableCatchAll = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">允许 Catch All 邮件转送到外部邮箱</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableCatchToOut" <% if ei.enableCatchToOut = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">Catch All 邮件后发送错误回复邮件</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="catchAllNeedBack" <% if ei.catchAllNeedBack = true then response.write "checked"%>>
        </td>
    </tr>
    </table><br>
<a name="killattack"></a>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用连接攻击保护功能&nbsp;&nbsp;&nbsp;<a href="showkillattack.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableKillAttacker" <% if ei.enableKillAttacker = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用连接攻击自动识别保护功能</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableAutoKillAttacker" <% if ei.enableAutoKillAttacker = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">界定是否为连接攻击行为的5分钟内最大连接数量</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="autoKillAttackerConnectMaxNumber" class="textbox" size="5" maxlength="4" value="<%=ei.autoKillAttackerConnectMaxNumber %>">
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">限制可能是攻击连接的接入概率为</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <select name="autoKillAttackerConnectRate" class="drpdwn">
<%
i = 0
do while i <= 100
	if i = ei.autoKillAttackerConnectRate then
		response.write "<option value='" & i & "' selected>" & i & "%</option>"
	else
		response.write "<option value='" & i & "'>" & i & "%</option>"
	end if

	i = i + 1
loop
%>
        </select>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">限制接入概率持续的时间为</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="autoKillAttackerExpiresMinute" class="textbox" size="5" maxlength="4" value="<%=ei.autoKillAttackerExpiresMinute %>">&nbsp;分钟&nbsp;&nbsp;(注意: 0为无限长)
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">允许应用于IMAP4服务</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableKillAttacker_With_IMAP4" <% if ei.enableKillAttacker_With_IMAP4 = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">相关事件保存到日志中</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="Enable_AttackerSaveLog" <% if ei.Enable_AttackerSaveLog = true then response.write "checked"%>>
        </td>
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用日志记录功能</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="LogSave" <% if ei.LogSave = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用日志自动删除功能</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableLogAutoRemove" <% if ei.enableLogAutoRemove = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">自动删除几天以前的日志</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="logAutoRemoveDay" class="textbox" size="5" maxlength="3" value="<%=ei.logAutoRemoveDay %>">&nbsp;天
        </td>
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">限定WebMail管理员的接入IP地址(或IP段)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableWebAdminIPLimit" <% if ei.enableWebAdminIPLimit = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">WebMail管理员的接入IP地址(或IP段)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="webAdminIP" class="textbox" maxlength="30" value="<%=ei.webAdminIP %>">&nbsp;(支持通配符输入)
        </td>
    </tr>
    </table><br>
	<a name="creditdomains"></a> 
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用接收邮件域控制功能 (<font color="#FF3333">一般不选中</font>)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableRec_Cortrol" <% if ei.enableRec_Cortrol = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">允许接收信任域发来的邮件&nbsp;&nbsp;&nbsp;<a href="creditdomains.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableRec_BelieveDomains" <% if ei.enableRec_BelieveDomains = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">允许接收已发域(曾经收到过本系统所发邮件)的邮件</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableRec_SendDomains" <% if ei.enableRec_SendDomains = true then response.write "checked"%>>
        </td>
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用依据投诉信息过滤垃圾邮件功能 (建议选中)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enable_IndictSpam" <% if ei.enable_IndictSpam = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用非垃圾邮件发送方确认功能 (建议选中)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="Enable_NoSpam_Affirm" <% if ei.Enable_NoSpam_Affirm = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
	<a name="AttachmentExName_Filter"></a> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用邮件附件名称过滤功能&nbsp;&nbsp;&nbsp;<a href="exnamefilter.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enable_AttachmentExName_Filter" <% if ei.enable_AttachmentExName_Filter = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用邮件内容过滤功能&nbsp;&nbsp;&nbsp;<a href="systemfilter.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableSystemFilter" <% if ei.enableSystemFilter = true then response.write "checked"%>>
        </td>
    </tr>
    <tr bgcolor="#ffffff">
	<a name="systemfilter"></a> 
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用邮件防病毒功能 (仅限企业版本有效)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableScanVirus" <% if ei.enableScanVirus = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">发现病毒邮件时 (仅限企业版本有效)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
		<select name="NotifyModeForVirus" class="drpdwn">
<%
if ei.NotifyModeForVirus = 0 then
	Response.Write "<option value='0' selected>不通知发件人</option>"
	Response.Write "<option value='1'>通知内部发件人</option>"
	Response.Write "<option value='2'>通知发件人</option>"
elseif ei.NotifyModeForVirus = 2 then
	Response.Write "<option value='0'>不通知发件人</option>"
	Response.Write "<option value='1'>通知内部发件人</option>"
	Response.Write "<option value='2' selected>通知发件人</option>"
else
	Response.Write "<option value='0'>不通知发件人</option>"
	Response.Write "<option value='1' selected>通知内部发件人</option>"
	Response.Write "<option value='2'>通知发件人</option>"
end if
%>
		</select>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用快速查毒功能 (仅限企业版本有效)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableSpeedScanVirus" <% if ei.enableSpeedScanVirus = true then response.write "checked"%>>
        </td>
    </tr>
    </table><br>
	<a name="collmail"></a> 
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用邮件采集功能&nbsp;&nbsp;&nbsp;<a href="systemcollmail.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableSystemCollectionMail" <% if ei.EnableSystemCollectionMail = true then response.write "checked"%>>
        </td>
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用用户外发邮件自动限制功能<br>(使用此功能前, 必须要启用SMTP发信认证功能)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableAutoStopUserOutGoing" <% if ei.enableAutoStopUserOutGoing = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">界定是否外发邮件过多的1小时最大外发邮件数量</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="autoStopUserOutGoingMaxNumber" class="textbox" size="5" maxlength="4" value="<%=ei.autoStopUserOutGoingMaxNumber %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">暂停其外发邮件功能的持续时间为</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="autoStopUserOutGoingExpiresMinute" class="textbox" size="5" maxlength="4" value="<%=ei.autoStopUserOutGoingExpiresMinute %>">&nbsp;分钟
        </td>
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left"><%=s_lang_0003 %>&nbsp;&nbsp;&nbsp;<a href="greylisting.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableGreylisting" <% if ei.EnableGreylisting = true then response.write "checked"%>>
        </td>
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left"><%=s_lang_0025 %>&nbsp;&nbsp;&nbsp;<a href="trustemail.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableMailFromTrustEmail" <% if ei.EnableMailFromTrustEmail = true then response.write "checked"%>>
        </td>
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用SMTP域名验证功能</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableSmtpDomainCheck" <% if ei.EnableSmtpDomainCheck = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">检查发送者邮件地址中域名的 A、MX、SPF 记录</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableCheckMailFromIP" <% if ei.EnableCheckMailFromIP = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left"><%=s_lang_0030 %></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableCheckMailHeader" <% if ei.EnableCheckMailHeader = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">使用 DNS 检查发送者邮件地址中域名的有效性</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableCheckMailFromDomainIsGood" <% if ei.EnableCheckMailFromDomainIsGood = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">检查 HELO/EHLO 主机名的 A、MX、SPF 记录</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableCheckHeloIP" <% if ei.EnableCheckHeloIP = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">检查主机名失败后, 检查发送者邮件地址中域名的 A、MX、SPF 记录</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableCheckMailFromIP_WhenCheckHeloIP_Error" <% if ei.EnableCheckMailFromIP_WhenCheckHeloIP_Error = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">IP地址匹配方式:</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <select name="CheckIPClass" class="drpdwn">
		<option value="0">A Class</option>
		<option value="1">B Class</option>
		<option value="2" selected>C Class</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">未通过检查邮件的处理方式:</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <select name="EnableSmtpCheckError2Trash" class="drpdwn">
		<option value="0" selected>放入垃圾箱</option>
		<option value="1">拒收</option>
        </select>
        </td>
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left">防止利用系统内帐号进行中继发送</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="DisableRelayEmail" <% if ei.DisableRelayEmail = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">处理方式:</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <select name="DisableRelayEmail_Mode" class="drpdwn">
		<option value="0" selected>拒收</option>
		<option value="1">修改邮件头</option>
        </select>
        </td>
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left"><%=s_lang_0159 %></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="Enable_MailRecall" <% if ei.Enable_MailRecall = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left"><%=s_lang_0160 %></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="MailRecall_SaveDays" class="textbox" size="3" maxlength="2" value="<%=ei.MailRecall_SaveDays %>">
        </td>
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用垃圾箱邮件统计信功能</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="Enable_System_TrashMsg" <% if ei.Enable_System_TrashMsg = true then response.write "checked"%>>
        </td>
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left"><%=s_lang_0205 %></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="Enable_SendOutMonitor" <% if ei.Enable_SendOutMonitor = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left"><%=s_lang_0206 %></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="SendOutMonitor_SaveDays" class="textbox" size="3" maxlength="2" value="<%=ei.SendOutMonitor_SaveDays %>">
        </td>
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left"><%=s_lang_0207 %></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="Enable_SendOut_Auto_Monitor" <% if ei.Enable_SendOut_Auto_Monitor = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left"><%=s_lang_0208 %></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="Auto_Monitor_Start_Max_SendNum" class="textbox" size="5" maxlength="3" value="<%=ei.Auto_Monitor_Start_Max_SendNum %>">
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left"><%=s_lang_0209 %></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="Auto_Monitor_Keep_Days" class="textbox" size="5" maxlength="2" value="<%=ei.Auto_Monitor_Keep_Days %>">
        </td>
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left"><%=s_lang_0297 %></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="ZATT_Is_Enable" <% if ei.ZATT_Is_Enable = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left"><%=s_lang_0298 %></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="ZATT_URL" size="50" class="textbox" value="<%=ei.ZATT_URL %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left"><%=s_lang_0299 %></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="ZATT_Validity_Days" class="textbox" size="4" maxlength="3" value="<%=ei.ZATT_Validity_Days %>">
        </td>
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left"><%=s_lang_0316 %></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableVerification" <% if ei.EnableVerification = true then response.write "checked"%>>
        </td>
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left">启用文件归档功能</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableArchive" <% if wms.EnableArchive = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">文件归档最大限额</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="Archive_MaxNumber" class="textbox" size="10" maxlength="8" value="<%=wms.Archive_MaxNumber %>">
        </td>
    </tr>
    </table><br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">收集系统错误退信的类型:</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <select name="Collection_SysError_Mail_Mode" class="drpdwn">
		<option value="0" selected>不收集</option>
		<option value="1">全部收集</option>
		<option value="2">只收集退往系统内的邮件</option>
		<option value="3">只收集退往系统外的邮件</option>
        </select>
        </td>
    </tr>
	<tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_2 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left">收集系统错误退信的邮件地址:</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
		<input type="text" name="Collection_SysError_Mail_To" class="textbox" size="30" maxlength="64" value="<%=ei.Collection_SysError_Mail_To %>">
		</td>
	</tr>
    <tr bgcolor="#ffffff">
      <td colspan="3" align="center"><br>
	<table width="100%"><tr>
	<td align="right">
		<input type="button" value=" 保存 " name="save2" onclick="javascript:sub()" class="Bsbttn" disabled>&nbsp;&nbsp;
		<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td></tr>
	</table>
      </td>
    </tr>
  </table>
  </FORM>
  <br>
</BODY>
</HTML>

<%
set ei = nothing
set eu = nothing
set wms = nothing
%>
