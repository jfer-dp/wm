<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.sysinfo")
ei.Load

dim adm
set adm = server.createobject("easymail.adminmsg")
adm.Load

needSet_TrashMsg_Text = false

sp = InStr(1, adm.TrashMsg_Text, "%URL%=http")
if sp > 0 then
	ep = InStr(sp + 10, adm.TrashMsg_Text, Chr(13), 0)
end if

if sp > 0 and ep > 0 then
	sp = sp + 10

	if InStr(1, LCase(Mid(adm.TrashMsg_Text, sp, ep - sp)), "://localhost/") > 0 then
		needSet_TrashMsg_Text = true
	end if
end if


if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ei.ZATT_URL = trim(request("ZATT_URL"))
	ei.DNS = trim(request("DNS"))

	if trim(request("EnableGreylisting")) <> "" then
		ei.EnableGreylisting = true
	else
		ei.EnableGreylisting = false
	end if

	if trim(request("enableOpenRelay")) <> "" then
		ei.enableOpenRelay = true
	else
		ei.enableOpenRelay = false
	end if

	if trim(request("enableRec_Cortrol")) <> "" then
		ei.enableRec_Cortrol = true
	else
		ei.enableRec_Cortrol = false
	end if

	if trim(request("onlyRcptToSystem")) <> "" then
		ei.onlyRcptToSystem = true
	else
		ei.onlyRcptToSystem = false
	end if

	if trim(request("onlyMailFromSystem")) <> "" then
		ei.onlyMailFromSystem = true
	else
		ei.onlyMailFromSystem = false
	end if

	if trim(request("useAuth")) <> "" then
		ei.useAuth = true
	else
		ei.useAuth = false
	end if


	if trim(request("EnableSmtpDomainCheck")) <> "" then
		ei.EnableSmtpDomainCheck = true
	else
		ei.EnableSmtpDomainCheck = false
	end if

	if trim(request("EnableCheckMailFromDomainIsGood")) <> "" then
		ei.EnableCheckMailFromDomainIsGood = true
	else
		ei.EnableCheckMailFromDomainIsGood = false
	end if

	if trim(request("EnableCheckMailFromIP")) <> "" then
		ei.EnableCheckMailFromIP = true
	else
		ei.EnableCheckMailFromIP = false
	end if

	if trim(request("EnableCheckMailHeader")) <> "" then
		ei.EnableCheckMailHeader = true
	else
		ei.EnableCheckMailHeader = false
	end if

	if trim(request("EnableCheckHeloIP")) <> "" then
		ei.EnableCheckHeloIP = true
	else
		ei.EnableCheckHeloIP = false
	end if

	if trim(request("EnableCheckMailFromIP_WhenCheckHeloIP_Error")) <> "" then
		ei.EnableCheckMailFromIP_WhenCheckHeloIP_Error = true
	else
		ei.EnableCheckMailFromIP_WhenCheckHeloIP_Error = false
	end if

	if IsNumeric(trim(request("CheckIPClass"))) = true then
		ei.CheckIPClass = CLng(trim(request("CheckIPClass")))
	end if

	if IsNumeric(trim(request("EnableSmtpCheckError2Trash"))) = true then
		if trim(request("EnableSmtpCheckError2Trash")) = "1" then
			ei.EnableSmtpCheckError2Trash = false
		else
			ei.EnableSmtpCheckError2Trash = true
		end if
	end if

	if trim(request("EnableKeywordFilter")) <> "" then
		ei.EnableKeywordFilter = true
	else
		ei.EnableKeywordFilter = false
	end if

	if trim(request("enable_AttachmentExName_Filter")) <> "" then
		ei.enable_AttachmentExName_Filter = true
	else
		ei.enable_AttachmentExName_Filter = false
	end if

	if trim(request("enableSystemFilter")) <> "" then
		ei.enableSystemFilter = true
	else
		ei.enableSystemFilter = false
	end if

	if trim(request("enableAutoStopUserOutGoing")) <> "" then
		ei.enableAutoStopUserOutGoing = true
	else
		ei.enableAutoStopUserOutGoing = false
	end if

	if IsNumeric(trim(request("autoStopUserOutGoingMaxNumber"))) = true then
		ei.autoStopUserOutGoingMaxNumber = CLng(trim(request("autoStopUserOutGoingMaxNumber")))
	end if

	if IsNumeric(trim(request("autoStopUserOutGoingExpiresMinute"))) = true then
		ei.autoStopUserOutGoingExpiresMinute = CLng(trim(request("autoStopUserOutGoingExpiresMinute")))
	end if

	if trim(request("DisableRelayEmail")) <> "" then
		ei.DisableRelayEmail = true
	else
		ei.DisableRelayEmail = false
	end if

	if IsNumeric(trim(request("DisableRelayEmail_Mode"))) = true then
		ei.DisableRelayEmail_Mode = CLng(trim(request("DisableRelayEmail_Mode")))
	end if

	if trim(request("enableKillAttacker")) <> "" then
		ei.enableKillAttacker = true
	else
		ei.enableKillAttacker = false
	end if

	if trim(request("Enable_WhiteList_IP")) <> "" then
		ei.Enable_WhiteList_IP = true
	else
		ei.Enable_WhiteList_IP = false
	end if

	if trim(request("Enable_WhiteList_Domain")) <> "" then
		ei.Enable_WhiteList_Domain = true
	else
		ei.Enable_WhiteList_Domain = false
	end if

	if trim(request("Enable_RBL")) <> "" then
		ei.Enable_RBL = true
	else
		ei.Enable_RBL = false
	end if

	ei.Save

	adm.TrashMsg_Subject = trim(request("TrashMsg_Subject"))
	adm.TrashMsg_Text = trim(request("TrashMsg_Text"))
	adm.Save

	set ei = nothing
	set adm = nothing

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=impset.asp"
end if
%>

<HTML>
<HEAD>
<%=s_lang_meta %>
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/hwem.css">
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
function gosub() {
	document.f1.submit();
}

function window_onload() {
	document.f1.CheckIPClass.value = "<%=ei.CheckIPClass %>";
	document.f1.EnableSmtpCheckError2Trash.value = "<%
if ei.EnableSmtpCheckError2Trash = true then
	Response.Write "0"
else
	Response.Write "1"
end if
%>";
	document.f1.DisableRelayEmail_Mode.value = "<%=ei.DisableRelayEmail_Mode %>";

	document.f1.save1.disabled = false;
	document.f1.save2.disabled = false;
}
//-->
</SCRIPT>


<BODY LANGUAGE=javascript onload="return window_onload()">
<br>
<FORM ACTION="impset.asp" METHOD="POST" NAME="f1">
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="30%" height="28">&nbsp;</td>
      <td width="70%" align=right>
	<input type="button" value=" 保存 " name="save1" onclick="javascript:gosub()" class="Bsbttn" disabled>&nbsp;&nbsp;
	<input type="button" value=" 退出 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">&nbsp;&nbsp;&nbsp;
      </td>
    </tr>
  </table>
	<br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
		<td height="25" colspan=2 bgcolor="<%=MY_COLOR_3 %>" align="left"> 
		<%=s_lang_0298 %>：<input type="text" name="ZATT_URL" size="70" class="textbox" value="<%=ei.ZATT_URL %>">
		</td>
	</td>
	</tr>
    <tr bgcolor="#FFF8D3">
	<td valign=center colspan=2 height="22" style="border:<%
if ei.ZATT_URL <> "http://localhost/downatt.asp" then
	Response.Write MY_COLOR_1
else
	Response.Write "#ff0000"
end if
%> solid 1px;">
	本项需填写可访问的http地址, 最后必须由 <font color="#FF3333">/downatt.asp</font> 结尾, 比如: http://mail.domain.com/downatt.asp<br>
	<font color="#FF3333">重要</font>: 此项设置不当, 会造成用户所发送的链接式附件无法下载.
	</td>
    </tr>
	<tr bgcolor="#ffffff"><td height="15" colspan=2></td></tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-top:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">DNS地址 (修改后, 需要重启邮件服务程序才能生效)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="DNS" class="textbox" size="30" maxlength="64" value="<%=ei.DNS %>">
        </td>
    </tr>
    <tr bgcolor="#FFF8D3">
	<td valign=center colspan=2 height="22" style="border:<%
if Len(ei.DNS) > 0 then
	Response.Write MY_COLOR_1
else
	Response.Write "#ff0000"
end if
%> solid 1px;">
	DNS地址设置非常重要, 将直接影响对外网发信的成败. 您应设置两个有效并且不相同的DNS地址(不能超过两个), 中间用 , 号分隔.
	</td>
    </tr>
	<tr bgcolor="#ffffff"><td height="15" colspan=2></td></tr>
    <tr>
	<td colspan=2 height="25" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
	<table width="100%" border="0" cellspacing="0">
    <tr>
	<td height="25">
	设置垃圾箱邮件统计信
	</td>
    </tr>
    <tr> 
	<td align="center" height="25" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        主题:&nbsp;&nbsp;<input name="TrashMsg_Subject" type="text" value="<%=adm.TrashMsg_Subject %>" size="50" class='textbox'>
	</td>
    </tr>
    <tr> 
	<td align="center">
		<textarea name="TrashMsg_Text" cols="66" rows="5" class='textarea'><%=server.htmlencode(adm.TrashMsg_Text) %></textarea>
	</td>
    </tr>
    </table>
	</td>
    </tr>
	<tr bgcolor="#FFF8D3">
	<td valign=center colspan=2 height="22" style="border:<%
if needSet_TrashMsg_Text = true then
	Response.Write "#ff0000"
else
	Response.Write MY_COLOR_1
end if
%> solid 1px;">
	您需要修改第一行 <font color="#FF3333">%URL%=</font> 后的内容为邮件系统可以由外部访问的http地址, 并且最后需要由 <font color="#FF3333">/trashmsg.asp</font> 结尾. 比如: http://www.domain.com/mail/trashmsg.asp 或 http://mail.domain.com/trashmsg.asp<br>
	<font color="#FF3333">重要</font>: 此内容设置不当, 会影响用户对垃圾箱邮件进行管理.
	</td>
	</tr>
	<tr bgcolor="#ffffff"><td height="15" colspan=2></td></tr>
<%
if ei.enableOpenRelay = false then
%>
	<tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
		<div align="left">不检查MAIL FROM命令与信头中FROM地址的一致性 (<font color="#FF3333">建议选中</font>)</div>
		</td>
		<td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
		<input type="checkbox" name="enableOpenRelay" <% if ei.enableOpenRelay = true then response.write "checked"%>>
		</td>
	</tr>
	<tr bgcolor="#FFF8D3">
	<td valign=center colspan=2 height="22" style="border:#ff0000 solid 1px;">
	使用检查一致性的功能后, 系统会拒收一些不规范的邮件, 从而造成无法接收一些邮局发来的邮件.
	</td>
	</tr>
	<tr bgcolor="#ffffff"><td height="15" colspan=2></td></tr>
<%
else
	Response.Write "<input type='hidden' name='enableOpenRelay' value='1'>" & Chr(13)
end if

if ei.enableRec_Cortrol = true then
%>
	<tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left">启用接收邮件域控制功能 (<font color="#FF3333">一般不选中</font>)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableRec_Cortrol" <% if ei.enableRec_Cortrol = true then response.write "checked"%>>
        </td>
    </tr>
	<tr bgcolor="#FFF8D3">
	<td valign=center colspan=2 height="22" style="border:#ff0000 solid 1px;">
	启用此功能后, 系统将只接收已设置的信任域发来的邮件, 以及已发送至的域名发来的邮件. 这会造成无法接收一些邮局发来的邮件.
	</td>
	</tr>
	<tr bgcolor="#ffffff"><td height="15" colspan=2></td></tr>
<%
end if

if ei.onlyRcptToSystem = true then
%>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left">只允许发信给本系统用户 (<font color="#FF3333">一般不选中</font>)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="onlyRcptToSystem" <% if ei.onlyRcptToSystem = true then response.write "checked"%>>
        </td>
    </tr>
	<tr bgcolor="#FFF8D3">
	<td valign=center colspan=2 height="22" style="border:#ff0000 solid 1px;">
	启用此功能后, 系统将无法对外网发送邮件.
	</td>
	</tr>
	<tr bgcolor="#ffffff"><td height="15" colspan=2></td></tr>
<%
end if

if ei.onlyMailFromSystem = false then
%>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left">只允许本系统用户发信 (<font color="#FF3333">建议选中</font>)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="onlyMailFromSystem" <% if ei.onlyMailFromSystem = true then response.write "checked"%>>
        </td>
    </tr>
	<tr bgcolor="#FFF8D3">
	<td valign=center colspan=2 height="22" style="border:#ff0000 solid 1px;">
	禁用此功能后, 系统将有可能会被他人利用, 以对外网发送垃圾邮件.
	</td>
	</tr>
	<tr bgcolor="#ffffff"><td height="15" colspan=2></td></tr>
<%
else
	Response.Write "<input type='hidden' name='onlyMailFromSystem' value='1'>" & Chr(13)
end if

if ei.useAuth = false then
%>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left">启用SMTP发信身份认证功能 (<font color="#FF3333">建议选中</font>)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="useAuth" <% if ei.useAuth = true then response.write "checked"%>>
        </td>
    </tr>
	<tr bgcolor="#FFF8D3">
	<td valign=center colspan=2 height="22" style="border:#ff0000 solid 1px;">
	禁用此功能后, 系统将有可能会被他人利用, 以对外网发送垃圾邮件.
	</td>
	</tr>
	<tr bgcolor="#ffffff"><td height="15" colspan=2></td></tr>
<%
else
	Response.Write "<input type='hidden' name='useAuth' value='1'>" & Chr(13)
end if
%>
  </table>
<br>
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-top:1px <%=MY_COLOR_1 %> solid;">
		<tr>
		<td colspan=3 height=27 valign=center align=center bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
		设置系统反垃圾邮件功能 (接收处理)
		</td>
		</tr>
    <tr>
      <td height="25" width="40%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
<a name="wlip"></a>
        <div align="left">启用IP白名单&nbsp;&nbsp;&nbsp;<a href="wl_ip.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="Enable_WhiteList_IP" <% if ei.Enable_WhiteList_IP = true then response.write "checked"%>>
		</td>
    </tr>
    <tr>
      <td height="25" width="40%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
<a name="wldomain"></a>
        <div align="left">启用域名白名单&nbsp;&nbsp;&nbsp;<a href="wl_domain.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="Enable_WhiteList_Domain" <% if ei.Enable_WhiteList_Domain = true then response.write "checked"%>>
	&nbsp;<font color="#FF3333">建议使用</font>
		</td>
    </tr>
    </table>
<br>
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-top:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用SMTP域名验证功能</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableSmtpDomainCheck" <% if ei.EnableSmtpDomainCheck = true then response.write "checked"%>>
	&nbsp;<font color="#FF3333">建议使用</font>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">检查发送者邮件地址中域名的 A、MX、SPF 记录</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableCheckMailFromIP" <% if ei.EnableCheckMailFromIP = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left"><%=s_lang_0030 %></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableCheckMailHeader" <% if ei.EnableCheckMailHeader = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">使用 DNS 检查发送者邮件地址中域名的有效性</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableCheckMailFromDomainIsGood" <% if ei.EnableCheckMailFromDomainIsGood = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">检查 HELO/EHLO 主机名的 A、MX、SPF 记录</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableCheckHeloIP" <% if ei.EnableCheckHeloIP = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">检查主机名失败后, 检查发送者邮件地址中域名的 A、MX、SPF 记录</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableCheckMailFromIP_WhenCheckHeloIP_Error" <% if ei.EnableCheckMailFromIP_WhenCheckHeloIP_Error = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">IP地址匹配方式:</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <select name="CheckIPClass" class="drpdwn">
		<option value="0">A Class</option>
		<option value="1">B Class</option>
		<option value="2" selected>C Class</option>
        </select>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">未通过检查邮件的处理方式:</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <select name="EnableSmtpCheckError2Trash" class="drpdwn">
		<option value="0" selected>放入垃圾箱</option>
		<option value="1">拒收</option>
        </select>
        </td>
    </tr>
    </table>
<br>
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-top:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td height="25" width="40%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left"><%=s_lang_0003 %>&nbsp;&nbsp;&nbsp;<a href="greylisting.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableGreylisting" <% if ei.EnableGreylisting = true then response.write "checked"%>>
	&nbsp;<font color="#FF3333">建议使用</font>
		</td>
    </tr>
    </table>
<br>
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-top:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td height="25" width="40%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
<a name="rbl"></a>
        <div align="left">启用 RBL (Real-time Black List)&nbsp;&nbsp;&nbsp;<a href="rbl.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="Enable_RBL" <% if ei.Enable_RBL = true then response.write "checked"%>>
		</td>
    </tr>
    </table>
<br>
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-top:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td height="25" width="40%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用连接攻击保护功能</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableKillAttacker" <% if ei.enableKillAttacker = true then response.write "checked"%>>
        </td>
      <td align="left" width="53%" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>&nbsp;</td>
    </tr>
    <tr>
      <td height="25" width="40%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用邮件附件名称过滤功能&nbsp;&nbsp;&nbsp;<a href="exnamefilter.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" width="7%" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enable_AttachmentExName_Filter" <% if ei.enable_AttachmentExName_Filter = true then response.write "checked"%>>
		</td>
      <td align="left" width="53%" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>&nbsp;</td>
    </tr>
    <tr>
      <td height="25" width="40%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用主题过滤功能&nbsp;&nbsp;&nbsp;<a href="keywords.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" width="7%" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableKeywordFilter" <% if ei.EnableKeywordFilter = true then response.write "checked"%>>
		</td>
      <td align="left" width="53%" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
	&nbsp;
		</td>
    </tr>
    <tr>
      <td height="25" width="40%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用邮件内容过滤功能&nbsp;&nbsp;&nbsp;<a href="systemfilter.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" width="7%" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableSystemFilter" <% if ei.enableSystemFilter = true then response.write "checked"%>>
		</td>
      <td align="left" width="53%" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
	&nbsp;
		</td>
    </tr>
	</table>
<br><br>
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-top:1px <%=MY_COLOR_1 %> solid;">
		<tr>
		<td colspan=3 height=27 valign=center align=center bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
		设置系统反垃圾邮件功能 (发送处理)
		</td>
		</tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">启用用户外发邮件自动限制功能<br>(使用此功能前, 必须要启用SMTP发信认证功能)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableAutoStopUserOutGoing" <% if ei.enableAutoStopUserOutGoing = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">界定是否外发邮件过多的1小时最大外发邮件数量</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="autoStopUserOutGoingMaxNumber" class="textbox" size="5" maxlength="4" value="<%=ei.autoStopUserOutGoingMaxNumber %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">暂停其外发邮件功能的持续时间为</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="autoStopUserOutGoingExpiresMinute" class="textbox" size="5" maxlength="4" value="<%=ei.autoStopUserOutGoingExpiresMinute %>">&nbsp;分钟
        </td>
    </tr>
	</table>
<br>
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-top:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left">防止利用系统内帐号进行中继发送</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="DisableRelayEmail" <% if ei.DisableRelayEmail = true then response.write "checked"%>>
	&nbsp;<font color="#FF3333">建议使用</font>
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">处理方式:</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <select name="DisableRelayEmail_Mode" class="drpdwn">
		<option value="0" selected>拒收</option>
		<option value="1">修改邮件头</option>
        </select>
        </td>
    </tr>
    </table>
<br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="30%" height="28">&nbsp;</td>
      <td width="70%" align=right>
	<input type="button" value=" 保存 " name="save2" onclick="javascript:gosub()" class="Bsbttn" disabled>&nbsp;&nbsp;
	<input type="button" value=" 退出 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">&nbsp;&nbsp;&nbsp;
      </td>
    </tr>
  </table>
<br>
  </FORM>
</BODY>
</HTML>

<%
set ei = nothing
set adm = nothing
%>
