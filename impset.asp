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
	<input type="button" value=" ���� " name="save1" onclick="javascript:gosub()" class="Bsbttn" disabled>&nbsp;&nbsp;
	<input type="button" value=" �˳� " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">&nbsp;&nbsp;&nbsp;
      </td>
    </tr>
  </table>
	<br>
	<table width="90%"  border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr>
		<td height="25" colspan=2 bgcolor="<%=MY_COLOR_3 %>" align="left"> 
		<%=s_lang_0298 %>��<input type="text" name="ZATT_URL" size="70" class="textbox" value="<%=ei.ZATT_URL %>">
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
	��������д�ɷ��ʵ�http��ַ, �������� <font color="#FF3333">/downatt.asp</font> ��β, ����: http://mail.domain.com/downatt.asp<br>
	<font color="#FF3333">��Ҫ</font>: �������ò���, ������û������͵�����ʽ�����޷�����.
	</td>
    </tr>
	<tr bgcolor="#ffffff"><td height="15" colspan=2></td></tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-top:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">DNS��ַ (�޸ĺ�, ��Ҫ�����ʼ�������������Ч)</div>
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
	DNS��ַ���÷ǳ���Ҫ, ��ֱ��Ӱ����������ŵĳɰ�. ��Ӧ����������Ч���Ҳ���ͬ��DNS��ַ(���ܳ�������), �м��� , �ŷָ�.
	</td>
    </tr>
	<tr bgcolor="#ffffff"><td height="15" colspan=2></td></tr>
    <tr>
	<td colspan=2 height="25" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
	<table width="100%" border="0" cellspacing="0">
    <tr>
	<td height="25">
	�����������ʼ�ͳ����
	</td>
    </tr>
    <tr> 
	<td align="center" height="25" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        ����:&nbsp;&nbsp;<input name="TrashMsg_Subject" type="text" value="<%=adm.TrashMsg_Subject %>" size="50" class='textbox'>
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
	����Ҫ�޸ĵ�һ�� <font color="#FF3333">%URL%=</font> �������Ϊ�ʼ�ϵͳ�������ⲿ���ʵ�http��ַ, ���������Ҫ�� <font color="#FF3333">/trashmsg.asp</font> ��β. ����: http://www.domain.com/mail/trashmsg.asp �� http://mail.domain.com/trashmsg.asp<br>
	<font color="#FF3333">��Ҫ</font>: ���������ò���, ��Ӱ���û����������ʼ����й���.
	</td>
	</tr>
	<tr bgcolor="#ffffff"><td height="15" colspan=2></td></tr>
<%
if ei.enableOpenRelay = false then
%>
	<tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
		<div align="left">�����MAIL FROM��������ͷ��FROM��ַ��һ���� (<font color="#FF3333">����ѡ��</font>)</div>
		</td>
		<td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
		<input type="checkbox" name="enableOpenRelay" <% if ei.enableOpenRelay = true then response.write "checked"%>>
		</td>
	</tr>
	<tr bgcolor="#FFF8D3">
	<td valign=center colspan=2 height="22" style="border:#ff0000 solid 1px;">
	ʹ�ü��һ���ԵĹ��ܺ�, ϵͳ�����һЩ���淶���ʼ�, �Ӷ�����޷�����һЩ�ʾַ������ʼ�.
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
        <div align="left">���ý����ʼ�����ƹ��� (<font color="#FF3333">һ�㲻ѡ��</font>)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableRec_Cortrol" <% if ei.enableRec_Cortrol = true then response.write "checked"%>>
        </td>
    </tr>
	<tr bgcolor="#FFF8D3">
	<td valign=center colspan=2 height="22" style="border:#ff0000 solid 1px;">
	���ô˹��ܺ�, ϵͳ��ֻ���������õ������������ʼ�, �Լ��ѷ������������������ʼ�. �������޷�����һЩ�ʾַ������ʼ�.
	</td>
	</tr>
	<tr bgcolor="#ffffff"><td height="15" colspan=2></td></tr>
<%
end if

if ei.onlyRcptToSystem = true then
%>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left">ֻ�����Ÿ���ϵͳ�û� (<font color="#FF3333">һ�㲻ѡ��</font>)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="onlyRcptToSystem" <% if ei.onlyRcptToSystem = true then response.write "checked"%>>
        </td>
    </tr>
	<tr bgcolor="#FFF8D3">
	<td valign=center colspan=2 height="22" style="border:#ff0000 solid 1px;">
	���ô˹��ܺ�, ϵͳ���޷������������ʼ�.
	</td>
	</tr>
	<tr bgcolor="#ffffff"><td height="15" colspan=2></td></tr>
<%
end if

if ei.onlyMailFromSystem = false then
%>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left">ֻ����ϵͳ�û����� (<font color="#FF3333">����ѡ��</font>)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="onlyMailFromSystem" <% if ei.onlyMailFromSystem = true then response.write "checked"%>>
        </td>
    </tr>
	<tr bgcolor="#FFF8D3">
	<td valign=center colspan=2 height="22" style="border:#ff0000 solid 1px;">
	���ô˹��ܺ�, ϵͳ���п��ܻᱻ��������, �Զ��������������ʼ�.
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
        <div align="left">����SMTP���������֤���� (<font color="#FF3333">����ѡ��</font>)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="useAuth" <% if ei.useAuth = true then response.write "checked"%>>
        </td>
    </tr>
	<tr bgcolor="#FFF8D3">
	<td valign=center colspan=2 height="22" style="border:#ff0000 solid 1px;">
	���ô˹��ܺ�, ϵͳ���п��ܻᱻ��������, �Զ��������������ʼ�.
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
		����ϵͳ�������ʼ����� (���մ���)
		</td>
		</tr>
    <tr>
      <td height="25" width="40%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
<a name="wlip"></a>
        <div align="left">����IP������&nbsp;&nbsp;&nbsp;<a href="wl_ip.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="Enable_WhiteList_IP" <% if ei.Enable_WhiteList_IP = true then response.write "checked"%>>
		</td>
    </tr>
    <tr>
      <td height="25" width="40%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
<a name="wldomain"></a>
        <div align="left">��������������&nbsp;&nbsp;&nbsp;<a href="wl_domain.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="Enable_WhiteList_Domain" <% if ei.Enable_WhiteList_Domain = true then response.write "checked"%>>
	&nbsp;<font color="#FF3333">����ʹ��</font>
		</td>
    </tr>
    </table>
<br>
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-top:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">����SMTP������֤����</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableSmtpDomainCheck" <% if ei.EnableSmtpDomainCheck = true then response.write "checked"%>>
	&nbsp;<font color="#FF3333">����ʹ��</font>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">��鷢�����ʼ���ַ�������� A��MX��SPF ��¼</div>
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
        <div align="left">ʹ�� DNS ��鷢�����ʼ���ַ����������Ч��</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableCheckMailFromDomainIsGood" <% if ei.EnableCheckMailFromDomainIsGood = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">��� HELO/EHLO �������� A��MX��SPF ��¼</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableCheckHeloIP" <% if ei.EnableCheckHeloIP = true then response.write "checked"%>>
        </td>
    </tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">���������ʧ�ܺ�, ��鷢�����ʼ���ַ�������� A��MX��SPF ��¼</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="EnableCheckMailFromIP_WhenCheckHeloIP_Error" <% if ei.EnableCheckMailFromIP_WhenCheckHeloIP_Error = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">IP��ַƥ�䷽ʽ:</div>
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
        <div align="left">δͨ������ʼ��Ĵ���ʽ:</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <select name="EnableSmtpCheckError2Trash" class="drpdwn">
		<option value="0" selected>����������</option>
		<option value="1">����</option>
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
	&nbsp;<font color="#FF3333">����ʹ��</font>
		</td>
    </tr>
    </table>
<br>
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-top:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td height="25" width="40%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
<a name="rbl"></a>
        <div align="left">���� RBL (Real-time Black List)&nbsp;&nbsp;&nbsp;<a href="rbl.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
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
        <div align="left">�������ӹ�����������</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableKillAttacker" <% if ei.enableKillAttacker = true then response.write "checked"%>>
        </td>
      <td align="left" width="53%" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>&nbsp;</td>
    </tr>
    <tr>
      <td height="25" width="40%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">�����ʼ��������ƹ��˹���&nbsp;&nbsp;&nbsp;<a href="exnamefilter.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
      </td>
      <td align="left" width="7%" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enable_AttachmentExName_Filter" <% if ei.enable_AttachmentExName_Filter = true then response.write "checked"%>>
		</td>
      <td align="left" width="53%" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>&nbsp;</td>
    </tr>
    <tr>
      <td height="25" width="40%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">����������˹���&nbsp;&nbsp;&nbsp;<a href="keywords.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
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
        <div align="left">�����ʼ����ݹ��˹���&nbsp;&nbsp;&nbsp;<a href="systemfilter.asp?<%=getGRSN() %>"><img src="images\ugo.gif" border="0" align="absbottom"></a></div>
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
		����ϵͳ�������ʼ����� (���ʹ���)
		</td>
		</tr>
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">�����û��ⷢ�ʼ��Զ����ƹ���<br>(ʹ�ô˹���ǰ, ����Ҫ����SMTP������֤����)</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="enableAutoStopUserOutGoing" <% if ei.enableAutoStopUserOutGoing = true then response.write "checked"%>>
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">�綨�Ƿ��ⷢ�ʼ������1Сʱ����ⷢ�ʼ�����</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="autoStopUserOutGoingMaxNumber" class="textbox" size="5" maxlength="4" value="<%=ei.autoStopUserOutGoingMaxNumber %>">
        </td>
    </tr>
    <tr> 
      <td height="25" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">��ͣ���ⷢ�ʼ����ܵĳ���ʱ��Ϊ</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="text" name="autoStopUserOutGoingExpiresMinute" class="textbox" size="5" maxlength="4" value="<%=ei.autoStopUserOutGoingExpiresMinute %>">&nbsp;����
        </td>
    </tr>
	</table>
<br>
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-top:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <div align="left">��ֹ����ϵͳ���ʺŽ����м̷���</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <input type="checkbox" name="DisableRelayEmail" <% if ei.DisableRelayEmail = true then response.write "checked"%>>
	&nbsp;<font color="#FF3333">����ʹ��</font>
        </td>
    </tr>
    <tr> 
      <td height="25" width="55%" bgcolor="<%=MY_COLOR_3 %>" align="right" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
        <div align="left">����ʽ:</div>
      </td>
      <td align="left" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
        <select name="DisableRelayEmail_Mode" class="drpdwn">
		<option value="0" selected>����</option>
		<option value="1">�޸��ʼ�ͷ</option>
        </select>
        </td>
    </tr>
    </table>
<br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="30%" height="28">&nbsp;</td>
      <td width="70%" align=right>
	<input type="button" value=" ���� " name="save2" onclick="javascript:gosub()" class="Bsbttn" disabled>&nbsp;&nbsp;
	<input type="button" value=" �˳� " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">&nbsp;&nbsp;&nbsp;
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
