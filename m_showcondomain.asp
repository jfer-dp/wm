<!--#include file="passinc.asp" -->

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.domain")
ei.Load

dim mam
set mam = server.createobject("easymail.AdminManager")
mam.Load

'-----------------------------------------
dim eu
set eu = Application("em")
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
function window_onload() {
	document.f1.save1.disabled = false;
}

function domainname_onchange() {
	location.href = "m_showcondomain.asp?<%=getGRSN() %>&selectdomain=" + document.f1.domainname.value;
}

function mycancel() {
	location.href="right.asp?<%=getGRSN() %>";
}

function findIt()
{
	if (document.f1.finddomain.value.length > 0)
		location.href = "m_showcondomain.asp?<%=getGRSN() %>&selectdomain=" + document.f1.finddomain.value;
}
//-->
</SCRIPT>


<BODY LANGUAGE=javascript onload="return window_onload()">
<br>
<FORM ACTION="m_savecondomain.asp" METHOD="POST" NAME="f1">
  <table width="97%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="5%" height="25">&nbsp;</td>
      <td width="65%"><b>ѡ������</b>:&nbsp;<select name="domainname" class="drpdwn" LANGUAGE=javascript onchange="return domainname_onchange()">
<%
i = 0

allnum = ei.GetCount()

do while i < allnum
	domain = ei.GetDomain(i)

	if domain <> trim(request("selectdomain")) then
		response.write "<option value='" & server.htmlencode(domain) & "'>" & server.htmlencode(domain) & "</option>" & Chr(13)
	else
		curdomain = domain
		response.write "<option value='" & server.htmlencode(domain) & "' selected>" & server.htmlencode(domain) & "</option>" & Chr(13)
	end if

	domain = NULL

	i = i + 1
loop


if curdomain = "" then
	curdomain = ei.GetDomain(0)
end if
%>
</select>
</td>
      <td width="30%"><input type="checkbox" name="Enable_DAdminAllotSize" value="checkbox" <% if mam.Enable_DAdminAllotSize = true then response.write "checked"%>>���������Ա����ռ�
</td>
    </tr>
  </table>
<br>
	<table width="97%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
      <td colspan="8" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>��������</b></font></div>
      </td>
    </tr>
    <tr bgcolor="<%=MY_COLOR_2 %>">
      <td width="8%" nowrap align="center" height="25" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">�Ƿ�<br>��ʾ</td>
      <td width="33%" align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">����</td>
      <td width="11%" nowrap align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">����<br>�û���</td>
      <td width="11%" align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">����û���</td>
      <td width="5%" align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">���ռ���(K)</td>
      <td width="5%" align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">�ѷ���ռ���(K)</td>
      <td width="19%" align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">�����Ա</td>
      <td width="8%" align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">��������<br>YYYYMMDD</td>
    </tr>
<%
ei.GetControlMsgEx curdomain, isshow, maxuser, dmanager, maxsize, allsize, expire

response.write "<tr><td align='center' height='25' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='checkshow'"

if isshow = true then
	response.write " checked"
end if

response.write "></td><td style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input name='domain' type='text' value='" & curdomain & "' size='35' maxlength='64' readonly class='textbox'></td>"
response.write "<td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & ei.GetUserNumberInDomain(curdomain) & "</td>"
response.write "<td style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='text' name='maxuser' maxlength='5' value='" & maxuser & "' class='textbox' size='8'></td>"

if maxsize = 0 then
	response.write "<td style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='text' name='maxsize' maxlength='8' value='����' class='textbox' size='8'></td>"
else
	response.write "<td style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='text' name='maxsize' maxlength='8' value='" & maxsize & "' class='textbox' size='8'></td>"
end if

if allsize = 0 then
	response.write "<td style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='text' name='allsize' maxlength='8' value='����' readonly class='textbox' size='8'></td>"
else
	response.write "<td style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='text' name='allsize' maxlength='8' value='" & allsize & "' readonly class='textbox' size='8'></td>"
end if


response.write "<td style='border-bottom:1px " & MY_COLOR_1 & " solid;'><select name='username' class='drpdwn'><option value=''> [��] </option>"

i = 0
allnum = eu.GetUsersCount
do while i < allnum
	eu.GetUserByIndex i, name, domain, comment

	if name = dmanager then
		response.write "<option value='" & server.htmlencode(name) & "' selected>" & server.htmlencode(name) & "</option>"
	else
		response.write "<option value='" & server.htmlencode(name) & "'>" & server.htmlencode(name) & "</option>"
	end if

	name = NULL
	domain = NULL
	comment = NULL

	i = i + 1
loop

response.write "</select></td>"

if expire = 0 then
	response.write "<td style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='text' name='expire' maxlength='8' value='������' class='textbox' size='8'></td></tr>" & Chr(13)
else
	response.write "<td style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='text' name='expire' maxlength='8' value='" & expire & "' class='textbox' size='8'></td></tr>" & Chr(13)
end if


isshow = NULL
maxuser = NULL
dmanager = NULL
maxsize = NULL
allsize = NULL
expire = NULL
%>
    <tr>
	<td colspan="5" align="left" bgcolor="#ffffff"><br>
<input type="text" name="finddomain" class="textbox">
<input type="button" value="��������" onclick="javascript:findIt()" class="sbttn">
	</td>
      <td colspan="3" align="right" bgcolor="#ffffff"><br>
        <input name="save1" type="submit" value=" ���� " class="Bsbttn" Disabled>&nbsp;&nbsp;
        <input type="button" value=" ȡ�� " onClick="javascript:mycancel()" class="Bsbttn">
      </td>
    </tr>
  </table>
    <br><br>
    <br>
  <div align="center">
    <table width="97%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr> 
        <td width="18%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'><font color="#FF3333">&nbsp;�Ƿ���ʾ</font></td>
        <td width="82%"> δ��ѡ�е�����������������û���������ʱ�������б���. <br>
          <br>
        </td>
      </tr>
      <tr> 
        <td valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'><font color="#FF3333">&nbsp;����û���</font></td>
        <td>��ָ��ǰ�����û����������������ֵʱ, ����������ͨ��Webҳ���ڴ����������û�, ������Ա�ɲ��ܴ����ƴ����û�, 
          ���, "�����û���"���������"����û���". <br>
          <br>
        </td>
      </tr>
      <tr> 
        <td valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
        <td><font color="#FF3333"> </font>������ķ������Ƿ�����Internet��ʱ, ��������ÿ��������"����û���"��������. 
		<br>&nbsp;
        </td>
      </tr>
      <tr> 
        <td valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
        <td><font color="#FF3333"> </font>��Ϊָ���������Ա��Ȩʹ����������Ŀռ��С, ������û���޶�������ռ���ǰ(�������й���Ա�Ҳ�Ϊϵͳ����Ա)��Ӧ�����������Ա������ռ�.
		</td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
<input name="curdomain" type="hidden" value="<%=curdomain %>">
</FORM>
<br>
</BODY>
</HTML>

<%
set eu = nothing
set ei = nothing
set mam = nothing
%>
