<!--#include file="passinc.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim dm
set dm = server.createobject("easymail.Domain")
dm.Load

'-----------------------------------------
dim ei
set ei = server.createobject("easymail.DomainDefaultMailBoxSize")
ei.Load
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>
<SCRIPT LANGUAGE=javascript>
<!--
function domainname_onchange() {
	location.href = "showddsize.asp?<%=getGRSN() %>&selectdomain=" + document.f1.domainname.value;
}


function changemyselect_onclick() {
	if (document.f1.ksize.disabled == true)
		document.f1.ksize.disabled = false;
	else
		document.f1.ksize.disabled = true;
}


function all2system()
{
	if (confirm("�Ƿ��������ʹ��ϵͳȱʡ�������С? "))
	{
		document.f1.cleanall.value = "yes";
		document.f1.submit();
	}
}

function gosub()
{
	if (document.f1.changemyselect.checked == false && document.f1.ksize.value == "")
		alert("�������!");
	else
		document.f1.submit();
}
//-->
</SCRIPT>


<BODY>
<br><br>
<FORM ACTION="saveddsize.asp" METHOD=POST NAME="f1">
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="5%" height="25">&nbsp;</td>
      <td width="36%"><b>ѡ������</b>:&nbsp;<select name="domainname" class="drpdwn" LANGUAGE=javascript onchange="return domainname_onchange()">
<%
i = 0
allnum = dm.GetCount()

do while i < allnum
	domain = dm.GetDomain(i)

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
	curdomain = dm.GetDomain(0)
end if

haveitdm = ei.haveit(curdomain)

ksize = ei.Get(curdomain)
%>
</select>
</td>
      <td width="22%"><a href="javascript:all2system()">ȫ��ʹ��ϵͳȱʡ�������С</a></td>
    </tr>
  </table>
<br>
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>">
	<td height="30" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b><%=server.htmlencode(curdomain) %> ��ȱʡ�����С</b></font>
		</div>
      </td>
    </tr>
    <tr><td height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid;">
&nbsp;<input type="checkbox" name="changemyselect" LANGUAGE=javascript onclick="return changemyselect_onclick()"<%if haveitdm = false then response.write " checked" end if %>>����ʹ��ϵͳȱʡ�������С
	</td></tr>
    <tr><td height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid;">
&nbsp;&nbsp;<font color="#FF3333"><%=server.htmlencode(curdomain) %></font> ��ȱʡ�����С:&nbsp;<input name="ksize" type="text" value="<%if ksize > 0 then response.write ksize end if %>" size="10" maxlength="8" class='textbox'<%if haveitdm = false then response.write " disabled" end if %>>&nbsp;K
	</td></tr>
    <tr> 
	<td align="right" bgcolor="#ffffff">
	<br><input type="button" value=" ���� " class="Bsbttn" onclick="javascript:gosub()">&nbsp;&nbsp;
	<input type="button" value=" ȡ�� " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
  </table>
<input name="cleanall" type="hidden" value="">
<input name="curdomain" type="hidden" value="<%=curdomain %>">
</FORM>
<br><br>
  <div align="center">
    <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr>
		<td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
		<td width="94%">������Ϊÿ����ָ����ͬ��ȱʡ�����С, Ҳ����ʹ��ϵͳȱʡ�������С.
		<br><br>��ʹ�ô˹��ܺ�, ���û�ͨ����������ҳ���ڲ�ͬ����ע��������ʱ, ��������������СҲ����ͬ.
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
<br>
</BODY>
</HTML>

<%
curdomain = NULL
text = NULL

set ei = nothing
set dm = nothing
%>
