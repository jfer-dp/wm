<!--#include file="passinc.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

dim ei
set ei = server.createobject("easymail.HandPoint2")
'-----------------------------------------
ei.Load

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ei.RemoveAll
	ei.Add trim(request("allmsgs"))
	ei.Save
	set ei = nothing

	response.redirect "ok.asp?" & getGRSN() & "&gourl=handpoint2.asp"
end if
%>


<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<script LANGUAGE=javascript>
<!--
function sub()
{
	var tempstr = "";
	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		tempstr = tempstr + document.f1.listall[i].value + "\f";
	}

	document.f1.allmsgs.value = tempstr;
	document.f1.submit();
}

function delout()
{
	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].selected == true)
		{
			document.f1.listall.remove(i);

			if (i < document.f1.listall.length)
				document.f1.listall[i].selected = true;
			else
			{
				if (i - 1 >= 0)
					document.f1.listall[i - 1].selected = true;
			}

			break;
		}
	}
}

function add()
{
	if (document.f1.domain.value.length > 0 && document.f1.domain_address.value.length > 0)
	{
		if (haveit() == false)
		{
			var oOption = document.createElement("OPTION");
			oOption.text = document.f1.domain.value + " --" + "> " + document.f1.domain_address.value;
			oOption.value = document.f1.domain.value + "\t" + document.f1.domain_address.value;
<%
if isMSIE = true then
%>
			document.f1.listall.add(oOption);
<%
else
%>
			document.f1.listall.appendChild(oOption);
<%
end if
%>
			return ;
		}
	}

	alert("�������!");
}

function haveit()
{
	var tempstr = document.f1.domain.value + "\t" + document.f1.domain_address.value;

	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].value == tempstr)
			return true;
	}

	return false;
}
//-->
</script>

<BODY>
<br>
<FORM ACTION="#" METHOD="POST" NAME="f1">
<input type="hidden" name="allmsgs">
<div align="center"><br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="15%" height="28">&nbsp;</td>
      <td width="32%"><a href="showsysinfo.asp?<%=getGRSN() %>#handpoint2">����������</a></td>
      <td colspan="23"><a href="right.asp?<%=getGRSN() %>">����</a></td>
      <td width="30%"><font class="s" color="<%=MY_COLOR_4 %>"><b>�ⷢ��ַ�������</b></font></td>
    </tr>
  </table>
</div>
  <div align="center">
  <table align="center" border="0" width="90%" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr> 
      <td height="94" rowspan="2" width="45%">
        <div align="center">
          &nbsp;<select name="listall" size="13" class="drpdwn" style="width: 480;">
<%
i = 0
allnum = ei.Count

do while i < allnum
	ei.Get i, domain, domain_address
	Response.Write "<option value=""" & server.htmlencode(domain) & Chr(9) & server.htmlencode(domain_address) & """>" & server.htmlencode(domain) & " --> " & server.htmlencode(domain_address) & "</option>"

	domain = NULL
	domain_address = NULL

	i = i + 1
loop
%>
          </select>
        </div>
      </td>
    </tr>
    <tr> 
      <td height="47" width="10%"> 
        <div align="center"> 
          <input type="button" value="ɾ��" class="sbttn" LANGUAGE=javascript onclick="delout()">
        </div>
      </td>
    </tr>
    <tr>
      <td height="20" colspan="3">
<tr><td height="30" align="left" nowrap>&nbsp;����:
<input type="input" name="domain" class="textbox" maxlength="64">&nbsp;&nbsp;&nbsp;������:
<input type="input" name="domain_address" class="textbox" maxlength="64">&nbsp;&nbsp;<input type="button" value=" ��� " class="sbttn" LANGUAGE=javascript onclick="add()">
      </td></tr>
    <tr>
    <tr>
      <td height="20" colspan="3" align="right"><br><hr size="1" color="<%=MY_COLOR_1 %>">
    <input type="button" value=" ���� " LANGUAGE=javascript onclick="sub()" class="Bsbttn">&nbsp;&nbsp;
    <input type="button" value=" �˳� " LANGUAGE=javascript onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
      </td></tr>
    <tr>
  </table>
<br><br><br>
  <div align="center">
    <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr>
        <td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
        <td width="94%">ϵͳ���ⷢ�ʼ�ʱ�ĵ�ַ���������. �����Ӧ����: һ��ר�����м�̨��ͬ���ʼ�ϵͳ, ���ʼ�ϵͳʹ�ò�ͬ������, ����û������ DNS �����.<br>
        <br>����������֧��ͨ�����ʽ. (*: ���ⳤ�ȵ��κ�����.  ?: һ���ַ����κ�����.)<br><br>
��������ǽ�ָ���������ʼ�����(��: anydomain.com), ָ�������ַ(��: 10.96.0.55), ������ϵͳ��Ҫ���͵����ʼ���������ʱ(��: user@anydomain.com), ���ʼ��ᱻֱ�ӷ��͵�ָ����ַ(��: 10.96.0.55).
<br><br>
����, �˹���Ҳ������Ϊ�򵥵��ⷢ�ʼ��м̷��������趨. �����÷��������ⲿ�������ʼ�, ���ɷ����� 202.103.72.99 ��ת��ʱ, ������������: ������: * (* ����ͨ���, ��ָ���е�����) ������: 202.103.72.99
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
<br><br>
</FORM>
  </div>
</BODY>
</HTML>


<%
set ei = nothing
%>
