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
set ei = server.createobject("easymail.stakeout")
'-----------------------------------------
ei.Load

dim eu
set eu = Application("em")
%>


<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<script LANGUAGE=javascript>
<!--
function addnew()
{
	var i = 0;
	var al = "";

	for (i; i < document.f1.selectusers.length; i++)
	{
		al = al + document.f1.selectusers[i].value + '\t';
	}

	form1.addlist.value = al;

	form1.submit();
}


function isinlist(name)
{
	var i = 0;
	for (i; i < document.f1.selectusers.length; i++)
	{
		if (document.f1.selectusers[i].value == name)
		{
			return true;
		}
	}
	
	return false;
}


function addinFromText()
{
	if (document.f1.addtext.value != "" && isinlist(document.f1.addtext.value) == false)
	{
		var oOption = document.createElement("OPTION");
		oOption.text = document.f1.addtext.value;
		oOption.value = document.f1.addtext.value;
<%
if isMSIE = true then
%>
		document.f1.selectusers.add(oOption);
<%
else
%>
		document.f1.selectusers.appendChild(oOption);
<%
end if
%>

		document.f1.addtext.value = "";
	}
}


function addin()
{
	var i = 0;
	for (i; i < document.f1.selectalluser.length; i++)
	{
		if (document.f1.selectalluser[i].selected == true)
		{
			if (isinlist(document.f1.selectalluser[i].value) == false)
			{
				var oOption = document.createElement("OPTION");
				oOption.text = document.f1.selectalluser[i].value;
				oOption.value = document.f1.selectalluser[i].value;
<%
if isMSIE = true then
%>
				document.f1.selectusers.add(oOption);
<%
else
%>
				document.f1.selectusers.appendChild(oOption);
<%
end if
%>
			}
		}
	}
}


function delout()
{
	var i = 0;
	for (i; i < document.f1.selectusers.length; i++)
	{
		if (document.f1.selectusers[i].selected == true)
		{
			document.f1.selectusers.remove(i);
			i--;
		}
	}
}

//-->
</script>

<BODY>
<br>
<form name="f1">
<div align="center"><br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="15%" height="28">&nbsp;</td>
      <td width="32%"><a href="showsysinfo.asp?<%=getGRSN() %>#showstakeout">启动项设置</a></td>
      <td colspan="23"><a href="right.asp?<%=getGRSN() %>">返回</a></td>
      <td width="30%"><font class="s" color="<%=MY_COLOR_4 %>"><b>邮件监控设置</b></font></td>
    </tr>
  </table>
</div>
  <div align="center">
  <table align="center" border="0" width="90%" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr>
      <td height="20" colspan="3">&nbsp;
      </td></tr>
    <tr>
      <td height="20" rowspan="2" width="45%"> 
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>所有用户</b></font></div>
      </td>
      <td height="20" rowspan="2"> 
        <div align="center"> </div>
        <div align="center"> </div>
      </td>
      <td height="20" rowspan="2" width="45%"> 
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>被监控用户</b></font></div>
      </td>
    </tr>
    <tr> </tr>
    <tr> 
      <td height="94" rowspan="2" width="45%"> 
        <div align="center">
          <select name="selectalluser" size="14" class="drpdwn" style="width: 200;" multiple>
            <%
i = 0
allnum = eu.GetUsersCount

do while i < allnum
	eu.GetUserByIndex i, name, domain, comment

	response.write "<option value='" & name & "'>" & name & "</option>"

	name = NULL
	domain = NULL
	comment = NULL

	i = i + 1
loop
%> 
          </select>
        </div>
		<br>
		<div align="center">
		<input type="text" name="addtext" class="textbox" maxlength="64">&nbsp;&nbsp;<input type="button" value="==&gt;&gt;" class="sbttn" LANGUAGE=javascript onclick="addinFromText()">
		</div>
		</td>
      <td height="47" width="10%"> 
        <div align="center"> 
          <input type="button" name="Button" value="  ==&gt;  " class="sbttn" LANGUAGE=javascript onclick="addin()">
        </div>
      </td>
      <td height="94" rowspan="2" width="45%"> 
        <div align="center"> 
          <select name="selectusers" size="16" class="drpdwn" style="width: 200;" multiple>
            <%
i = 0
allnum = ei.getcount

do while i < allnum
	response.write "<option value='" & ei.GetNameByIndex(i) & "'>" & ei.GetNameByIndex(i) & "</option>"
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
          <input type="button" name="Submit2" value="  &lt;==  " class="sbttn" LANGUAGE=javascript onclick="delout()">
        </div>
      </td>
    </tr>
    <tr>
      <td height="20" colspan="3">&nbsp;
      </td></tr>
    <tr>
    <tr>
      <td height="20" colspan="3" align="right">
    <input type="button" value=" 确认 " LANGUAGE=javascript onclick="addnew()" class="Bsbttn">&nbsp;&nbsp;
    <input type="button" value=" 退出 " LANGUAGE=javascript onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
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
        <td width="94%">邮件监控功能可以让管理员对"被监控用户"进行监控, 其所发出以及收到的邮件将会在管理员的邮箱中保留一份拷贝.
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
</form>
<form action="savestakeout.asp" method=post id=form1 name=form1>
<input type="hidden" name="addlist">
</FORM>
  </div>
<br>
</BODY>
</HTML>


<%
set ei = nothing
set eu = nothing
%>
