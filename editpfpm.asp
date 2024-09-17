<!--#include file="passinc.asp" -->

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

fileid = trim(request("fileid"))


dim pf
set pf = server.createobject("easymail.PubFolderManager")

pf.load fileid

if pf.admin <> Session("wem") then
	if isadmin() = false then
		set pf = nothing
		response.redirect "noadmin.asp"
	end if
end if

if Len(pf.admin) < 1 then
	set pf = nothing
	Response.Redirect "err.asp?" & getGRSN() & "&gourl=showallpf.asp"
end if

dim filename
dim admin
dim permission
dim name
dim createtime
dim count
dim maxid
dim maxitem
dim itemmaxsize

pf.GetFolderInfo filename, admin, permission, name, createtime, count, maxid, maxitem, itemmaxsize

filename = NULL
admin = NULL
permission = NULL
createtime = NULL
count = NULL
maxid = NULL
maxitem = NULL
itemmaxsize = NULL

set pf = nothing


dim pfvl
set pfvl = server.createobject("easymail.PubFolderViewLimit")
pfvl.Load fileid

if trim(request("save")) = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	pfvl.RemoveAll
	pfvl.AddAll(trim(request("allmsgs")))
	pfvl.Save

	name = NULL
	set pfvl = nothing

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("editpfpm.asp?fileid=" & fileid)
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
		tempstr = tempstr + document.f1.listall[i].value;

		if (i != document.f1.listall.length - 1)
			tempstr = tempstr + "\f";
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
	if (document.f1.s_msg.value.length < 1)
		return ;

	if (haveit() == false)
	{
		var oOption = document.createElement("OPTION");
		oOption.text = convVPF(document.f1.s_mode.value, document.f1.s_msg.value);
		oOption.value = document.f1.s_mode.value + document.f1.s_msg.value;
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
	}
}

function haveit()
{
	var tempstr = document.f1.s_mode.value + document.f1.s_msg.value;

	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].value == tempstr)
			return true;
	}

	return false;
}

function convVPF(v_mode, v_msg)
{
	var tmpstr = "";

	if (v_mode == 0)
		tmpstr = "允许(域名)";
	else if (v_mode == 1)
		tmpstr = "允许(邮件地址)";
	else if (v_mode == 2)
		tmpstr = "拒绝(域名)";
	else if (v_mode == 3)
		tmpstr = "拒绝(邮件地址)";

	if (tmpstr.length > 0)
		return tmpstr + ": " + v_msg;
	else
		return "";
}
//-->
</script>

<body>
<form name="f1" method="post" action="editpfpm.asp">
<input type="hidden" name="allmsgs">
<input type="hidden" name="save" value="1">
<input type="hidden" name="fileid" value="<%=fileid %>">
<br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
	<td height="28" width="100%" align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>公共文件夹访问权限设置: <%=server.htmlencode(name) %></b></font></td>
	</tr>
  </table>
  <table align="center" border="0" width="90%" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr>
	<td height="94" width="80%" rowspan="2" align="center">
	<br>
<select name="listall" size="12" class="drpdwn" style="width: 480;">
<%
i = 0
allnum = pfvl.Count

do while i < allnum
	pfvl.Get i, vpf_mode, vpf_msg
	Response.Write "<option value=""" & CStr(vpf_mode) & server.htmlencode(vpf_msg) & """>" & server.htmlencode(convVPF(vpf_mode, vpf_msg)) & "</option>" & Chr(13)

	vpf_mode = NULL
	vpf_msg = NULL

	i = i + 1
loop
%>
</select>
	</tr>
    <tr> 
      <td height="47" width="10%"> 
        <div align="center"> 
          <input type="button" value="删除" class="sbttn" LANGUAGE=javascript onclick="delout()">
        </div>
      </td>
    </tr>
	<tr>
	<td height="20">
<tr><td height="30" align="left" nowrap>&nbsp;
<select name="s_mode" class=drpdwn>
<option value="0">允许域名</option>
<option value="1">允许邮件地址</option>
<option value="2">拒绝域名</option>
<option value="3">拒绝邮件地址</option>
</select>
<input type="input" name="s_msg" class="textbox" size="30" maxlength="128">&nbsp;<input type="button" value=" 添加 " class="sbttn" LANGUAGE=javascript onclick="add()">
      </td></tr>
    <tr>
    <tr>
      <td colspan="2" height="20" align="right"><br><hr size="1" color="<%=MY_COLOR_1 %>">
    <input type="button" value=" 保存 " LANGUAGE=javascript onclick="sub()" class="Bsbttn">&nbsp;&nbsp;
    <input type="button" value=" 退出 " LANGUAGE=javascript onclick="javascript:location.href='showallpf.asp?<%=getGRSN() %>';" class="Bsbttn">&nbsp;
      </td></tr>
    <tr>
  </table>
</form>
<br><br>
  <div align="center">
    <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr> 
        <td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
        <td width="94%">未设置权限时, 即为所有用户都允许访问.<br>
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
name = NULL

set pfvl = nothing


function convVPF(v_mode, v_msg)
	tmpstr = ""

	if v_mode = 0 then
		tmpstr = "允许(域名)"
	elseif v_mode = 1 then
		tmpstr = "允许(邮件地址)"
	elseif v_mode = 2 then
		tmpstr = "拒绝(域名)"
	elseif v_mode = 3 then
		tmpstr = "拒绝(邮件地址)"
	end if

	convVPF = tmpstr & ": " & v_msg
end function
%>
