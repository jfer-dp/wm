<!--#include file="passinc.asp" -->

<%
fileid = trim(request("fileid"))


dim pf
set pf = server.createobject("easymail.PubFolderManager")

pf.load fileid

if LCase(pf.admin) <> LCase(Session("wem")) then
	if isadmin() = false then
		set pf = nothing
		response.redirect "noadmin.asp"
	end if
end if


dim eu
set eu = Application("em")


if trim(request("save")) = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	pfname = trim(request("pfname"))
	pfadmin = trim(request("pfadmin"))
	mpermission = trim(request("mpermission"))
	pfmaxitems = trim(request("pfmaxitems"))
	pfmaxlength = trim(request("pfmaxlength"))

	if IsNumeric(mpermission) = true and IsNumeric(pfmaxitems) = true and IsNumeric(pfmaxlength) = true then
		pf.admin = pfadmin
		pf.foldername = pfname
		pf.permission = CInt(mpermission)
		pf.MaxItems = CInt(pfmaxitems)
		pf.itemmaxsize = CLng(pfmaxlength)

		pf.save

		set pf = nothing
		Response.Redirect "showallpf.asp?" & getGRSN()
	end if

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
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>
<script LANGUAGE=javascript>
<!--
function gosub()
{
	if (document.f1.pfname.value != "" && document.f1.pfadmin.value != "" && document.f1.pfmaxitems.value != "" && document.f1.pfmaxitems.value >= 0 && document.f1.pfmaxlength.value != "" && document.f1.pfmaxlength.value > 0)
		document.f1.submit();
}
//-->
</script>

<body>
<form name="f1" method="post" action="editpf.asp">
<input type="hidden" name="save" value="1">
<input type="hidden" name="fileid" value="<%=fileid %>">
<br><br><br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" nowrap style="border:1px <%=MY_COLOR_1 %> solid;"><p align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>编辑公共文件夹属性: <%=server.htmlencode(name) %></b></font></td>
	</tr>
	<tr>
      <td width="35%" align="right" bgcolor="<%=MY_COLOR_2 %>" height='28' style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><font class="s" color="<%=MY_COLOR_4 %>"><b>公共文件夹名称&nbsp;:&nbsp;</b></font></td>
      <td width="65%" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><input type="text" name="pfname" class="textbox" value="<%=name %>"></td>
    </tr>
    <tr>
      <td width="35%" align="right" bgcolor="<%=MY_COLOR_2 %>" height='28' style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><font class="s" color="<%=MY_COLOR_4 %>"><b>管理员&nbsp;:&nbsp;</b></font></td>
      <td width="65%" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
		<select name="pfadmin" class="drpdwn" size="1">
<%
i = 0
allnum = eu.GetUsersCount

do while i < allnum
	eu.GetUserByIndex i, name, domain, comment

	if name <> admin then
		response.write "<option value='" & name & "'>" & name & "</option>"
	else
		response.write "<option value='" & name & "' selected>" & name & "</option>"
	end if

	name = NULL
	domain = NULL
	comment = NULL

	i = i + 1
loop
%>
	</select>
	</td>
    </tr>
    <tr>
      <td width="35%" align="right" bgcolor="<%=MY_COLOR_2 %>" height='28' style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><font class="s" color="<%=MY_COLOR_4 %>"><b>权限&nbsp;:&nbsp;</b></font></td>
      <td width="65%" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
		<select name="mpermission" class="drpdwn" size="1">
<%
i = 0

do while i < 5
	if permission <> i then
		response.write "<option value='" & i & "'>" & getPermissionStr(i) & "</option>"
	else
		response.write "<option value='" & i & "' selected>" & getPermissionStr(i) & "</option>"
	end if

	i = i + 1
loop
%>
	</select>
	</td>
    </tr>
    <tr>
      <td width="35%" align="right" bgcolor="<%=MY_COLOR_2 %>" height='28' style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><font class="s" color="<%=MY_COLOR_4 %>"><b>当前贴子总数&nbsp;:&nbsp;</b></font></td>
      <td width="65%" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><%=count %></td>
    </tr>
    <tr>
      <td width="35%" align="right" bgcolor="<%=MY_COLOR_2 %>" height='28' style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><font class="s" color="<%=MY_COLOR_4 %>"><b>允许的最大贴数&nbsp;:&nbsp;</b></font></td>
      <td width="65%" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><input type="text" name="pfmaxitems" maxlength="4" class="textbox" value="<%=maxitem %>"></td>
    </tr>
    <tr>
      <td width="35%" align="right" bgcolor="<%=MY_COLOR_2 %>" height='28' style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><font class="s" color="<%=MY_COLOR_4 %>"><b>贴子的最大长度&nbsp;:&nbsp;</b></font></td>
      <td width="65%" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><input type="text" name="pfmaxlength" maxlength="9" class="textbox" value="<%=itemmaxsize %>"> Byte</td>
    </tr>
	<tr>
      <td colspan="7" align="right" bgcolor="#ffffff"><br>
	<input type="button" value=" 保存 " onclick="javascript:gosub();" class="Bsbttn">&nbsp;
	<input type="button" value=" 返回 " onclick="javascript:location.href='showallpf.asp?<%=getGRSN() %>';" class="Bsbttn">
      </td>
    </tr>
  </table>
<%
filename = NULL
admin = NULL
permission = NULL
name = NULL
createtime = NULL
count = NULL
maxid = NULL
maxitem = NULL
itemmaxsize = NULL
%>
</table>
</form>
</BODY>
</HTML>

<%
set pf = nothing
set eu = nothing

function getTimeStr(otime)
	getTimeStr = mid(otime, 1, 4) & "-"
	getTimeStr = getTimeStr & mid(otime, 5, 2) & "-"
	getTimeStr = getTimeStr & mid(otime, 7, 2) & "&nbsp;"
	getTimeStr = getTimeStr & mid(otime, 9, 2) & ":"
	getTimeStr = getTimeStr & mid(otime, 11, 2) & ":"
	getTimeStr = getTimeStr & mid(otime, 13, 2)
end function

function getShowSize(bytesize)
	if bytesize < 1000 then
		getShowSize = bytesize & "字节"
	else
		getShowSize = CLng(bytesize/1000) & "K"
	end if
end function

function getSortStr(sortnum)
	if sortnum = 0 then
		getSortStr = "发贴时间"
	elseif sortnum = 1 then
		getSortStr = "发贴人"
	elseif sortnum = 2 then
		getSortStr = "主题"
	elseif sortnum = 3 then
		getSortStr = "长度"
	elseif sortnum = 4 then
		getSortStr = "点击数"
	end if
end function

function getPermissionStr(pm)
	if pm = 0 then
		getPermissionStr = "发贴人可以修改"
	elseif pm = 1 then
		getPermissionStr = "发贴人可以修改或删除"
	elseif pm = 2 then
		getPermissionStr = "只有管理员才可以修改或删除"
	elseif pm = 3 then
		getPermissionStr = "不允许加新贴"
	elseif pm = 4 then
		getPermissionStr = "完全锁住(不许加新贴, 也不许修改或删除)"
	end if
end function
%>
