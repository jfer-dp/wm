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

showindex = 1
if IsNumeric(trim(request("showindex"))) = true then
	showindex = CLng(trim(request("showindex")))
end if

dim mam
set mam = server.createobject("easymail.AdminManager")
mam.Load


savecolorsvalue = trim(request("savecolorsvalue"))
colname = trim(request("colname"))
mode = trim(request("mode"))

if mode = "setcurcolor" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim ei
	set ei = server.createobject("easymail.UserWeb")
	ei.Load Session("wem")

	isok = mam.SetSysColor(showindex, colname, ei.OwnColor)
	mam.Save

	set ei = nothing
	set mam = nothing

	if isok = true then
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("syscolor.asp?showindex=" & showindex)
	else
		Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("syscolor.asp?showindex=" & showindex)
	end if
end if

if savecolorsvalue <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	isok = mam.SetSysColor(showindex, colname, savecolorsvalue)
	mam.Save

	set mam = nothing

	if isok = true then
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("syscolor.asp?showindex=" & showindex)
	else
		Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("syscolor.asp?showindex=" & showindex)
	end if
end if


mam.GetSysColor showindex, cur_sc_name, cur_sc_color

if Len(cur_sc_color) <> 66 then
	cur_sc_color = csbi_color_str_default
end if

colorallnum = 12
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
function gocurcolor() {
	if (confirm("确实要使用当前自订制颜色吗?") == false)
		return ;

	if (document.f1.colname.value.length < 1)
	{
		alert("名称不可为空");
		document.f1.colname.focus();
		return ;
	}

	document.f1.mode.value = "setcurcolor";
	document.f1.submit();
}

function godefault() {
	if (confirm("确实要使用缺省颜色吗?") == false)
		return ;

	if (document.f1.colname.value.length < 1)
	{
		alert("名称不可为空");
		document.f1.colname.focus();
		return ;
	}

	document.f1.savecolorsvalue.value = "<%=csbi_color_str_default %>";
	document.f1.submit();
}

function gosub() {
	if (document.f1.colname.value.length < 1)
	{
		alert("名称不可为空");
		document.f1.colname.focus();
		return ;
	}

	document.f1.savecolorsvalue.value = "";
<%
i = 1

do while i < colorallnum
%>
	if (colorvalue_isok(document.f1.colorvalue<%=i %>.value) == true)
		document.f1.savecolorsvalue.value = document.f1.savecolorsvalue.value + document.f1.colorvalue<%=i %>.value.substr(1);
	else
	{
		alert("输入错误");
		document.f1.colorvalue<%=i %>.focus();
		return ;
	}
<%
    i = i + 1
loop
%>

	document.f1.submit();
}

function isokchar(s)
{
	var okChar = "0123456789abcdefABCDEF";
	var i,c;

	for (i = 0; i < s.length; i++)
	{
		c = s.charAt(i);

		if (okChar.indexOf(c) == -1)
			return false;
	}

	return true;
}

function colorvalue_isok(cstr) {
	if (cstr.length == 7 && cstr.charAt(0) == '#' && isokchar(cstr.substr(1)) == true)
		return true;

	return false;
}

function selcolor_onchange() {
	location.href = "syscolor.asp?showindex=" + document.f1.selcolor.value + "&<%=getGRSN() %>";
}

function set_color(colornum) {
<%
i = 1

do while i < colorallnum
%>
	if (colornum == <%=i %>)
	{
		if (colorvalue_isok(document.f1.colorvalue<%=i %>.value) == true)
			showcolor<%=i %>.bgColor = document.f1.colorvalue<%=i %>.value;
	}
<%
    i = i + 1
loop
%>
}

function editcolor(colornum) {
<%
if isMSIE = true then
%>
	var remote = null;
	var popurl = "";

<%
i = 1

do while i < colorallnum
%>
	if (colornum == <%=i %>)
	{
		popurl = "popcolor.asp?<%=getGRSN() %>&selcol=<%=i %>&cv=";

		if (colorvalue_isok(document.f1.colorvalue<%=i %>.value) == true)
			popurl = popurl + document.f1.colorvalue<%=i %>.value.substr(1);

		remote = window.open(popurl, "", "top=150; left=190; height=223,width=330,scrollbars=yes,resizable=yes,status=no,toolbar=no,menubar=no,location=no");

		if (remote)
			remote.opener = this;
	}
<%
    i = i + 1
loop

else
%>
	document.getElementById("showcolor" + colornum).bgColor = document.getElementById("colorvalue" + colornum).value;
<%
end if
%>
}
//-->
</SCRIPT>

<BODY>
<br><br>
<FORM ACTION="syscolor.asp" METHOD="POST" NAME="f1">
<input name="savecolorsvalue" type="hidden">
<input name="mode" type="hidden">
<input name="showindex" type="hidden" value="<%=showindex %>">
<table width="85%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
  <tr bgcolor="<%=MY_COLOR_2 %>">
	<td height="30" colspan="3" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
	<div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>颜色订制</b></font>&nbsp;&nbsp;&nbsp;
<select name="selcolor" class="drpdwn" LANGUAGE=javascript onchange="return selcolor_onchange()">
<%
i = 1

do while i < 10
	mam.GetSysColor i, sc_name, sc_color

	if i <> showindex then
		if sc_name = "" then
			Response.Write "<option value='" & i & "'>[空]</option>" & Chr(13)
		else
			Response.Write "<option value='" & i & "'>" & server.htmlencode(sc_name) & "</option>" & Chr(13)
		end if
	else
		if sc_name = "" then
			Response.Write "<option value='" & i & "' selected>[空]</option>" & Chr(13)
		else
			Response.Write "<option value='" & i & "' selected>" & server.htmlencode(sc_name) & "</option>" & Chr(13)
		end if
	end if

	sc_name = NULL
	sc_color = NULL

	i = i + 1
loop
%>
</select></div>
	</td>
	</tr>

	<tr><td bgcolor="#ffffff" width="20%" height="27" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">&nbsp;名称</td>
	<td align="left" colspan="2" bgcolor="#ffffff" nowrap height="32" width="30%" style="border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="text" name="colname" size="20" maxlength="128" class='textbox' value="<%=cur_sc_name %>">
	&nbsp;&nbsp;&nbsp;<input type="button" value="使用自订制颜色" onclick="javascript:gocurcolor()" class="Bsbttn">
	</td>
	</tr>
<%
i = 1

do while i < colorallnum
	gcsbi = getColorStringByIndex(cur_sc_color, i)

	if Len(gcsbi) = 6 then
		gcsbi = "#" & gcsbi
	end if
%>
	<tr><td bgcolor="#ffffff" width="32%" height="27" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">&nbsp;<%=getColorName(i) %></td>
	<td align="center" bgcolor="#ffffff" nowrap width="28%" style="border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table height="30" width="90" border="1" style="CURSOR: pointer">
	<tr><td id="showcolor<%=i %>" bgcolor="<%=gcsbi %>" onclick="javascript:editcolor(<%=i %>);">&nbsp;</td></tr>
	</table>
	</td>
	<td bgcolor="#ffffff" width="40%" style="border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="text" id="colorvalue<%=i %>" name="colorvalue<%=i %>" size="10" maxlength="7" class='textbox' value="<%=gcsbi %>">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' value=' 编辑 ' onclick='javascript:editcolor(<%=i %>);' class='sbttn'>
	</td>
	</tr>
<%
    i = i + 1
loop
%>
	<tr bgcolor="#ffffff">
	<td colspan="3" align="right" height="26"><br>
	<input type="button" value="使用缺省颜色" onclick="javascript:godefault()" class="Bsbttn">&nbsp;&nbsp;&nbsp;
	<input type="button" value=" 保存 " onclick="javascript:gosub()" class="Bsbttn">&nbsp;&nbsp;&nbsp;
	<input type="button" value=" 退出 " LANGUAGE=javascript onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
</table>
</form>
<br><br>
  <div align="center">
    <table width="85%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr>
		<td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
		<td width="94%">管理员可以创建9种不同颜色的界面, 以供普通用户选用.
		<br><br>普通用户可以在 "选项 | 邮箱配置 | 颜色设置" 中进行选取.
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
<br><br>
</BODY>
</HTML>

<%
cur_sc_name = NULL
cur_sc_color = NULL

set mam = nothing


function getColorName(color_show_index)
	getColorName = "背景 (" & color_show_index & ")"

	if color_show_index = 5 then
		getColorName = "发信和读信背景 (" & color_show_index & ")"
	elseif color_show_index = 6 then
		getColorName = "横线 (" & color_show_index & ")"
	elseif color_show_index = 7 then
		getColorName = "快速地址列表字颜色 (" & color_show_index & ")"
	elseif color_show_index = 8 then
		getColorName = "树形菜单1背景 (" & color_show_index & ")"
	elseif color_show_index = 9 then
		getColorName = "菜单标题字颜色 (" & color_show_index & ")"
	elseif color_show_index = 10 then
		getColorName = "树形菜单2背景 (" & color_show_index & ")"
	elseif color_show_index = 11 then
		getColorName = "树形菜单2字颜色 (" & color_show_index & ")"
	end if
end function
%>
