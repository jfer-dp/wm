<!--#include file="passinc.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim mht
set mht = server.createobject("easymail.Hint")
mht.Load

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if trim(request("mode")) = "add" then
		if trim(request("add_disabled")) <> "" then
			add_disabled = true
		else
			add_disabled = false
		end if

		if trim(request("add_pop")) <> "" then
			add_pop = true
		else
			add_pop = false
		end if

		if trim(request("add_usegrsn")) <> "" then
			add_usegrsn = false
		else
			add_usegrsn = true
		end if

		add_expire = 0
		tmp_expire = trim(request("add_expire"))
		if Len(tmp_expire) > 0 and IsNumeric(tmp_expire) = true then
			if CLng(tmp_expire) > 19720101 then
				add_expire = CLng(tmp_expire)
			end if
		end if

		add_expProc = 0
		if IsNumeric(trim(request("add_expProc"))) = true then
			add_expProc = CLng(trim(request("add_expProc")))
		end if

		if mht.Add(add_disabled, add_usegrsn, add_pop, add_expire, add_expProc, trim(request("add_msg")), trim(request("add_url"))) = true then
			isok = true
			mht.Save
		else
			isok = false
		end if

		set mht = nothing

		if isok = true then
			response.redirect "ok.asp?" & getGRSN() & "&gourl=syshint.asp"
		else
			response.redirect "err.asp?" & getGRSN() & "&gourl=syshint.asp"
		end if
	end if

	if trim(request("mode")) = "save" then
		if trim(request("EnableHint")) <> "" then
			EnableHint = true
		else
			EnableHint = false
		end if

		mht.EnableHint = EnableHint
		mht.Save

		set mht = nothing
		response.redirect "ok.asp?" & getGRSN() & "&gourl=syshint.asp"
	end if
end if

if Request.ServerVariables("REQUEST_METHOD") = "GET" then
	if IsNumeric(trim(request("delindex"))) = true then
		mht.RemoveByIndex CLng(trim(request("delindex")))
		mht.Save
	end if
end if
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
function add() {
	if (document.f1.add_msg.value.length > 0)
	{
		document.f1.mode.value = "add";
		document.f1.submit();
	}
}

function save() {
	document.f1.mode.value = "save";
	document.f1.submit();
}
//-->
</SCRIPT>

<BODY>
<br>
<FORM ACTION="syshint.asp" METHOD="POST" NAME="f1">
<input type="hidden" name="mode">
	<table width="97%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr>
	<td colspan="8" height="35" width="50%" bgcolor="#ffffff"><table><tr bgcolor="<%=MY_COLOR_2 %>"><td nowrap style="border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<input type="checkbox" name="EnableHint" value="checkbox" <% if mht.EnableHint = true then response.write "checked"%>>启用提醒功能&nbsp;</td>
	<td bgcolor="#ffffff">&nbsp;&nbsp;&nbsp;<input type="button" value=" 保存 " onClick="javascript:save()" class="Bsbttn"></td></tr></table></td>
    </tr>
    <tr bgcolor="<%=MY_COLOR_2 %>">
      <td colspan="8" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>提醒设置</b></font></div>
      </td>
    </tr>
    <tr bgcolor="<%=MY_COLOR_2 %>">
      <td width="4%" nowrap align="center" height="25" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">删除</td>
      <td width="6%" nowrap align="center" height="25" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">状态</td>
      <td width="6%" nowrap align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">弹出<br>窗口</td>
      <td width="28%" nowrap align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">提示内容</td>
      <td width="28%" nowrap align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">链接</td>
      <td width="6%" nowrap align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">缓存<br>页面</td>
      <td width="12%" nowrap align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">期满日期<br>YYYYMMDD</td>
      <td width="10%" nowrap align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">到期处<br>理方式</td>
    </tr>
<%
i = 0
allnum = mht.Count

do while i < allnum
	mht.Get i, isDisabled, isUseGrsn, isPOP, expire, expProc, msg, url

	response.write "<tr><td align='center' height='25' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><a href='syshint.asp?" & getGRSN() & "&delindex=" & i & "'><img src='images/remove.gif' border='0'></a></td>"

	if isDisabled = false then
		response.write "<td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>显示</td>"
	else
		response.write "<td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>禁止</td>"
	end if

	if isPOP = true then
		response.write "<td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>是</td>"
	else
		response.write "<td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>否</td>"
	end if

	response.write "<td style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='text' name='msg" & i & "' maxlength='510' value='" & msg & "' class='textbox' size='25' readonly></td>"
	response.write "<td style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='text' name='url" & i & "' maxlength='510' value='" & url & "' class='textbox' size='25' readonly></td>"

	if isUseGrsn = true then
		response.write "<td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>否</td>"
	else
		response.write "<td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>是</td>"
	end if

	if expire = 0 then
		response.write "<td style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='text' name='expire" & i & "' maxlength='8' value='不限制' class='textbox' size='8' readonly></td>"
	else
		response.write "<td style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='text' name='expire" & i & "' maxlength='8' value='" & expire & "' class='textbox' size='8' readonly></td>"
	end if

	if expProc = 0 then
		response.write "<td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>删除</td>"
	else
		response.write "<td nowrap align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>禁用</td>"
	end if

	response.write "</tr>" & Chr(13)


	isDisabled = NULL
	isUseGrsn = NULL
	isPOP = NULL
	expire = NULL
	expProc = NULL
	msg = NULL
	url = NULL

	i = i + 1
loop
%>
    <tr>
      <td colspan="8" align="right" bgcolor="#ffffff" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><br>
        <input type="button" value=" 返回 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
      </td>
    </tr>
  </table>
    <br><br><br>
	<table width="97%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
      <td colspan="8" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>新建提醒</b></font></div>
      </td>
    </tr>
    <tr bgcolor="<%=MY_COLOR_2 %>">
      <td width="6%" nowrap align="center" height="25" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">禁止<br>显示</td>
      <td width="6%" nowrap align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">弹出<br>窗口</td>
      <td width="30%" nowrap align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">提示内容</td>
      <td width="30%" nowrap align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">链接</td>
      <td width="6%" nowrap align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">缓存<br>页面</td>
      <td width="12%" nowrap align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">期满日期<br>YYYYMMDD</td>
      <td width="10%" nowrap align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">到期处<br>理方式</td>
    </tr>
    <tr>
	<td align="center" height='25' style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><input type='checkbox' name='add_disabled'></td>
	<td align="center" height='25' style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><input type='checkbox' name='add_pop'></td>
	<td style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><input type='text' name='add_msg' maxlength='510' class='textbox' size='27'></td>
	<td style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><input type='text' name='add_url' maxlength='510' class='textbox' size='27'></td>
	<td align='center' style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><input type='checkbox' name='add_usegrsn'></td>
	<td style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><input type='text' name='add_expire' maxlength='8' class='textbox' size='8'></td>
	<td style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
	<select name='add_expProc' class=drpdwn size='1'>
	<option value='0' selected>删除</option>
	<option value='1'>禁用</option>
	</select>
	</td></tr>
    <tr>
      <td colspan="8" align="center" bgcolor="#ffffff"><br>
        <input type="button" value=" 添加 " onClick="javascript:add()" class="Bsbttn">
      </td>
    </tr>
  </table>
  <br><br><br>
  <div align="center">
    <table width="97%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr> 
        <td width="7%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
        <td width="93%"> 设置在用户登录邮箱后随机显示的提醒信息<br>
        </td>
      </tr>
        <td colspan="2" height="10">
        </td>
      </tr>
    </table><br><br>
  </div>
</FORM>
</BODY>
</HTML>

<%
set mht = nothing
%>
