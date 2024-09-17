<!--#include file="passinc.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim sh
set sh = server.createobject("easymail.SignHold")
'-----------------------------------------
sh.Load

allnum = sh.count


mode = trim(request("mode"))

if mode <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	sh.RemoveAll

	i = 0
	if mode = "save" then
		do while i < allnum + 1
			if trim(request("kill" & i)) <> "" then
				sh.Add trim(request("kill" & i))
			end if

		    i = i + 1
		loop
	elseif mode = "add" then
		do while i < allnum + 1
			if trim(request("kill" & i)) <> "" then
				sh.Add trim(request("kill" & i))
			end if

		    i = i + 1
		loop
	elseif mode = "del" then
		do while i < allnum + 1
			if trim(request("check" & i)) = "" and trim(request("kill" & i)) <> "" then
				sh.Add trim(request("kill" & i))
			end if

		    i = i + 1
		loop
	end if

	sh.Save


	if trim(request("mode")) <> "add" then
		set sh = nothing
		response.redirect "ok.asp?" & getGRSN() & "&gourl=signhold.asp"
	end if
end if

allnum = sh.count
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>


<SCRIPT LANGUAGE=javascript>
<!--
function del() {
	if (ischeck() == true)
	{
		document.f1.mode.value = "del";
		document.f1.submit();
	}
}

function add() {
	document.f1.mode.value = "add";
	document.f1.submit();
}

function save() {
	document.f1.mode.value = "save";
	document.f1.submit();
}

function ischeck() {
	var i = 0;
	var theObj;

	for(; i<<%=allnum %>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}
//-->
</SCRIPT>


<BODY>
<br>
<FORM ACTION="signhold.asp" METHOD="POST" NAME="f1">
<input type="hidden" name="mode">
<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr>
	<td colspan="2" align="right" bgcolor="#ffffff">
	<br>
	<input type="button" value=" 添加 " onclick="javascript:add()" class="Bsbttn">&nbsp;
	<input type="button" value=" 删除 " onclick="javascript:del()" class="Bsbttn">&nbsp;
	<input type="button" value=" 保存 " onclick="javascript:save()" class="Bsbttn">&nbsp;
	<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	<br>&nbsp;
	</td>
	</tr>
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>保留帐号设置</b></font></div>
      </td>
    </tr>
<%
i = 0

do while i < allnum
	response.write "<tr><td width='9%' align='center' height='25' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "'></td><td width='91%' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input name='kill" & i & "' type='text' value='" & sh.Get(i) & "' size='70' maxlength='128' class='textbox'></td></tr>"
	i = i + 1
loop

if request("mode") <> "" then
	response.write "<tr><td width='9%' align='center' height='25' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "'></td><td width='91%' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input name='kill" & i & "' type='text' size='70' maxlength='128' class='textbox'></td></tr>"
end if
%>
	<tr>
	<td colspan="2" align="right" bgcolor="#ffffff">
	<br>
	<input type="button" value=" 添加 " onclick="javascript:add()" class="Bsbttn">&nbsp;
	<input type="button" value=" 删除 " onclick="javascript:del()" class="Bsbttn">&nbsp;
	<input type="button" value=" 保存 " onclick="javascript:save()" class="Bsbttn">&nbsp;
	<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
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
		<td width="94%">保留帐号功能, 可确保所设的名称不会被用户公开申请到, 管理员以及域管理员不受此限制.
		<br>此功能可用来保护一些重要的邮箱名称资源(如: master, webmaster, poster, administrator等), 即使现在暂时不使用也不会被用户公开申请到.
		<br>支持通配符方式. (*: 任意长度的任何内容.&nbsp;&nbsp;?: 一个字符的任何内容.)
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10">
        </td>
      </tr>
    </table>
  </div>
  </FORM>
<br>
<div style="position: absolute; left: 50; top: 20;">
<table><tr bgcolor="<%=MY_COLOR_2 %>"><td nowrap style="border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
&nbsp;<a href="webadmin.asp?<%=getGRSN() %>#signhold" style="text-transform: none; text-decoration: none;"><font class="s" color="<%=MY_COLOR_4 %>"><b>启动项设置</b></font>&nbsp;<img src="images\ugo.gif" border="0" align="absbottom"></a>&nbsp;
</td></tr></table>
</div>
</BODY>
</HTML>

<%
set sh = nothing
%>
