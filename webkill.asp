<!--#include file="passinc.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.WebKill")
'-----------------------------------------
ei.Load

allnum = ei.count


mode = trim(request("mode"))

if mode <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ei.RemoveAll

	i = 0
	if mode = "save" then
		do while i < allnum + 1
			if trim(request("kill" & i)) <> "" then
				ei.Add trim(request("kill" & i))
			end if

		    i = i + 1
		loop
	elseif mode = "add" then
		do while i < allnum + 1
			if trim(request("kill" & i)) <> "" then
				ei.Add trim(request("kill" & i))
			end if

		    i = i + 1
		loop
	elseif mode = "del" then
		do while i < allnum + 1
			if trim(request("check" & i)) = "" and trim(request("kill" & i)) <> "" then
				ei.Add trim(request("kill" & i))
			end if

		    i = i + 1
		loop
	end if

	ei.Save


	if trim(request("mode")) <> "add" then
		set ei = nothing
		response.redirect "ok.asp?" & getGRSN() & "&gourl=webkill.asp"
	end if
end if

allnum = ei.count
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
<FORM ACTION="webkill.asp" METHOD="POST" NAME="f1">
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
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>拒绝Web登录设置 (请输入IP地址)</b></font></div>
      </td>
    </tr>
<%
i = 0

do while i < allnum
	response.write "<tr><td width='9%' align='center' height='25' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "'></td><td width='91%' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input name='kill" & i & "' type='text' value='" & ei.Get(i) & "' size='70' maxlength='64' class='textbox'></td></tr>"
	i = i + 1
loop

if request("mode") <> "" then
	response.write "<tr><td width='9%' align='center' height='25' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "'></td><td width='91%' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input name='kill" & i & "' type='text' size='70' maxlength='32' class='textbox'></td></tr>"
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
        <td width="94%">您可以通过输入IP地址或是IP段 (如: 61.121.*) 的方式拒绝来自这些IP对WebMail的访问 (注意: 如输入"*"时, 将拒绝所有IP对WebMail的访问, 但管理员帐号将不受此限制). 
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
</BODY>
</HTML>

<%
set ei = nothing
%>
