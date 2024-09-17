<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.domain")
'-----------------------------------------
ei.Load

allnum = ei.getcount

d_move = trim(request("move"))
d_id = trim(request("id"))

if d_id <> "" and IsNumeric(d_id) = true then
	if d_move = "up" then
		ei.moveUP CLng(d_id)
		ei.Save
	elseif d_move = "down" then
		ei.moveDOWN CLng(d_id)
		ei.Save
	end if

	set ei = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=showdomain.asp"
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

function upit(dmid) {
	location.href = "showdomain.asp?move=up&<%=getGRSN() %>&id=" + dmid;
}

function downit(dmid) {
	location.href = "showdomain.asp?move=down&<%=getGRSN() %>&id=" + dmid;
}

function findDomain() {
	if (document.f1.searchstr.value != "")
	{
		var isfind = false;
		var i = 0;
		var theObj;

		for (;i < <%=allnum %>; i++)
		{
			theObj = eval("document.f1.domain" + i);
			if (theObj.value.length == document.f1.searchstr.value.length && theObj.value.toLowerCase().indexOf(document.f1.searchstr.value.toLowerCase(), 0) == 0)
			{
				isfind = true;
				theObj.focus();
				theObj.select();
			}
		}

		if (isfind == false)
			alert("域名没有找到.");
	}
}
//-->
</SCRIPT>


<BODY>
<br>
<FORM ACTION="savedomain.asp" METHOD="POST" NAME="f1">
	<input type="hidden" name="mode">
	<table width="95%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr>
      <td colspan="2" align="left" bgcolor="#ffffff">
	<table width="100%">
	<tr><td width="50%" height="40">
	<input type="text" name="searchstr" class="textbox" size="20">
	<input type="button" value="域查找" onclick="javascript:findDomain();" class="sbttn">
	</td>
	<td align="right">
	<input type="button" value=" 添加 " onclick="javascript:add()" class="Bsbttn">&nbsp;
	<input type="button" value=" 删除 " onclick="javascript:del()" class="Bsbttn">&nbsp;
	<input type="button" value=" 保存 " onclick="javascript:save()" class="Bsbttn">&nbsp;
	<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td></tr>
	</table>
	</td>
    </tr>
    <tr bgcolor="<%=MY_COLOR_2 %>">
      <td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>域名管理</b></font></div>
      </td>
    </tr>
<%
i = 0

do while i < allnum
	Response.Write "<tr><td width='9%' align='center' height='25' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "'></td><td width='91%' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input name='domain" & i & "' id='domain" & i & "' type='text' value='" & ei.GetDomain(i) & "' size='65' maxlength='64' class='textbox'>"

	if allnum = 1 then
		Response.Write "&nbsp;&nbsp;&nbsp;(<font color=""#FF3333"">主域</font>)</td></tr>" & Chr(13)
	elseif i = 0 then
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='javascript:downit(" & i & ")'><img src='images\arrow_down.gif' border='0' align='absmiddle' alt='下移'></a>&nbsp;&nbsp;(<font color=""#FF3333"">主域</font>)</td></tr>" & Chr(13)
	elseif i = allnum - 1 then
		Response.Write "&nbsp;&nbsp;&nbsp;<a href='javascript:upit(" & i & ")'><img src='images\arrow_up.gif' border='0' align='absmiddle' alt='上移'></a>&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>" & Chr(13)
	else
		Response.Write "&nbsp;&nbsp;&nbsp;<a href='javascript:upit(" & i & ")'><img src='images\arrow_up.gif' border='0' align='absmiddle' alt='上移'></a>&nbsp;&nbsp;<a href='javascript:downit(" & i & ")'><img src='images\arrow_down.gif' border='0' align='absmiddle' alt='下移'></a></td></tr>" & Chr(13)
	end if

	i = i + 1
loop

if request("mode") <> "" then
	response.write "<tr><td width='9%' align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "'></td><td width='91%' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input name='domain" & i & "' type='text' size='65' maxlength='64' class='textbox'></td></tr>"
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
  </FORM>
<br><br><br>
  <div align="center">
    <table width="95%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr> 
        <td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
        <td width="94%"><%=s_lang_0031 %>.
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
set ei = nothing
%>
