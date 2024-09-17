<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim sysinfo
set sysinfo = server.createobject("easymail.sysinfo")
sysinfo.Load

dim ei
set ei = server.createobject("easymail.AuthTrustIP")
'-----------------------------------------
ei.Load

allnum = ei.count


mode = trim(request("mode"))

if mode <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ei.RemoveAll

	i = 0
	if mode = "save" then
		do while i < allnum + 1
			if trim(request("listitem" & i)) <> "" then
				ei.Add trim(request("listitem" & i))
			end if

		    i = i + 1
		loop
	elseif mode = "add" then
		do while i < allnum + 1
			if trim(request("listitem" & i)) <> "" then
				ei.Add trim(request("listitem" & i))
			end if

		    i = i + 1
		loop
	elseif mode = "del" then
		do while i < allnum + 1
			if trim(request("check" & i)) = "" and trim(request("listitem" & i)) <> "" then
				ei.Add trim(request("listitem" & i))
			end if

		    i = i + 1
		loop
	end if

	ei.Save

	if trim(request("EnableAuthTrustIP")) <> "" then
		if sysinfo.EnableAuthTrustIP = false then
			sysinfo.EnableAuthTrustIP = true
			sysinfo.Save
		end if
	else
		if sysinfo.EnableAuthTrustIP = true then
			sysinfo.EnableAuthTrustIP = false
			sysinfo.Save
		end if
	end if

	if trim(request("mode")) <> "add" then
		set ei = nothing
		set sysinfo = nothing
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=authtrustip.asp"
	end if
end if

allnum = ei.count
%>

<HTML>
<HEAD>
<%=s_lang_meta %>
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
<FORM ACTION="authtrustip.asp" METHOD="POST" NAME="f1">
<input type="hidden" name="mode">
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="2%" height="25">&nbsp;</td>
	<td width="74%"><input type="checkbox" id="EnableAuthTrustIP" name="EnableAuthTrustIP" value="checkbox"<% if sysinfo.EnableAuthTrustIP = true then Response.Write " checked"%>><%=s_lang_0028 %></td>
	<td width="24%" nowrap><b><%=s_lang_0026 %></b></td>
    </tr>
  </table>
<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr>
      <td colspan="2" align="right" bgcolor="#ffffff"> 
	<br>
	<input type="button" value=" <%=s_lang_add %> " onclick="javascript:add()" class="Bsbttn">&nbsp;
	<input type="button" value=" <%=s_lang_del %> " onclick="javascript:del()" class="Bsbttn">&nbsp;
	<input type="button" value=" <%=s_lang_save %> " onclick="javascript:save()" class="Bsbttn">&nbsp;
	<input type="button" value=" <%=s_lang_return %> " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	<br>&nbsp;
	</td>
    </tr>
	<tr bgcolor="<%=MY_COLOR_2 %>">
      <td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0026 %></b></font></div>
      </td>
    </tr>
<%
i = 0

do while i < allnum
	response.write "<tr><td width='9%' align='center' height='25' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "'></td><td width='91%' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input name='listitem" & i & "' type='text' value='" & ei.Get(i) & "' size='70' maxlength='64' class='textbox'></td></tr>"
	i = i + 1
loop

if request("mode") <> "" then
	response.write "<tr><td width='9%' align='center' height='25' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "'></td><td width='91%' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input name='listitem" & i & "' type='text' size='70' maxlength='32' class='textbox'></td></tr>"
end if
%>
    <tr> 
      <td colspan="2" align="right" bgcolor="#ffffff"> 
	<br>
	<input type="button" value=" <%=s_lang_add %> " onclick="javascript:add()" class="Bsbttn">&nbsp;
	<input type="button" value=" <%=s_lang_del %> " onclick="javascript:del()" class="Bsbttn">&nbsp;
	<input type="button" value=" <%=s_lang_save %> " onclick="javascript:save()" class="Bsbttn">&nbsp;
	<input type="button" value=" <%=s_lang_return %> " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
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
        <td width="94%"><%=s_lang_0027 %>.<br><br><%=s_lang_tpf %>
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
set sysinfo = nothing
%>
