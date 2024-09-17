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

mode = trim(request("wmode"))

if mode <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim gltip
	set gltip = server.createobject("easymail.GLTrustIP")
	gltip.Load

	if mode = "save" then
		if trim(request("EnableGreylisting")) <> "" then
			sysinfo.EnableGreylisting = true
		else
			sysinfo.EnableGreylisting = false
		end if

		if IsNumeric(trim(request("GL_RejectMinutes"))) = true then
			sysinfo.GL_RejectMinutes = CLng(trim(request("GL_RejectMinutes")))
		end if

		if IsNumeric(trim(request("GL_LiveHours"))) = true then
			sysinfo.GL_LiveHours = CLng(trim(request("GL_LiveHours")))
		end if

		if IsNumeric(trim(request("GL_DynamicTrustIP_LiveDays"))) = true then
			sysinfo.GL_DynamicTrustIP_LiveDays = CLng(trim(request("GL_DynamicTrustIP_LiveDays")))
		end if

		if IsNumeric(trim(request("GL_DynamicBadIP_LiveDays"))) = true then
			sysinfo.GL_DynamicBadIP_LiveDays = CLng(trim(request("GL_DynamicBadIP_LiveDays")))
		end if

		sysinfo.Save
	elseif mode = "cleantip" then
		gltip.RemoveAll_GL_DynamicTrustIP
		gltip.Save
	elseif mode = "cleanbip" then
		gltip.RemoveAll_GL_DynamicBadIP
		gltip.Save
	end if

	set sysinfo = nothing
	set gltip = nothing

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=greylisting.asp"
end if
%>

<HTML>
<HEAD>
<%=s_lang_meta %>
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
function cleanTIP() {
	if (confirm("<%=s_lang_0001 %>") == false)
		return ;

	document.f1.wmode.value = "cleantip";
	document.f1.submit();
}

function cleanBIP() {
	if (confirm("<%=s_lang_0002 %>") == false)
		return ;

	document.f1.wmode.value = "cleanbip";
	document.f1.submit();
}

function save() {
	document.f1.wmode.value = "save";
	document.f1.submit();
}

function set_default() {
	document.f1.GL_RejectMinutes.value = 5;
	document.f1.GL_LiveHours.value = 3;
	document.f1.GL_DynamicTrustIP_LiveDays.value = 3;
	document.f1.GL_DynamicBadIP_LiveDays.value = 7;
}

function set_tip() {
	location.href = "gl_tip.asp";
}

function set_tsender() {
	location.href = "gl_tsender.asp";
}
//-->
</SCRIPT>


<BODY>
<br>
<FORM ACTION="greylisting.asp" METHOD="POST" NAME="f1">
<input type="hidden" name="wmode">
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="2%" height="25">&nbsp;</td>
	<td width="47%"><input type="checkbox" id="EnableGreylisting" name="EnableGreylisting" value="checkbox"<% if sysinfo.EnableGreylisting = true then Response.Write " checked"%>><%=s_lang_0003 %></td>
	<td width="29%"><a href="right.asp?<%=getGRSN() %>"><b><%=s_lang_return %></b></a></td>
	<td width="22%" nowrap><b><%=s_lang_0005 %></b></td>
    </tr>
  </table>
<br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr><td height="28" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;&nbsp;<%=s_lang_0006 %>:
	<input type="text" name="GL_RejectMinutes" id="GL_RejectMinutes" class="textbox" size="5" maxlength="2" value="<%=sysinfo.GL_RejectMinutes %>">&nbsp;<%=s_lang_minute %>
	</td>
    </tr>
	<tr><td height="28" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;&nbsp;<%=s_lang_0007 %>:
	<input type="text" name="GL_LiveHours" id="GL_LiveHours" class="textbox" size="5" maxlength="1" value="<%=sysinfo.GL_LiveHours %>">&nbsp;<%=s_lang_hour %>
	</td>
    </tr>
	<tr><td height="28" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;&nbsp;<%=s_lang_0008 %>:
	<input type="text" name="GL_DynamicTrustIP_LiveDays" id="GL_DynamicTrustIP_LiveDays" class="textbox" size="5" maxlength="1" value="<%=sysinfo.GL_DynamicTrustIP_LiveDays %>">&nbsp;<%=s_lang_day %>
	</td>
    </tr>
	<tr><td height="28" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;&nbsp;<%=s_lang_0009 %>:
	<input type="text" name="GL_DynamicBadIP_LiveDays" id="GL_DynamicBadIP_LiveDays" class="textbox" size="5" maxlength="2" value="<%=sysinfo.GL_DynamicBadIP_LiveDays %>">&nbsp;<%=s_lang_day %>
	</td>
    </tr>
	<tr><td align="center" height="32" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	<input type="button" value="<%=s_lang_0012 %>" class="sbttn" style="WIDTH: 96px" language=javascript onClick="set_default()">
	</td>
    </tr>
	<tr><td align="center" height="32" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	<input type="button" value="<%=s_lang_0013 %>" class="sbttn" style="WIDTH: 130px" language=javascript onClick="set_tip()">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="button" value="<%=s_lang_0014 %>" class="sbttn" style="WIDTH: 130px" language=javascript onClick="set_tsender()">
	</td>
    </tr>
	<tr><td align="center" height="32" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	<input type="button" value="<%=s_lang_0004 %>" class="sbttn" style="WIDTH: 130px" language=javascript onClick="cleanTIP()">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="button" value="<%=s_lang_0010 %>" class="sbttn" style="WIDTH: 130px" language=javascript onClick="cleanBIP()">
	</td>
    </tr>
    <tr> 
      <td height="50" colspan="2" align="right" bgcolor="#ffffff" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	<br>
	<input type="button" value=" <%=s_lang_save %> " class="Bsbttn" language=javascript onClick="save()">&nbsp;
	<input type="button" value=" <%=s_lang_return %> " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
      </td>
    </tr>
  </table>
  </FORM>
<br><br><br>
  <div align="center">
    <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr> 
        <td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
		<td width="94%"><%=s_lang_0011 %>
		<br>
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
set sysinfo = nothing
%>
