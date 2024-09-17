<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" -->

<%
dim si
set si = server.createobject("easymail.sysinfo")
si.Load
is_EnableAutoReply = si.EnableAutoReply
set si = nothing

dim ei
set ei = server.createobject("easymail.UserMessages")
ei.Load Session("wem")

ei.GetAutoReplyDateLimit sy, sm, sd, ey, em, ed
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<STYLE type=text/css>
<!--
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.cont_td {white-space:nowrap; height:28px; padding-left:14px; padding-right:4px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function window_onload()
{
<%
if is_EnableAutoReply = true then
%>
	document.f1.sy.value = "<%=sy %>";
	document.f1.sm.value = "<%=sm %>";
	document.f1.sd.value = "<%=sd %>";
	document.f1.ey.value = "<%=ey %>";
	document.f1.em.value = "<%=em %>";
	document.f1.ed.value = "<%=ed %>";
<%
end if
%>
}

function changetime() {
	if (document.getElementById("checkauto").checked == false)
	{
		f1.sy.value = "0";
		f1.sm.value = "0";
		f1.sd.value = "0";
		f1.ey.value = "0";
		f1.em.value = "0";
		f1.ed.value = "0";
	}
}

function gosub()
{
<%
if is_EnableAutoReply = true then
%>
	if (document.f1.text.value.length > 4090)
		document.f1.text.value = document.f1.text.value.substring(0, 4090);

	var isok = false;
	var nowdate = new Date(<%=Year(now()) & "," & Month(now()) & "," & Day(now()) %>);
	var sdate = new Date(f1.sy.value, f1.sm.value, f1.sd.value);
	var edate = new Date(f1.ey.value, f1.em.value, f1.ed.value);

	if (document.getElementById("checkauto").checked == true)
	{
		if (f1.sy.value == "0" && f1.sm.value == "0" && f1.sd.value == "0" && f1.ey.value == "0" && f1.em.value == "0" && f1.ed.value == "0")
			isok = true;
		else
		{
			if (edate >= nowdate && edate >= sdate)
				isok = true;
			else
				alert("<%=b_lang_345 %>");
		}
	}
	else
	{
		isok = true;

		if (edate >= nowdate && edate >= sdate)
			document.getElementById("checkauto").checked = true;
	}

	if (isok == true)
		document.f1.submit();
<%
else
%>
	document.f1.submit();
<%
end if
%>
}
//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload()">
<FORM ACTION="saveusersetup.asp" METHOD="POST" NAME="f1">
<%
if is_EnableAutoReply = true then
%>
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_176 %>
</td></tr>
<tr><td class="block_top_td" style="height:4px; _height:6px;"></td></tr>
<tr><td align="left" class="cont_td" style="border-bottom:1px solid #A5B6C8; padding-left:10px;">
	<input type="checkbox" name="checkauto" id="checkauto" onclick="return changetime()" <%
if ei.UseAutoReply = true then
	Response.Write "checked"
end if
%>>
	<%=b_lang_177 %>
	</td></tr>
	<tr><td align="left" class="cont_td" style="border-bottom:1px solid #A5B6C8;">
	<%=b_lang_178 %><%=s_lang_mh %>
<select name="sy" class="drpdwn">
<option value="0">----</option>
<%
	now_temp = Year(Now()) - 1

	i = now_temp
	do while i < now_temp + 10
		Response.Write "<option value='" & i & "'>" & i & "</option>"
		i = i + 1
	loop
%>
</select>
<select name="sm" class="drpdwn">
<option value="0">--</option>
<%
	i = 1
	do while i < 13
		Response.Write "<option value='" & i & "'>" & i & "</option>"
		i = i + 1
	loop
%>
</select>
<select name="sd" class="drpdwn">
<option value="0">--</option>
<%
	i = 1
	do while i < 32
		Response.Write "<option value='" & i & "'>" & i & "</option>"
		i = i + 1
	loop
%>
</select>
-
<select name="ey" class="drpdwn">
<option value="0">----</option>
<%
	now_temp = Year(Now()) - 1

	i = now_temp
	do while i < now_temp + 10
		Response.Write "<option value='" & i & "'>" & i & "</option>"
		i = i + 1
	loop
%>
</select>
<select name="em" class="drpdwn">
<option value="0">--</option>
<%
	i = 1
	do while i < 13
		Response.Write "<option value='" & i & "'>" & i & "</option>"
		i = i + 1
	loop
%>
</select>
<select name="ed" class="drpdwn">
<option value="0">--</option>
<%
	i = 1
	do while i < 32
		Response.Write "<option value='" & i & "'>" & i & "</option>"
		i = i + 1
	loop
%>
</select>
	</td></tr>

	<tr><td align="left" class="cont_td">
	<%=b_lang_065 %><%=s_lang_mh %>
	<input name="subject" type="text" size="60" value="<%=ei.AutoReplySubject %>" class="n_textbox">
	</td></tr>
	<tr><td align="left" class="cont_td" style="border-bottom:1px solid #A5B6C8; padding-bottom:4px;">
	<textarea name="text" cols="80" rows="8" class="n_textarea"><%=ei.AutoReplyText %></textarea>
	</td></tr>
</table>
<br><br>
<%
end if
%>

<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_179 %>
</td></tr>
<tr><td class="block_top_td" style="height:4px; _height:6px;"></td></tr>

<tr><td align="left" class="cont_td" style="border-bottom:1px solid #A5B6C8; padding-left:10px;">
	<input type="checkbox" name="checkautoforward" <%
if ei.UseAutoForward = true then
	Response.Write "checked"
end if
%>>
	<%=b_lang_180 %>
	</td></tr>
	<tr><td align="left" class="cont_td" style="border-bottom:1px solid #A5B6C8; padding-left:10px;">
	<input type="checkbox" name="checklocalsave" <%
if ei.LocalSave = true then
	Response.Write "checked"
end if
%>>
	<%=b_lang_181 %>
	</td></tr>
	<tr><td align="left" class="cont_td" style="border-bottom:1px solid #A5B6C8;">
	<%=b_lang_182 %><%=s_lang_mh %>
	<input type="text" name="AutoForwardTo" size="20" maxlength="128" value="<%=ei.AutoForwardTo %>" class="n_textbox">
	&nbsp;&nbsp;<font color="#901111">*</font><font color="#444444"><%=b_lang_183 %></font>
	</td></tr>

<tr><td class="block_top_td" style="height:4px;"></td></tr>

<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px;">
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
</td></tr>
</table>
</FORM>

<div style="position:absolute; left:12px; top:10px;">
<a href="help.asp#showusersetup" target="_blank"><img src="images/help.gif" border="0" title="<%=s_lang_help %>"></a></div>
</BODY>
</HTML>

<%
sy = NULL
sm = NULL
sd = NULL
ey = NULL
em = NULL
ed = NULL

set ei = nothing
%>
