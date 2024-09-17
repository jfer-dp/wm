<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
if isadmin() = false then
	Response.Redirect "noadmin.asp"
end if
%>

<%
gourl = trim(request("gourl"))

dim ei
set ei = server.createobject("easymail.UserWorkTimer")
ei.Load_Templet


if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if trim(request("setdefault")) = "true" then
		ei.New_Templet
		set ei = nothing

		if gourl = "" then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("uwt.asp")
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("uwt.asp") & "&pgourl=" & Server.URLEncode(gourl)
		end if
	end if


	if trim(request("is_update_disabled_user")) <> "" then
		ei.is_update_disabled_user = true
	else
		ei.is_update_disabled_user = false
	end if

	if trim(request("is_update_limitout")) <> "" then
		ei.is_update_limitout = true
	else
		ei.is_update_limitout = false
	end if

	if trim(request("is_update_sendout_day_limit")) <> "" then
		ei.is_update_sendout_day_limit = true
	else
		ei.is_update_sendout_day_limit = false
	end if

	if trim(request("is_update_disabled_auto_reply")) <> "" then
		ei.is_update_disabled_auto_reply = true
	else
		ei.is_update_disabled_auto_reply = false
	end if

	if trim(request("is_update_disabled_auto_forward")) <> "" then
		ei.is_update_disabled_auto_forward = true
	else
		ei.is_update_disabled_auto_forward = false
	end if

	if trim(request("is_update_password_strong")) <> "" then
		ei.is_update_password_strong = true
	else
		ei.is_update_password_strong = false
	end if

	if trim(request("is_update_auto_move_to_arc")) <> "" then
		ei.is_update_auto_move_to_arc = true
	else
		ei.is_update_auto_move_to_arc = false
	end if

	if trim(request("is_update_user_in_monitor")) <> "" then
		ei.is_update_user_in_monitor = true
	else
		ei.is_update_user_in_monitor = false
	end if

	if trim(request("is_update_user_out_monitor")) <> "" then
		ei.is_update_user_out_monitor = true
	else
		ei.is_update_user_out_monitor = false
	end if

	if trim(request("Enable_In_Monitor")) <> "" then
		ei.Enable_In_Monitor = true
	else
		ei.Enable_In_Monitor = false
	end if

	if trim(request("Enable_Out_Monitor")) <> "" then
		ei.Enable_Out_Monitor = true
	else
		ei.Enable_Out_Monitor = false
	end if

	if trim(request("Enable_Out_Monitor_Only_OutSystem")) <> "" then
		ei.Enable_Out_Monitor_Only_OutSystem = true
	else
		ei.Enable_Out_Monitor_Only_OutSystem = false
	end if

	if IsNumeric(trim(request("In_Monitor_KSize"))) = true then
		ei.In_Monitor_KSize = CLng(trim(request("In_Monitor_KSize")))
	end if

	if IsNumeric(trim(request("In_Monitor_SaveDays"))) = true then
		ei.In_Monitor_SaveDays = CLng(trim(request("In_Monitor_SaveDays")))
	end if

	if IsNumeric(trim(request("Out_Monitor_KSize"))) = true then
		ei.Out_Monitor_KSize = CLng(trim(request("Out_Monitor_KSize")))

		if ei.Out_Monitor_KSize < 1 then
			ei.Out_Monitor_KSize = 1
		end if
	end if

	if IsNumeric(trim(request("Out_Monitor_SaveDays"))) = true then
		ei.Out_Monitor_SaveDays = CLng(trim(request("Out_Monitor_SaveDays")))
	end if

	ei.disabled_user_over = Remove_ZG(trim(request("disabled_user_over")))

	ei.limitout_over = Remove_ZG(trim(request("limitout_over")))

	ei.sendout_day_limit_over = Remove_ZG(trim(request("sendout_day_limit_over")))

	ei.disabled_auto_reply_over = Remove_ZG(trim(request("disabled_auto_reply_over")))

	ei.disabled_auto_forward_over = Remove_ZG(trim(request("disabled_auto_forward_over")))

	ei.user_in_monitor_over = Remove_ZG(trim(request("user_in_monitor_over")))

	ei.user_out_monitor_over = Remove_ZG(trim(request("user_out_monitor_over")))

	if IsNumeric(trim(request("sendout_day_limit_mailnum"))) = true then
		ei.sendout_day_limit_mailnum = CLng(trim(request("sendout_day_limit_mailnum")))
	end if

	if IsNumeric(trim(request("password_strong"))) = true then
		ei.password_strong = CLng(trim(request("password_strong")))
	end if

	if IsNumeric(trim(request("auto_move_to_arc"))) = true then
		ei.auto_move_to_arc = CLng(trim(request("auto_move_to_arc")))
	end if

	ei.Save_Templet
	set ei = nothing

	if gourl = "" then
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("uwt.asp")
	else
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("uwt.asp") & "&pgourl=" & Server.URLEncode(gourl)
	end if
end if
%>

<HTML>
<HEAD>
<%=s_lang_meta %>
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
<SCRIPT language=javascript src="images/cal/popcalendar.js"></SCRIPT>
<SCRIPT Language="JavaScript">dateFormat='yyyy-mm-dd'</SCRIPT>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
function godef() {
	document.f1.setdefault.value = "true"
	document.f1.submit();
}

function gosub(){
	if (document.f1.Enable_Out_Monitor.checked == false)
		document.f1.user_out_monitor_over.value = "0";
	else if (document.f1.user_out_monitor_over_forever.checked == true)
		document.f1.user_out_monitor_over.value = "1";

	if (document.f1.user_out_monitor_over.value.length == 0)
		document.f1.user_out_monitor_over.value = "1";


	if (document.f1.Enable_In_Monitor.checked == false)
		document.f1.user_in_monitor_over.value = "0";
	else if (document.f1.user_in_monitor_over_forever.checked == true)
		document.f1.user_in_monitor_over.value = "1";

	if (document.f1.user_in_monitor_over.value.length == 0)
		document.f1.user_in_monitor_over.value = "1";

	if (document.f1.In_Monitor_KSize_no_limit.checked == true)
		document.f1.In_Monitor_KSize.value = "0";


	if (document.f1.disabled_auto_reply_over_on_off.checked == false)
		document.f1.disabled_auto_reply_over.value = "0";
	else if (document.f1.disabled_auto_reply_over_forever.checked == true)
		document.f1.disabled_auto_reply_over.value = "1";

	if (document.f1.disabled_auto_reply_over.value.length == 0)
		document.f1.disabled_auto_reply_over.value = "1";


	if (document.f1.disabled_auto_forward_over_on_off.checked == false)
		document.f1.disabled_auto_forward_over.value = "0";
	else if (document.f1.disabled_auto_forward_over_forever.checked == true)
		document.f1.disabled_auto_forward_over.value = "1";

	if (document.f1.disabled_auto_forward_over.value.length == 0)
		document.f1.disabled_auto_forward_over.value = "1";


	if (document.f1.sendout_day_limit_over_on_off.checked == false)
		document.f1.sendout_day_limit_over.value = "0";
	else if (document.f1.sendout_day_limit_over_forever.checked == true)
		document.f1.sendout_day_limit_over.value = "1";

	if (document.f1.sendout_day_limit_over.value.length == 0)
		document.f1.sendout_day_limit_over.value = "1";


	if (document.f1.limitout_over_on_off.checked == false)
		document.f1.limitout_over.value = "0";
	else if (document.f1.limitout_over_forever.checked == true)
		document.f1.limitout_over.value = "1";

	if (document.f1.limitout_over.value.length == 0)
		document.f1.limitout_over.value = "1";


	if (document.f1.disabled_user_over_on_off.checked == false)
		document.f1.disabled_user_over.value = "0";
	else if (document.f1.disabled_user_over_forever.checked == true)
		document.f1.disabled_user_over.value = "1";

	if (document.f1.disabled_user_over.value.length == 0)
		document.f1.disabled_user_over.value = "1";

	document.f1.submit();
}

function window_onload() {
	init();

	document.f1.password_strong.value = "<%=ei.password_strong %>";
	document.f1.auto_move_to_arc.value = "<%=ei.auto_move_to_arc %>";

	show_it('Enable_Out_Monitor');
	show_it('disabled_auto_reply_over_on_off');
	show_it('disabled_auto_forward_over_on_off');
	show_it('sendout_day_limit_over_on_off');
	show_it('limitout_over_on_off');
	show_it('disabled_user_over_on_off');
	show_it('Enable_In_Monitor');
}

function goback()
{
<% if gourl = "" then %>
	history.back();
<% else %>
	location.href = "<%=gourl %>";
<% end if %>
}

function show_it(name) {
	var show_span = document.getElementById(name + "_span")

	if (document.getElementById(name).checked == true)
		show_span.style.display = "inline";
	else
		show_span.style.display = "none";
}
//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload()">
<FORM ACTION="#" METHOD="POST" NAME="f1">
<input name="setdefault" type="hidden">
<input name="gourl" type="hidden" value="<%=gourl %>">
<br><br>
  <table width="75%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr> 
	<td width="50%" bgcolor="#ffffff">
	<table border="0" cellspacing="0"><tr bgcolor="<%=MY_COLOR_2 %>"><td nowrap height="24" style="border:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0227 %></b></font>&nbsp;
	</td></tr></table>
	</td>
      <td colspan="2" align="right" bgcolor="#ffffff">
	<input type="button" value=" <%=s_lang_save %> " LANGUAGE=javascript onclick="gosub()" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value="<%=s_lang_0228 %>" LANGUAGE=javascript onclick="godef()" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" <%=s_lang_return %> " onclick="goback()" class="Bsbttn">
      </td>
    </tr>
  </table>
<br>
<table width="75%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
	<td colspan=2 align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0229 %></b>
	</td></tr>
	<tr>
	<td width="18%" height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" name="is_update_user_out_monitor"<% if ei.is_update_user_out_monitor = true then Response.Write " checked" %>><%=s_lang_0230 %>
	</td>
	<td align=left style="border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" id="Enable_Out_Monitor" name="Enable_Out_Monitor" onclick="javascript:show_it('Enable_Out_Monitor')"<% if ei.Enable_Out_Monitor = true then Response.Write " checked" %>><%=s_lang_0231 %><br>
<span id="Enable_Out_Monitor_span" name="Enable_Out_Monitor_span">
	&nbsp;<input type="checkbox" id="Enable_Out_Monitor_Only_OutSystem" name="Enable_Out_Monitor_Only_OutSystem"<% if ei.Enable_Out_Monitor_Only_OutSystem = true then Response.Write " checked" %>><%=s_lang_0232 %><br>
	&nbsp;<img src="aw.gif" border="0">&nbsp;<%=s_lang_0234 %>&nbsp;<input type="text" name="Out_Monitor_KSize" id="Out_Monitor_KSize" class='textbox' value="<%=ei.Out_Monitor_KSize %>" size="10" maxlength="7">&nbsp;K <%=s_lang_0233 %><br>
	&nbsp;<img src="aw.gif" border="0">&nbsp;<%=s_lang_0235 %>:&nbsp;<input type="text" name="Out_Monitor_SaveDays" id="Out_Monitor_SaveDays" class='textbox' value="<%=ei.Out_Monitor_SaveDays %>" size="5" maxlength="2">&nbsp;<%=s_lang_day %><br>
	&nbsp;<img src="aw.gif" border="0">&nbsp;<%=s_lang_0236 %>: <input type="text" name="user_out_monitor_over" id="user_out_monitor_over" class='textbox' value="<% if IsNull(ei.user_out_monitor_over) = false and Len(ei.user_out_monitor_over) = 8 then Response.Write Show_Som_Date(ei.user_out_monitor_over) %>" readonly size="12">
<script language='javascript'> 
<!--
if (!document.layers) {
	document.write("<img align=absmiddle style='CURSOR:pointer' src='images/cal/calendar.gif' onclick='popUpCalendar(this, document.f1.user_out_monitor_over, dateFormat,-1,-1)' alt='Select'>");
}
//-->
</script>
	&nbsp;<input type="checkbox" id="user_out_monitor_over_forever" name="user_out_monitor_over_forever"<% if ei.user_out_monitor_over = "1" then Response.Write " checked" %>><%=s_lang_0237 %><br>
</span>
	</td></tr>
</table>
<br>
<table width="75%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
	<td colspan=2 align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0238 %></b>
	</td></tr>
	<tr>
	<td width="18%" height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" name="is_update_sendout_day_limit"<% if ei.is_update_sendout_day_limit = true then Response.Write " checked" %>><%=s_lang_0230 %>
	</td>
	<td align=left style="border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" id="sendout_day_limit_over_on_off" name="sendout_day_limit_over_on_off" onclick="javascript:show_it('sendout_day_limit_over_on_off')"<% if ei.sendout_day_limit_over = "1" or Len(ei.sendout_day_limit_over) = 8 then Response.Write " checked" %>><%=s_lang_0238 %><br>
<span id="sendout_day_limit_over_on_off_span" name="sendout_day_limit_over_on_off_span">
	&nbsp;<img src="aw.gif" border="0">&nbsp;<%=s_lang_0239 %>:&nbsp;<input type="text" name="sendout_day_limit_mailnum" id="sendout_day_limit_mailnum" class='textbox' value="<%=ei.sendout_day_limit_mailnum %>" size="6" maxlength="5"><br>
	&nbsp;<img src="aw.gif" border="0">&nbsp;<%=s_lang_0236 %>: <input type="text" name="sendout_day_limit_over" id="sendout_day_limit_over" class='textbox' value="<% if IsNull(ei.sendout_day_limit_over) = false and Len(ei.sendout_day_limit_over) = 8 then Response.Write Show_Som_Date(ei.sendout_day_limit_over) %>" readonly size="12">
<script language='javascript'> 
<!--
if (!document.layers) {
	document.write("<img align=absmiddle style='CURSOR:pointer' src='images/cal/calendar.gif' onclick='popUpCalendar(this, document.f1.sendout_day_limit_over, dateFormat,-1,-1)' alt='Select'>");
}
//-->
</script>
	&nbsp;<input type="checkbox" id="sendout_day_limit_over_forever" name="sendout_day_limit_over_forever"<% if ei.sendout_day_limit_over = "1" then Response.Write " checked" %>><%=s_lang_0237 %><br>
</span>
	</td></tr>
</table>
<br>
<table width="75%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
	<td colspan=2 align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0240 %></b>
	</td></tr>
	<tr>
	<td width="18%" height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" name="is_update_password_strong"<% if ei.is_update_password_strong = true then Response.Write " checked" %>><%=s_lang_0230 %>
	</td>
	<td align=left style="border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
&nbsp;<select id="password_strong" name="password_strong" class="drpdwn" size="1">
<option value="0"><%=s_lang_0241 %></option>
<option value="1"><%=s_lang_0242 %></option>
<option value="2"><%=s_lang_0243 %></option>
<option value="3"><%=s_lang_0244 %></option>
</select>
	</td></tr>
</table>
<br>
<table width="75%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
	<td colspan=2 align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0245 %></b>
	</td></tr>
	<tr>
	<td width="18%" height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" name="is_update_disabled_auto_reply"<% if ei.is_update_disabled_auto_reply = true then Response.Write " checked" %>><%=s_lang_0230 %>
	</td>
	<td align=left style="border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" id="disabled_auto_reply_over_on_off" name="disabled_auto_reply_over_on_off" onclick="javascript:show_it('disabled_auto_reply_over_on_off')"<% if ei.disabled_auto_reply_over = "1" or Len(ei.disabled_auto_reply_over) = 8 then Response.Write " checked" %>><%=s_lang_0245 %><br>
<span id="disabled_auto_reply_over_on_off_span" name="disabled_auto_reply_over_on_off_span">
	&nbsp;<img src="aw.gif" border="0">&nbsp;<%=s_lang_0236 %>: <input type="text" name="disabled_auto_reply_over" id="disabled_auto_reply_over" class='textbox' value="<% if IsNull(ei.disabled_auto_reply_over) = false and Len(ei.disabled_auto_reply_over) = 8 then Response.Write Show_Som_Date(ei.disabled_auto_reply_over) %>" readonly size="12">
<script language='javascript'> 
<!--
if (!document.layers) {
	document.write("<img align=absmiddle style='CURSOR:pointer' src='images/cal/calendar.gif' onclick='popUpCalendar(this, document.f1.disabled_auto_reply_over, dateFormat,-1,-1)' alt='Select'>");
}
//-->
</script>
	&nbsp;<input type="checkbox" id="disabled_auto_reply_over_forever" name="disabled_auto_reply_over_forever"<% if ei.disabled_auto_reply_over = "1" then Response.Write " checked" %>><%=s_lang_0237 %><br>
</span>
	</td></tr>
</table>
<br>
<table width="75%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
	<td colspan=2 align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0246 %></b>
	</td></tr>
	<tr>
	<td width="18%" height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" name="is_update_disabled_auto_forward"<% if ei.is_update_disabled_auto_forward = true then Response.Write " checked" %>><%=s_lang_0230 %>
	</td>
	<td align=left style="border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" id="disabled_auto_forward_over_on_off" name="disabled_auto_forward_over_on_off" onclick="javascript:show_it('disabled_auto_forward_over_on_off')"<% if ei.disabled_auto_forward_over = "1" or Len(ei.disabled_auto_forward_over) = 8 then Response.Write " checked" %>><%=s_lang_0246 %><br>
<span id="disabled_auto_forward_over_on_off_span" name="disabled_auto_forward_over_on_off_span">
	&nbsp;<img src="aw.gif" border="0">&nbsp;<%=s_lang_0236 %>: <input type="text" name="disabled_auto_forward_over" id="disabled_auto_forward_over" class='textbox' value="<% if IsNull(ei.disabled_auto_forward_over) = false and Len(ei.disabled_auto_forward_over) = 8 then Response.Write Show_Som_Date(ei.disabled_auto_forward_over) %>" readonly size="12">
<script language='javascript'> 
<!--
if (!document.layers) {
	document.write("<img align=absmiddle style='CURSOR:pointer' src='images/cal/calendar.gif' onclick='popUpCalendar(this, document.f1.disabled_auto_forward_over, dateFormat,-1,-1)' alt='Select'>");
}
//-->
</script>
	&nbsp;<input type="checkbox" id="disabled_auto_forward_over_forever" name="disabled_auto_forward_over_forever"<% if ei.disabled_auto_forward_over = "1" then Response.Write " checked" %>><%=s_lang_0237 %><br>
</span>
	</td></tr>
</table>
<br>
<table width="75%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
	<td colspan=2 align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0247 %></b>
	</td></tr>
	<tr>
	<td width="18%" height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" name="is_update_limitout"<% if ei.is_update_limitout = true then Response.Write " checked" %>><%=s_lang_0230 %>
	</td>
	<td align=left style="border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" id="limitout_over_on_off" name="limitout_over_on_off" onclick="javascript:show_it('limitout_over_on_off')"<% if ei.limitout_over = "1" or Len(ei.limitout_over) = 8 then Response.Write " checked" %>><%=s_lang_0247 %><br>
<span id="limitout_over_on_off_span" name="limitout_over_on_off_span">
	&nbsp;<img src="aw.gif" border="0">&nbsp;<%=s_lang_0236 %>: <input type="text" name="limitout_over" id="limitout_over" class='textbox' value="<% if IsNull(ei.limitout_over) = false and Len(ei.limitout_over) = 8 then Response.Write Show_Som_Date(ei.limitout_over) %>" readonly size="12">
<script language='javascript'> 
<!--
if (!document.layers) {
	document.write("<img align=absmiddle style='CURSOR:pointer' src='images/cal/calendar.gif' onclick='popUpCalendar(this, document.f1.limitout_over, dateFormat,-1,-1)' alt='Select'>");
}
//-->
</script>
	&nbsp;<input type="checkbox" id="limitout_over_forever" name="limitout_over_forever"<% if ei.limitout_over = "1" then Response.Write " checked" %>><%=s_lang_0237 %><br>
</span>
	</td></tr>
</table>
<br>
<table width="75%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
	<td colspan=2 align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0248 %></b>
	</td></tr>
	<tr>
	<td width="18%" height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" name="is_update_disabled_user"<% if ei.is_update_disabled_user = true then Response.Write " checked" %>><%=s_lang_0230 %>
	</td>
	<td align=left style="border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" id="disabled_user_over_on_off" name="disabled_user_over_on_off" onclick="javascript:show_it('disabled_user_over_on_off')"<% if ei.disabled_user_over = "1" or Len(ei.disabled_user_over) = 8 then Response.Write " checked" %>><%=s_lang_0248 %><br>
<span id="disabled_user_over_on_off_span" name="disabled_user_over_on_off_span">
	&nbsp;<img src="aw.gif" border="0">&nbsp;<%=s_lang_0236 %>: <input type="text" name="disabled_user_over" id="disabled_user_over" class='textbox' value="<% if IsNull(ei.disabled_user_over) = false and Len(ei.disabled_user_over) = 8 then Response.Write Show_Som_Date(ei.disabled_user_over) %>" readonly size="12">
<script language='javascript'> 
<!--
if (!document.layers) {
	document.write("<img align=absmiddle style='CURSOR:pointer' src='images/cal/calendar.gif' onclick='popUpCalendar(this, document.f1.disabled_user_over, dateFormat,-1,-1)' alt='Select'>");
}
//-->
</script>
	&nbsp;<input type="checkbox" id="disabled_user_over_forever" name="disabled_user_over_forever"<% if ei.disabled_user_over = "1" then Response.Write " checked" %>><%=s_lang_0237 %><br>
</span>
	</td></tr>
</table>
<br>
<table width="75%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
	<td colspan=2 align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0249 %></b>
	</td></tr>
	<tr>
	<td width="18%" height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" name="is_update_user_in_monitor"<% if ei.is_update_user_in_monitor = true then Response.Write " checked" %>><%=s_lang_0230 %>
	</td>
	<td align=left style="border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" id="Enable_In_Monitor" name="Enable_In_Monitor" onclick="javascript:show_it('Enable_In_Monitor')"<% if ei.Enable_In_Monitor = true then Response.Write " checked" %>><%=s_lang_0250 %><br>
<span id="Enable_In_Monitor_span" name="Enable_In_Monitor_span">
	&nbsp;<img src="aw.gif" border="0">&nbsp;<%=s_lang_0234 %>&nbsp;<input type="text" name="In_Monitor_KSize" id="In_Monitor_KSize" class='textbox' value="<% if ei.In_Monitor_KSize > 0 then Response.Write ei.In_Monitor_KSize %>" size="10" maxlength="7">&nbsp;K <%=s_lang_0233 %>&nbsp;&nbsp;<input type="checkbox" id="In_Monitor_KSize_no_limit" name="In_Monitor_KSize_no_limit"<% if ei.In_Monitor_KSize = 0 then Response.Write " checked" %>><%=s_lang_0251 %><br>
	&nbsp;<img src="aw.gif" border="0">&nbsp;<%=s_lang_0235 %>:&nbsp;<input type="text" name="In_Monitor_SaveDays" id="In_Monitor_SaveDays" class='textbox' value="<%=ei.In_Monitor_SaveDays %>" size="5" maxlength="2">&nbsp;<%=s_lang_day %><br>
	&nbsp;<img src="aw.gif" border="0">&nbsp;<%=s_lang_0236 %>: <input type="text" name="user_in_monitor_over" id="user_in_monitor_over" class='textbox' value="<% if IsNull(ei.user_in_monitor_over) = false and Len(ei.user_in_monitor_over) = 8 then Response.Write Show_Som_Date(ei.user_in_monitor_over) %>" readonly size="12">
<script language='javascript'> 
<!--
if (!document.layers) {
	document.write("<img align=absmiddle style='CURSOR:pointer' src='images/cal/calendar.gif' onclick='popUpCalendar(this, document.f1.user_in_monitor_over, dateFormat,-1,-1)' alt='Select'>");
}
//-->
</script>
	&nbsp;<input type="checkbox" id="user_in_monitor_over_forever" name="user_in_monitor_over_forever"<% if ei.user_in_monitor_over = "1" then Response.Write " checked" %>><%=s_lang_0237 %><br>
</span>
	</td></tr>
</table>
<br>
<table width="75%" border="0" align="center" cellspacing="0">
	<tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
	<td colspan=2 align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b>强制自动归档</b>
	</td></tr>
	<tr>
	<td width="18%" height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" name="is_update_auto_move_to_arc"<% if ei.is_update_auto_move_to_arc = true then Response.Write " checked" %>><%=s_lang_0230 %>
	</td>
	<td align=left style="border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
<div>&nbsp;<select id="auto_move_to_arc" name="auto_move_to_arc" class="drpdwn" size="1">
<option value="0">关闭强制自动归档功能</option>
<option value="1">超过最大邮件数30%时强制自动归档</option>
<option value="2">超过最大邮件数50%时强制自动归档</option>
<option value="3">超过最大邮件数70%时强制自动归档</option>
</select></div><div style="padding-top:5px;">&nbsp;<font color="666666">(不建议对多数用户开启此项功能，否则用户堆积的邮件可能会将服务器硬盘撑满)</font></div>
	</td></tr>
</table>
<br>
  <table width="75%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr> 
      <td colspan="2" align="right" bgcolor="#ffffff">
	<input type="button" value=" <%=s_lang_save %> " LANGUAGE=javascript onclick="gosub()" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value="<%=s_lang_0228 %>" LANGUAGE=javascript onclick="godef()" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" <%=s_lang_return %> " onclick="goback()" class="Bsbttn">
      </td>
    </tr>
  </table>
	<input type="hidden" name="gourl" value="<%=gourl %>">
  </FORM>
<br>
  <div align="center">
    <table width="75%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
		<tr>
		<td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
		<td width="94%"><%=s_lang_0252 %><br><br>
		</td>
		</tr>
    </table>
  </div>
<br><br>
</BODY>
</HTML>

<%
set ei = nothing

function Show_Som_Date(ostr)
	if Len(ostr) = 8 then
		Show_Som_Date = Mid(ostr, 1, 4) & "-" & Mid(ostr, 5, 2) & "-" & Mid(ostr, 7, 2)
	end if
end function

function Remove_ZG(ostr)
	Remove_ZG = replace(ostr , "-", "")
end function
%>
