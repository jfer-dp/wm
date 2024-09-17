<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
if isadmin() = false then
	Response.Redirect "noadmin.asp"
end if
%>

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

gourl = trim(request("gourl"))
user = trim(request("user"))

dim ei
set ei = server.createobject("easymail.UserWorkTimer")
ei.Load_User user
ei.Load_User_Filter user

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim mt
	set mt = server.createobject("easymail.WMethod")

	set app_em = Application("em")

	if trim(request("settmp")) = "true" then
		ei.Load_Templet
		ei.Set_Templet_To_User user

		if ei.is_update_disabled_user = true then
			if mt.Is_Disabled_User(user) = false then
				if ei.disabled_user_over = "1" or Len(ei.disabled_user_over) = 8 then
					app_em.ForbidUserByName user, true
				end if
			else
				if ei.disabled_user_over = "0" then
					app_em.ForbidUserByName user, false
				end if
			end if
		end if

		if ei.is_update_limitout = true then
			if mt.Is_Limitout_User(user) = false then
				if ei.limitout_over = "1" or Len(ei.limitout_over) = 8 then
					app_em.SetLimitOut user, true
				end if
			else
				if ei.limitout_over = "0" then
					app_em.SetLimitOut user, false
				end if
			end if
		end if

		set app_em = nothing
		set mt = nothing
		set ei = nothing

		if gourl = "" then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("uwtuser.asp?user=" & user)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("uwtuser.asp?user=" & user) & "&pgourl=" & Server.URLEncode(gourl)
		end if
	end if

	ei.RemoveAll_User_Filter
	ei.Add_User_Filter trim(request("allmsgs"))
	ei.Save_User_Filter

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

	if IsNumeric(trim(request("mail_filter_ksize"))) = true then
		ei.mail_filter_ksize = CLng(trim(request("mail_filter_ksize")))
	end if


	if ei.disabled_user_over <> Remove_ZG(trim(request("disabled_user_over"))) then
		ei.disabled_user_over = Remove_ZG(trim(request("disabled_user_over")))

		if mt.Is_Disabled_User(user) = false then
			if ei.disabled_user_over = "1" or Len(ei.disabled_user_over) = 8 then
				app_em.ForbidUserByName user, true
			end if
		else
			if ei.disabled_user_over = "0" then
				app_em.ForbidUserByName user, false
			end if
		end if
	end if


	if ei.limitout_over <> Remove_ZG(trim(request("limitout_over"))) then
		ei.limitout_over = Remove_ZG(trim(request("limitout_over")))

		if mt.Is_Limitout_User(user) = false then
			if ei.limitout_over = "1" or Len(ei.limitout_over) = 8 then
				app_em.SetLimitOut user, true
			end if
		else
			if ei.limitout_over = "0" then
				app_em.SetLimitOut user, false
			end if
		end if
	end if


	ei.mail_filter_over = Remove_ZG(trim(request("mail_filter_over")))

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

	ei.Save_User

	set app_em = nothing
	set mt = nothing
	set ei = nothing

	if gourl = "" then
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("uwtuser.asp?user=" & user)
	else
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("uwtuser.asp?user=" & user) & "&pgourl=" & Server.URLEncode(gourl)
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


	if (document.f1.mail_filter_over_on_off.checked == false)
		document.f1.mail_filter_over.value = "0";
	else if (document.f1.mail_filter_over_forever.checked == true)
		document.f1.mail_filter_over.value = "1";

	if (document.f1.mail_filter_over.value.length == 0)
		document.f1.mail_filter_over.value = "1";

	if (document.f1.mail_filter_ksize_no_limit.checked == true)
		document.f1.mail_filter_ksize.value = "0";


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

	var tempstr = "";
	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		tempstr = tempstr + document.f1.listall[i].value + "\f";
	}

	document.f1.allmsgs.value = tempstr;

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
	show_it('mail_filter_over_on_off');
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

function select_content_onchange()
{
	var i = 0;
	for (i; i < document.f1.s_mode.length; i++)
	{
		document.f1.s_mode.remove(i);
		i--;
	}

	if (document.f1.s_content.value != "3")
	{
		add_s_mode("1", "<%=s_lang_0253 %>");
		add_s_mode("2", "<%=s_lang_0254 %>");
		add_s_mode("3", "<%=s_lang_0255 %>");
		add_s_mode("4", "<%=s_lang_0256 %>");
		add_s_mode("5", "<%=s_lang_0257 %>");
	}
	else
	{
		add_s_mode("1", "<%=s_lang_0258 %>");
		add_s_mode("6", "<%=s_lang_0259 %>");
		add_s_mode("7", "<%=s_lang_0260 %>");
	}
}

function delout()
{
	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].selected == true)
		{
			document.f1.listall.remove(i);

			if (i < document.f1.listall.length)
				document.f1.listall[i].selected = true;
			else
			{
				if (i - 1 >= 0)
					document.f1.listall[i - 1].selected = true;
			}

			break;
		}
	}
}

function add()
{
	if (haveit() == false)
	{
		if (document.f1.s_content.value == 3)
		{
			if (isNaN(parseInt(document.f1.s_msg.value)) == true)
			{
				alert("<%=s_lang_inputerr %>!");
				return ;
			}
			else
				document.f1.s_msg.value = parseInt(document.f1.s_msg.value);
		}

		var oOption = document.createElement("OPTION");
		oOption.text = getFilterStr(document.f1.s_content.value, document.f1.s_mode.value, document.f1.s_msg.value);
		oOption.value = document.f1.s_content.value + "\t" + document.f1.s_mode.value + "\t" + document.f1.s_msg.value;
<%
if isMSIE = true then
%>
		document.f1.listall.add(oOption);
<%
else
%>
		document.f1.listall.appendChild(oOption);
<%
end if
%>
	}
}

function haveit()
{
	var tempstr = document.f1.s_content.value + "\t" + document.f1.s_mode.value + "\t" + document.f1.s_msg.value;

	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].value == tempstr)
			return true;
	}

	return false;
}

function getFilterStr(f_content, f_mode, f_msg)
{
	var retstr;
	if (f_content == 1)
		retstr = "From";
	else if (f_content == 2)
		retstr = "Sender";
	else if (f_content == 3)
		retstr = "Size";
	else if (f_content == 4)
		retstr = "Subject";
	else if (f_content == 5)
		retstr = "Header";
	else if (f_content == 6)
		retstr = "Body";
	else if (f_content == 7)
		retstr = "To";
	else if (f_content == 8)
		retstr = "Cc";
	else if (f_content == 9)
		retstr = "Reply-To";
	else if (f_content == 10)
		retstr = "Boundary";
	else
		return "";

	if (f_mode == 1)
		retstr = retstr + " <%=s_lang_0253 %> ";
	else if (f_mode == 2)
		retstr = retstr + " <%=s_lang_0254 %> ";
	else if (f_mode == 3)
		retstr = retstr + " <%=s_lang_0255 %> ";
	else if (f_mode == 4)
		retstr = retstr + " <%=s_lang_0256 %> ";
	else if (f_mode == 5)
		retstr = retstr + " <%=s_lang_0261 %> ";
	else if (f_mode == 6)
		retstr = retstr + " <%=s_lang_0259 %> ";
	else if (f_mode == 7)
		retstr = retstr + " <%=s_lang_0260 %> ";
	else
		return "";

	if (f_msg.length == 0)
		retstr = retstr + "[Empty]";
	else
		retstr = retstr + f_msg;

	return retstr;
}

function add_s_mode(a_value, a_text)
{
	var oOption = document.createElement("OPTION");
	oOption.text = a_text;
	oOption.value = a_value;
<%
if isMSIE = true then
%>
	document.f1.s_mode.add(oOption);
<%
else
%>
	document.f1.s_mode.appendChild(oOption);
<%
end if
%>
}

function set_sys_tmp() {
	document.f1.settmp.value = "true"
	document.f1.submit();
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
<input name="settmp" type="hidden">
<input type="hidden" name="allmsgs">
<input type="hidden" name="user" value="<%=user %>">
<input name="gourl" type="hidden" value="<%=gourl %>">
<br><br>
  <table width="75%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr> 
	<td width="50%" bgcolor="#ffffff">
	<table border="0" cellspacing="0"><tr bgcolor="<%=MY_COLOR_2 %>"><td nowrap height="24" style="border:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<font class="s" color="<%=MY_COLOR_4 %>"><%=s_lang_0262 %>: <b><%=user %></b></font>&nbsp;
	</td></tr></table>
	</td>
      <td align="right" bgcolor="#ffffff">
	<input type="button" value=" <%=s_lang_save %> " LANGUAGE=javascript onclick="gosub()" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value="<%=s_lang_0263 %>" LANGUAGE=javascript onclick="set_sys_tmp()" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" <%=s_lang_return %> " onclick="goback()" class="Bsbttn">
      </td>
    </tr>
  </table>
<br>
<table width="75%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
	<td align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0229 %></b>
	</td></tr>
	<tr>
	<td align=left style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
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
	<td align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0238 %></b>
	</td></tr>
	<tr>
	<td align=left style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
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
	<td align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0240 %></b>
	</td></tr>
	<tr>
	<td align=left style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
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
	<td align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0245 %></b>
	</td></tr>
	<tr>
	<td align=left style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" id="disabled_auto_reply_over_on_off" name="disabled_auto_reply_over_on_off" onclick="javascript:show_it('disabled_auto_reply_over_on_off')"<% if ei.disabled_auto_reply_over = "1" or Len(ei.disabled_auto_reply_over) = 8 then Response.Write " checked" %>><%=s_lang_0245 %><br>
<span id="disabled_auto_reply_over_on_off_span" name="_span">
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
	<td align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0246 %></b>
	</td></tr>
	<tr>
	<td align=left style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
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
	<td colspan=2 align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0264 %></b>
	</td></tr>
	<tr>
	<td colspan=2 align=left style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" id="mail_filter_over_on_off" name="mail_filter_over_on_off" onclick="javascript:show_it('mail_filter_over_on_off')"<% if ei.mail_filter_over = "1" or Len(ei.mail_filter_over) = 8 then Response.Write " checked" %>><%=s_lang_0265 %><br>
<span id="mail_filter_over_on_off_span" name="mail_filter_over_on_off_span">
	&nbsp;<img src="aw.gif" border="0">&nbsp;<%=s_lang_0266 %>&nbsp;<input type="text" name="mail_filter_ksize" id="mail_filter_ksize" class='textbox' value="<% if ei.mail_filter_ksize > 0 then Response.Write ei.mail_filter_ksize %>" size="10" maxlength="7">&nbsp;K <%=s_lang_0233 %>&nbsp;&nbsp;<input type="checkbox" id="mail_filter_ksize_no_limit" name="mail_filter_ksize_no_limit"<% if ei.mail_filter_ksize = 0 then Response.Write " checked" %>><%=s_lang_0251 %><br>
	&nbsp;<img src="aw.gif" border="0">&nbsp;<%=s_lang_0236 %>: <input type="text" name="mail_filter_over" id="mail_filter_over" class='textbox' value="<% if IsNull(ei.mail_filter_over) = false and Len(ei.mail_filter_over) = 8 then Response.Write Show_Som_Date(ei.mail_filter_over) %>" readonly size="12">
<script language='javascript'> 
<!--
if (!document.layers) {
	document.write("<img align=absmiddle style='CURSOR:pointer' src='images/cal/calendar.gif' onclick='popUpCalendar(this, document.f1.mail_filter_over, dateFormat,-1,-1)' alt='Select'>");
}
//-->
</script>
	&nbsp;<input type="checkbox" id="mail_filter_over_forever" name="mail_filter_over_forever"<% if ei.mail_filter_over = "1" then Response.Write " checked" %>><%=s_lang_0237 %><br>
	<table border="0" cellspacing="0"><tr><td width="500">
	<select name="listall" size="7" class="drpdwn" style="width: 480;">
<%
i = 0
allnum = ei.Count_User_Filter

do while i < allnum
	ei.Get_User_Filter i, as_content, as_mode, as_msg
	Response.Write "<option value=""" & as_content & Chr(9) & as_mode & Chr(9) & server.htmlencode(as_msg) & """>" & server.htmlencode(getFilterStr(as_content, as_mode, as_msg)) & "</option>" & Chr(13)

	as_content = NULL
	as_mode = NULL
	as_msg = NULL

	i = i + 1
loop
%>
	</select>
	</td>
	<td>
	<input type="button" value="<%=s_lang_del %>" class="sbttn" LANGUAGE=javascript onclick="delout()">
	</td></tr>
<tr><td colspan=2 height="30" align="left" nowrap>&nbsp;
<select name="s_content" class=drpdwn LANGUAGE=javascript onchange="select_content_onchange()">
<option value="1">From</option>
<option value="2">Sender</option>
<option value="3">Size</option>
<option value="4">Subject</option>
<% if IsEnterpriseVersion = true then %>
<option value="5">Header</option>
<option value="6">Body</option>
<option value="7">To</option>
<option value="8">Cc</option>
<option value="9">Reply-To</option>
<option value="10">Boundary</option>
<% end if %>
</select>
<select name="s_mode" class=drpdwn>
<option value="1"><%=s_lang_0253 %></option>
<option value="2"><%=s_lang_0254 %></option>
<option value="3"><%=s_lang_0255 %></option>
<option value="4"><%=s_lang_0256 %></option>
<option value="5"><%=s_lang_0257 %></option>
</select>
<input type="input" name="s_msg" class="textbox" size="30" maxlength="100">&nbsp;<input type="button" value=" <%=s_lang_add %> " class="sbttn" LANGUAGE=javascript onclick="add()">
	</td></tr></table>
</span>
	</td></tr>
</table>
<br>
<table width="75%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
	<td align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0247 %></b>
	</td></tr>
	<tr>
	<td align=left style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
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
	<td align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0248 %></b>
	</td></tr>
	<tr>
	<td align=left style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
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
	<td align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b><%=s_lang_0249 %></b>
	</td></tr>
	<tr>
	<td align=left style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
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
	<td align="center" height="28" style="border:1px <%=MY_COLOR_1 %> solid;"><b>强制自动归档</b>
	</td></tr>
	<tr>
	<td align=left style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
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
	<input type="button" value="<%=s_lang_0263 %>" LANGUAGE=javascript onclick="set_sys_tmp()" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" <%=s_lang_return %> " onclick="goback()" class="Bsbttn">
      </td>
    </tr>
  </table>
	<input type="hidden" name="gourl" value="<%=gourl %>">
  </FORM>
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

function getFilterStr(f_content, f_mode, f_msg)
	if f_content = 1 then
		getFilterStr = "From"
	elseif f_content = 2 then
		getFilterStr = "Sender"
	elseif f_content = 3 then
		getFilterStr = "Size"
	elseif f_content = 4 then
		getFilterStr = "Subject"
	elseif f_content = 5 then
		getFilterStr = "Header"
	elseif f_content = 6 then
		getFilterStr = "Body"
	elseif f_content = 7 then
		getFilterStr = "To"
	elseif f_content = 8 then
		getFilterStr = "Cc"
	elseif f_content = 9 then
		getFilterStr = "Reply-To"
	elseif f_content = 10 then
		getFilterStr = "Boundary"
	else
		Exit Function
	end if

	if f_mode = 1 then
		getFilterStr = getFilterStr & " " & s_lang_0253 & " "
	elseif f_mode = 2 then
		getFilterStr = getFilterStr & " " & s_lang_0254 & " "
	elseif f_mode = 3 then
		getFilterStr = getFilterStr & " " & s_lang_0255 & " "
	elseif f_mode = 4 then
		getFilterStr = getFilterStr & " " & s_lang_0256 & " "
	elseif f_mode = 5 then
		getFilterStr = getFilterStr & " " & s_lang_0257 & " "
	elseif f_mode = 6 then
		getFilterStr = getFilterStr & " " & s_lang_0259 & " "
	elseif f_mode = 7 then
		getFilterStr = getFilterStr & " " & s_lang_0260 & " "
	else
		getFilterStr = ""
		Exit Function
	end if

	if f_msg = "" then
		getFilterStr = getFilterStr & "[Empty]"
	else
		getFilterStr = getFilterStr & f_msg
	end if
end function
%>
