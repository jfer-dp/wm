<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
filename = trim(request("filename"))
gourl = trim(request("gourl"))

dim ei
set ei = server.createobject("easymail.RecallInfoManager")

if Len(filename) > 5 and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	mode = trim(request("mode"))

	if mode <> "del" then
		ei.Load Session("wem"), filename
	end if

	isok = true
	if mode = "del" then
		isok = ei.Del(Session("wem"), filename)
	elseif mode = "rccheck" then
		themax = ei.count

		i = 0
		do while i <= themax
			if trim(request("check" & i)) <> "" then
				if Len(recall_str) = 0 then
					recall_str = i
				else
					recall_str = recall_str & "," & i
				end if
			end if 

		    i = i + 1
		loop

		isok = ei.Recall(recall_str)
	else
		isok = ei.Recall(mode)
	end if

	set ei = nothing

	if isok = true then
		Response.Redirect "ok.asp?gourl=" & Server.URLEncode(gourl)
	else
		Response.Redirect "err.asp?gourl=" & Server.URLEncode(gourl)
	end if
end if

ei.Load Session("wem"), filename

allnum = ei.Count
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
td {padding-left:3px; padding-right:3px;}
.td_left {height:25px; background:#EFF7FF; color:#104A7B; text-align:right; white-space:nowrap;}
.td_right {border-bottom:1px solid #8CA5B5;}

.in_td {height:25px; text-align:center; white-space:nowrap; background:#EFF7FF; border-right:1px solid #8CA5B5; border-bottom:1px solid #8CA5B5;}
-->
</STYLE>
</HEAD>

<BODY>
<FORM ACTION="recallinfo.asp" METHOD="POST" name="f1">
<INPUT NAME="mode" TYPE="hidden">
<INPUT NAME="filename" TYPE="hidden" value="<%=filename %>">
<INPUT NAME="gourl" TYPE="hidden" value="<%=gourl %>">
<table width="90%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr><td align="left" height="28" style="padding-left:4px;">
	<a class='wwm_btnDownload btn_blue' href="javascript:back()"><< <%=s_lang_return %></a>&nbsp;
	<a class='wwm_btnDownload btn_blue' href="javascript:recall_check()"><%=s_lang_0143 %></a>&nbsp;
	<a class='wwm_btnDownload btn_blue' href="javascript:recall_all()"><%=s_lang_0144 %></a>&nbsp;
	<a class='wwm_btnDownload btn_blue' href="javascript:del()"><%=s_lang_del %></a>
	</td></tr>
</table>
<br>
<table width="90%" border="0" align="center" cellspacing="0" style='border:1px solid #336699;'>
	<tr>
	<td width="14%" class="td_left"><%=s_lang_0128 %><%=s_lang_mh %></td>
	<td class="td_right">
	<table width="100%" border="0" cellspacing="0" align="center"><tr><td><%=get_date_showstr(ei.Date) %></td><td align="right"><img src="images/<%
if ei.IsEnd = true then
	Response.Write "rc_end.gif"" title=""" & s_lang_0132
else
	Response.Write "rc_noend.gif"" title=""" & s_lang_0133
end if
%>" align='absmiddle' border='0'></td></tr>
	</table>
	</td></tr>

	<tr><td class="td_left"><%=s_lang_0145 %><%=s_lang_mh %></td>
	<td class="td_right"><%
if ei.Priority = 2 then
	Response.Write "<font color='#901111'>" & s_lang_0130 & "</font>"
elseif ei.Priority = 1 then
	Response.Write s_lang_0131
else
	Response.Write s_lang_0146
end if
%>&nbsp;
	</td></tr>

	<tr><td class="td_left"><%=s_lang_0147 %><%=s_lang_mh %></td>
	<td class="td_right"><%=server.htmlencode(ei.FromName) %>&nbsp;
	</td></tr>

	<tr><td class="td_left"><%=s_lang_0148 %><%=s_lang_mh %></td>
	<td class="td_right"><%=server.htmlencode(ei.FromEmail) %>&nbsp;
	</td></tr>

	<tr><td class="td_left" style='border-bottom:1px solid #8CA5B5;'><%=s_lang_0127 %><%=s_lang_mh %></td>
	<td class="td_right" style='border-bottom:1px solid #8CA5B5; word-break:break-all; word-wrap:break-word;'><%=server.htmlencode(ei.subject) %>&nbsp;</td>
	</tr>

	<tr><td colspan="2" style="padding:0px;">
<table width="100%" border="0" cellspacing="0" align="center">
<tr style='background-color:#EFF7FF; color:#444;'>
<td width="5%" class="in_td"><input type="checkbox" onclick="checkall(this)"></td>
<td width='68%' class="in_td"><%=s_lang_0149 %></td>
<td width='20%' class="in_td"><%=s_lang_0126 %></td>
<td align="center" width='7%' bgcolor="#EFF7FF" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><%=s_lang_0150 %></td>
</tr>
<%
i = 0
allnum = ei.count

do while i < allnum
	ei.Get i, rc_mode, rc_to, rc_state, rc_msg

	if i + 1 < allnum then
		bottom_line = " style='border-bottom:1px " & MY_COLOR_1 & " solid;'"
	else
		bottom_line = ""
	end if

	Response.Write "<tr id='tr_" & i & "' onmouseover='m_over(this);' onmouseout='m_out(this);'>"

	if rc_state = 2 or rc_state = 3 or rc_state = 5 then
		Response.Write "<td align='center' nowrap" & bottom_line & "><input type='checkbox' disabled id='check" & i & "'></td>"
	else
		Response.Write "<td align='center' nowrap" & bottom_line & "><input type='checkbox' id='check" & i & "' name='check" & i & "' value='" & i & "' onclick='ck_select(this);'></td>"
	end if

	Response.Write "<td nowrap height='24'" & bottom_line & ">" & server.htmlencode(rc_to)

	if rc_mode = 1 then
		Response.Write "&nbsp;[<font color='#901111'>" & s_lang_0151 & "</font>]</td>"
	elseif rc_mode = 2 then
		Response.Write "&nbsp;[<font color='#901111'>" & s_lang_0087 & "</font>]</td>"
	else
		Response.Write "</td>"
	end if

	Response.Write "<td align='center' nowrap" & bottom_line & ">" & get_state_showstr(rc_state) & "</td>"

	if rc_state = 2 or rc_state = 3 or rc_state = 5 then
		Response.Write "<td align='center' nowrap" & bottom_line & ">&nbsp;</td>"
	else
		Response.Write "<td align='center' nowrap" & bottom_line & "><a href=""javascript:recall_one(" & i & ")""><img src='images/recallone.gif' border='0' title='" & s_lang_0150 & "'></a></td>"
	end if

	Response.Write "</tr>" & Chr(13)

	rc_mode = NULL
	rc_to = NULL
	rc_state = NULL
	rc_msg = NULL

	i = i + 1
loop
%>
</table>
</td></tr>
</table>
</FORM>

<table width="92%" border="0" style="margin-top:16px;"><tr><td align="right">
<a href="#"><img src='images/gotop.gif' border='0' title="<%=s_lang_0152 %>"></a>
</td></tr>
</table>

<table width="90%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px solid #8CA5B5; margin-top:50px; padding-bottom:10px;'>
	<tr>
	<td valign="top" style="padding:4px; width:30px;"><font color="#901111">&nbsp;<img src='images/remind.gif' border='0' align='absmiddle'>&nbsp;</font></td>
	<td style="padding:4px; color:#444;"><%=s_lang_0161 %><br>
		<%=s_lang_0162 %><br>
		<%=s_lang_0163 %>
	</td></tr>
</table>

<script type="text/javascript">
<!-- 
function back() {
<% if gourl = "" then %>
	history.back();
<% else %>
	location.href = "<%=gourl %>&<%=getGRSN() %>";
<% end if %>
}

function checkall(tgobj) {
	var theObj;
	for(var i = 0; i < <%=i %>; i++)
	{
		theObj = document.getElementById("check" + i);
		if (theObj != null && theObj.disabled == false)
		{
			theObj.checked = tgobj.checked;
			ck_select(theObj);
		}
	}
}

function ischeck() {
	var theObj;

	for(var i = 0; i < <%=i %>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null && theObj.disabled == false)
		{
			if (theObj.checked == true)
				return true;
		}
	}

	return false;
}

function recall_one(id) {
	if (confirm("<%=s_lang_0142 %>") == false)
		return ;

	document.f1.mode.value = id;
	document.f1.submit();
}

function recall_check() {
	if (ischeck() == true)
	{
		if (confirm("<%=s_lang_0142 %>") == false)
			return ;

		document.f1.mode.value = "rccheck";
		document.f1.submit();
	}
}

function recall_all() {
	var i = 0;
	var theObj;
	var allstr = "";

	for(; i < <%=allnum %>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null && theObj.disabled == false)
		{
			if (allstr.length == 0)
				allstr = theObj.value;
			else
				allstr = allstr + "," + theObj.value;
		}
	}

	if (allstr.length > 0)
	{
		if (confirm("<%=s_lang_0142 %>") == false)
			return ;

		document.f1.mode.value = allstr;
		document.f1.submit();
	}
}

function del() {
	if (confirm("<%=s_lang_0115 %>") == false)
		return ;

	document.f1.mode.value = "del";
	document.f1.submit();
}

function m_over(tag_obj)
{
	var theObj = eval("document.f1.check" + tag_obj.id.substr(3));

	if (theObj != null)
	{
		if (theObj.checked == false)
			tag_obj.style.backgroundColor = "#ecf9ff";
	}
}

function m_out(tag_obj)
{
	var theObj = eval("document.f1.check" + tag_obj.id.substr(3));

	if (theObj != null)
	{
		if (theObj.checked == false)
			tag_obj.style.backgroundColor = "white";
	}
}

function ck_select(tag_obj) {
	if (tag_obj.checked == true)
		document.getElementById("tr_" + tag_obj.id.substr(5)).style.background = "#93BEE2";
	else
		document.getElementById("tr_" + tag_obj.id.substr(5)).style.background = "white";
}
// -->
</script>

</BODY>
</HTML>

<%
set ei = nothing

function get_date_showstr(show_date_str)
	if Len(show_date_str) = 14 then
		tmp_month = Mid(show_date_str, 5, 2)
		if Mid(tmp_month, 1, 1) = "0" then
			tmp_month = Mid(tmp_month, 2, 1)
		end if

		tmp_day = Mid(show_date_str, 7, 2)
		if Mid(tmp_day, 1, 1) = "0" then
			tmp_day = Mid(tmp_day, 2, 1)
		end if

		get_date_showstr = Mid(show_date_str, 1, 4) & s_lang_0139 & tmp_month & s_lang_0140 & tmp_day & s_lang_0141 & " " & Mid(show_date_str, 9, 2) & ":" & Mid(show_date_str, 11, 2) & ":" & Mid(show_date_str, 13, 2)
	else
		get_date_showstr = ""
	end if
end function

function get_state_showstr(in_state)
	if in_state = 1 then
		get_state_showstr = s_lang_0153
	elseif in_state = 2 then
		get_state_showstr = s_lang_0154
	elseif in_state = 3 then
		get_state_showstr = s_lang_0155
	elseif in_state = 4 then
		get_state_showstr = s_lang_0156
	elseif in_state = 5 then
		get_state_showstr = s_lang_0157
	else
		get_state_showstr = s_lang_0158
	end if
end function
%>
