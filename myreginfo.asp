<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
RegIsEmpty = true
errline = -1

dim ei
set ei = server.createobject("easymail.MoreRegInfo")
ei.LoadRegInfo Session("wem")

allnum = ei.Count_RegInfo

if allnum > 0 then
	ei.LoadSetting
	allnum_setting = ei.Count_Setting

	if allnum_setting = allnum then
		RegIsEmpty = false
	end if

	ei.LoadRegInfo Session("wem")
end if


if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ei.RemoveAll_RegInfo

	if RegIsEmpty = true then
		ei.LoadSetting
		allnum_setting = ei.Count_Setting
	else
		allnum_setting = allnum
	end if

	i = 0

	do while i < allnum_setting
		rgline = trim(request("rgline" & i))

		if errline = -1 and Len(rgline) < 3 then
			errline = i
			Exit Do
		end if

		if errline = -1 then
			if ei.AddLine_RegInfo(rgline) = false then
				errline = i
				Exit Do
			end if
		end if 

	    i = i + 1
	loop

	if errline = -1 then
		ei.SaveRegInfo
		set ei = nothing
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=myreginfo.asp"
	else
		set ei = nothing
		Response.Redirect "err.asp?errstr=" & s_lang_0447 & errline + 1 & s_lang_0448 & "&" & getGRSN() & "&gourl=myreginfo.asp"
	end if
end if

if allnum < 1 or RegIsEmpty = true then
	ei.LoadSetting
	allnum = ei.Count_Setting
end if

Set em = Application("em")
em.GetUserByName3 Session("wem"), gu_name, gu_domain, gu_comment, gu_forbid, gu_lasttime, gu_amode, gu_limitout, gu_expiresday, gu_monitor

gu_comment = NULL
gu_forbid = NULL
gu_lasttime = NULL
gu_domain = NULL
gu_name = NULL
gu_amode = NULL
gu_limitout = NULL
gu_monitor = NULL

set em = nothing
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
html {overflow: scroll; overflow-x: hidden; overflow-y: auto !important;}
.cont_td {height:24px; text-align:left; border-bottom:1px solid #A5B6C8; padding-left:12px; padding-right:5px; word-break:break-all; word-wrap:break-word;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function GetStringRealLength(tstr) {
	var reallen = 0;

	for (var i = 0; i < tstr.length; i++)
	{
		if (escape(tstr.charAt(i)).length < 4)
			reallen++;
		else
			reallen = reallen + 2;
	}

	return reallen;
}

function checksub() {
<%
i = 0

do while i < allnum
	if RegIsEmpty = false then
		ei.Get_RegInfo i, s_name, s_sel, s_len, s_msg
	else
		ei.Get_Setting i, s_name, s_sel, s_len
	end if

	Response.Write "	if (checkone(" & i & ", """ & s_name & """, " & s_sel & ", " & s_len & ") == false) return ;" & Chr(13)

	s_name = NULL
	s_sel = NULL
	s_len = NULL

	if RegIsEmpty = false then
		s_msg = NULL
	end if

	i = i + 1
loop
%>
	document.f1.submit();
}

function checkone(cnum, cname, csel, clen) {
	var isok = true;
	inputObj = eval("document.all(\"input\" + cnum)");

	if (csel == 0 && GetStringRealLength(inputObj.value) >= clen)
		isok = false;
	else if (csel == 1 && GetStringRealLength(inputObj.value) != clen)
		isok = false;
	else if (csel == 2 && GetStringRealLength(inputObj.value) <= clen)
		isok = false;

	var tempstr = "";
	if (isok == false)
	{
		alert("<%=s_lang_inputerr %>");
		inputObj.focus();
	}
	else
	{
		tempstr = cname + '\t' + csel + '\t' + clen + '\t' + inputObj.value;

		inputObj = eval("document.all(\"rgline\" + cnum)");
		inputObj.value = tempstr;
	}

	return isok;
}

function exall()
{
	var show_span = document.getElementById("showall_span")

	if (show_span.style.display == "none")
	{
		show_span.style.display = "inline";
		document.getElementById("btex").innerHTML = "<%=s_lang_0057 %>";
	}
	else
	{
		show_span.style.display = "none";
		document.getElementById("btex").innerHTML = "<%=s_lang_0058 %>";
	}
}
//-->
</SCRIPT>

<body>
<form action="myreginfo.asp?<%=getGRSN() %>" name="f1" METHOD="POST">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=s_lang_0449 %>
</td></tr>
<tr><td class="block_top_td" style="height:6px; _height:8px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
<%
if gu_expiresday <> "" then
%>
	<tr><td align='left' class='cont_td'>
	<%=s_lang_0450 %><%=s_lang_mh %><font color='#901111'><%=getYear(gu_expiresday) & "-" & getMonth(gu_expiresday) & "-" & getDay(gu_expiresday) %></font></td>
	</td></tr>
<%
end if

dim ul
set ul = server.createobject("easymail.UserLog")
ul.Load Session("wem")
i = ul.Count - 1
write_span = false

do while i >= 0
	ul.Get i, datestr, mode, exstr
%>
	<tr><td class='cont_td'><%
if mode = 1 then
	Response.Write server.htmlencode(s_lang_0281 & exstr) & "&nbsp;&nbsp;&nbsp;&nbsp;[" & get_date_showstr(datestr) & " " & get_time_showstr(datestr) & "]" & Chr(13)
else
	Response.Write server.htmlencode(s_lang_0282 & exstr) & "&nbsp;&nbsp;&nbsp;&nbsp;[" & get_date_showstr(datestr) & " " & get_time_showstr(datestr) & "]" & Chr(13)
end if
%>
	</td></tr>
<%
	if i <= (ul.Count - 10) and write_span = false then
		write_span = true
		Response.Write "<tr><td><span id='showall_span' style='display:none'><table width='100%' border='0' cellspacing='0' bgcolor='white'>" & Chr(13)
	end if

	datestr = NULL
	mode = NULL
	exstr = NULL

	i = i - 1
loop

if write_span = true then
	Response.Write "</table></span>"
end if

set ul = nothing

allnum = ei.Count_RegInfo
i = 0

if allnum > 0 then
%>
	</td></tr></table>
	</td></tr></table>
<br><br>
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=s_lang_0451 %>
</td></tr>
<tr><td class="block_top_td" style="height:6px; _height:8px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
<%
end if

do while i < allnum
	if RegIsEmpty = false then
		ei.Get_RegInfo i, s_name, s_sel, s_len, s_msg
	else
		ei.Get_Setting i, s_name, s_sel, s_len
	end if

	tinput = trim(request("input" & i))

	if Request.ServerVariables("REQUEST_METHOD") = "GET" then
		if tinput = "" then
			tinput = s_msg
		end if
	end if

	if errline <> i then
		if s_sel = 0 then
			Response.Write "<tr><td align='left' class='cont_td'>" & server.htmlencode(s_name) & s_lang_mh
			Response.Write "<input name='input" & i & "' type='text' class='n_textbox' size='30' maxlength='" & s_len - 1 & "' value='" & tinput & "'><input name='rgline" & i & "' type='hidden'></td></tr>" & Chr(13)
		elseif s_sel = 1 then
			Response.Write "<tr><td align='left' class='cont_td'>" & server.htmlencode(s_name) & s_lang_mh
			Response.Write "<input name='input" & i & "' type='text' class='n_textbox' size='30' maxlength='" & s_len & "' value='" & tinput & "'>&nbsp;*<input name='rgline" & i & "' type='hidden'></td></tr>" & Chr(13)
		elseif s_sel = 2 then
			Response.Write "<tr><td align='left' class='cont_td'>" & server.htmlencode(s_name) & s_lang_mh
			Response.Write "<input name='input" & i & "' type='text' class='n_textbox' size='30' value='" & tinput & "' maxlength='128'>&nbsp;*<input name='rgline" & i & "' type='hidden'></td></tr>" & Chr(13)
		end if
	else
		if s_sel = 0 then
			Response.Write "<tr><td align='left' class='cont_td'><font color='#FF3333'>" & server.htmlencode(s_name) & s_lang_mh
			Response.Write "<input name='input" & i & "' type='text' class='n_textbox' size='30' maxlength='" & s_len - 1 & "' value='" & tinput & "'><input name='rgline" & i & "' type='hidden'></td></tr>" & Chr(13)
		elseif s_sel = 1 then
			Response.Write "<tr><td align='left' class='cont_td'><font color='#FF3333'>" & server.htmlencode(s_name) & s_lang_mh
			Response.Write "<input name='input" & i & "' type='text' class='n_textbox' size='30' maxlength='" & s_len & "' value='" & tinput & "'>&nbsp;*<input name='rgline" & i & "' type='hidden'></td></tr>" & Chr(13)
		elseif s_sel = 2 then
			Response.Write "<tr><td align='left' class='cont_td'><font color='#FF3333'>" & server.htmlencode(s_name) & s_lang_mh
			Response.Write "<input name='input" & i & "' type='text' class='n_textbox' size='30' value='" & tinput & "' maxlength='128'>&nbsp;*<input name='rgline" & i & "' type='hidden'></td></tr>" & Chr(13)
		end if
	end if

	s_name = NULL
	s_sel = NULL
	s_len = NULL

	if RegIsEmpty = false then
		s_msg = NULL
	end if

	i = i + 1
loop
%></td></tr>
</table>

</td></tr>
<tr><td height="40" bgcolor="white" align="left"><br>
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<%
if allnum > 0 then
%>
<a class='wwm_btnDownload btn_blue' href="javascript:checksub();"><%=s_lang_save %></a>
<%
end if

if write_span = true then
	Response.Write "<a id='btex' class='wwm_btnDownload btn_blue' href='javascript:exall();'>" & s_lang_0058 & "</a>"
end if
%>
</td></tr>
</table>
</form>

<% if IsEnterpriseVersion = true then %>
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style="border:1px #A5B6C8 solid; margin-top:20px;">
	<tr><td align="center" height="32">
<a class='wwm_btnDownload btn_gray' href="ldap.asp?<%=getGRSN() %>"><%=s_lang_0032 %></a>
&nbsp;&nbsp;&nbsp;
<a class='wwm_btnDownload btn_gray' href="ldappw.asp?<%=getGRSN() %>"><%=s_lang_0033 %></a>
	</td></tr>
</table>
<% end if %>
</body>
</html>


<%
gu_expiresday = NULL

set ei = nothing


function getYear(exday)
	getYear = Mid(Cstr(exday), 1, 4)
end function

function getMonth(exday)
	getMonth = Mid(Cstr(exday), 5, 2)
end function

function getDay(exday)
	getDay = Mid(Cstr(exday), 7, 2)
end function

function get_date_showstr(show_date_str)
	if Len(show_date_str) = 14 or Len(show_date_str) = 12 then
		tmp_month = Mid(show_date_str, 5, 2)
		if Mid(tmp_month, 1, 1) = "0" then
			tmp_month = Mid(tmp_month, 2, 1)
		end if

		tmp_day = Mid(show_date_str, 7, 2)
		if Mid(tmp_day, 1, 1) = "0" then
			tmp_day = Mid(tmp_day, 2, 1)
		end if

		get_date_showstr = Mid(show_date_str, 1, 4) & s_lang_0139 & tmp_month & s_lang_0140 & tmp_day & s_lang_0141
	else
		get_date_showstr = ""
	end if
end function

function get_time_showstr(show_date_str)
	if Len(show_date_str) = 14 or Len(show_date_str) = 12 then
		t_time_hour = CInt(Mid(show_date_str, 9, 2))

		dim t_hour_name
		if t_time_hour >= 0 and t_time_hour < 6 then
			t_hour_name = s_lang_0268
		elseif t_time_hour >= 6 and t_time_hour < 12 then
			t_hour_name = s_lang_0269
		elseif t_time_hour >= 12 and t_time_hour < 14 then
			t_hour_name = s_lang_0283
			if t_time_hour > 12 then
				t_time_hour = t_time_hour - 12
			end if
		elseif t_time_hour >= 14 and t_time_hour < 18 then
			t_hour_name = s_lang_0270
			t_time_hour = t_time_hour - 12
		elseif t_time_hour >= 18 then
			t_hour_name = s_lang_0271
			t_time_hour = t_time_hour - 12
		end if

		ts_time_hour = t_time_hour

		get_time_showstr = t_hour_name & ts_time_hour& ":" & Mid(show_date_str, 11, 2)
	else
		get_time_showstr = ""
	end if
end function
%>
