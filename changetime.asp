<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" -->

<%
filename = trim(request("filename"))
timestr = trim(request("timestr"))

if Request.ServerVariables("REQUEST_METHOD") = "GET" then
	sgourl = "showmail.asp?filename=" & filename & "&" & getGRSN()
	pgourl = Server.URLEncode(trim(request("gourl")))
end if

if Request.ServerVariables("REQUEST_METHOD") = "POST" and timestr <> "" then
	sgourl = Server.URLEncode(trim(request("sgourl")))
	pgourl = trim(request("pgourl"))

	dim ei
	set ei = server.createobject("easymail.emmail")
	ei.LoadText Session("wem"), filename

	dim isok
	isok = ei.SetNewTimeForTimerMail(timestr)
	set ei = nothing

	if isok = true then
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & sgourl & "&pgourl=" & pgourl
	else
		Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & sgourl & "&pgourl=" & pgourl
	end if
end if
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
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function cutz(inval)
{
	var rval = "";
	for (var i = 0; i < inval.length; i++)
	{
		if (inval.charAt(i) != '0')
			break;
	}
	rval = inval.substring(i);
	return rval;
}

function timerSending() {
	var err = "<%=b_lang_345 %>";
	var nowdate = new Date(<%=Year(now()) & "," & Month(now()) - 1 & "," & Day(now()) & "," & Hour(now()) & "," & Minute(now()) %>);
	var mydate = new Date(document.f1.t_year.value, document.f1.t_month.value - 1, document.f1.t_day.value, document.f1.t_hour.value, 1);

	var nmonth = document.f1.t_month.value;
	var nday = document.f1.t_day.value;
	var nhour = document.f1.t_hour.value;

	if (document.f1.t_year.value == "" || document.f1.t_year.value > 9999 || document.f1.t_year.value < <%=Year(now()) %>)
	{
		alert(err);
		document.f1.t_year.focus();
		return ;
	}

	if (nmonth == "" || nmonth > 12 || nmonth < 1)
	{
		alert(err);
		document.f1.t_month.focus();
		return ;
	}

	if (nday == "" || nday > 31 || nday < 1)
	{
		alert(err);
		document.f1.t_day.focus();
		return ;
	}

	if (nhour == "" || nhour > 23 || nhour < 0)
	{
		alert(err);
		document.f1.t_hour.focus();
		return ;
	}

	if (mydate > nowdate)
	{
		if (document.f1.t_month.value < 10)
			nmonth = "0" + cutz(document.f1.t_month.value);

		if (document.f1.t_day.value < 10)
			nday = "0" + cutz(document.f1.t_day.value);

		if (document.f1.t_hour.value < 10)
			nhour = "0" + cutz(document.f1.t_hour.value);

		if (nhour == "0")
			nhour = "00"

		document.f1.timestr.value = document.f1.t_year.value + nmonth + nday + nhour;
	}
	else
	{
		alert("<%=b_lang_346 %>");
		return ;
	}

	document.f1.submit();
}
//-->
</script>

<body>
<form name="f1" method="post" action="changetime.asp">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_347 %>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr>
	<td width="18%" height="35" align="right">
	<%=b_lang_348 & s_lang_mh %>
	</td>
	<td align="left" style="padding-left:4px;">
<select name="t_year" class="drpdwn">
<%
	now_temp = Year(Now())

	i = now_temp
	do while i < now_temp + 10
		Response.Write "<option value='" & i & "'>" & i & b_lang_028 & "</option>"
		i = i + 1
	loop
%>
</select>&nbsp;
<select name="t_month" class="drpdwn">
<%
	now_temp = Month(Now())
	i = 1
	do while i < 13
		if i <> now_temp then
			Response.Write "<option value='" & i & "'>" & i & b_lang_029 & "</option>"
		else
			Response.Write "<option value='" & i & "' selected>" & i & b_lang_029 & "</option>"
		end if
		i = i + 1
	loop
%>
</select>&nbsp;
<select name="t_day" class="drpdwn">
<%
	now_temp = Day(Now())
	i = 1
	do while i < 32
		if i <> now_temp then
			Response.Write "<option value='" & i & "'>" & i & b_lang_030 & "</option>"
		else
			Response.Write "<option value='" & i & "' selected>" & i & b_lang_030 & "</option>"
		end if
		i = i + 1
	loop
%>
</select>&nbsp;
<select name="t_hour" class="drpdwn">
<%
	i = 0
	do while i < 24
		Response.Write "<option value='" & i & "'>" & i & b_lang_031 & "</option>"
		i = i + 1
	loop
%>
</select>
	</td>
	</tr>
</table>
	</td>
	</tr>
<tr><td class="block_top_td" style="height:10px;"></td></tr>

<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="javascript:history.back();"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:timerSending();"><%=s_lang_save %></a>
</td></tr>
</table>

<input type="hidden" name="timestr">
<input type="hidden" name="filename" value="<%=filename %>">
<input type="hidden" name="sgourl" value="<%=sgourl %>">
<input type="hidden" name="pgourl" value="<%=pgourl %>">
</form>
</BODY>
</HTML>
