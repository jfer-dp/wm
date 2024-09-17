<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
fileid = trim(request("fileid"))
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<STYLE type=text/css>
<!--
.title_tr {white-space:nowrap; background:#f2f4f6; height:24px;}
.title_td {text-align:center; white-space:nowrap; border:1px solid #A5B6C8;}
.font_top_title {font-size:11pt; color:#104A7B; font-weight:bold;}

.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.table_min_width {width:660px; font-size:0pt; height:0px; width:0px; border:0px;}
.cont_tr {background:white; height:30px;}
.cont_td {height:30px; border-bottom:1px solid #A5B6C8; padding-right:4px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function msearch() {
	document.f2.mailsearch.value = "\t\t" + document.f2.shead.value + "\t" + document.f2.smode.value + "\t" + document.f2.stext.value + "\t\tDate\t" + document.f2.sdatemode.value + "\t" + document.f2.syear.value + document.f2.smonth.value + document.f2.sday.value + "\t\tSize\t" + document.f2.ssizemode.value + "\t" + document.f2.ssize.value + "\t\tKey\t" + document.f2.skeytext.value + "\t\tHit\t" + document.f2.shit.value + "\t" + document.f2.shittext.value + "\t\t";
	document.f2.submit();
}
//-->
</script>

<BODY>
<form name="f2" method="post" action="searchpf.asp?fileid=<%=fileid %>">
<input type="hidden" name="mailsearch">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_236 %>
</td></tr>
<tr><td class="block_top_td" style="height:6px; _height:8px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr class="cont_tr">
	<td width="18%" align="right" class="cont_td">
	<select name="shead" class="drpdwn" size="1">
		<option value="Subject" selected><%=a_lang_237 %></option>
		<option value="PostName"><%=a_lang_269 %></option>
	</select><%=s_lang_mh %>
	</td>
	<td align="left" class="cont_td">
	<select name="smode" class="drpdwn" size="1">
		<option value="1" selected><%=a_lang_240 %></option>
		<option value="2"><%=a_lang_241 %></option>
		<option value="3"><%=a_lang_242 %></option>
	</select>
	<input type="text" name="stext" class='n_textbox' size="30">
	</td></tr>

	<tr class="cont_tr">
	<td align="right" class="cont_td"><%=a_lang_270 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td">
	<select name="skeymode" class="drpdwn" size="1">
		<option value="1" selected><%=a_lang_240 %></option>
	</select>
	<input type="text" name="skeytext" class='n_textbox' size="30">
	</td></tr>

	<tr class="cont_tr">
	<td align="right" class="cont_td"><%=a_lang_243 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td">
	<select name="syear" class="drpdwn" size="1">
<%
mc = 1990

do while mc <= Year(Now)

	if mc = Year(Now) then
		Response.Write "<option value='" & mc & "' selected>" & mc & a_lang_244 & "</option>"
	else
		Response.Write "<option value='" & mc & "'>" & mc & a_lang_244 & "</option>"
	end if

	mc = mc + 1
loop
%> 
	</select>
	<select name="smonth" class="drpdwn" size="1">
<%
mc = 1

do while mc <= 12

	if mc < 10 then
		smc = "0" & CStr(mc)
	else
		smc = CStr(mc)
	end if

	if mc = Month(Now) then
		Response.Write "<option value='" & smc & "' selected>" & mc & a_lang_245 & "</option>"
	else
		Response.Write "<option value='" & smc & "'>" & mc & a_lang_245 & "</option>"
	end if

	mc = mc + 1
loop
%> 
	</select>
	<select name="sday" class="drpdwn" size="1">
<%
mc = 1

do while mc <= 31

	if mc < 10 then
		smc = "0" & CStr(mc)
	else
		smc = CStr(mc)
	end if

	if mc = Day(Now) then
		Response.Write "<option value='" & smc & "' selected>" & mc & a_lang_246 & "</option>"
	else
		Response.Write "<option value='" & smc & "'>" & mc & a_lang_246 & "</option>"
	end if

	mc = mc + 1
loop
%>
	</select>
	<select name="sdatemode" class="drpdwn" size="1">
		<option value="1" selected><%=a_lang_247 %></option>
		<option value="2"><%=a_lang_248 %></option>
		<option value="3"><%=a_lang_249 %></option>
	</select>
	</td></tr>

	<tr class="cont_tr">
	<td align="right" class="cont_td"><%=a_lang_250 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td" style="color:#444;">
	<select name="ssizemode" class="drpdwn" size="1">
		<option value="1" selected><%=a_lang_251 %></option>
		<option value="2"><%=a_lang_252 %></option>
		<option value="3"><%=a_lang_253 %></option>
	</select>
	<input type="text" name="ssize" class='n_textbox'>&nbsp;(<%=a_lang_254 %>)
	</td></tr>

	<tr class="cont_tr">
	<td align="right" class="cont_td"><%=a_lang_271 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td">
	<select name="shit" class="drpdwn" size="1">
		<option value="1" selected><%=a_lang_251 %></option>
		<option value="2"><%=a_lang_252 %></option>
		<option value="3"><%=a_lang_253 %></option>
	</select>
	<input type="text" name="shittext" maxlength="4" class='n_textbox'>
	</td></tr>

	<tr class="cont_tr"><td colspan="2" height="50" bgcolor="white" align="left">
<a class='wwm_btnDownload btn_blue' href="javascript:history.back();"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:msearch();"><%=a_lang_266 %></a>
	</td></tr>
</table>

</td></tr>
</table>
</form>

<table width="92%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px <%=MY_COLOR_1 %> solid; margin-top:90px; padding-bottom:10px;'>
	<tr>
	<td valign="top" style="padding:4px; width:82px;"><font color="#901111">*<%=a_lang_267 %></font></td>
	<td style="padding:4px; color:#444;">
	<%=a_lang_268 %>
	</td></tr>
</table>
</BODY>
</HTML>

<%
set pf = nothing
%>
