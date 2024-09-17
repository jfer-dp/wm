<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
Session("SearchStr") = ""

dim pf
set pf = server.createobject("easymail.PerFolders")
pf.Load Session("wem")
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
.font_top_title {font-size:11pt; font-weight:bold;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.table_min_width {width:660px; font-size:0pt; height:0px; width:0px; border:0px;}
.cont_tr {background:white; height:30px;}
.cont_td_head {border-bottom:1px solid #A5B6C8; padding-right:4px; color:black;}
.cont_td {height:30px; border-bottom:1px solid #A5B6C8; padding-right:4px; color:#444444;}
-->
</STYLE>
</HEAD>

<script type="text/javascript" src="images/sc_left.js"></script>

<SCRIPT LANGUAGE=javascript>
<!--
function msearch() {
	document.f2.mailsearch.value = "\t\t" + document.f2.shead.value + "\t" + document.f2.smode.value + "\t" + document.f2.stext.value + "\t\tRecDate\t" + document.f2.sdatemode.value + "\t" + document.f2.syear.value + document.f2.smonth.value + document.f2.sday.value + "\t\tSize\t" + document.f2.ssizemode.value + "\t" + document.f2.ssize.value + "\t\tRead\t1\t" + document.f2.sread.value + "\t\tFolders\t" + getcheck() + "\t";
	document.f2.submit();
}

function getcheck() {
	var i = 0;
	var theObj;
	var str = "";

	for(; i<<%= CStr(4 + pf.FolderCount)%>; i++)
	{
		theObj = eval("document.f2.check" + i);
		if (theObj.checked == true)
			str = str + theObj.value + "\t";
	}

	return str;
}

function clearall() {
	var i = 0;
	var theObj;

	for(; i<<%= CStr(4 + pf.FolderCount)%>; i++)
	{
		theObj = eval("document.f2.check" + i);
		theObj.checked = "";
	}
}

function checkall() {
	var i = 0;
	var theObj;

	for(; i<<%= CStr(4 + pf.FolderCount)%>; i++)
	{
		theObj = eval("document.f2.check" + i);
		theObj.checked = "check";
	}
}
//-->
</SCRIPT>

<BODY>
<form name="f2" method="post" action="searchlistmail.asp">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_236 %>
</td></tr>
<tr><td class="block_top_td" style="height:6px; _height:8px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr class="cont_tr">
	<td width="18%" align="right" class="cont_td_head">
	<select name="shead" class="drpdwn" size="1">
		<option value="Subject" selected><%=a_lang_237 %></option>
		<option value="FromMail"><%=a_lang_238 %></option>
		<option value="FromName"><%=a_lang_239 %></option>
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
	<td align="right" class="cont_td_head"><%=a_lang_243 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td">
	<select name="syear" class="drpdwn" size="1">
<%
mc = 1990
do while mc <= Year(Now)
	if mc = Year(Now) then
		response.write "<option value='" & mc & "' selected>" & mc & a_lang_244 & "</option>"
	else
		response.write "<option value='" & mc & "'>" & mc & a_lang_244 & "</option>"
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
		response.write "<option value='" & smc & "' selected>" & mc & a_lang_245 & "</option>"
	else
		response.write "<option value='" & smc & "'>" & mc & a_lang_245 & "</option>"
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
		response.write "<option value='" & smc & "' selected>" & mc & a_lang_246 & "</option>"
	else
		response.write "<option value='" & smc & "'>" & mc & a_lang_246 & "</option>"
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
	<td align="right" class="cont_td_head"><%=a_lang_250 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td">
	<select name="ssizemode" class="drpdwn" size="1">
		<option value="1" selected><%=a_lang_251 %></option>
		<option value="2"><%=a_lang_252 %></option>
		<option value="3"><%=a_lang_253 %></option>
	</select>
	<input type="text" name="ssize" class='n_textbox'>&nbsp;(<%=a_lang_254 %>)
	</td></tr>

	<tr class="cont_tr">
	<td align="right" class="cont_td_head"><%=a_lang_255 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td">
	<select name="sread" class="drpdwn" size="1">
		<option value="1" selected><%=a_lang_256 %></option>
		<option value="2"><%=a_lang_257 %></option>
		<option value="3"><%=a_lang_258 %></option>
	</select>
	</td></tr>

	<tr class="cont_tr">
	<td align="right" class="cont_td_head"><%=a_lang_259 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td" style="padding-bottom:4px;">
		<input type='checkbox' name='check0' value="in" checked><%=a_lang_260 %>&nbsp;&nbsp; 
		<input type='checkbox' name='check1' value="out" checked><%=a_lang_261 %>&nbsp;&nbsp; 
		<input type='checkbox' name='check2' value="sed" checked><%=a_lang_262 %>&nbsp;&nbsp; 
		<input type='checkbox' name='check3' value="del" checked><%=a_lang_263 %>&nbsp;&nbsp;
<%
allnum = pf.FolderCount

i = 0
do while i < allnum
	Response.Write "<input type='checkbox' name='check" & i+4 & "' value=""" & pf.GetFolderName(i) & """>" & server.htmlencode(pf.GetFolderName(i)) & "&nbsp;&nbsp;"

	i = i + 1
loop
%> 
		<input type="hidden" name="mailsearch">
		<br><br>
		<a class='wwm_btnDownload btn_gray' href='javascript:checkall()'><%=a_lang_264 %></a>&nbsp;
		<a class='wwm_btnDownload btn_gray' href='javascript:clearall()'><%=a_lang_265 %></a>
	</td></tr>

	<tr class="cont_tr"><td colspan="2" height="50" bgcolor="white" align="left">
	<a class='wwm_btnDownload btn_blue' href="javascript:msearch();"><%=a_lang_266 %></a>
	</td></tr>
	<tr><td class="block_top_td" colspan="2"><div class="table_min_width"></div></td></tr>
</table>
</td></tr>
</table>

<table width="92%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px <%=MY_COLOR_1 %> solid; margin-top:90px; padding-bottom:10px;'>
	<tr>
	<td valign="top" style="padding:4px; width:82px;"><font color="#901111">*<%=a_lang_267 %></font></td>
	<td style="padding:4px; color:#444;">
	<%=a_lang_268 %>
	</td></tr>
</table>
</form>
</BODY>
</HTML>

<%
set pf = nothing
%>
