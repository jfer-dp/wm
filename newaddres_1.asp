<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
dim pf
set pf = server.createobject("easymail.PerFolders")
pf.Load Session("wem")

if Len(trim(request("errstr"))) > 0 then
	errstr = a_lang_312
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
.cont_td {height:24px; border-bottom:1px solid #A5B6C8; padding-left:4px; padding-right:4px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function msearch() {
	document.f2.folders.value = "\t\tSize\t1\t0\t\tFolders\t" + getcheck() + "\t";
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
<form name="f2" method="post" action="newaddres_2.asp">
<input type="hidden" name="folders">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_313 %>
</td></tr>
<tr><td class="block_top_td" style="height:6px; _height:8px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr>
	<td width="12%" nowrap align="right" class="cont_td">
	<%=a_lang_314 %><%=s_lang_mh %>
	</td>
	<td align="left" class="cont_td" style="padding-bottom:6px;">
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
<a class='wwm_btnDownload btn_gray' href="javascript:checkall();"><%=a_lang_264 %></a>
<a class='wwm_btnDownload btn_gray' href="javascript:clearall();"><%=a_lang_265 %></a>
	</td></tr>

	<tr>
	<td align="right" nowrap class="cont_td" style="height:50px;">
	<%=a_lang_315 %><%=s_lang_mh %>
	</td>
	<td align="left" class="cont_td">
<%=a_lang_316 %><%=s_lang_mh %>
<select name="s_shownum" class="drpdwn" size="1">
<option value="1" selected><%=a_lang_317 %></option>
<option value="2"><%=a_lang_318 %></option>
<option value="3"><%=a_lang_319 %></option>
<option value="4"><%=a_lang_320 %></option>
<option value="5"><%=a_lang_321 %></option>
</select><br>
<%=a_lang_322 %><%=s_lang_mh %>
<select name="s_showdays" class="drpdwn" size="1">
<option value="7" selected><%=a_lang_323 %></option>
<option value="30"><%=a_lang_324 %></option>
<option value="90"><%=a_lang_325 %></option>
<option value="365"><%=a_lang_326 %></option>
</select>
	</td></tr>
<%
if Len(errstr) > 0 then
%>
	<tr><td colspan="2" align="left" class="cont_td" style="border:1px #ff0000 solid; background-color:#FFF8D3;">
	<%=errstr %>fsafafa
	</td></tr>
<%
end if
%>
</table>
	</td></tr>

<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px;">
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:msearch();"><%=a_lang_327 %></a>
</td></tr>
</table>

<table width="92%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px #A5B6C8 solid; margin-top:80px; padding-bottom:10px;'>
	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:4px; color:#444444;">
	<%=a_lang_328 %><br>
	</td>
	</tr>
</table>
</form>
</BODY>
</HTML>

<%
set pf = nothing
%>
