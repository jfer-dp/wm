<!--#include file="passinc.asp" --> 

<%
if isadmin() = false then
	Response.Redirect "noadmin.asp"
end if

isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

if trim(request("isrtf")) = "" or trim(request("isrtf")) = "True" then
	isrtf = true
else
	isrtf = false
end if


dim eam
set eam = server.createobject("easymail.AdminManager")
eam.Load

issave = trim(Request("issave"))

if issave = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	eam.FlySheet = trim(Request("RichEdit_Text"))
	eam.FlySheetHtml = trim(Request("RichEdit_Html"))
	eam.Save

	set eam = nothing

	response.redirect "ok.asp?" & getGRSN() & "&gourl=flysheet.asp?isrtf=" & isrtf
end if
%>

<html>
<head>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
<script language="JavaScript" type="text/javascript" src="rte/wrte1.js"></script>
<script language="JavaScript" type="text/javascript" src="rte/wrte2.js"></script>
</head>

<SCRIPT LANGUAGE=javascript>
<!--
<%
if isrtf = true then
%>
initRTE("./rte/images/", "./rte/", "", false);
<%
end if
%>

function window_onload() {
}

function save_onclick() {
<%
if isrtf = true then
%>
	updateRTE('richedit');

	document.f1.RichEdit_Text.value = getText(document.f1.richedit.value);
	document.f1.RichEdit_Html.value = document.f1.richedit.value;
<%
end if
%>
	document.f1.submit();
}
//-->
</SCRIPT>


<BODY LANGUAGE=javascript onload="return window_onload()">
<br>
<br>
<form method="post" action="flysheet.asp" name="f1">
<input type="hidden" name="issave" value="1">
<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>">
<td width="43%" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid;"><p align="center">&nbsp;
</td>
<td width="32%" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>广&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;告</b></font>
</td>
<td width="25%" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
<%
if isrtf = true then
%>
<a href="flysheet.asp?<%=getGRSN() %>&isrtf=False"><b>文本格式</b></a>
<%
else
%>
<a href="flysheet.asp?<%=getGRSN() %>&isrtf=True"><b>超文本格式</b></a>
<%
end if
%></td>
	</tr>
	<tr>
	<td height="340" colspan="3">
<%
if isrtf = false then
%>
<textarea cols="<%
if isMSIE = true then
	Response.Write "75"" rows=""22"""
else
	Response.Write "65"" rows=""17"""
end if
%>" wrap="soft" name="RichEdit_Text" class="textarea">
<%
t = eam.FlySheet
response.write t
%></textarea>
<%
else
%>
<script language="JavaScript" type="text/javascript">
<!--
<%
	if eam.FlySheetHtml <> "" then
		Response.Write "writeRichText('richedit', RemoveScript('" & RTESafe(eam.FlySheetHtml) & "'), 545, 282, true, false);"
	else
		html_text = replace(eam.FlySheet, "'", "&#39;")
		html_text = replace(html_text, "<", "&lt;")
		html_text = replace(html_text, ">", "&gt;")
		html_text = replace(html_text, Chr(13) & Chr(10), "<br>")
		html_text = replace(html_text, Chr(10) & Chr(13), "<br>")
		html_text = replace(html_text, Chr(13), "<br>")
		html_text = replace(html_text, Chr(10), "<br>")
		html_text = replace(html_text, "\", "\\")
		Response.Write "writeRichText('richedit', '" & html_text & "', 545, 282, true, false);"
	end if
%>
//-->
</script>
<%
end if
%>
	</td>
  </tr>
	<tr>
	<td height="5" colspan="3">&nbsp;
	</td></tr>
	<tr>
	<td height="15" align="right" colspan="3">
<INPUT type="button" value=" 保存 " LANGUAGE=javascript onclick="save_onclick()" class="Bsbttn">
&nbsp;&nbsp;
<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">&nbsp;&nbsp;
	</td></tr>
</table>
<br><br><br>
  <div align="center">
    <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr> 
        <td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
        <td width="94%">广告中的内容将会被追加到每一封通过浏览器所发邮件的尾部. <br>
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
<input name="RichEdit_Html" type="hidden">
<input name="isrtf" type="hidden" value="<%=isrtf %>">
<%
if isrtf = true then
%>
<div style="position:absolute; top:10; left:10; z-index:15; visibility:hidden">
<textarea name="RichEdit_Text" cols="0" rows="0"></textarea>
</div>
<%
end if
%>
</FORM>
<br>
</body>
</html>

<%
set eam = nothing


function RTESafe(strText)
	dim tmpString
	tmpString = replace(strText, "'", "&#39;")
	tmpString = replace(tmpString, Chr(10), "")
	tmpString = replace(tmpString, Chr(13), " ")
	tmpString = replace(tmpString, "&lt;", "&#11;")
	tmpString = replace(tmpString, "<", "&lt;")
	RTESafe = replace(tmpString, "\", "\\")
end function
%>
