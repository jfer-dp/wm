<!--#include file="passinc.asp" --> 

<%
if isadmin() = false then
	dim mam
	set mam = server.createobject("easymail.AdminManager")
	mam.Load

	if mam.Enable_DomainAdmin_SetAdvertisingMsg = false then
		set mam = nothing
		response.redirect "noadmin.asp"
	end if

	set mam = nothing
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

dim dm
set dm = server.createobject("easymail.Domain")
dm.Load

if dm.GetUserManagerDomainCount(Session("wem")) < 1 then
	if isadmin() = false then
		set dm = nothing
		response.redirect "noadmin.asp"
	end if
end if


'-----------------------------------------
dim ei
set ei = server.createobject("easymail.Domain_Advertising_Msg")
ei.Load
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
<script language="JavaScript" type="text/javascript" src="rte/wrte1.js"></script>
<script language="JavaScript" type="text/javascript" src="rte/wrte2.js"></script>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
<%
if isrtf = true then
%>
initRTE("./rte/images/", "./rte/", "", false);
<%
end if
%>

function domainname_onchange() {
	location.href = "showadv.asp?<%=getGRSN() %>&selectdomain=" + document.f1.domainname.value + "&isrtf=<%=isrtf %>";
}


function changemyselect_onclick() {
<%
if isrtf = false then
%>
	if (document.f1.RichEdit_Text.disabled == true)
		document.f1.RichEdit_Text.disabled = false;
	else
		document.f1.RichEdit_Text.disabled = true;
<%
end if
%>
}


function all2system()
{
	if (confirm("是否清除所有域的广告内容, 而使用系统广告内容? "))
	{
		document.f1.cleanall.value = "yes";
		document.f1.submit();
	}
}
//-->
</SCRIPT>


<BODY LANGUAGE=javascript onload="return window_onload()">
<br><br>
<FORM ACTION="saveadv.asp" METHOD=POST NAME="f1">
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="5%" height="25">&nbsp;</td>
      <td width="40%"><b>选择域名</b>:&nbsp;<select name="domainname" class="drpdwn" LANGUAGE=javascript onchange="return domainname_onchange()">
<%
i = 0

if isadmin() = false then
	allnum = dm.GetUserManagerDomainCount(Session("wem"))

	do while i < allnum
		domain = dm.GetUserManagerDomain(Session("wem"), i)

		if domain <> trim(request("selectdomain")) then
			response.write "<option value='" & server.htmlencode(domain) & "'>" & server.htmlencode(domain) & "</option>" & Chr(13)
		else
			curdomain = domain
			response.write "<option value='" & server.htmlencode(domain) & "' selected>" & server.htmlencode(domain) & "</option>" & Chr(13)
		end if

		domain = NULL

		i = i + 1
	loop
else
	allnum = dm.GetCount()

	do while i < allnum
		domain = dm.GetDomain(i)

		if domain <> trim(request("selectdomain")) then
			response.write "<option value='" & server.htmlencode(domain) & "'>" & server.htmlencode(domain) & "</option>" & Chr(13)
		else
			curdomain = domain
			response.write "<option value='" & server.htmlencode(domain) & "' selected>" & server.htmlencode(domain) & "</option>" & Chr(13)
		end if

		domain = NULL

		i = i + 1
	loop
end if


if curdomain = "" then
	if isadmin() = false then
		curdomain = dm.GetUserManagerDomain(Session("wem"), 0)
	else
		curdomain = dm.GetDomain(0)
	end if
end if

haveitdm = ei.haveit(curdomain)

ei.Get curdomain, text, htmltext
%>
</select>
</td>
      <td width="18%"><a href="javascript:all2system()">全部使用系统广告内容</a></td>
    </tr>
  </table>
<br>
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="30" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b><%=server.htmlencode(curdomain) %> 域广告内容</b></font>
		</div>
      </td>
    </tr>
    <tr><td height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<input type="checkbox" name="changemyselect" LANGUAGE=javascript onclick="return changemyselect_onclick()"<%if haveitdm = false then response.write " checked" end if %>>此域使用系统广告内容&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%
if isrtf = true then
%>
<a href="showadv.asp?<%=getGRSN() %>&selectdomain=<%=Server.URLEncode(curdomain) %>&isrtf=False"><b>文本格式</b></a>
<%
else
%>
<a href="showadv.asp?<%=getGRSN() %>&selectdomain=<%=Server.URLEncode(curdomain) %>&isrtf=True"><b>超文本格式</b></a>
<%
end if
%>
	</td>
	<td align="right" style="border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<input type="button" value=" 保存 " class="Bsbttn" LANGUAGE="javascript" onclick="gosub();">&nbsp;&nbsp;
<%
if isadmin() = false then
%>
	<input type="button" value=" 取消 " onclick="javascript:location.href='domainright.asp?<%=getGRSN() %>';" class="Bsbttn">&nbsp;&nbsp;
<%
else
%>
	<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">&nbsp;&nbsp;
<%
end if
%>
	</td></tr>
	<tr>
	<td height="340" colspan="2">
<%
if isrtf = false then
%>
<textarea cols="<%
if isMSIE = true then
	Response.Write "75"" rows=""22"""
else
	Response.Write "65"" rows=""17"""
end if
%>" wrap="soft" name="RichEdit_Text" class="textarea"<%if haveitdm = false then response.write " disabled" end if %>>
<%=text%></textarea>
<%
else
%>
<script language="JavaScript" type="text/javascript">
<!--
<%
	if htmltext <> "" then
		Response.Write "writeRichText('richedit', RemoveScript('" & RTESafe(htmltext) & "'), 545, 282, true, false);"
	else
		html_text = replace(text, "'", "&#39;")
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
  </table>
<input name="cleanall" type="hidden" value="">
<input name="curdomain" type="hidden" value="<%=curdomain %>">
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
  <div align="center">
    <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr>
		<td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
		<td width="94%">您可以为每个域创建不同的域广告内容, 也可以使用系统广告内容.
		<br><br>广告中的内容将会被追加到每一封本域内用户通过浏览器所发邮件的尾部.
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
<br>
</BODY>

<SCRIPT LANGUAGE=javascript>
<!--
function window_onload() {
}

function gosub() {
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
</HTML>

<%
curdomain = NULL
text = NULL
htmltext = NULL

set ei = nothing
set dm = nothing


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
