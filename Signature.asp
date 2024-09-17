<!--#include file="passinc.asp" --> 
<!--#include file="language-2.asp" -->

<%
if trim(request("isrtf")) = "" or trim(request("isrtf")) = "True" then
	isrtf = true
else
	isrtf = false
end if

if IsNumeric(trim(request("sindex"))) = true then
	sindex = CInt(trim(request("sindex")))
else
	sindex = -1
end if

mode = trim(request("mode"))
newtitle = trim(request("newtitle"))
RichEdit_Text = trim(request("RichEdit_Text"))
RichEdit_Html = trim(request("RichEdit_Html"))

dim esm
set esm = server.createobject("easymail.SignManager")
esm.Load Session("wem")

allnum = esm.count
isfull = false
if allnum >= CInt(Application("em_MaxSigns")) then
	isfull = true
end if

if isfull = false and newtitle <> "" and RichEdit_Text <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	esm.Add newtitle, RichEdit_Text, RichEdit_Html
	esm.Save
	set esm = nothing

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("Signature.asp?sindex=" & allnum & "&isrtf=" & isrtf)
end if


if mode = "edit" and trim(request("settitle")) <> "" and RichEdit_Text <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	esm.Modify sindex, trim(request("settitle")), RichEdit_Text, RichEdit_Html
	esm.Save
	set esm = nothing

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("Signature.asp?sindex=" & sindex & "&isrtf=" & isrtf)
end if


if mode = "del" then
	esm.DeleteByIndex sindex
	esm.Save
	set esm = nothing

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("Signature.asp?isrtf=" & isrtf)
end if


dim ei
set ei = server.createobject("easymail.UserWeb")
ei.Load Session("wem")

if sindex = -1 and mode <> "new" then
	if ei.defaultSign > -1 and allnum > 0 then
		esm.Get ei.defaultSign, stitle, stext, shtmltext
		sindex = ei.defaultSign
	else
		if isfull = true then
			esm.Get 0, stitle, stext, shtmltext
			sindex = 0
		end if
	end if

	stitle = NULL
	stext = NULL
	shtmltext = NULL
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
html {overflow: scroll; overflow-x: hidden; overflow-y: auto !important;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.cont_td {white-space:nowrap; height:26px; padding-left:8px; padding-right:8px;}
-->
</STYLE>
</HEAD>

<script language="JavaScript" type="text/javascript" src="rte/wrte1.js"></script>
<script language="JavaScript" type="text/javascript" src="rte/wrte2.js"></script>

<script type="text/javascript">
<!--
<%
if isrtf = true then
%>
initRTE("./rte/images/", "./rte/", "", false);
<%
end if
%>

function signselect_onchange() {
<%
if isfull = false then
%>
	if (document.f1.sindex.selectedIndex == 0)
		location.href = "Signature.asp?<%=getGRSN() %>&mode=new&isrtf=<%=isrtf %>";
	else
<%
end if
%>
		location.href = "Signature.asp?<%=getGRSN() %>&isrtf=<%=isrtf %>&sindex=" + document.f1.sindex.value;
}

function sdel() {
	if (confirm("<%=b_lang_036 %>") == false)
		return ;

	location.href = "Signature.asp?<%=getGRSN() %>&isrtf=<%=isrtf %>&mode=del&sindex=<%=sindex %>";
}
//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload()">
<FORM ACTION="Signature.asp" METHOD=POST NAME="f1">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_184 %>
</td></tr>
<tr><td class="block_top_td" style="height:12px; _height:14px;"></td></tr>
<tr><td align="left" class="cont_td">
<%=b_lang_185 %><%=s_lang_mh %>
<select name="sindex" class="drpdwn" onchange="javascript:signselect_onchange();">
<%
if isfull = false then
%>
<option value="">---<%=b_lang_186 %>---</option>
<%
end if

i = 0
do while i < allnum
	esm.Get i, stitle, stext, shtmltext

	if i <> sindex then
		Response.Write "<option value='" & i & "'>" & server.htmlencode(stitle) & "</option>" & Chr(13)
	else
		Response.Write "<option value='" & i & "' selected>" & server.htmlencode(stitle) & "</option>" & Chr(13)
	end if

	stitle = NULL
	stext = NULL
	shtmltext = NULL

	i = i + 1
loop

if sindex > -1 then
	esm.Get sindex, stitle, stext, shtmltext
end if
%>
</select>
	</td></tr>
<%
if isfull = false and sindex = -1 then
%>
	<tr><td align="left" class="cont_td">
	<%=b_lang_187 %><%=s_lang_mh %>
	<input type="text" name="newtitle" class='n_textbox' size="40" maxlength="64">
	</td></tr>
<%
else
%>
<input type="hidden" name="settitle" maxlength="64" value="<%=stitle %>">
<%
end if
%>
	<tr><td class="cont_td">
<%
if isrtf = false then
%>
<textarea name="RichEdit_Text" cols="100" rows="20" class='n_textarea'><%=stext %></textarea>
<%
else
%>
<script language="JavaScript" type="text/javascript">
<!--
<%
	if shtmltext <> "" then
		Response.Write "writeRichText('richedit', RemoveScript('" & RTESafe(shtmltext) & "'), 545, 231, true, false);"
	else
		if stext <> "" then
			html_text = replace(stext, "'", "&#39;")
			html_text = replace(html_text, "<", "&lt;")
			html_text = replace(html_text, ">", "&gt;")
			html_text = replace(html_text, Chr(13) & Chr(10), "<br>")
			html_text = replace(html_text, Chr(10) & Chr(13), "<br>")
			html_text = replace(html_text, Chr(13), "<br>")
			html_text = replace(html_text, Chr(10), "<br>")
			html_text = replace(html_text, "\", "\\")
		else
			html_text = ""
		end if

		Response.Write "writeRichText('richedit', '" & html_text & "', 545, 240, true, false);"
	end if
%>
//-->
</script>
<%
end if
%>
	</td></tr>
<tr><td class="block_top_td" style="height:8px;"></td></tr>

<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
<%
if sindex > -1 then
%>
<a class='wwm_btnDownload btn_blue' href="javascript:sdel();"><%=s_lang_del %></a>
<%
end if

if isrtf = true then
%>
<a class='wwm_btnDownload btn_blue' href="Signature.asp?<%=getGRSN() %>&mode=<%=mode %>&isrtf=False&sindex=<%=sindex %>"><%=b_lang_188 %></a>
<%
else
%>
<a class='wwm_btnDownload btn_blue' href="Signature.asp?<%=getGRSN() %>&mode=<%=mode %>&isrtf=True&sindex=<%=sindex %>"><%=b_lang_189 %></a>
<%
end if
%>
</td></tr>
</table>

<%
if isrtf = true then
%>
<div style="display:none;"><textarea name="RichEdit_Text" cols="0" rows="0"></textarea></div>
<%
end if
%>
<input name="RichEdit_Html" type="hidden">
<input name="mode" type="hidden">
<input name="isrtf" type="hidden" value="<%=isrtf %>">
</FORM>

<table width="92%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px #A5B6C8 solid; margin-top:30px; padding-bottom:10px;'>
	<tr>
	<td style="padding:4px; color:#444444;">
<%=b_lang_190 %><%=s_lang_mh %><br>
<font color="#901111">$DATE$</font><%=s_lang_mh %> <%=b_lang_191 %><br>
<font color="#901111">$TIME$</font><%=s_lang_mh %> <%=b_lang_192 %><br>
	</td>
	</tr>
</table>
</BODY>

<SCRIPT LANGUAGE=javascript>
<!--
function window_onload() {
}

function gosub()
{
<%
if isrtf = true then
%>
	updateRTE('richedit');

	document.f1.RichEdit_Text.value = getText(document.f1.richedit.value);
	document.f1.RichEdit_Html.value = document.f1.richedit.value;
<%
end if

if isfull = false and sindex = -1 then
%>
	if (document.f1.newtitle.value == "")
	{
		alert("<%=b_lang_193 %>");
		document.f1.newtitle.focus();
		return ;
	}

	document.f1.submit();
<%
else
%>
	document.f1.mode.value = "edit";
	document.f1.submit();
<%
end if
%>
}
//-->
</SCRIPT>
</HTML>

<%
set esm = nothing
set ei = nothing

stitle = NULL
stext = NULL
shtmltext = NULL


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
