<!--#include file="passinc.asp" --> 
<!--#include file="language-2.asp" --> 

<%
dim arex
set arex = server.createobject("easymail.AutoReplyEx")
arex.Load Session("wem")
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
.cont_td {white-space:nowrap; height:26px; padding-left:14px; padding-right:4px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function arindex_onchange() {
	location.href = "showautoreplyex.asp?<%=getGRSN() %>&selectar=" + document.f1.arindex.value + "&returl=" + document.f1.returl.value;
}

function onsub() {
	if (document.f1.arindex.value != -1)
		document.f1.submit();
	else
	{
		if (document.f1.mare_name.value != "" && (document.f1.mare_subject.value != "" || document.f1.mare_text.value != ""))
			document.f1.submit();
	}
}

function ondel() {
	if (document.f1.arindex.value != -1)
	{
		document.f1.isdel.value = "yes";
		document.f1.submit();
	}
}
//-->
</SCRIPT>

<BODY>
<FORM ACTION="saveautoreplyex.asp" METHOD=POST NAME="f1">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_096 %>
</td></tr>
<tr><td class="block_top_td" style="height:10px; _height:12px;"></td></tr>
<tr><td align="left" class="cont_td" style="padding-bottom:4px;">
<%=b_lang_097 %><%=s_lang_mh %>
<select name="arindex" class="drpdwn" LANGUAGE=javascript onchange="return arindex_onchange()">
<option value='-1'><%=b_lang_098 %></option>
<%
i = 0
allnum = arex.count
curselect = -1

do while i < allnum
	arex.Get i, are_name, are_subject, are_text

	if IsNumeric(trim(request("selectar"))) = true then
		if i <> CInt(trim(request("selectar"))) then
			response.write "<option value='" & i & "'>" & server.htmlencode(are_name) & "</option>" & Chr(13)
		else
			curselect = i
			response.write "<option value='" & i & "' selected>" & server.htmlencode(are_name) & "</option>" & Chr(13)
		end if
	else
		response.write "<option value='" & i & "'>" & server.htmlencode(are_name) & "</option>" & Chr(13)
	end if

	are_name = NULL
	are_subject = NULL
	are_text = NULL

	i = i + 1
loop


if curselect = -1 then
	are_name = ""
	are_subject = ""
	are_text = ""
else
	arex.Get curselect, are_name, are_subject, are_text
end if
%>
</select>
	</td></tr>
<%
if curselect = -1 then
%>
	<tr><td align="left" class="cont_td">
	<%=b_lang_099 %><%=s_lang_mh %>
	<input name="mare_name" type="text" value="" size="30" maxlength="30" class='n_textbox'>
	</td></tr>
<%
end if
%>
	<tr><td align="left" class="cont_td">
	<%=b_lang_100 %><%=s_lang_mh %>
	<input name="mare_subject" type="text" value="<%=are_subject %>" size="60" maxlength="120" class='n_textbox'>
	</td></tr>

	<tr><td align="left" class="cont_td">
	<textarea name="mare_text" cols="80" rows="8" class='n_textarea'><%=are_text %></textarea>
	</td></tr>

<tr><td class="block_top_td" style="height:10px;"></td></tr>

<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<%
if trim(request("returl")) = "" then
%>
	<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<%
else
%>
	<a class='wwm_btnDownload btn_blue' href="<%=trim(request("returl")) %>?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<%
end if
%>


<a class='wwm_btnDownload btn_blue' href="javascript:onsub();"><%=s_lang_save %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:ondel();"><%=s_lang_del %></a>
</td></tr>
</table>

<input type="hidden" name="isdel">
<input type="hidden" name="returl" value="<%=trim(request("returl")) %>">
</FORM>


<table width="92%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px #A5B6C8 solid; margin-top:50px; padding-bottom:10px;'>
	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:4px; color:#444444;">
您可以在上述的增强型自动回复信息中使用宏变量:<br>
<font color="#901111">%date%</font> 表示当前日期<br>
<font color="#901111">%time%</font> 表示当前时间<br>
<font color="#901111">%sendname%</font> 表示来信的发件人名称<br>
<font color="#901111">%sendmail%</font> 表示来信的发件人邮件地址<br>
<font color="#901111">%subject%</font> 表示来信的标题<br><br>
<font color="#901111">注意</font>: 此功能可在<a href="userfiltermail.asp?<%=getGRSN() %>">邮件过滤</a>中进行调用.
	</td>
	</tr>
</table>
</BODY>
</HTML>

<%
are_name = NULL
are_subject = NULL
are_text = NULL

set arex = nothing
%>
