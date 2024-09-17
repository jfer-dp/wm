<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" --> 

<%
foldername = trim(request("foldername"))

dim perfolders
set perfolders = server.createobject("easymail.PerFolders")

perfolders.Load Session("wem")

if trim(request("save")) = "1" and perfolders.CanSetWithReceiveOutMail(foldername) = true and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if trim(request("enableRecOutMail")) = "" then
		perfolders.Set_EnableRecOutMail foldername, false
	else
		perfolders.Set_EnableRecOutMail foldername, true
	end if

	if trim(request("enableAutoForward")) = "" then
		perfolders.Set_EnableAutoForward foldername, false
	else
		perfolders.Set_EnableAutoForward foldername, true
	end if

	if trim(request("enableLocalSave")) = "" then
		perfolders.Set_EnableLocalSave foldername, false
	else
		perfolders.Set_EnableLocalSave foldername, true
	end if

	if trim(request("enableSave2InBox")) = "" then
		perfolders.Set_EnableSave2InBox foldername, false
	else
		perfolders.Set_EnableSave2InBox foldername, true
	end if

	perfolders.Set_AutoReplyEx_Name foldername, trim(request("autoreplyex_name"))
	perfolders.Set_AutoForward_Mail foldername, trim(request("af_mail"))
	perfolders.Save

	set perfolders = nothing
	Response.Redirect "viewmailbox.asp?" & getGRSN()
end if

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
.cont_td {height:24px; white-space:nowrap; border-bottom:1px solid #A5B6C8; padding-left:2px; padding-right:2px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function gosub()
{
	document.f1.submit();
}
//-->
</script>

<body>
<form name="f1" method="post" action="pf_setupfolder.asp">
<input type="hidden" name="foldername" value="<%=foldername %>">
<input type="hidden" name="save" value="1">
<%
if foldername = "att" then
	foldername = b_lang_000
end if
%>
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px; padding-right:6px; word-break:break-all; word-wrap:break-word;">
<%=b_lang_001 %>&nbsp;(<%=server.htmlencode(foldername) %>)
<%
if perfolders.CanSetWithReceiveOutMail(foldername) = false then
	Response.Write "&nbsp;<font color='#444444' style='font-size:12px; font-weight:normal;'>" & b_lang_002 & "</font>"
end if
%>
</td></tr>
<tr><td class="block_top_td" style="height:4px; _height:6px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr>
	<td width="30%" align="right" class="cont_td"><%=b_lang_003 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td">
<input type="checkbox" name="enableRecOutMail" value="checkbox" <% if perfolders.Get_EnableRecOutMail(foldername) = true then response.write "checked"%>> <%=b_lang_004 %>
	</td>
	</tr>

	<tr>
	<td align="right" class="cont_td"><%=b_lang_005 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td"><input type="checkbox" name="enableSave2InBox" value="checkbox" <% if perfolders.Get_EnableSave2InBox(foldername) = true then response.write "checked"%>>
	</td>
	</tr>

	<tr>
    <td align="right" class="cont_td"><%=b_lang_006 %><%=s_lang_mh %></td>
    <td align="left" class="cont_td"><select name="autoreplyex_name" class="drpdwn">
<option value=''><%=b_lang_007 %></option>
<%
i = 0
allnum = arex.count

fn_arename = perfolders.Get_AutoReplyEx_Name(foldername)

do while i < allnum
	arex.Get i, are_name, are_subject, are_text

	if fn_arename <> are_name then
		response.write "<option value='" & server.htmlencode(are_name) & "'>" & server.htmlencode(are_name) & "</option>" & Chr(13)
	else
		response.write "<option value='" & server.htmlencode(are_name) & "' selected>" & server.htmlencode(are_name) & "</option>" & Chr(13)
	end if

	are_name = NULL
	are_subject = NULL
	are_text = NULL

	i = i + 1
loop
%>
</select>
	</td>
	</tr>

	<tr>
	<td align="right" class="cont_td"><%=b_lang_008 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td"><input type="checkbox" name="enableAutoForward" value="checkbox" <% if perfolders.Get_EnableAutoForward(foldername) = true then response.write "checked"%>>
	</td>
	</tr>

	<tr>
    <td align="right" class="cont_td"><%=b_lang_009 %><%=s_lang_mh %></td>
    <td align="left" class="cont_td"><input type="text" name="af_mail" value="<%=perfolders.Get_AutoForward_Mail(foldername) %>" maxlength="64" size="40" class="n_textbox"></td>
	</tr>

	<tr>
	<td align="right" class="cont_td"><%=b_lang_010 %><%=s_lang_mh %></td>
	<td align="left" class="cont_td"><input type="checkbox" name="enableLocalSave" value="checkbox" <% if perfolders.Get_EnableLocalSave(foldername) = true then response.write "checked"%>></td>
	</tr>

<tr><td align="left" colspan="2" style="background-color:white; padding-right:16px; padding-top:18px; padding-bottom:10px;">
<a class='wwm_btnDownload btn_blue' href="javascript:history.back();"><< <%=s_lang_return %></a>
<%
if perfolders.CanSetWithReceiveOutMail(foldername) = true then
%>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
<%
end if
%>
	</td></tr>
</table>
</td></tr>
</table>

<table width="89%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px #A5B6C8 solid; margin-top:80px; padding-bottom:10px;'>
	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:6px; color:#444444;">
	������˽���ļ���Ϊ��������ⲿ�ʼ�ʱ, �൱�ڴ�����һ����������ʼ���ַ. 
	<br>��: ������ʵ�ʼ���ַΪ: user@domain.com ʱ, ��������һ����������ⲿ�ʼ���˽���ļ���: friend, ��������ӵ����һ�����ʼ���ַ: friend~user@domain.com
	<br>����ʼ���ַ��: ˽���ļ�������(<font color="#901111">friend</font>) + <font color="#901111">~</font> + ��ʵ���ʼ���ַ(<font color="#901111">user@domain.com</font>) ���.
	����д�� friend~user@domain.com ���ʼ����ᱻ������˽���ļ��� friend ��(������ǿ��Ҫ���䱣�����ռ�����).
	</td>
	</tr>

	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:6px; color:#444444;">
	�������ڽ����ⲿ�ʼ�ʱ, ˽���ļ������Ƶĳ��Ȳ����Գ���32���ֽ�, ���������в����԰���������: < @ > ~ �����������ַ�.
	</td>
	</tr>

	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:6px; color:#444444;">
	������Ϊÿһ����������ʼ���ַ���ò�ͬ���Զ�ת���Լ��Զ��ظ���Ϣ.
	</td>
	</tr>
</table>
</form>
</BODY>
</HTML>

<%
set perfolders = nothing
set arex = nothing
%>
