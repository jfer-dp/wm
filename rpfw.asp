<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" --> 

<%
dim ei
set ei = server.createobject("easymail.emmail")
ei.Load_RP_FW Session("wem"), Session("mail"), ""

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ei.ReplyTemplet = trim(request("rptext"))
	ei.ForwardTemplet = trim(request("fwtext"))

	ei.Save_RP_FW
	set ei = nothing

	Response.Redirect "ok.asp?gourl=rpfw.asp"
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
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function gosub()
{
	if (document.f1.rptext.value.length > 4090)
		document.f1.rptext.value = document.f1.rptext.value.substring(0, 4090);

	if (document.f1.fwtext.value.length > 4090)
		document.f1.fwtext.value = document.f1.fwtext.value.substring(0, 4090);

	document.f1.submit();
}
//-->
</SCRIPT>

<BODY>
<FORM ACTION="rpfw.asp" METHOD="POST" NAME="f1">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_055 %>
</td></tr>
<tr><td class="block_top_td" style="height:8px; _height:10px;"></td></tr>

<tr><td align="left" style="padding-left:6px;">
<textarea name="rptext" cols="80" rows="8" class="n_textarea"><%=ei.ReplyTemplet %></textarea>
</td></tr>

<tr><td align="left" style="background-color:white; padding-top:10px; padding-bottom:10px;">
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
</td></tr>
</table>

<table width="92%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px #A5B6C8 solid; margin-top:20px; padding-bottom:10px;'>
	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:4px; color:#444444;">
�ظ�ģ������ʹ�õĺ����:<br>
<font color="#901111">$QFROMNAME$</font> : ԭ�ʼ�����������<br>
<font color="#901111">$QFROMADDR$</font> : ԭ�ʼ��������ʼ���ַ<br>
<font color="#901111">$QDATE$</font> : ��������<br>
<font color="#901111">$QTIME$</font> : ����ʱ��<br>
<font color="#901111">$QSUBJ$</font> : ��������<br>
<font color="#901111">$QUOTES$</font> : ����ԭ�ʼ�����<br>
<font color="#901111">$NAME$</font> : ��������<br>
<font color="#901111">$ADDR$</font> : �������ʼ���ַ<br>
<font color="#901111">$DATE$</font> : ��ǰ����<br>
<font color="#901111">$TIME$</font> : ��ǰʱ��<br>
	</td>
	</tr>
</table>

<br><br>
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_056 %>
</td></tr>
<tr><td class="block_top_td" style="height:8px; _height:10px;"></td></tr>

<tr><td align="left" style="padding-left:6px;">
<textarea name="fwtext" cols="80" rows="8" class="n_textarea"><%=ei.ForwardTemplet %></textarea>
</td></tr>

<tr><td align="left" style="background-color:white; padding-top:10px; padding-bottom:10px;">
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
</td></tr>
</table>

<table width="92%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px #A5B6C8 solid; margin-top:20px; padding-bottom:10px;'>
	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:4px; color:#444444;">
ת��ģ������ʹ�õĺ����:<br>
<font color="#901111">$QFROMNAME$</font> : ԭ�ʼ�����������<br>
<font color="#901111">$QFROMADDR$</font> : ԭ�ʼ��������ʼ���ַ<br>
<font color="#901111">$QDATE$</font> : ��������<br>
<font color="#901111">$QTIME$</font> : ����ʱ��<br>
<font color="#901111">$QSUBJ$</font> : ��������<br>
<font color="#901111">$QTEXT$</font> : ԭ�ʼ�����<br>
<font color="#901111">$NAME$</font> : ��������<br>
<font color="#901111">$ADDR$</font> : �������ʼ���ַ<br>
<font color="#901111">$DATE$</font> : ��ǰ����<br>
<font color="#901111">$TIME$</font> : ��ǰʱ��<br>
	</td>
	</tr>
</table>

</FORM>
</BODY>
</HTML>

<%
set ei = nothing
%>
