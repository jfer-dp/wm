<!--#include file="passinc.asp" -->
<!--#include file="language.asp" -->

<%
errstr = trim(request("errstr"))
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
</HEAD>

<BODY>
<table width="90%" align="center" border="0" cellspacing="0" cellpadding="0" style="margin-top:40px;">
	<tr style="background:#EFF7FF; color:#104A7B;">
<%
if errstr = "" then
%>
    <td height="30" style="border:1px solid #8CA5B5;">&nbsp;<img src="images/error.gif" align="absmiddle" border="0">&nbsp;<%=s_lang_0558 %><%=s_lang_mh %><%=s_lang_0559 %><%=s_lang_jh %></td>
<%
else
%>
    <td height="30" style="border:1px solid #8CA5B5;">&nbsp;<img src="images/error.gif" align="absmiddle" border="0">&nbsp;<%=s_lang_0558 %><%=s_lang_mh %><%=errstr %><%=s_lang_jh %></td>
<%
end if
%>
	</tr>
	<tr><td style="border-bottom:1px solid #8CA5B5; height:24px;">&nbsp;</td></tr>
	<tr><td style="height:24px;">&nbsp;</td></tr>
	<tr><td align="right" style="padding-right:30px;">
<a class="wwm_btnDownload btn_blue" href="javascript:history.back();"><%=s_lang_0313 %></a>
	</td></tr></table>
</body>
</html>
