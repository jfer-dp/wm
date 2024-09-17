<%
Response.CacheControl = "no-cache" 
%>

<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" -->

<%
if isadmin() = false then
	Response.Redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.MailboxBanjia")
ei.Load

bj_is_null = true

if ei.IsRun = false and ei.Count < 1 and Len(ei.Server_Name) < 1 then
	bj_is_null = true
else
	bj_is_null = false
end if

if bj_is_null = false and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	p_mode = trim(request("mode"))

	if p_mode = "start" then
		ei.IsRun = true
	elseif p_mode = "stop" then
		ei.IsRun = false
	elseif p_mode = "del" then
		ei.DeleteBJ
	end if

	ei.Save
	set ei = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("banjia.asp")
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
html {overflow: scroll; overflow-x: hidden; overflow-y: auto !important;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.td_line_l {text-align:right; white-space:nowrap; background-color:#EFF7FF; border-bottom:1px #A5B6C8 solid; height:30px; color:#303030;}
.td_line_r {text-align:left; background-color:white; border-bottom:1px #A5B6C8 solid; height:30px; padding-left:6px;}
</STYLE>
</HEAD>

<script type="text/javascript">
function window_onload() {
}

function iFrameHeight() {
	var ifm= document.getElementById("iframepage");
	var subWeb = document.frames ? document.frames["iframepage"].document : ifm.contentDocument;
	if(ifm != null && subWeb != null) {
		if (subWeb.body.offsetHeight < 400)
			ifm.height = subWeb.body.offsetHeight;
		else
			ifm.height = 400;
	}
}

function start() {
	document.fm1.mode.value = "start";
	document.fm1.submit();
}

function stop() {
	document.fm1.mode.value = "stop";
	document.fm1.submit();
}

function del() {
	if (confirm("<%=b_lang_036 %>") == false)
		return ;

	document.fm1.mode.value = "del";
	document.fm1.submit();
}

function edit() {
<%
if bj_is_null = false then
	if ei.IsRun = false then
%>
	location.href = "bj_new.asp?mode=edit&<%=getGRSN() %>"
<%
	else
%>
	alert("<%=b_lang_373 %>");
<%
	end if
end if
%>
}

function del_ok_and_edit() {
<%
if bj_is_null = false then
	if ei.IsRun = false then
%>
	location.href = "bj_new.asp?mode=delokedit&<%=getGRSN() %>"
<%
	else
%>
	alert("<%=b_lang_373 %>");
<%
	end if
end if
%>
}
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload()">
<form method="post" action="#" name="fm1">
<input name="mode" type="hidden">
<table width="92%" border="0" align="center" cellspacing="0">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_374 %>
</td></tr>
<tr><td class="block_top_td" style="height:18px; _height:20px;"></td></tr>
</table>
<%
if bj_is_null = true then
%>
<table width="88%" border="0" align="center" cellspacing="0">
<tr><td valign="center" align="center"><a class='wwm_btnDownload btn_blue' href="bj_new.asp?<%=getGRSN() %>"><%=b_lang_375 %></a></td></tr>
</table>
<%
else
%>
<table width="88%" border="0" align="center" cellspacing="0">
	<tr>
	<td nowrap align="left" width="25%" style="padding-left:14px;"><font color="#606060"><%=b_lang_376 %></font><%
if ei.IsRun = true then
	Response.Write "<font color='#007947'>" & b_lang_377 & "</font>"
else
	Response.Write b_lang_378
end if
%></td>
	<td nowrap align="left" width="25%"><font color="#606060"><%=b_lang_379 %></font><%=server.htmlencode(ei.Server_Name) %></td>
	<td nowrap align="left" width="25%"><font color="#606060"><%=b_lang_380 %></font><%=server.htmlencode(ei.Server_Port) %></td>
	<td nowrap align="left" width="25%"><font color="#606060"><%=b_lang_381 %></font><%=server.htmlencode(ei.All_Password) %></td>
	</tr>
<tr><td colspan="4" style="border-bottom:1px #deab8a solid; font-size:6px; font-weight:bold; color:#093665; padding-left:6px;">&nbsp;</td></tr>
	<tr><td colspan="4" style="padding-left:12px; padding-right:12px;">
<iframe src="bj_iframe.asp?<%=getGRSN() %>" id="iframepage" name="iframepage" scrolling="AUTO" frameBorder=0 width="100%" onLoad="iFrameHeight()" style="padding-top:14px;"></iframe>
	</td></tr>
</table>
<%
end if
%>
<table width="92%" border="0" align="center" cellspacing="0">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">&nbsp;</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>

<tr><td align="right">
<a class='wwm_btnDownload btn_blue' href="right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
&nbsp;&nbsp;<a class='wwm_btnDownload btn_blue' href="banjia.asp?<%=getGRSN() %>"><%=b_lang_391 %></a>
<%
if bj_is_null = false then
	if ei.IsRun = false then
%>
&nbsp;&nbsp;<a class='wwm_btnDownload btn_blue' href="javascript:start();"><%=b_lang_382 %></a>
<%
	else
%>
&nbsp;&nbsp;<a class='wwm_btnDownload btn_blue' href="javascript:stop();"><%=b_lang_383 %></a>
<%
	end if
%>
&nbsp;&nbsp;<a class='wwm_btnDownload btn_blue' href="javascript:edit();"><%=s_lang_modify %></a>
&nbsp;&nbsp;<a class='wwm_btnDownload btn_blue' href="javascript:del_ok_and_edit();"><%=b_lang_384 %></a>
&nbsp;&nbsp;<a class='wwm_btnDownload btn_blue' href="javascript:del();"><%=s_lang_del %></a>
<%
end if
%>
</td></tr>
</table>
<br>
</Form>
</BODY>
</HTML>

<%
set ei = nothing
%>
