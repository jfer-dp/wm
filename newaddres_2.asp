<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
wmode = trim(request("wmode"))
themax = trim(request("themax"))

if Len(themax) > 0 and IsNumeric(themax) = true then
	themax = CLng(themax)
else
	themax = 0
end if

if Request.ServerVariables("REQUEST_METHOD") = "POST" and wmode = "true" and themax > 0 then
	dim ads
	set ads = server.createobject("easymail.Addresses")
	ads.Load Session("wem")

	isok = false
	i = 0

	do while i < themax
		if trim(request("check" & i)) <> "" then
			if ads.Simple_Add_Email(trim(request("nid" & i)), trim(request("mail" & i))) = true then
				isok = true
			end if
		end if

		i = i + 1
	loop

	if isok = true then
		ads.Save
	end if

	set ads = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=ads_brow.asp"
end if


folders = trim(request("folders"))

s_shownum = trim(request("s_shownum"))
if Len(s_shownum) > 0 and IsNumeric(s_shownum) = true then
	s_shownum = CLng(s_shownum)
else
	s_shownum = 1
end if

s_showdays = trim(request("s_showdays"))
if Len(s_showdays) > 0 and IsNumeric(s_showdays) = true then
	s_showdays = CLng(s_showdays)
else
	s_showdays = 7
end if

dim infolist
set infolist = server.createobject("easymail.InfoList")
il_allnum = 0

if Request.ServerVariables("REQUEST_METHOD") = "POST" and Len(folders) > 0 then
	infolist.searchstring = folders

	infolist.ForNewAddress_RepeatNumber = s_shownum
	infolist.ForNewAddress_Days = s_showdays

	infolist.LoadMailBox_ForNewAddress Session("wem")

	il_allnum = infolist.getMailsCount
end if

if il_allnum = 0 then
	set infolist = nothing
	Response.Redirect "newaddres_1.asp?errstr=noadd&" & getGRSN()
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
.title_tr {white-space:nowrap; background:#f2f4f6; height:24px;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.st_l, .st_r {height:24px; text-align:center; white-space:nowrap; border-left:1px solid #A5B6C8; border-top:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8;}
.st_r {border-right:1px solid #A5B6C8;}
.cont_tr {background:white; height:24px;}
.cont_td {height:24px; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function gosub() {
	if (ischeck() == true)
	{
		document.f1.wmode.value = "true";
		document.f1.submit();
	}
}

function ischeck() {
	var i = 0;
	var theObj;

	for(; i<<%=il_allnum %>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}

function checkall(check) {
	var i = 0;
	var theObj;

	for(; i<<%=il_allnum %>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			theObj.checked = check;
	}
}

function allcheck_onclick() {
	if (document.f1.allcheck.checked == true)
		checkall(true);
	else
		checkall(false);
}
//-->
</SCRIPT>

<BODY>
<form action="newaddres_2.asp" method=post name="f1">
<input type="hidden" name="wmode">
<input type="hidden" name="themax" value="<%=il_allnum %>">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_329 %>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr class="title_tr">
	<td width="8%" class="st_l"><input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()"></td>
	<td width="46%" class="st_l"><%=a_lang_330 %></td>
	<td width="46%" class="st_r"><%=a_lang_331 %></td>
<%
i = 0
do while i < il_allnum
	infolist.getMailInfo i, idname, isread, priority, sendMail, sendName, subject, size, etime

	Response.Write "<tr class='cont_tr'>"
	Response.Write "<td align='center' class='cont_td'><input type='checkbox' name='check" & i & "' value='" & i & "'></td>"

	Response.Write "<td align='left' class='cont_td'><input type='text' size='35' class='n_textbox' name='mail" & i & "' value=""" & sendMail & """></td>"
	Response.Write "<td align='left' class='cont_td'><input type='text' size='35' class='n_textbox' name='nid" & i & "' value=""" & sendName & """></td>"
	Response.Write "</tr>" & Chr(13)

	idname = NULL
	isread = NULL
	priority = NULL
	sendMail = NULL
	sendName = NULL
	subject = NULL
	size = NULL
	etime = NULL

	i = i + 1
loop
%>
</table>
	</td></tr>

<tr><td align="left" style="background-color:white; padding-top:16px; padding-bottom:10px;">
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="newaddres_1.asp?<%=getGRSN() %>"><%=a_lang_332 %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=a_lang_333 %></a>
</td></tr>
</table>

</FORM>

<table width="92%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px #A5B6C8 solid; margin-top:80px; padding-bottom:10px;'>
	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:4px; color:#444444;">
	<%=a_lang_334 %><br>
	</td>
	</tr>
</table>
</BODY>
</HTML>

<%
set infolist = nothing
%>
