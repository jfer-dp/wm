<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" -->

<%
fid = trim(request("fid"))
gourl = trim(request("gourl"))
curdomain = Mid(Session("mail"), InStr(Session("mail"), "@") + 1)

ismanager = false
if isadmin() = true then
	ismanager = true
end if

dim poll
set poll = server.createobject("easymail.Poll")
poll.LoadOne fid

if poll.PI_HaveThisDomain(curdomain) = false and ismanager = false then
	set poll = nothing
	response.redirect "noadmin.asp"
end if

if poll.PI_Can_Poll(Session("wem")) = false or poll.PI_IsEnd = true then
	set poll = nothing
	Response.Redirect "poll_showone.asp?fid=" & fid & "&" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if

PI_Limit_Choose_Number = poll.PI_Limit_Choose_Number
allnum = poll.PI_ChooseCount

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	isok = true
	i = 0

	if PI_Limit_Choose_Number = 1 then
		if trim(request("sel")) <> "" then
			poll.PI_Poll_To Session("wem"), curdomain, Chr(9) & trim(request("sel"))
		else
			isok = false
		end if
	elseif PI_Limit_Choose_Number > 1 then
		votestr = ""
		do while i < allnum
			if trim(request("sel" & i)) <> "" then
				votestr = votestr & Chr(9) & i
			end if 

		    i = i + 1
		loop

		if votestr <> "" then
			poll.PI_Poll_To Session("wem"), curdomain, votestr
		else
			isok = false
		end if
	end if

	set poll = nothing

	if isok = true then
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
	else
		Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
	end if
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
.font_g {font-size:12px; color:#444444; font-weight:normal;}
.cont_td {height:27px; bgcolor:white; border-left:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px;}
.cont_td_word {height:27px; bgcolor:white; border-left:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px; word-break:break-all; word-wrap:break-word;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
var maxnumber = <%=PI_Limit_Choose_Number %>;

function checknum() {
	var i = 0;
	var theObj;
	var retnum = 0;

	for(; i < <%=allnum %>; i++)
	{
		theObj = eval("document.f1.sel" + i);

		if (theObj != null)
			if (theObj.checked == true)
				retnum++;
	}

	return retnum;
}

function gosub() {
<%
if PI_Limit_Choose_Number > 1 then
%>
	if (checknum() != maxnumber)
		alert("<%=b_lang_052 %>" + maxnumber + "<%=b_lang_053 %>");
	else
<%
end if
%>
	document.f1.submit();
}
//-->
</script>

<BODY>
<FORM ACTION="poll_vote.asp" METHOD="POST" NAME="f1">
<input name="fid" type="hidden" value="<%=fid %>">
<input name="gourl" type="hidden" value="<%=gourl %>">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px; padding-right:6px; word-break:break-all; word-wrap:break-word;">
<%
Response.Write server.htmlencode(poll.PI_Title)

if PI_Limit_Choose_Number > 1 then
	Response.Write "&nbsp;<font class='font_g'>(" & b_lang_052 & PI_Limit_Choose_Number & b_lang_053 & ")</font>"
end if
%>
</td></tr>
<tr><td align="center">

<%
i = 0

do while i < allnum
	poll.PI_GetNameAndNumber i, v_name, v_num

	if PI_Limit_Choose_Number = 1 then
%>
	<tr><td height="27" align="left" style="border-bottom:1px solid #A5B6C8;">
	&nbsp;<input name="sel" type="radio" value="<%=i %>"><%=server.htmlencode(v_name) %>
	</td></tr>
<%
	else
%>
	<tr><td height="27" align="left" style="border-bottom:1px solid #A5B6C8;">
	&nbsp;<input name="sel<%=i %>" type="checkbox" value="<%=i %>"><%=server.htmlencode(v_name) %>
	</td></tr>
<%
	end if

	i = i + 1

	v_name = NULL
	v_num = NULL
loop
%>
	<tr><td align="left" height="26"><br>
<a class='wwm_btnDownload btn_blue' href="javascript:history.back();"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=b_lang_054 %></a>
	</td></tr>
</table>
</form>
</BODY>
</HTML>

<%
set poll = nothing
%>
