<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" -->

<%
if isadmin() = false then
	Response.Redirect "noadmin.asp"
end if
%>

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if
%>

<%
dim ei

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	set ei = server.createobject("easymail.MailboxBanjia")
	ei.Load
	ei.DeleteBJ

	ei.Load
	ei.Server_Name = trim(request("Server_Name"))
	ei.All_Password = trim(request("All_Password"))

	if IsNumeric(trim(request("Server_Port"))) = true then
		ei.Server_Port = CLng(trim(request("Server_Port")))
	end if

	ei.Add_All trim(request("bjusers"))
	ei.Save

	set ei = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("banjia.asp")
end if

is_edit = false
is_delok_edit = false

if Request.ServerVariables("REQUEST_METHOD") = "GET" then
	if LCase(trim(request("mode"))) = "edit" then
		is_edit = true

		set ei = server.createobject("easymail.MailboxBanjia")
		ei.Load
	elseif LCase(trim(request("mode"))) = "delokedit" then
		is_delok_edit = true

		set ei = server.createobject("easymail.MailboxBanjia")
		ei.Load
	end if
end if


dim dm
set dm = server.createobject("easymail.domain")
dm.Load
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<STYLE type=text/css>
html {overflow: scroll; overflow-x: hidden; overflow-y: auto !important;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.td_line_l {text-align:right; white-space:nowrap; background-color:#EFF7FF; border-bottom:1px #A5B6C8 solid; height:30px; color:#303030;}
.td_line_r {text-align:left; background-color:white; border-bottom:1px #A5B6C8 solid; height:30px; padding-left:6px;}
</STYLE>
</HEAD>

<script type="text/javascript" src="images/jquery.min.js"></script>

<script type="text/javascript">
function isinlist(name)
{
	var i = 0;
	for (; i < document.f1.bjuserlist.length; i++)
	{
		if (document.f1.bjuserlist[i].value == name)
			return true;
	}

	return false;
}

function add() {
	var i = 0;
	for (; i < document.f1.sysuserlist.length; i++)
	{
		if (document.f1.sysuserlist[i].selected == true)
		{
			if (isinlist(document.f1.sysuserlist[i].value) == false)
			{
				var oOption = document.createElement("OPTION");
				oOption.text = document.f1.sysuserlist[i].value;
				oOption.value = document.f1.sysuserlist[i].value;
<%
if isMSIE = true then
%>
				document.f1.bjuserlist.add(oOption);
<%
else
%>
				document.f1.bjuserlist.appendChild(oOption);
<%
end if
%>
			}
		}
	}
}

function del()
{
	var i = 0;
	for (; i < document.f1.bjuserlist.length; i++)
	{
		if (document.f1.bjuserlist[i].selected == true)
		{
			document.f1.bjuserlist.remove(i);
			i--;
		}
	}
}

function gosub() {
	if (document.getElementById("Server_Name").value.length < 5)
	{
		alert("<%=s_lang_inputerr %>");
		document.getElementById("Server_Name").focus();
		return ;
	}

	if (document.getElementById("Server_Port").value.length < 2)
	{
		alert("<%=s_lang_inputerr %>");
		document.getElementById("Server_Port").focus();
		return ;
	}

	if (document.getElementById("All_Password").value.length < 1)
	{
		alert("<%=s_lang_inputerr %>");
		document.getElementById("All_Password").focus();
		return ;
	}

	var bjuserstr = "";
	var i = 0;
	for (; i < document.f1.bjuserlist.length; i++)
	{
		if (i == 0)
			bjuserstr = document.f1.bjuserlist[i].value;
		else
			bjuserstr = bjuserstr + "\t" + document.f1.bjuserlist[i].value;
	}

	document.f1.bjusers.value = bjuserstr;
	document.f1.submit();
}

function domainname_onchange() {
	var theObj = document.getElementById("domainname");
	var post_date = "domain=" + theObj.value + "&<%=getGRSN() %>";
	send_star(post_date);
}

function send_star(post_date)
{
$.ajax({
	type:"POST",
	url:"ajgetuser.asp",
	data:post_date,
	success:function(data){
		var theObj;
		theObj = document.getElementById("div_select");

		if (theObj != null)
			theObj.innerHTML = data;
	},
	error:function(){
	}
});
}
</SCRIPT>

<BODY>
<form method="post" action="#" name="f1">
<input name="bjusers" type="hidden">
<table width="92%" border="0" align="center" cellspacing="0">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_374 %> - 
<%
if is_edit = true or is_delok_edit = true then
	Response.Write s_lang_modify
else
	Response.Write b_lang_385
end if
%>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
</table>

<table width="88%" border="0" align="center" cellspacing="0">
	<tr>
	<td align="left" style="padding-left:12px;"><%=b_lang_379 %>&nbsp;<input type="text" id="Server_Name" name="Server_Name" class='n_textbox' size="32" maxlength="64"<%
if is_edit = true or is_delok_edit = true then
	Response.Write " value='" & server.htmlencode(ei.Server_Name) & "'"
end if
%>></td>
	</tr>
	<tr>
	<td align="left" style="padding-left:12px;"><%=b_lang_380 %>&nbsp;<input type="text" id="Server_Port" name="Server_Port" class='n_textbox' style="text-align:right;" size="6" maxlength="5"<%
if is_edit = true or is_delok_edit = true then
	Response.Write " value='" & server.htmlencode(ei.Server_Port) & "'"
else
	Response.Write " value='110'"
end if
%>></td>
	</tr>
	<tr>
	<td align="left" style="padding-left:12px;"><%=b_lang_381 %>&nbsp;<input type="text" id="All_Password" name="All_Password" class='n_textbox' size="32" maxlength="64"<%
if is_edit = true or is_delok_edit = true then
	Response.Write " value='" & server.htmlencode(ei.All_Password) & "'"
end if
%>></td>
	</tr>
<tr><td style="border-bottom:1px #deab8a solid; font-size:10px; font-weight:bold; color:#093665; padding-left:6px;">&nbsp;</td></tr>
</table>

<table width="88%" border="0" align="center" cellspacing="0">
	<tr>
	<td align="left" style="padding-left:12px;" colspan="3">
<select id="domainname" name="domainname" class="drpdwn" style="margin-top:12px; margin-bottom:4px;" LANGUAGE=javascript onchange="domainname_onchange()">
<option value="">[<%=b_lang_386 %>]</option>
<option value="all"><%=b_lang_387 %></option>
<%
i = 0
allnum = dm.GetCount()

do while i < allnum
	domain = dm.GetDomain(i)
	Response.Write "<option value='" & server.htmlencode(domain) & "'>" & server.htmlencode(domain) & "</option>" & Chr(13)

	domain = NULL
	i = i + 1
loop
%>
</select>
	</td>
	</tr>

  <tr valign=top> 
	<td style="padding-left:12px;">
<div id="div_select">
	<select class="drpdwn" style="WIDTH:300px" multiple size=22 id="sysuserlist" name="sysuserlist" width="300">
	</select>
</div>
	</td>
    <td valign="middle" align="center"> 
	<table cellspacing=0 cellpadding=0>
	<tr><td>
	<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="add()" type=button value="<%=s_lang_add %> >>">
	</td></tr>
	<tr><td><br>
	<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="del()" type=button value="<< <%=s_lang_del %>">
	</td></tr>
	</table>
	</td>
	<td>
	<select class="drpdwn" style="WIDTH:300px" multiple size=22 id="bjuserlist" name="bjuserlist" width="300">
<%
if is_edit = true or is_delok_edit = true then
	i = 0
	allnum = ei.Count

	do while i < allnum
		ei.Get i, s_name, s_state

		if is_delok_edit = true then
			if s_state <> 1 then
				Response.Write "<option value=""" & server.htmlencode(s_name) & """>" & server.htmlencode(s_name) & "</option>" & Chr(13)
			end if
		else
			Response.Write "<option value=""" & server.htmlencode(s_name) & """>" & server.htmlencode(s_name) & "</option>" & Chr(13)
		end if

		s_name = NULL
		s_state = NULL
		i = i + 1
	loop
end if
%>
	</select>
	</td>
  </tr>
</table>

<table width="92%" border="0" align="center" cellspacing="0">
<tr><td class="block_top_td" style="height:1px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">&nbsp;</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>

<tr><td align="right">
<a class='wwm_btnDownload btn_blue' href="banjia.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>&nbsp;&nbsp;
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
</td></tr>
</table>
<br>
</Form>
</BODY>
</HTML>

<%
if is_edit = true or is_delok_edit = true then
	set ei = nothing
end if

set dm = nothing
%>
