<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" --> 

<%
dim ei
set ei = server.createobject("easymail.usermessages")
ei.Load Session("wem")

allnum = ei.GetRejectEMailCount

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if trim(request("checkrej")) <> "" then
		ei.UseReject = true
	else
		ei.UseReject = false
	end if

	ei.RemoveAllRejectEMail

	dim msg
	msg = trim(request("allmsgs"))

	if Len(msg) > 0 or Len(addname) > 0 then
		dim item
		dim ss
		dim se
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(9))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				ei.AddRejectEMail item
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	ei.SaveReject
	set ei = nothing

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=showuserkill.asp"
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
.sbttn {font-family:<%=s_lang_font %>;font-size:9pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5;color:#000066;height:19px;text-decoration:none;cursor:pointer}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.cont_td {height:24px; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript" src="images/mglobal.js"></script>

<script type="text/javascript">
<!--
function sub()
{
	var tempstr = "";
	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		tempstr = tempstr + document.f1.listall[i].value + "\t";
	}

	document.f1.allmsgs.value = tempstr;
	document.f1.action = "#";
	document.f1.method = "POST";
	document.f1.submit();
}

function delout()
{
	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].selected == true)
		{
			document.f1.listall.remove(i);
			i--;
		}
	}
}

function add()
{
	if (document.f1.addmsg.value.indexOf("\t") != -1)
	{
		alert("<%=s_lang_inputerr %>");
		document.f1.addmsg.focus();
		return ;
	}

	if (document.f1.addmsg.value.length > 0)
	{
		if (haveit() == false)
		{
			var oOption = document.createElement("OPTION");
			oOption.text = document.f1.addmsg.value;
			oOption.value = document.f1.addmsg.value;

			if (ie == false)
				document.getElementById("listall").appendChild(oOption);
			else
				document.getElementById("listall").add(oOption);

			return ;
		}
		else
			return ;
	}

	alert("<%=s_lang_inputerr %>");
}

function haveit()
{
	var tempstr = document.f1.addmsg.value;

	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].value == tempstr)
			return true;
	}

	return false;
}

function mdown()
{
	var i = 0;
	var findit = -1;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].selected == true)
		{
			findit = i;
			break;
		}
	}

	for (i = 0; i < document.f1.listall.length; i++)
	{
		if (findit != i)
			document.f1.listall[i].selected = false;
	}

	if (findit > -1 && findit < document.f1.listall.length - 1)
	{
		var tempstr = document.f1.listall[findit + 1].text;
		document.f1.listall[findit + 1].text = document.f1.listall[findit].text;
		document.f1.listall[findit].text = tempstr;

		tempstr = document.f1.listall[findit + 1].value;
		document.f1.listall[findit + 1].value = document.f1.listall[findit].value;
		document.f1.listall[findit].value = tempstr;

		document.f1.listall[findit + 1].selected = true;
		document.f1.listall[findit].selected = false;
	}
}

function mup()
{
	var i = 0;
	var findit = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].selected == true)
		{
			findit = i;
			break;
		}
	}

	for (i = 0; i < document.f1.listall.length; i++)
	{
		if (findit != i)
			document.f1.listall[i].selected = false;
	}

	if (findit > 0)
	{
		var tempstr = document.f1.listall[findit - 1].text;
		document.f1.listall[findit - 1].text = document.f1.listall[findit].text;
		document.f1.listall[findit].text = tempstr;

		tempstr = document.f1.listall[findit - 1].value;
		document.f1.listall[findit - 1].value = document.f1.listall[findit].value;
		document.f1.listall[findit].value = tempstr;

		document.f1.listall[findit - 1].selected = true;
		document.f1.listall[findit].selected = false;
	}
}

function goent() {
	if (ie != false && event.keyCode == 13)
	{
		event.keyCode = 9;
		add();
	}
}
//-->
</SCRIPT>

<BODY>
<FORM NAME="f1">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_158 %>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="left">
	&nbsp;<input type="checkbox" name="checkrej" <%
if ei.UseReject = true then
	Response.Write "checked"
end if
%>>
	<%=b_lang_159 %>
</td></tr>
<tr><td class="block_top_td" style="height:4px; _height:6px;"></td></tr>

<tr><td align="left" style="border-top:1px solid #A5B6C8; padding-left:8px;">
	<table border="0" width="80%" cellspacing="0" bgcolor="white">
	<tr valign=bottom>
	<td nowrap align="center" height="26" width="30%"><%=b_lang_160 %><%=s_lang_mh %></td>
	<td nowrap width="20%"></td>
	<td nowrap align="center" width="50%"><%=b_lang_161 %><%=s_lang_mh %></td>
	</tr>
	<tr valign=top> 
	<td>
	<input maxlength="64" size="30" name="addmsg" class='n_textbox' onkeydown="goent()">
	</td>
	<td align=middle> 
		<table cellspacing=0 cellpadding=0>
		<tr>
		<td>
		<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="add()" type=button value="<%=s_lang_add %> >>">
		</td>
		</tr>
		<tr>
		<td><br>
		<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="delout()" type=button value="<< <%=s_lang_del %>">
		</td>
		</tr>
		<tr> 
		<td><br><br>
		<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="mup()" type=button value="<%=b_lang_013 %>">
		</td>
		</tr>
		<tr> 
		<td><br>
		<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="mdown()" type=button value="<%=b_lang_014 %>">
		</td>
		</tr>
		</table>
	</td>
	<td>
	<select class="drpdwn" style="WIDTH: 320px;" multiple size=10 id="listall" name="listall">
<%
i = 0

do while i < allnum
	tmsg = ei.GetRejectEMail(i)
	Response.Write "<option value=""" & server.htmlencode(tmsg) & """>" & server.htmlencode(tmsg) & "</option>" & Chr(13)

	tmsg = NULL

	i = i + 1
loop
%>
	</select>
	</td></tr>
</table>
</td></tr>

<tr><td class="block_top_td" style="height:8px;"></td></tr>

<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:sub();"><%=s_lang_save %></a>
</td></tr>
</table>

<table width="92%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px #A5B6C8 solid; margin-top:50px; padding-bottom:10px;'>
	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:4px; color:#444444;">
	<%=b_lang_162 %><br><%=s_lang_tpf %><br>
	</td>
	</tr>
</table>
<input type="hidden" name="allmsgs">
</FORM>

<div style="position:absolute; left:12px; top:10px;">
<a href="help.asp#showuserkill" target="_blank"><img src="images/help.gif" border="0" title="<%=s_lang_help %>"></a></div>
</BODY>
</HTML>

<%
set ei = nothing
%>
