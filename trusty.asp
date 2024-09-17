<!--#include file="passinc.asp" --> 
<!--#include file="language-2.asp" -->

<%
dim ei
set ei = server.createobject("easymail.Trusty")
ei.Load Session("wem")

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ei.RemoveAll

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
				ei.Add item
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if


	ei.Save
	set ei = nothing

	response.redirect "ok.asp?" & getGRSN() & "&gourl=trusty.asp"
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

function goent() {
	if (ie != false && event.keyCode == 13)
	{
		event.keyCode = 9;
		add();
	}
}
//-->
</script>

<BODY>
<FORM NAME="f1">
<input type="hidden" name="allmsgs">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_231 %>
</td></tr>
<tr><td class="block_top_td" style="height:4px; _height:6px;"></td></tr>

<tr><td align="left" style="padding-left:8px;">
	<table border="0" width="86%" cellspacing="0" bgcolor="white">
	<tr valign=bottom>
	<td nowrap align="center" height="26" width="30%"><%=b_lang_232 %><%=s_lang_mh %></td>
	<td nowrap width="20%"></td>
	<td nowrap align="center" width="50%"><%=b_lang_231 %><%=s_lang_mh %></td>
	</tr>
	<tr valign=top> 
	<td>
	<input maxlength=120 size=30 name="addmsg" class='n_textbox' onkeydown="goent()">
	</td>
	<td align=middle> 
		<table cellspacing=0 cellpadding=0>
		<tr> 
		<td>
		<input class="sbttn" style="WIDTH: 90px" LANGUAGE=javascript onclick="add()" type=button value="<%=s_lang_add %> >>">
		</td>
		</tr>
		<tr> 
		<td><br>
		<input class="sbttn" style="WIDTH: 90px" LANGUAGE=javascript onclick="delout()" type=button value="<< <%=s_lang_del %>">
		</td>
		</tr>
		</table>
	</td>
	<td>
	<select class="drpdwn" style="WIDTH:320px;" multiple size=10 id="listall" name="listall">
<%
i = 0
allnum = ei.Count

do while i < allnum
	tmsg = ei.Get(i)
	Response.Write "<option value=""" & server.htmlencode(tmsg) & """>" & server.htmlencode(tmsg) & "</option>" & Chr(13)

	tmsg = NULL
	i = i + 1
loop
%>
	</select>
	</td></tr>
</table>
<tr><td class="block_top_td" style="height:10px;"></td></tr>

<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:sub();"><%=s_lang_save %></a>
</td></tr>
</table>

<table width="92%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px #A5B6C8 solid; margin-top:50px; padding-bottom:10px;'>
	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:4px; color:#444444;"><%=b_lang_233 %><br>
	<%=s_lang_tpf %><br></td>
	</tr>
</table>
</FORM>
</BODY>
</HTML>

<%
set ei = nothing
%>
