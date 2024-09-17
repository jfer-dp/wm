<!--#include file="passinc.asp" --> 
<!--#include file="language-2.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.WhiteListIP")
ei.Load

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

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("wl_ip.asp")
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
-->
</STYLE>
</HEAD>

<script type="text/javascript" src="images/mglobal.js"></script>

<script LANGUAGE=javascript>
<!--
function sub()
{
	var tempstr = "";
	var i = 0;
	for (i; i < document.getElementById("f1").listall.length; i++)
	{
		tempstr = tempstr + document.getElementById("f1").listall[i].value + "\t";
	}

	document.getElementById("f1").allmsgs.value = tempstr;
	document.getElementById("f1").submit();
}

function delout()
{
	var i = 0;
	for (i; i < document.getElementById("f1").listall.length; i++)
	{
		if (document.getElementById("f1").listall[i].selected == true)
		{
			document.getElementById("f1").listall.remove(i);
			i--;
		}
	}
}

function add()
{
	if (document.getElementById("f1").addmsg.value.indexOf("\t") != -1)
	{
		alert("<%=s_lang_inputerr %>");
		document.getElementById("f1").addmsg.focus();
		return ;
	}

	if (document.getElementById("f1").addmsg.value.length > 0)
	{
		if (haveit() == false)
		{
			var oOption = document.createElement("OPTION");
			oOption.text = document.getElementById("f1").addmsg.value;
			oOption.value = document.getElementById("f1").addmsg.value;

			if (ie == false)
				document.getElementById("f1").listall.appendChild(oOption);
			else
				document.getElementById("f1").listall.add(oOption);

			document.getElementById("f1").addmsg.value = "";
			document.getElementById("f1").addmsg.focus();
			return ;
		}
		else
			return ;
	}

	alert("<%=s_lang_inputerr %>");
	document.getElementById("f1").addmsg.focus();
}

function haveit()
{
	var tempstr = document.getElementById("f1").addmsg.value;

	var i = 0;
	for (i; i < document.getElementById("f1").listall.length; i++)
	{
		if (document.getElementById("f1").listall[i].value == tempstr)
			return true;
	}

	return false;
}

function goent() {
	if (ie == false)
		return ;

	if (event.keyCode == 13)
	{
		event.keyCode = 9;
		add();
	}
}
//-->
</script>

<BODY>
<FORM ACTION="#" METHOD="POST" ID="f1" NAME="f1">
<input type="hidden" name="allmsgs">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_349 %>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="center">

<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white">
  <tr valign=bottom align="center">
	<td height="10" style="color:#444444;">&nbsp;<%=b_lang_350 & s_lang_mh %></td>
	<td></td>
	<td style="color:#444444;">&nbsp;<%=b_lang_349 & s_lang_mh %></td>
  </tr>
  <tr valign=top align="center"> 
	<td>
	&nbsp;<input maxlength="64" size="28" name="addmsg" class='n_textbox' onkeydown="goent()">
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
	<tr><td></td></tr>
	<tr><td></td></tr>
	</table>
	</td>
	<td>
	<select class="drpdwn" style="WIDTH: 305px" multiple size=20 name="listall" width="305">
<%
i = 0
allnum = ei.Count

do while i < allnum
	Response.Write "<option value=""" & server.htmlencode(ei.Get(i)) & """>" & server.htmlencode(ei.Get(i)) & "</option>" & Chr(13)
	i = i + 1
loop
%>
	</select>
	</td>
  </tr>
  </table>
<tr><td class="block_top_td" style="height:14px;"></td></tr>

<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:2px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="impset.asp?<%=getGRSN() %>#wlip"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:sub();"><%=s_lang_save %></a>
</td></tr>
</table>
</FORM>
</BODY>
</HTML>

<%
set ei = nothing
%>
