<!--#include file="passinc.asp" --> 
<!--#include file="language-1.asp" -->

<%
dim ei
set ei = server.createobject("easymail.CheckGoodMail")
ei.Load Session("wem")

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if IsNumeric(trim(request("minMatchNumber"))) = true then
		ei.minMatchNumber = CLng(trim(request("minMatchNumber")))
	end if

	if trim(request("isEnableCheckGoodMail")) <> "" then
		ei.isEnableCheckGoodMail = true
	else
		ei.isEnableCheckGoodMail = false
	end if

	ei.RemoveAll
	ei.Add trim(request("allmsgs"))

	ei.Save
	set ei = nothing

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=cgfilter.asp"
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
.cont_td {height:28px; text-align:left; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px;}
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
		tempstr = tempstr + document.f1.listall[i].value + "\f";
	}

	document.f1.allmsgs.value = tempstr;
	document.f1.submit();
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

function delout()
{
	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].selected == true)
		{
			document.f1.listall.remove(i);

			if (i < document.f1.listall.length)
				document.f1.listall[i].selected = true;
			else
			{
				if (i - 1 >= 0)
					document.f1.listall[i - 1].selected = true;
			}

			break;
		}
	}
}

function add()
{
	if (haveit() == false)
	{
		if (document.f1.s_msg.value.length < 1)
		{
			alert("<%=s_lang_inputerr %>");
			return ;
		}

		var oOption = document.createElement("OPTION");
		oOption.text = "<%=a_lang_136 %>: " + document.f1.s_msg.value;
		oOption.value = document.f1.s_msg.value;

		if (ie == false)
			document.getElementById("listall").appendChild(oOption);
		else
			document.getElementById("listall").add(oOption);

		document.f1.s_msg.value = "";
	}
}

function haveit()
{
	var tempstr = document.f1.s_msg.value;

	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].value == tempstr)
			return true;
	}

	return false;
}
//-->
</script>

<BODY>
<FORM ACTION="#" METHOD="POST" NAME="f1">
<input type="hidden" name="allmsgs">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_137 %>
</td></tr>
<tr><td class="block_top_td" style="height:6px; _height:8px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="cont_td">
<input type="checkbox" name="isEnableCheckGoodMail" value="checkbox"<% if ei.isEnableCheckGoodMail = true then Response.Write " checked"%>><%=a_lang_138 %>
</td></tr>

<tr><td class="cont_td">
&nbsp;<%=a_lang_139 %> <input type="input" name="minMatchNumber" class="n_textbox" size="4" maxlength="2" value="<%=ei.minMatchNumber %>"> <%=a_lang_140 %>
</td></tr>

<tr><td align="left" style="padding-top:8px; padding-bottom:8px; padding-left:8px;">
	<table border="0" width="100%" cellspacing="0">
	<tr> 
	<td align="left" width="30%">
	<select id="listall" name="listall" size="10" class="drpdwn" style="width:480px;">
<%
i = 0
allnum = ei.Count

do while i < allnum
	as_msg = ei.Get(i)

	Response.Write "<option value=""" & server.htmlencode(as_msg) & """>" & server.htmlencode(a_lang_136 & ": " & as_msg) & "</option>" & Chr(13)
	as_msg = NULL

	i = i + 1
loop
%>
	</select>
	</td>
	<td align="left" style="padding-left:12px;">
	<input type="button" value="<%=a_lang_141 %>" class="sbttn" style="width:50px;" LANGUAGE=javascript onclick="mup()">
	<br><br><br>
	<input type="button" value="<%=a_lang_142 %>" class="sbttn" style="width:50px;" LANGUAGE=javascript onclick="delout()">
	<br><br><br>
	<input type="button" value="<%=a_lang_143 %>" class="sbttn" style="width:50px;" LANGUAGE=javascript onclick="mdown()">
	</td>
	</tr>
	</table>
</td></tr>

	<tr>
	<td nowrap style="height:28px; text-align:left; border-top:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px;">
	<table border="0" width="100%" cellspacing="0">
	<tr><td nowrap width="20%">
	&nbsp;<%=a_lang_136 %><%=s_lang_mh %><input type="input" name="s_msg" class="n_textbox" size="30" maxlength="126">
	</td>
	<td nowrap style="padding-left:4px;">
<a class='wwm_btnDownload btn_gray' href="javascript:add();"><%=s_lang_add %></a>
	</td></tr>
	</table>
	</td></tr>

	<tr><td align="left" style="background-color:white; padding-right:16px; padding-top:16px;">
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:sub();"><%=s_lang_save %></a>
	</td></tr>
</table>

</td></tr>
</table>

<table width="92%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px #A5B6C8 solid; margin-top:50px; padding-bottom:10px;'>
	<tr>
	<td valign="top" style="padding:4px; padding-left:8px; width:22px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:4px; color:#444444;">
	<%=a_lang_144 %><br>
	<font color="#901111"><%=a_lang_145 %></font>: <%=a_lang_146 %><a href="userspamguard.asp?<%=getGRSN() %>"><%=a_lang_147 %>
	<br><br><%=s_lang_tpf %><br>
	</td>
	</tr>
</table>
</FORM>
</BODY>
</HTML>

<%
set ei = nothing
%>
