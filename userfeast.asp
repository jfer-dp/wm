<!--#include file="passinc.asp" --> 
<!--#include file="language-2.asp" -->

<%
dim ecalset
set ecalset = server.createobject("easymail.CalOptions")
ecalset.Load Session("wem")

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ecalset.RemoveAll
	ecalset.RemoveAllNL

	dim msg
	dim item
	dim ss
	dim se

	msg = trim(request("allmsgs"))
	if Len(msg) > 0 then
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(9))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				ecalset.Add item
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	msg = trim(request("allmsgsNL"))
	if Len(msg) > 0 then
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(9))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				ecalset.AddNL item
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	ecalset.Save

	set ecalset = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=userfeast.asp"
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

<script type="text/javascript">
<!--
function gosub()
{
	var tempstr = "";
	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		tempstr = tempstr + document.f1.listall[i].value + "\t";
	}
	document.f1.allmsgs.value = tempstr;

	tempstr = "";
	i = 0;
	for (i; i < document.f1.listallNL.length; i++)
	{
		tempstr = tempstr + document.f1.listallNL[i].value + "\t";
	}
	document.f1.allmsgsNL.value = tempstr;

	document.f1.action = "userfeast.asp";
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
			oOption.text = document.f1.gl_month.value + document.f1.gl_day.value + " " + document.f1.addmsg.value;
			oOption.value = document.f1.gl_month.value + document.f1.gl_day.value + " " + document.f1.addmsg.value;

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
	document.f1.addmsg.focus();
}

function haveit()
{
	var tempstr = document.f1.gl_month.value + document.f1.gl_day.value + " "

	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].value.substr(0, 5) == tempstr)
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

function deloutNL()
{
	var i = 0;
	for (i; i < document.f1.listallNL.length; i++)
	{
		if (document.f1.listallNL[i].selected == true)
		{
			document.f1.listallNL.remove(i);
			i--;
		}
	}
}

function addNL()
{
	if (document.f1.addmsgNL.value.indexOf("\t") != -1)
	{
		alert("<%=s_lang_inputerr %>");
		document.f1.addmsgNL.focus();
		return ;
	}

	if (document.f1.addmsgNL.value.length > 0)
	{
		if (haveitNL() == false)
		{
			var oOption = document.createElement("OPTION");
			oOption.text = document.f1.nl_month.value + document.f1.nl_day.value + " " + document.f1.addmsgNL.value;
			oOption.value = document.f1.nl_month.value + document.f1.nl_day.value + " " + document.f1.addmsgNL.value;

			if (ie == false)
				document.getElementById("listallNL").appendChild(oOption);
			else
				document.getElementById("listallNL").add(oOption);

			return ;
		}
		else
			return ;
	}

	alert("<%=s_lang_inputerr %>");
	document.f1.addmsgNL.focus();
}

function haveitNL()
{
	var tempstr = document.f1.nl_month.value + document.f1.nl_day.value + " "

	var i = 0;
	for (i; i < document.f1.listallNL.length; i++)
	{
		if (document.f1.listallNL[i].value.substr(0, 5) == tempstr)
			return true;
	}

	return false;
}

function goentNL() {
	if (ie != false && event.keyCode == 13)
	{
		event.keyCode = 9;
		addNL();
	}
}
//-->
</SCRIPT>

<BODY>
<FORM NAME="f1">
<input type="hidden" name="allmsgs">
<input type="hidden" name="allmsgsNL">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_236 %>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="left">

<table align="left" width="100%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr valign=top> 
	<td nowrap align="left" width="30%">
	&nbsp;<select name="gl_month" class="drpdwn">
<%
i = 1

do while i < 13
	if i < 10 then
		Response.Write "<option value=""0" & i & """>" & i & b_lang_029 & "</option>" & Chr(13)
	else
		Response.Write "<option value=""" & i & """>" & i & b_lang_029 & "</option>" & Chr(13)
	end if

	i = i + 1
loop
%>
	</select><select name="gl_day" class="drpdwn">
<%
i = 1

do while i < 32
	if i < 10 then
		Response.Write "<option value=""0" & i & """>" & i & b_lang_030 & "</option>" & Chr(13)
	else
		Response.Write "<option value=""" & i & """>" & i & b_lang_030 & "</option>" & Chr(13)
	end if

	i = i + 1
loop
%>
	</select>
	&nbsp;<input maxlength=30 size=12 name="addmsg" class='n_textbox' onkeydown="goent()"><br>
	&nbsp;<font color="#444444">(<%=b_lang_237 %>)</font>
	</td>
	<td nowrap align="left" width="13%"> 
		<table align="center" width="100%" border="0" align="center" cellspacing="0" bgcolor="white">
		<tr><td>
		<input class="sbttn" style="WIDTH:70px" LANGUAGE=javascript onclick="add()" type=button value="<%=s_lang_add %> >>">
		</td></tr>
		<tr>
		<td><br>
		<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="delout()" type=button value="<< <%=s_lang_del %>">
		</td></tr>
		</table>
	</td>
	<td align="left">
	<select class="drpdwn" style="WIDTH:280px;" multiple size=7 id="listall" name="listall">
<%
i = 0
allnum = ecalset.Count

do while i < allnum
	ecalset.Get i, mm, dd, fname

	tmsg = server.htmlencode(convFeast(mm, dd, fname))
	Response.Write "<option value=""" & tmsg & """>" & tmsg & "</option>" & Chr(13)

	tmsg = NULL
	mm = NULL
	dd = NULL
	fname = NULL

	i = i + 1
loop
%>
	</select>
	</td>
	</tr>
</table>
</td></tr>
</table>

<br>
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_238 %>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="left">

<table align="left" width="100%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr valign=top> 
	<td nowrap align="left" width="30%">
	&nbsp;<select name="nl_month" class="drpdwn">
<%
i = 1

do while i < 13
	if i < 10 then
		Response.Write "<option value=""0" & i & """>" & i & b_lang_029 & "</option>" & Chr(13)
	else
		Response.Write "<option value=""" & i & """>" & i & b_lang_029 & "</option>" & Chr(13)
	end if

	i = i + 1
loop
%>
	</select><select name="nl_day" class="drpdwn">
<%
i = 1

do while i < 32
	if i < 10 then
		Response.Write "<option value=""0" & i & """>" & i & b_lang_030 & "</option>" & Chr(13)
	else
		Response.Write "<option value=""" & i & """>" & i & b_lang_030 & "</option>" & Chr(13)
	end if

	i = i + 1
loop
%>
	</select>
	&nbsp;<input maxlength=30 size=12 name="addmsgNL" class='n_textbox' onkeydown="goentNL()"><br>
	&nbsp;<font color="#444444">(<%=b_lang_239 %>)</font>
	</td>
	<td nowrap align="left" width="13%">
		<table align="center" width="100%" border="0" align="center" cellspacing="0" bgcolor="white">
		<tr> 
		<td>
		<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="addNL()" type=button value="<%=s_lang_add %> >>">
		</td>
		</tr>
		<tr> 
		<td><br>
		<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="deloutNL()" type=button value="<< <%=s_lang_del %>">
		</td>
		</tr>
		</table>
	</td>
	<td>
	<select class="drpdwn" style="WIDTH:280px;" multiple size=7 id="listallNL" name="listallNL">
<%
i = 0
allnum = ecalset.CountNL

do while i < allnum
	ecalset.GetNL i, mm, dd, fname

	tmsg = server.htmlencode(convFeast(mm, dd, fname))
	Response.Write "<option value=""" & tmsg & """>" & tmsg & "</option>" & Chr(13)

	tmsg = NULL
	mm = NULL
	dd = NULL
	fname = NULL

	i = i + 1
loop
%>
	</select>
	</td>
	</tr>
</table>
</td></tr>

<tr><td class="block_top_td" style="height:8px;"></td></tr>

<tr><td align="left" style="background-color:white; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
</td></tr>
</table>

</FORM>
</BODY>
</HTML>

<%
set ecalset = nothing


function convFeast(mm, dd, fname)
	tmpstr = ""
	if mm < 10 then
		tmpstr = "0" & mm
	else
		tmpstr = mm
	end if

	if dd < 10 then
		tmpstr = tmpstr & "0" & dd
	else
		tmpstr = tmpstr & dd
	end if

	convFeast = tmpstr & " " & fname
end function
%>
