<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

dim ei
set ei = server.createobject("easymail.CheckMyIPinRBL")
ei.Load

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ei.RemoveAll

	if trim(request("EnableCheckRBL")) <> "" then
		ei.IsEnabled = true
	else
		ei.IsEnabled = false
	end if

	dim msg
	dim item
	dim ss
	dim se

	msg = trim(request("allmsgsemail"))

	if Len(msg) > 0 then
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(9))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				ei.Add_Email item
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	msg = trim(request("allmsgsip"))

	if Len(msg) > 0 then
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(9))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				ei.Add_IP item
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	msg = trim(request("allmsgshost"))

	if Len(msg) > 0 then
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(9))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				ei.Add_Host item
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	ei.Save
	set ei = nothing

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=checkrbl.asp"
end if
%>


<HTML>
<HEAD>
<%=s_lang_meta %>
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<script LANGUAGE=javascript>
<!--
function sub()
{
	var tempstr = "";
	var i = 0;
	for (i; i < document.f1.listallemail.length; i++)
	{
		tempstr = tempstr + document.f1.listallemail[i].value + "\t";
	}
	document.f1.allmsgsemail.value = tempstr;

	tempstr = "";
	i = 0;
	for (i; i < document.f1.listallip.length; i++)
	{
		tempstr = tempstr + document.f1.listallip[i].value + "\t";
	}
	document.f1.allmsgsip.value = tempstr;

	tempstr = "";
	i = 0;
	for (i; i < document.f1.listallhost.length; i++)
	{
		tempstr = tempstr + document.f1.listallhost[i].value + "\t";
	}
	document.f1.allmsgshost.value = tempstr;

	document.f1.submit();
}

function deloutemail()
{
	var i = 0;
	for (i; i < document.f1.listallemail.length; i++)
	{
		if (document.f1.listallemail[i].selected == true)
		{
			document.f1.listallemail.remove(i);
			i--;
		}
	}
}

function addemail()
{
	if (document.f1.addmsgemail.value.indexOf("\t") != -1)
	{
		alert("<%=s_lang_inputerr %>");
		document.f1.addmsgemail.focus();
		return ;
	}

	if (document.f1.addmsgemail.value.length > 0)
	{
		if (haveitemail() == false)
		{
			var oOption = document.createElement("OPTION");
			oOption.text = document.f1.addmsgemail.value;
			oOption.value = document.f1.addmsgemail.value;
<%
if isMSIE = true then
%>
			document.f1.listallemail.add(oOption);
<%
else
%>
			document.f1.listallemail.appendChild(oOption);
<%
end if
%>
			return ;
		}
		else
			return ;
	}

	alert("<%=s_lang_inputerr %>");
}

function haveitemail()
{
	var tempstr = document.f1.addmsgemail.value;

	var i = 0;
	for (i; i < document.f1.listallemail.length; i++)
	{
		if (document.f1.listallemail[i].value == tempstr)
			return true;
	}

	return false;
}

function goentemail() {
<%
if isMSIE = true then
%>
	if (event.keyCode == 13)
	{
		event.keyCode = 9;
		addemail();
	}
<%
end if
%>
}

function deloutip()
{
	var i = 0;
	for (i; i < document.f1.listallip.length; i++)
	{
		if (document.f1.listallip[i].selected == true)
		{
			document.f1.listallip.remove(i);
			i--;
		}
	}
}

function addip()
{
	if (document.f1.addmsgip.value.indexOf("\t") != -1)
	{
		alert("<%=s_lang_inputerr %>");
		document.f1.addmsgip.focus();
		return ;
	}

	if (document.f1.addmsgip.value.length > 0)
	{
		if (haveitip() == false)
		{
			var oOption = document.createElement("OPTION");
			oOption.text = document.f1.addmsgip.value;
			oOption.value = document.f1.addmsgip.value;
<%
if isMSIE = true then
%>
			document.f1.listallip.add(oOption);
<%
else
%>
			document.f1.listallip.appendChild(oOption);
<%
end if
%>
			return ;
		}
		else
			return ;
	}

	alert("<%=s_lang_inputerr %>");
}

function haveitip()
{
	var tempstr = document.f1.addmsgip.value;

	var i = 0;
	for (i; i < document.f1.listallip.length; i++)
	{
		if (document.f1.listallip[i].value == tempstr)
			return true;
	}

	return false;
}

function goentip() {
<%
if isMSIE = true then
%>
	if (event.keyCode == 13)
	{
		event.keyCode = 9;
		addip();
	}
<%
end if
%>
}

function delouthost()
{
	var i = 0;
	for (i; i < document.f1.listallhost.length; i++)
	{
		if (document.f1.listallhost[i].selected == true)
		{
			document.f1.listallhost.remove(i);
			i--;
		}
	}
}

function addhost()
{
	if (document.f1.addmsghost.value.indexOf("\t") != -1)
	{
		alert("<%=s_lang_inputerr %>");
		document.f1.addmsghost.focus();
		return ;
	}

	if (document.f1.addmsghost.value.length > 0)
	{
		if (haveithost() == false)
		{
			var oOption = document.createElement("OPTION");
			oOption.text = document.f1.addmsghost.value;
			oOption.value = document.f1.addmsghost.value;
<%
if isMSIE = true then
%>
			document.f1.listallhost.add(oOption);
<%
else
%>
			document.f1.listallhost.appendChild(oOption);
<%
end if
%>
			return ;
		}
		else
			return ;
	}

	alert("<%=s_lang_inputerr %>");
}

function haveithost()
{
	var tempstr = document.f1.addmsghost.value;

	var i = 0;
	for (i; i < document.f1.listallhost.length; i++)
	{
		if (document.f1.listallhost[i].value == tempstr)
			return true;
	}

	return false;
}

function goenthost() {
<%
if isMSIE = true then
%>
	if (event.keyCode == 13)
	{
		event.keyCode = 9;
		addhost();
	}
<%
end if
%>
}
//-->
</script>

<BODY>
<br>
<FORM ACTION="#" METHOD="POST" NAME="f1">
<input type="hidden" name="allmsgsemail">
<input type="hidden" name="allmsgsip">
<input type="hidden" name="allmsgshost">
<div align="center"><br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>" style="border:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="15%" height="28">&nbsp;</td>
      <td width="60%"><a href="right.asp?<%=getGRSN() %>"><%=s_lang_return %></a></td>
      <td><font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0560 %></b></font></td>
	</tr>
	<tr bgcolor="<%=MY_COLOR_3 %>">
	<td height="26" colspan="4" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	<input type="checkbox" name="EnableCheckRBL" id="EnableCheckRBL" <% if ei.IsEnabled = true then response.write "checked"%>>
	<%=s_lang_0562 %>
	</td>
	</tr>
  </table>
</div>
  <div align="center">
  <table align="center" border="0" width="90%" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid;">
  <tr valign=bottom>
	<td height="30">&nbsp;<%=s_lang_0563 %>:</td>
	<td></td>
	<td>&nbsp;<%=s_lang_0564 %>:</td>
  </tr>
  <tr valign=top> 
	<td>
	&nbsp;<input maxlength="64" size="23" name="addmsgemail" class='textbox' onkeydown="goentemail()">
	</td>
    <td align=middle> 
	<table cellspacing=0 cellpadding=0>
	<tr>
	<td>
	<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="addemail()" type=button value="<%=s_lang_add %> >>">
	</td>
	</tr>
	<tr> 
	<td><br>
	<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="deloutemail()" type=button value="<< <%=s_lang_del %>">
	</td>
	</tr>
	</table>
	</td>
	<td>
	<select class="drpdwn" style="WIDTH: 305px" multiple size=4 name="listallemail" width="305">
<%
i = 0
allnum = ei.Count_Email

do while i < allnum
	Response.Write "<option value=""" & server.htmlencode(ei.Get_Email(i)) & """>" & server.htmlencode(ei.Get_Email(i)) & "</option>" & Chr(13)
	i = i + 1
loop
%>
	</select>
	</td>
  </tr>
	<tr>
	<td height="15" colspan="3">
	</td></tr>
  </table>


  <table align="center" border="0" width="90%" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid;">
  <tr valign=bottom>
	<td height="30">&nbsp;<%=s_lang_0565 %>:</td>
	<td></td>
	<td>&nbsp;<%=s_lang_0566 %>:</td>
  </tr>
  <tr valign=top> 
	<td>
	&nbsp;<input maxlength="64" size="23" name="addmsgip" class='textbox' onkeydown="goentip()">
	</td>
    <td align=middle> 
	<table cellspacing=0 cellpadding=0>
	<tr>
	<td>
	<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="addip()" type=button value="<%=s_lang_add %> >>">
	</td>
	</tr>
	<tr> 
	<td><br>
	<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="deloutip()" type=button value="<< <%=s_lang_del %>">
	</td>
	</tr>
	</table>
	</td>
	<td>
	<select class="drpdwn" style="WIDTH: 305px" multiple size=6 name="listallip" width="305">
<%
i = 0
allnum = ei.Count_IP

do while i < allnum
	Response.Write "<option value=""" & server.htmlencode(ei.Get_IP(i)) & """>" & server.htmlencode(ei.Get_IP(i)) & "</option>" & Chr(13)
	i = i + 1
loop
%>
	</select>
	</td>
  </tr>
	<tr>
	<td height="15" colspan="3">
	</td></tr>
  </table>

  <table align="center" border="0" width="90%" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
  <tr valign=bottom>
	<td height="30">&nbsp;<%=s_lang_0567 %>:</td>
	<td></td>
	<td>&nbsp;<%=s_lang_0568 %>:</td>
  </tr>
  <tr valign=top> 
	<td>
	&nbsp;<input maxlength="64" size="23" name="addmsghost" class='textbox' onkeydown="goenthost()">
	</td>
    <td align=middle> 
	<table cellspacing=0 cellpadding=0>
	<tr>
	<td>
	<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="addhost()" type=button value="<%=s_lang_add %> >>">
	</td>
	</tr>
	<tr> 
	<td><br>
	<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="delouthost()" type=button value="<< <%=s_lang_del %>">
	</td>
	</tr>
	</table>
	</td>
	<td>
	<select class="drpdwn" style="WIDTH: 305px" multiple size=8 name="listallhost" width="305">
<%
i = 0
allnum = ei.Count_Host

do while i < allnum
	Response.Write "<option value=""" & server.htmlencode(ei.Get_Host(i)) & """>" & server.htmlencode(ei.Get_Host(i)) & "</option>" & Chr(13)
	i = i + 1
loop
%>
	</select>
	</td>
  </tr>
	<tr>
	<td height="15" colspan="3">
	</td></tr>
	<tr>
	<td height="50" colspan="3" align="right" bgcolor="white" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	<input type="button" value=" <%=s_lang_save %> " LANGUAGE=javascript onclick="sub()" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" <%=s_lang_return %> " LANGUAGE=javascript onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td></tr>
  </table>

<br><br>
  <div align="center">
    <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr> 
        <td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
        <td width="94%"><%=s_lang_0569 %>
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
</FORM>
</div>
<br>
</BODY>
</HTML>

<%
set ei = nothing
%>
