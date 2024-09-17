<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" --> 

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

dim shm
set shm = server.createobject("easymail.SH_Manager")
shm.Load

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	shm.AdminRemoveAll
	shm.SH_User_RemoveAll

	if trim(request("EnableSH")) <> "" then
		shm.IsEnabled = true
	else
		shm.IsEnabled = false
	end if

	if IsNumeric(trim(request("MaxHours"))) = true then
		shm.MaxHours = CLng(trim(request("MaxHours")))
	end if

	if IsNumeric(trim(request("MaxMails"))) = true then
		shm.MaxMails = CLng(trim(request("MaxMails")))
	end if

	if trim(request("AutoPass")) = "1" then
		shm.AutoPass = true
	else
		shm.AutoPass = false
	end if

	if IsNumeric(trim(request("SH_User_Mode"))) = true then
		shm.SH_User_Mode = CLng(trim(request("SH_User_Mode")))
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
				shm.AddAdmin item
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
				shm.Add_SH_User item
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	shm.Save
	set shm = nothing

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=setsh.asp"
end if
%>


<HTML>
<HEAD>
<%=s_lang_meta %>
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
<script type="text/javascript" src="images/jquery.min.js"></script>
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

function sh_onchange() {
	if (document.getElementById("SH_User_Mode").selectedIndex == 0)
	{
		$("#ulist_1").html("<%=s_lang_0603 %>:");
		$("#ulist_2").html("<%=s_lang_0604 %>:");
	}
	else
	{
		$("#ulist_1").html("<%=s_lang_0601 %>:");
		$("#ulist_2").html("<%=s_lang_0602 %>:");
	}
}
//-->
</script>

<BODY>
<br>
<FORM ACTION="#" METHOD="POST" NAME="f1">
<input type="hidden" name="allmsgsemail">
<input type="hidden" name="allmsgshost">
<div align="center"><br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>" style="border:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="15%" height="28">&nbsp;</td>
      <td width="60%"><a href="right.asp?<%=getGRSN() %>"><%=s_lang_return %></a></td>
      <td><font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0588 %></b></font></td>
	</tr>
	<tr bgcolor="<%=MY_COLOR_3 %>">
	<td height="26" colspan="4" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	<input type="checkbox" name="EnableSH" id="EnableSH" <% if shm.isEnabled = true then response.write "checked"%>>
	<%=s_lang_0598 %>
	</td>
	</tr>
  </table>
</div>
  <div align="center">
  <table align="center" border="0" width="90%" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid;">
  <tr>
	<td height="30" colspan="3" style="padding-top:12px; padding-left:8px; padding-bottom:12px;">
<%=s_lang_0607 %>:&nbsp;<input type="text" name="MaxHours" id="MaxHours" class='textbox' value="<%=shm.MaxHours %>" size="4" maxlength="3">&nbsp;<%=s_lang_0608 %><br>
<%=s_lang_0609 %>:&nbsp;<input type="text" name="MaxMails" id="MaxMails" class='textbox' value="<%=shm.MaxMails %>" size="5" maxlength="4">&nbsp;<%=s_lang_0610 %><br>
<%=s_lang_0611 %>:
<select name="AutoPass" id="AutoPass" class="drpdwn">
	<option value="0"<%
if shm.AutoPass = false then
	Response.Write " selected"
end if
%>><%=s_lang_0612 %></option>
	<option value="1"<%
if shm.AutoPass = true then
	Response.Write " selected"
end if
%>><%=s_lang_0613 %></option>
</select>
	</td>
  </tr>
  </table>

  <table align="center" border="0" width="90%" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid;">
  <tr valign=bottom>
	<td height="30" style="padding-left:8px;"><%=s_lang_0599 %>:</td>
	<td></td>
	<td style="padding-left:8px;"><%=s_lang_0600 %>:</td>
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
allnum = shm.AdminCount

do while i < allnum
	Response.Write "<option value=""" & server.htmlencode(shm.GetAdmin(i)) & """>" & server.htmlencode(shm.GetAdmin(i)) & "</option>" & Chr(13)
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
  <tr>
	<td height="30" colspan="3" style="padding-top:14px; padding-left:8px;">
<select name="SH_User_Mode" id="SH_User_Mode" class="drpdwn" onchange="javascript:sh_onchange()">
	<option value="1"<%
if shm.SH_User_Mode = 1 then
	Response.Write " selected"
end if
%>><%=s_lang_0605 %></option>
	<option value="2"<%
if shm.SH_User_Mode = 2 then
	Response.Write " selected"
end if
%>><%=s_lang_0606 %></option>
</select>
	</td>
  </tr>
  <tr valign=bottom>
	<td height="30" id="ulist_1" style="padding-left:8px;"><%
if shm.SH_User_Mode = 1 then
	Response.Write s_lang_0603
else
	Response.Write s_lang_0601
end if
%>:</td>
	<td></td>
	<td id="ulist_2" style="padding-left:8px;"><%
if shm.SH_User_Mode = 1 then
	Response.Write s_lang_0604
else
	Response.Write s_lang_0602
end if
%>:</td>
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
allnum = shm.SH_User_Count

do while i < allnum
	Response.Write "<option value=""" & server.htmlencode(shm.Get_SH_User(i)) & """>" & server.htmlencode(shm.Get_SH_User(i)) & "</option>" & Chr(13)
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
	<td height="50" colspan="3" align="right" bgcolor="white" style="border-top:1px <%=MY_COLOR_1 %> solid; padding-right:20px;">
	<input type="button" value=" <%=s_lang_save %> " LANGUAGE=javascript onclick="sub()" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" <%=s_lang_return %> " LANGUAGE=javascript onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td></tr>
  </table>
</FORM>
<br>
</BODY>
</HTML>

<%
set shm = nothing
%>
