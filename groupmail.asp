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
set ei = server.createobject("easymail.GroupMail")
'-----------------------------------------
ei.Load

dim sif
set sif = server.createobject("easymail.sysinfo")
sif.Load

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	needsave = false
	if trim(request("EnableSendWithGroupMail")) <> "" then
		if sif.EnableSendWithGroupMail = false then
			sif.EnableSendWithGroupMail = true
			needsave = true
		end if
	else
		if sif.EnableSendWithGroupMail = true then
			sif.EnableSendWithGroupMail = false
			needsave = true
		end if
	end if

	if trim(request("EnableAccreditForGroupMail")) <> "" then
		if sif.EnableAccreditForGroupMail = false then
			sif.EnableAccreditForGroupMail = true
			needsave = true
		end if
	else
		if sif.EnableAccreditForGroupMail = true then
			sif.EnableAccreditForGroupMail = false
			needsave = true
		end if
	end if

	if needsave = true then
		sif.Save
	end if

	set sif = nothing


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

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=groupmail.asp"
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
	for (i; i < document.f1.listall.length; i++)
	{
		tempstr = tempstr + document.f1.listall[i].value + "\t";
	}

	document.f1.allmsgs.value = tempstr;
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
<%
if isMSIE = true then
%>
			document.f1.listall.add(oOption);
<%
else
%>
			document.f1.listall.appendChild(oOption);
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
<%
if isMSIE = true then
%>
	if (event.keyCode == 13)
	{
		event.keyCode = 9;
		add();
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
<input type="hidden" name="allmsgs">
<div align="center"><br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="15%" height="28">&nbsp;</td>
      <td width="32%"><a href="showsysinfo.asp?<%=getGRSN() %>#groupmail"><%=s_lang_enable %></a></td>
      <td width="23%"><a href="right.asp?<%=getGRSN() %>"><%=s_lang_return %></a></td>
      <td width="30%"><font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0078 %></b></font></td>
	</tr>
	<tr bgcolor="<%=MY_COLOR_3 %>">
	<td height="26" colspan="4" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	<input type="checkbox" name="EnableSendWithGroupMail" id="EnableSendWithGroupMail" <% if sif.EnableSendWithGroupMail = true then response.write "checked"%>>
	<%=s_lang_0080 %>
	</td>
	</tr>
	<tr bgcolor="<%=MY_COLOR_3 %>">
	<td height="26" colspan="4" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	<input type="checkbox" name="EnableAccreditForGroupMail" id="EnableAccreditForGroupMail" <% if sif.EnableAccreditForGroupMail = true then response.write "checked"%>>
	<%=s_lang_0081 %>
	</td>
	</tr>
  </table>
</div>
  <div align="center">
  <table align="center" border="0" width="90%" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
  <tr valign=bottom>
	<td height="30">&nbsp;<%=s_lang_0082 %>:</td>
	<td></td>
	<td>&nbsp;<%=s_lang_0083 %>:</td>
  </tr>
  <tr valign=top> 
	<td>
	&nbsp;<input maxlength="64" size="23" name="addmsg" class='textbox' onkeydown="goent()">
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
	<select class="drpdwn" style="WIDTH: 305px" multiple size=10 name=listall width="305">
<%
i = 0
allnum = ei.Count

do while i < allnum
	Response.Write "<option value=""" & server.htmlencode(ei.Get(i)) & """>" & server.htmlencode(ei.Get(i)) & "</option>" & Chr(13)
	i = i + 1
loop
%>
	</select>
	</font> </td>
  </tr>
	<tr>
	<td height="20" colspan="3" align="right"><br><hr size="1" color="<%=MY_COLOR_1 %>">
	<input type="button" value=" <%=s_lang_save %> " LANGUAGE=javascript onclick="sub()" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" <%=s_lang_exit %> " LANGUAGE=javascript onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td></tr>
  </table>
<br><br><br>
  <div align="center">
    <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr> 
        <td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
        <td width="94%"><%=s_lang_0084 %>.<br>
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
set sif = nothing
set ei = nothing
%>
