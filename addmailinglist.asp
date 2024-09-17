<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" --> 

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

dim dm
set dm = server.createobject("easymail.Domain")
dm.Load

if dm.GetUserManagerDomainCount(Session("wem")) < 1 then
	if isadmin() = false then
		set dm = nothing
		response.redirect "noadmin.asp"
	end if
end if

dim ManagerDomainString
i = 0
if isadmin() = false then
	ManagerDomainString = Chr(9)
	allnum = dm.GetUserManagerDomainCount(Session("wem"))

	do while i < allnum
		domain = dm.GetUserManagerDomain(Session("wem"), i)
		ManagerDomainString = ManagerDomainString & LCase(domain) & Chr(9)
		domain = NULL

		i = i + 1
	loop
end if

id = trim(request("rid"))

dim ei
set ei = server.createobject("easymail.mailinglist")
'-----------------------------------------

ei.LoadOne id

dim isnew
isnew = true

if id <> "" then
	editname = id

	allnum = ei.ItemCount
	i = 0

	do while i < allnum
		tmp_name = server.htmlencode(ei.GetItemNameByIndex(i))
		inlist = inlist & "<option value='" & tmp_name & "'>" & tmp_name & "</option>" & Chr(13)

		tmp_name = NULL
		i = i + 1
	loop

	if Len(editname) > 0 then
		isnew = false
	end if
end if


dim eu
set eu = Application("em")
%>

<HTML>
<HEAD>
<%=s_lang_meta %>
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<script LANGUAGE=javascript>
<!--
function addnew()
{
	if (document.f1.selectlistusers.length < 1)
		alert("<%=s_lang_0095 %>");
	else if (document.f1.filename.value == "")
		alert("<%=s_lang_0096 %>");
	else
	{
		document.f1.mode.value = "addnew";
		var i = 0;
		var al = "";

		for (i; i < document.f1.selectlistusers.length; i++)
		{
			al = al + document.f1.selectlistusers[i].value + '\t';
		}

		document.f1.addlist.value = al;

		document.f1.addname.value = document.f1.filename.value;

		al = "";
		for (i = 0; i < document.f1.listall.length; i++)
		{
			al = al + document.f1.listall[i].value + "\t";
		}

		document.f1.accredit_list.value = al;
		document.f1.submit();
	}
}


function isinlist(name)
{
	var i = 0;
	for (i; i < document.f1.selectlistusers.length; i++)
	{
		if (document.f1.selectlistusers[i].value == name)
		{
			return true;
		}
	}

	return false;
}


function isinAllList(name)
{
	var i = 0;
	for (i; i < document.f1.selectalluser.length; i++)
	{
		if (document.f1.selectalluser[i].value == name)
			return true;
	}

	return false;
}


function addinFromText()
{
	if (document.f1.addtext.value != "" && isinlist(document.f1.addtext.value) == false)
	{
		if (isinAllList(document.f1.addtext.value) == true)
		{
			var oOption = document.createElement("OPTION");
			oOption.text = document.f1.addtext.value;
			oOption.value = document.f1.addtext.value;
<%
if isMSIE = true then
%>
			document.f1.selectlistusers.add(oOption);
<%
else
%>
			document.f1.selectlistusers.appendChild(oOption);
<%
end if
%>

			document.f1.addtext.value = "";
		}
		else
			alert("<%=s_lang_0097 %>");
	}
}


function addin()
{
	var i = 0;
	for (i; i < document.f1.selectalluser.length; i++)
	{
		if (document.f1.selectalluser[i].selected == true)
		{
			if (isinlist(document.f1.selectalluser[i].value) == false)
			{
				var oOption = document.createElement("OPTION");
				oOption.text = document.f1.selectalluser[i].value;
				oOption.value = document.f1.selectalluser[i].value;
<%
if isMSIE = true then
%>
				document.f1.selectlistusers.add(oOption);
<%
else
%>
				document.f1.selectlistusers.appendChild(oOption);
<%
end if
%>
			}
		}
	}
}


function delout()
{
	var i = 0;
	for (i; i < document.f1.selectlistusers.length; i++)
	{
		if (document.f1.selectlistusers[i].selected == true)
		{
			document.f1.selectlistusers.remove(i);
			i--;
		}
	}
}


function selectfn_onchange()
{
	document.f1.filename.value = document.f1.selectfilename.value;
}


function ac_delout()
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

function ac_add()
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
		ac_add();
	}
<%
end if
%>
}
//-->
</script>

<BODY>
<br>
<form action="savemailinglist.asp" method=post name="f1">
<div align="center"><br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
	<td width="25" height="25"></td>
      <td><a href="browmailinglist.asp?<%=getGRSN() %>"><%=s_lang_return %></a></td>
<% if isnew = true then %>
      <td width="24%"><b><%=s_lang_0085 %></b></td>
<% else %>
      <td width="40%"><b><%=s_lang_0098 %>:</b><%=editname %></td>
<% end if %>
    </tr>
  </table>
</div>
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr><td height="2"></td></tr>
	<tr>
	<td height="25">
	&nbsp;<font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0099 %></b>:</font>
	<input type="text" name="filename" maxlength="64" class="textbox" readonly value="<%=editname %>">
<% if isnew = true then %>
            <select name="selectfilename" class="drpdwn" LANGUAGE=javascript onchange="selectfn_onchange()">
            <option value=""></option>
              <%
ei.LoadLists

i = 0
allnum = eu.GetUsersCount

do while i < allnum
	eu.GetUserByIndex i, name, domain, comment

	showline = true
	if isadmin() = false then
		showline = false
		if InStr(1, ManagerDomainString, Chr(9) & LCase(domain) & Chr(9)) > 0 then
			showline = true
		end if
	end if

	if showline = true and ei.IsInMailingList(name) = FALSE then
		Response.Write "<option value='" & name & "'>" & name & "</option>" & Chr(13)
	end if

	name = NULL
	domain = NULL
	comment = NULL

	i = i + 1
loop
%> 
            </select>
<% end if %>
        </td>
      </tr>
	<tr bgcolor="<%=MY_COLOR_3 %>">
	<td height="26" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	<input type="checkbox" name="isDisabled" id="isDisabled" <% if ei.isDisabled = true then response.write "checked"%>>
	<%=s_lang_0173 %>
	</td>
	</tr>
	<tr bgcolor="<%=MY_COLOR_3 %>">
	<td height="26" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	<input type="checkbox" name="isSendWithMailingList" id="isSendWithMailingList" <% if ei.isSendWithMailingList = true then response.write "checked"%>>
	<%=s_lang_0100 %>
	</td>
	</tr>
	<tr bgcolor="<%=MY_COLOR_3 %>">
	<td height="26" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	<input type="checkbox" name="isPrivate" id="isPrivate" <% if ei.isPrivate = true then response.write "checked"%>>
	<%=s_lang_0101 %>
	</td>
	</tr>
	<tr bgcolor="<%=MY_COLOR_3 %>">
	<td height="26" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	<input type="checkbox" name="isShowToCc" id="isShowToCc" <% if ei.isShowToCc = true then response.write "checked"%>>
	<%=s_lang_0164 %>
	</td>
	</tr>
<%
if isadmin() = true then
%>
	<tr bgcolor="<%=MY_COLOR_3 %>">
	<td height="26" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;<%=s_lang_0092 %>:&nbsp;<input type="text" name="dManagerDomain" id="dManagerDomain" maxlength="64" class="textbox" value="<%=ei.dManagerDomain %>">
	</td>
	</tr>
<%
end if
%>
  </table>
</div>
  <div align="center">
  <table align="center" border="0" width="90%" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-top:1px <%=MY_COLOR_1 %> solid;">
  <tr valign=bottom>
	<td height="22">&nbsp;<%=s_lang_0082 %>:</td>
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
	<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="ac_add()" type=button value="<%=s_lang_add %> &gt;&gt;">
	</td>
	</tr>
	<tr> 
	<td><br>
	<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="ac_delout()" type=button value="&lt;&lt; <%=s_lang_del %>">
	</td>
	</tr>
	<tr><td></td></tr>
	<tr><td></td></tr>
	</table>
	</td>
	<td>
	<select class="drpdwn" style="WIDTH: 305px" multiple size=4 name=listall width="305">
<%
i = 0
allnum = ei.AccreditItemCount

do while i < allnum
	tmp_name = server.htmlencode(ei.GetAccreditItemNameByIndex(i))
	Response.Write "<option value=""" & tmp_name & """>" & tmp_name & "</option>" & Chr(13)
	tmp_name = NULL
	i = i + 1
loop
%>
	</select>
	</font> </td>
  </tr>
	<tr><td height="10" colspan="3" align="right"><hr size="1" color="<%=MY_COLOR_1 %>"></td></tr>
  </table>


  <table align="center" border="0" width="90%" bgcolor="<%=MY_COLOR_3 %>">
    <tr> 
      <td height="6" rowspan="2" width="45%"> 
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0102 %></b></font></div>
      </td>
      <td height="6" rowspan="2"> 
      </td>
      <td height="6" rowspan="2" width="45%"> 
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0103 %></b></font></div>
      </td>
    </tr>
    <tr><td></td></tr>
    <tr> 
      <td height="94" rowspan="2" width="45%"> 
        <div align="center"> 
          <select name="selectalluser" size="5" class="drpdwn" style="width: 200;" multiple>
            <%
i = 0
allnum = eu.GetUsersCount

do while i < allnum
	eu.GetUserByIndex i, name, domain, comment

	showline = true
	if isadmin() = false then
		showline = false
		if InStr(1, ManagerDomainString, Chr(9) & LCase(domain) & Chr(9)) > 0 then
			showline = true
		end if
	end if

	if showline = true then
		Response.Write "<option value='" & name & "'>" & name & "</option>" & Chr(13)
	end if

	name = NULL
	domain = NULL
	comment = NULL

	i = i + 1
loop
%> 
          </select>
        </div>
		<br>
		<div align="center">
		<input type="text" name="addtext" class="textbox" maxlength="64">&nbsp;&nbsp;<input type="button" value="<%=s_lang_add %> &gt;&gt;" class="sbttn" LANGUAGE=javascript onclick="addinFromText()">
		</div>
      </td>
      <td height="47" width="10%"> 
        <div align="center"> 
          <input type="button" name="Button" value="<%=s_lang_add %> &gt;&gt;" class="sbttn" LANGUAGE=javascript onclick="addin()">
        </div>
      </td>
      <td height="94" rowspan="2" width="45%"> 
        <div align="center"> 
          <select name="selectlistusers" size="8" class="drpdwn" style="width: 200;" multiple>
<%=inlist %>
          </select>
        </div>
      </td>
    </tr>
    <tr> 
      <td height="47" width="10%"> 
        <div align="center"> 
          <input type="button" name="Submit2" value="&lt;&lt; <%=s_lang_del %>" class="sbttn" LANGUAGE=javascript onclick="delout()">
        </div>
      </td>
    </tr>
	<tr><td height="10" colspan="3" align="right"><hr size="1" color="<%=MY_COLOR_1 %>"></td></tr>
    <tr>
	<td colspan="3" align="right">
    <input type="button" value=" <%=s_lang_0059 %> " LANGUAGE=javascript onclick="addnew()" class="Bsbttn">&nbsp;&nbsp;
    <input type="button" value=" <%=s_lang_return %> " LANGUAGE=javascript onclick="javascript:location.href='browmailinglist.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td></tr>
    <tr>
  </table>
<input type="hidden" name="mode">
<input type="hidden" name="addname">
<input type="hidden" name="addlist">
<input type="hidden" name="accredit_list">
<%
if isadmin() = false then
%>
<input type="hidden" name="dManagerDomain" value="<%=ei.dManagerDomain %>">
<%
end if
%>
</form>
<br>
</BODY>
</HTML>

<%
set ei = nothing
set eu = nothing
set dm = nothing
%>
