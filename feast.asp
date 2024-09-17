<!--#include file="passinc.asp" --> 

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

dim efst
selectdomain = trim(request("selectdomain"))

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	curdomain = trim(request("curdomain"))

	if Len(curdomain) = 0 and isadmin() = true then
		set efst = server.createobject("easymail.CalSystemFeast")
		efst.Load
	else
		allnum = dm.GetUserManagerDomainCount(Session("wem"))
		isok = false
		i = 0

		do while i < allnum
			if curdomain = dm.GetUserManagerDomain(Session("wem"), i) then
				isok = true
	            exit do
			end if

			i = i + 1
		loop

		if isok = true then
			set efst = server.createobject("easymail.CalDomainFeasts")
			efst.Load curdomain
		else
			set dm = nothing
			Response.Redirect "noadmin.asp"
		end if
	end if

	efst.RemoveAll
	efst.RemoveAllNL

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
				efst.Add item
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
				efst.AddNL item
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	efst.Save

	set efst = nothing
	set dm = nothing

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=feast.asp?selectdomain=" & Server.URLEncode(curdomain)
end if
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>
<SCRIPT LANGUAGE=javascript>
<!--
function domainname_onchange() {
	location.href = "feast.asp?<%=getGRSN() %>&selectdomain=" + document.f1.domainname.value;
}

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
		alert("输入错误!");
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

	alert("输入错误!");
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
		alert("输入错误!");
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
<%
if isMSIE = true then
%>
			document.f1.listallNL.add(oOption);
<%
else
%>
			document.f1.listallNL.appendChild(oOption);
<%
end if
%>
			return ;
		}
		else
			return ;
	}

	alert("输入错误!");
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
<%
if isMSIE = true then
%>
	if (event.keyCode == 13)
	{
		event.keyCode = 9;
		addNL();
	}
<%
end if
%>
}
//-->
</SCRIPT>


<BODY>
<br><br>
<FORM ACTION="feast.asp" METHOD=POST NAME="f1">
<input type="hidden" name="allmsgs">
<input type="hidden" name="allmsgsNL">
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>">
	<td height="30" align="center" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
	<font class="s" color="<%=MY_COLOR_4 %>"><b>节日设置</b></font>
	</td>
    </tr>
    <tr><td height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid;">
&nbsp;&nbsp;操作对象:&nbsp;<select name="domainname" class="drpdwn" LANGUAGE=javascript onchange="return domainname_onchange()">
<%
i = 0

if isadmin() = false then
	allnum = dm.GetUserManagerDomainCount(Session("wem"))

	do while i < allnum
		domain = dm.GetUserManagerDomain(Session("wem"), i)

		if domain <> selectdomain then
			response.write "<option value='" & server.htmlencode(domain) & "'>" & server.htmlencode(domain) & "</option>" & Chr(13)
		else
			curdomain = domain
			response.write "<option value='" & server.htmlencode(domain) & "' selected>" & server.htmlencode(domain) & "</option>" & Chr(13)
		end if

		domain = NULL

		i = i + 1
	loop
else
	Response.Write "<option>系统节日</option>" & Chr(13)
end if


'-----------------------------------------
if isadmin() = false then
	if curdomain = "" then
		curdomain = dm.GetUserManagerDomain(Session("wem"), 0)
		set efst = server.createobject("easymail.CalDomainFeasts")
		efst.Load curdomain
	else
		set efst = server.createobject("easymail.CalDomainFeasts")
		efst.Load curdomain
	end if
else
	set efst = server.createobject("easymail.CalSystemFeast")
	efst.Load
end if
%>
</select>
	</td></tr>
	<tr>
	<td height=22 valign=bottom colspan=2 align=center>
	<font class="s"><b>编辑公历节日</b></font><br>
	</td>
	</tr>
    <tr>
	<td height="30" style="border-bottom:1px <%=MY_COLOR_1 %> solid;">
<table>
  <tr valign=top> 
	<td>
		&nbsp;<select name="gl_month" class="drpdwn">
<%
i = 1

do while i < 13
	if i < 10 then
		Response.Write "<option value=""0" & i & """>" & i & "月</option>" & Chr(13)
	else
		Response.Write "<option value=""" & i & """>" & i & "月</option>" & Chr(13)
	end if

	i = i + 1
loop
%>
		</select><select name="gl_day" class="drpdwn">
<%
i = 1

do while i < 32
	if i < 10 then
		Response.Write "<option value=""0" & i & """>" & i & "日</option>" & Chr(13)
	else
		Response.Write "<option value=""" & i & """>" & i & "日</option>" & Chr(13)
	end if

	i = i + 1
loop
%>
		</select>
	&nbsp;<input maxlength=30 size=12 name="addmsg" class='textbox' onkeydown="goent()"><br>
	&nbsp;(输入公历节日名称)
	</td>
    <td align=middle> 
      <table cellspacing=0 cellpadding=0>
        <tr> 
          <td>
			<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="add()" type=button value="添加 >>">
		</td>
		</tr>
		<tr> 
			<td><br>
			<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="delout()" type=button value="<< 删除">
			</td>
		</tr>
	</table>
	</td>
	<td>
	<select class="drpdwn" style="WIDTH: 250px" multiple size=7 name=listall width="230">
<%
i = 0
allnum = efst.Count

do while i < allnum
	efst.Get i, mm, dd, fname

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
	</td>
    </tr>
	<tr>
	<td height=22 valign=bottom colspan=2 align=center>
	<font class="s"><b>编辑农历节日</b></font><br>
	</td>
	</tr>
    <tr>
	<td height="30" style="border-bottom:1px <%=MY_COLOR_1 %> solid;">
<table>
  <tr valign=top> 
	<td>
		&nbsp;<select name="nl_month" class="drpdwn">
<%
i = 1

do while i < 13
	if i < 10 then
		Response.Write "<option value=""0" & i & """>" & i & "月</option>" & Chr(13)
	else
		Response.Write "<option value=""" & i & """>" & i & "月</option>" & Chr(13)
	end if

	i = i + 1
loop
%>
		</select><select name="nl_day" class="drpdwn">
<%
i = 1

do while i < 31
	if i < 10 then
		Response.Write "<option value=""0" & i & """>" & i & "日</option>" & Chr(13)
	else
		Response.Write "<option value=""" & i & """>" & i & "日</option>" & Chr(13)
	end if

	i = i + 1
loop
%>
		</select>
	&nbsp;<input maxlength=30 size=12 name="addmsgNL" class='textbox' onkeydown="goentNL()"><br>
	&nbsp;(输入农历节日名称)
	</td>
    <td align=middle> 
      <table cellspacing=0 cellpadding=0>
        <tr> 
          <td>
			<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="addNL()" type=button value="添加 >>">
		</td>
		</tr>
		<tr> 
			<td><br>
			<input class="sbttn" style="WIDTH: 70px" LANGUAGE=javascript onclick="deloutNL()" type=button value="<< 删除">
			</td>
		</tr>
	</table>
	</td>
	<td>
	<select class="drpdwn" style="WIDTH: 250px" multiple size=7 name=listallNL width="230">
<%
i = 0
allnum = efst.CountNL

do while i < allnum
	efst.GetNL i, mm, dd, fname

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
	</td>
    </tr>
    <tr> 
	<td align="right" bgcolor="#ffffff">
	<br><input type="button" value=" 保存 " LANGUAGE=javascript onclick="gosub()" class="Bsbttn">&nbsp;&nbsp;
<%
if isadmin() = false then
%>
	<input type="button" value=" 返回 " onclick="javascript:location.href='domainright.asp?<%=getGRSN() %>';" class="Bsbttn">
<%
else
%>
	<input type="button" value=" 返回 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
<%
end if
%>
	</td>
	</tr>
  </table>
<input name="curdomain" type="hidden" value="<%=curdomain %>">
</FORM>
<br>
</BODY>
</HTML>

<%
set efst = nothing
set dm = nothing


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
