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

dim ei
set ei = server.createobject("easymail.KeyWordFilterManager")
ei.Load

dim eit
set eit = server.createobject("easymail.KeyWordFilterTrashManager")
eit.Load

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ei.RemoveAll
	am = trim(request("allmsgs"))

	ps = 1
	pe = 1

	do while pe >= ps
		pe = InStr(ps, am, Chr(12))

		if pe > 0 then
			ei.Add Mid(am, ps, pe - ps)

			pe = pe + 1
			ps = pe
		else
			Exit Do
		end if
	loop

	ei.Save


	eit.RemoveAll
	am = trim(request("allmsgs2"))

	ps = 1
	pe = 1

	do while pe >= ps
		pe = InStr(ps, am, Chr(12))

		if pe > 0 then
			eit.Add Mid(am, ps, pe - ps)

			pe = pe + 1
			ps = pe
		else
			Exit Do
		end if
	loop

	eit.Save

	set ei = nothing
	set eit = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=keywords.asp"
end if
%>


<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
<STYLE type=text/css>
<!--
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.jtextbox {border:1px solid #555;}
-->
</STYLE>
</HEAD>

<script LANGUAGE=javascript>
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

	tempstr = "";
	i = 0;
	for (i; i < document.f1.listall2.length; i++)
	{
		tempstr = tempstr + document.f1.listall2[i].value + "\f";
	}
	document.f1.allmsgs2.value = tempstr;

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
	if (document.f1.addinfo.value.indexOf("\f") != -1)
	{
		alert("<%=s_lang_inputerr %>");
		document.f1.addinfo.focus();
		return ;
	}

	if (document.f1.addinfo.value.length > 0)
	{
		if (haveit() == false)
		{
			var oOption = document.createElement("OPTION");
			oOption.text = document.f1.addinfo.value;
			oOption.value = document.f1.addinfo.value;
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
			document.f1.addinfo.value = "";
			document.f1.addinfo.focus();
			return ;
		}
		else
		{
			document.f1.addinfo.value = "";
			document.f1.addinfo.focus();
			return ;
		}
	}
}

function haveit()
{
	var tempstr = document.f1.addinfo.value;

	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].value == tempstr)
			return true;
	}

	return false;
}

function delout2()
{
	var i = 0;
	for (i; i < document.f1.listall2.length; i++)
	{
		if (document.f1.listall2[i].selected == true)
		{
			document.f1.listall2.remove(i);

			if (i < document.f1.listall2.length)
				document.f1.listall2[i].selected = true;
			else
			{
				if (i - 1 >= 0)
					document.f1.listall2[i - 1].selected = true;
			}

			break;
		}
	}
}

function add2()
{
	if (document.f1.addinfo2.value.indexOf("\f") != -1)
	{
		alert("<%=s_lang_inputerr %>");
		document.f1.addinfo2.focus();
		return ;
	}

	if (document.f1.addinfo2.value.length > 0)
	{
		if (haveit2() == false)
		{
			var oOption = document.createElement("OPTION");
			oOption.text = document.f1.addinfo2.value;
			oOption.value = document.f1.addinfo2.value;
<%
if isMSIE = true then
%>
			document.f1.listall2.add(oOption);
<%
else
%>
			document.f1.listall2.appendChild(oOption);
<%
end if
%>
			document.f1.addinfo2.value = "";
			document.f1.addinfo2.focus();
			return ;
		}
		else
		{
			document.f1.addinfo2.value = "";
			document.f1.addinfo2.focus();
			return ;
		}
	}
}

function haveit2()
{
	var tempstr = document.f1.addinfo2.value;

	var i = 0;
	for (i; i < document.f1.listall2.length; i++)
	{
		if (document.f1.listall2[i].value == tempstr)
			return true;
	}

	return false;
}
//-->
</script>

<BODY>
<br>
<FORM ACTION="#" METHOD="POST" NAME="f1">
<input type="hidden" name="allmsgs">
<input type="hidden" name="allmsgs2">
<div align="center"><br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="10%" height="28">&nbsp;</td>
      <td width="35%"><a href="showsysinfo.asp?<%=getGRSN() %>#keywords"><%=s_lang_enable %></a></td>
      <td colspan="32"><a href="right.asp?<%=getGRSN() %>"><%=s_lang_return %></a></td>
      <td width="23%"><font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0620 %></b></font></td>
    </tr>
  </table>
</div>
  <table align="center" border="0" width="86%" cellspacing="0">
	<tr><td height="12" colspan="2" style="padding-top:12px; padding-left:8px; padding-bottom:12px;"></td></tr>
	<tr><td colspan="2" style="border-bottom:2px #a7c5e2 solid; font-size:14px; color:#093665; padding-left:6px;">
	<%=s_lang_0621 %></td></tr>
    <tr><td style="padding-top:8px;" width="30%">
	&nbsp;<select name="listall" size="8" class="drpdwn" style="width:480px;">
<%
i = 0
allnum = ei.Count

do while i < allnum
	kstr = ei.GetInfo(i)
	Response.Write "<option value=""" & server.htmlencode(kstr) & """>" & server.htmlencode(kstr) & "</option>" & Chr(13)

	kstr = NULL
	i = i + 1
loop
%>
	</select>
	</td>
	<td style="padding-left:20px;">
	<input type="button" value="<%=s_lang_del %>" class="sbttn" LANGUAGE=javascript onclick="delout()">
	</td>
    </tr>
    <tr>
      <td style="padding-top:12px;" align="left" colspan="2">
	&nbsp;<input type="input" name="addinfo" class='jtextbox' maxlength="64" size="40">
	<input type="button" value="<%=s_lang_add %>" class="sbttn" LANGUAGE=javascript onclick="add()">
	</td></tr>
  </table>
<br>
  <table align="center" border="0" width="86%" cellspacing="0">
	<tr><td height="20" colspan="2" style="padding-top:12px; padding-left:8px; padding-bottom:12px;"></td></tr>
	<tr><td colspan="2" style="border-bottom:2px #a7c5e2 solid; font-size:14px; color:#093665; padding-left:6px;">
	<%=s_lang_0622 %></td></tr>
    <tr><td style="padding-top:8px;" width="30%">
	&nbsp;<select name="listall2" size="8" class="drpdwn" style="width:480px;">
<%
i = 0
allnum = eit.Count

do while i < allnum
	kstr = eit.GetInfo(i)
	Response.Write "<option value=""" & server.htmlencode(kstr) & """>" & server.htmlencode(kstr) & "</option>" & Chr(13)

	kstr = NULL
	i = i + 1
loop
%>
	</select>
	</td>
	<td style="padding-left:20px;">
	<input type="button" value="<%=s_lang_del %>" class="sbttn" LANGUAGE=javascript onclick="delout2()">
	</td>
    </tr>
    <tr>
      <td style="padding-top:12px;" align="left" colspan="2">
	&nbsp;<input type="input" name="addinfo2" class='jtextbox' maxlength="64" size="40">
	<input type="button" value="<%=s_lang_add %>" class="sbttn" LANGUAGE=javascript onclick="add2()">
	</td></tr>
	<tr>
	<td height="20" colspan="2" align="right"><br><hr size="1" color="<%=MY_COLOR_1 %>">
    <input type="button" value="<%=s_lang_save %>" LANGUAGE=javascript onclick="sub()" class="Bsbttn">&nbsp;&nbsp;
    <input type="button" value="<%=s_lang_return %>" LANGUAGE=javascript onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
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
        <td width="94%"><%=s_lang_0623 %>
        <br><br><%=s_lang_tpf %>
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
<br><br>
</FORM>
</BODY>
</HTML>

<%
set ei = nothing
set eit = nothing
%>
