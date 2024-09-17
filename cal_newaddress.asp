<!--#include file="passinc.asp" -->

<%
calid = trim(request("calid"))

preturl = trim(request("preturl"))
ppreturl = trim(request("ppreturl"))
returl = preturl

wmode = trim(request("wmode"))
themax = trim(request("themax"))

if Len(themax) > 0 and IsNumeric(themax) = true then
	themax = CLng(themax)
else
	themax = 0
end if

if Request.ServerVariables("REQUEST_METHOD") = "POST" and wmode = "1" and themax > 0 then
	dim ads
	set ads = server.createobject("easymail.Addresses")
	ads.Load Session("wem")

	isok = false
	i = 0

	do while i < themax
		if trim(request("check" & i)) <> "" then
			if ads.Simple_Add_Email(trim(request("nid" & i)), trim(request("mail" & i))) = true then
				isok = true
			end if
		end if

		i = i + 1
	loop

	if isok = true then
		ads.Save
	end if

	set ads = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl) & "&returl=" & Server.URLEncode(ppreturl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	else
		Response.Redirect "err.asp?" & getGRSN()
	end if
end if


dim ecal
set ecal = server.createobject("easymail.CalendarExtend")
ecal.Load Session("wem"), calid

il_allnum = ecal.Count
%>

<HTML>
<HEAD>
<TITLE>WinWebMail</TITLE>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<LINK href="images\hwem.css" rel=stylesheet>
<STYLE type=text/css>
<!--
.calendar_dayname {
	BORDER-TOP: #ffffc0 5px solid;
	BORDER-LEFT: #ffffc0 5px solid;
	BORDER-BOTTOM: #ffffc0 2px solid;
	FONT-WEIGHT: normal;
	color: #202020;
	BACKGROUND-COLOR: #ffffc0;
}
-->
</STYLE>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
function goback()
{
<%
if Len(preturl) > 0 then
%>
	location.href = "<%=preturl %>&returl=<%=Server.URLEncode(ppreturl) %>";
<%
else
%>
	history.back();
<%
end if
%>
}

function gosub() {
	if (ischeck() == true)
	{
		document.f1.wmode.value = "1";
		document.f1.submit();
	}
}

function ischeck() {
	var i = 0;
	var theObj;

	for(; i<<%=il_allnum %>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}

function checkall(check) {
	var i = 0;
	var theObj;

	for(; i<<%=il_allnum %>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			theObj.checked = check;
	}
}

function allcheck_onclick() {
	if (document.f1.allcheck.checked == true)
		checkall(true);
	else
		checkall(false);
}
//-->
</SCRIPT>

<BODY>
<br>
<form method="post" action="cal_newaddress.asp" name="f1">
<input type="hidden" name="returl" value="<%=returl %>">
<input type="hidden" name="preturl" value="<%=preturl %>">
<input type="hidden" name="ppreturl" value="<%=ppreturl %>">
<input type="hidden" name="calid" value="<%=calid %>">
<input type="hidden" name="wmode">
<input type="hidden" name="themax" value="<%=il_allnum %>">
  <table width="90%" border="0" align="center" cellspacing="0" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="2%" height="28">&nbsp;</td>
      <td><font class="s" color="<%=MY_COLOR_4 %>"><b>添加客人到我的地址簿</b></font></td>
    </tr>
  </table>

  <table width="90%" border="0" align="center">
    <tr>
      <td width="100%" class=calendar_dayname><font class="s">选择您想要添加到地址簿的收件人. 请确保每一个新联系人的昵称都是唯一的.</font></td>
    </tr>
  </table>
</td></tr>
</table>
<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
  <tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
    <td width="8%" align="center" height="25" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()"></td>
	<td width="46%" align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>电子邮件地址</b></font></td>
	<td width="46%" align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>昵称 (必需)</b></font></td>
<%
i = 0
do while i < il_allnum
	if ecal.MoveTo(i) = true then
		Response.Write "<tr>"
		Response.Write "<td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "' value='" & i & "'></td>"

		Response.Write "<td align='left' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='text' size='35' class='textbox' name='mail" & i & "' value=""" & ecal.ce_email & """></td>"

		Response.Write "<td align='left' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='text' size='35' class='textbox' name='nid" & i & "' value="""
		if Len(ecal.ce_myname) > 0 then
			Response.Write ecal.ce_myname
		else
			Response.Write ecal.ce_username
		end if
		Response.Write """></td>"

		Response.Write "</tr>" & Chr(13)
	end if

	i = i + 1
loop
%>
</table>
  <table width="90%" border="0" align="center">
    <tr valign="middle"> 
      <td colspan="2" height="40" bgcolor="#ffffff"><br>
        <div align="right"> 
          <input type=button value="添加到地址簿" class="Bsbttn" style="WIDTH: 112px" onClick="javascript:gosub()">&nbsp;&nbsp;
          <input type=button value=" 取消 " class="Bsbttn" style="WIDTH: 60px" onclick="javascript:goback();">&nbsp;&nbsp;&nbsp;
        </div>
      </td>
    </tr>
  </table>
  </FORM>
<br>
</BODY>
</HTML>

<%
set ecal = nothing
%>
