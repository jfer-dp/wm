<!--#include file="passinc.asp" -->

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

isMobile = false
dim http_user_agent
http_user_agent = LCase(Request.ServerVariables("HTTP_User-Agent"))
if InStr(http_user_agent, "applewebkit") > 0 or InStr(http_user_agent, "mobile") > 0 then
	if InStr(http_user_agent, "iphone") > 0 or InStr(http_user_agent, "ipod") > 0 or InStr(http_user_agent, "android") > 0 or InStr(http_user_agent, "ios") > 0 or InStr(http_user_agent, "ipad") > 0 then
		isMobile = true
	end if
end if

calid = trim(request("calid"))
msgname = trim(request("msgname"))
if Len(msgname) < 1 then
	msgname = Session("wem")
end if

preturl = trim(request("preturl"))
ppreturl = trim(request("ppreturl"))
returl = preturl

dim ecal
set ecal = server.createobject("easymail.Calendar")
ecal.Load Session("wem")

moveisok = ecal.MoveToID(calid)

if moveisok = false then
	set ecal = nothing
	Response.Redirect "err.asp?" & getGRSN()
end if

newemails = trim(request("newemails"))
if Len(calid) > 0 and Len(newemails) > 0 and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ecal.bi_has_invitation = true
	ecal.bi_notice_name = msgname
	ecal.bi_notice_email = Session("mail")

	isok = ecal.Set(calid)

	if isok = true then
		isok = ecal.Save()

		if isok = true then
			errnum = ecal.AddEmails(calid, newemails)
			isok = false
			if errnum = 0 then
				isok = true
			end if
		end if
	end if

	set ecal = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl) & "&returl=" & Server.URLEncode(ppreturl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	else
		if Len(returl) > 3 then
			if errnum > 0 then
				Response.Redirect "err.asp?" & getGRSN() & "&errstr=" & Server.URLEncode("邀请客人时有" & errnum & "处错误")
			else
				Response.Redirect "err.asp?" & getGRSN()
			end if
		else
			Response.Redirect "err.asp?" & getGRSN()
		end if
	end if
end if
%>

<html>
<head>
<TITLE>WinWebMail</TITLE>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<LINK href="images\hwem.css" rel=stylesheet>
</head>

<script language="JavaScript">
<!--
function goback()
{
	history.back();
}

function gosub()
{
	if (document.f1.newemails.value.length == 0)
	{
		alert("请输入收件人邮件地址.");
		document.f1.newemails.focus();
	}
	else
		document.f1.submit();
}

function popaddress()
{
	var remote = null;
	remote = window.open("selectadd.asp?mode=To&ofm=<%=Server.URLEncode("document.f1.newemails") %>&<%=getGRSN() %>", "", "top=80; left=150; height=345,width=496,scrollbars=yes,resizable=yes,status=no,toolbar=no,menubar=no,location=no");
}

<%
if Application("em_EnableEntAddress") = true then
%>
function eapop() {
	window.open("ea_pop.asp?mode=To&ofm=<%=Server.URLEncode("document.f1.newemails") %>&<%=getGRSN() %>", "", "top=80; left=130; height=330,width=510,scrollbars=yes,resizable=yes,status=no,toolbar=no,menubar=no,location=no");
}
<%
end if
%>
//-->
</script>


<BODY>
<br>
<form method="post" action="cal_moreguest.asp" name="f1">
<input type="hidden" name="returl" value="<%=returl %>">
<input type="hidden" name="preturl" value="<%=preturl %>">
<input type="hidden" name="ppreturl" value="<%=ppreturl %>">
<input type="hidden" name="calid" value="<%=calid %>">
  <table width="90%" border="0" align="center" cellspacing="0" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="2%" height="28">&nbsp;</td>
      <td><font class="s" color="<%=MY_COLOR_4 %>"><b>邀请更多的客人-<%=server.htmlencode(ecal.bi_name) %></b></font></td>
    </tr>
  </table><br>

  <table width="91%" border="0" align="center">
    <tr> 
      <td>
        <div align="center">
          <table width="100%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-top:1px <%=MY_COLOR_1 %> solid;">
			<tr>
				<td valign=center width="18%" height=27 align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
		<b>发件人</b>:&nbsp;
				</td>
				<td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
		<input type="text" name="msgname" class='textbox' value="<%=msgname %>" size="30" maxlength="100">
				</td>
			</tr>
			<tr>
				<td valign=center height=27 align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
		<b>收件人</b>:&nbsp;
				</td>
				<td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
				输入您想要发送电子请帖的电子邮件地址，用逗号间隔。<br>
<%
if isMobile = false then
%>
				[<a href="javascript:popaddress()">从地址簿选择收件人</a>]
<%
	if Application("em_EnableEntAddress") = true then
%>
&nbsp;[<a href="javascript:eapop()">从企业地址簿选择收件人</a>]
<%
	end if
%>
	<br>
<%
end if
%>
		<textarea name="newemails" cols="50" rows="6" class='textarea'></textarea>
				</td>
			</tr>
			<tr>
			<td colspan=2 height=50 valign=center align=right bgcolor="#ffffff"> 
      <input type="button" value="发送" style="WIDTH: 60px" onclick="javascript:gosub()" class="Bsbttn">&nbsp;&nbsp;
      <input type="button" value="取消" style="WIDTH: 60px" onclick="javascript:goback();" class="Bsbttn">&nbsp;&nbsp;&nbsp;
			</td>
			</tr>
		</table>
	</table>
</form>
<br>
</body>
</html>

<%
set ecal = nothing
%>
