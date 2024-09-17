<!--#include file="passinc.asp" -->

<%
calid = trim(request("calid"))
msgname = trim(request("msgname"))
if Len(msgname) < 1 then
	msgname = Session("wem")
end if

preturl = trim(request("preturl"))
ppreturl = trim(request("ppreturl"))
returl = preturl

dim ecal

if Len(calid) > 0 and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	smyes = false
	smwait = false
	smno = false

	if Len(trim(request("g_yes"))) > 0 then
		smyes = true
	end if

	if Len(trim(request("g_wait"))) > 0 then
		smwait = true
	end if

	if Len(trim(request("g_no"))) > 0 then
		smno = true
	end if

	Set mailsend = Server.CreateObject("easymail.MailSend")
	mailsend.createnew Session("wem"), Session("tid")
	mailsend.MailName = msgname
	mailsend.EM_Subject = trim(request("msgsubject"))
	mailsend.EM_Text = trim(request("msgtext"))

	set ecal = server.createobject("easymail.CalendarExtend")
	ecal.Load Session("wem"), calid

	sendto_str = ""
	i = 0
	do while i < ecal.Count
		if ecal.MoveTo(i) = true then
			if (ecal.ce_join = 1 and smyes = true) or (ecal.ce_join = 0 and smwait = true) or (ecal.ce_join = -1 and smno = true) then
				if LCase(ecal.ce_email) <> LCase(Session("mail")) then
					if Len(sendto_str) > 0 then
						sendto_str = sendto_str & "," & ecal.ce_email
					else
						sendto_str = ecal.ce_email
					end if
				end if
			end if
		end if

	    i = i + 1
	loop

	isok = false
	if Len(sendto_str) > 0 then
		mailsend.EM_To = sendto_str
		isok = mailsend.Send()
	end if

	set ecal = nothing
	set mailsend = nothing

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


set ecal = server.createobject("easymail.Calendar")
ecal.Load Session("wem")

moveisok = ecal.MoveToID(calid)

if moveisok = false then
	set ecal = nothing
	Response.Redirect "err.asp?" & getGRSN()
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

function gosub()
{
	if (document.f1.g_yes.checked == false && document.f1.g_wait.checked == false && document.f1.g_no.checked == false)
		alert("请选择收件人.");
	else
		document.f1.submit();
}
//-->
</script>


<BODY>
<br>
<form method="post" action="cal_sendmsg.asp" name="f1">
<input type="hidden" name="returl" value="<%=returl %>">
<input type="hidden" name="preturl" value="<%=preturl %>">
<input type="hidden" name="ppreturl" value="<%=ppreturl %>">
<input type="hidden" name="calid" value="<%=calid %>">
  <table width="90%" border="0" align="center" cellspacing="0" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="2%" height="28">&nbsp;</td>
      <td><font class="s" color="<%=MY_COLOR_4 %>"><b>给您的客人发信</b></font></td>
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
<input type="checkbox" name="g_yes" checked>将要出席的客人<br>
<input type="checkbox" name="g_wait" checked>仍未做决定的客人<br>
<input type="checkbox" name="g_no" checked>谢绝出席的客人<br>
				</td>
			</tr>
			<tr>
				<td valign=center height=27 align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
		<b>主题</b>:&nbsp;
				</td>
				<td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
		<input type="text" name="msgsubject" class='textbox' value="<%=ecal.bi_name %>" size="50" maxlength="500">
				</td>
			</tr>
			<tr>
				<td valign=center height=27 align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
		<b>内容</b>:&nbsp;
				</td>
				<td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
		<textarea name="msgtext" cols="55" rows="10" class='textarea'><%=ecal.bi_note %></textarea>
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
