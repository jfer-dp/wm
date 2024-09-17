<!--#include file="passinc.asp" -->

<%
email = trim(request("email"))
calid = trim(request("calid"))

purl = trim(request("purl"))
returl = trim(request("returl"))
preturl = trim(request("preturl"))
ppreturl = trim(request("ppreturl"))

if Len(purl) < 1 then
	purl = "preturl=" & Server.URLEncode(preturl) & "&ppreturl=" & Server.URLEncode(ppreturl)
end if

dim ecal
set ecal = server.createobject("easymail.CalendarExtend")
ecal.Load Session("wem"), calid
moveisok = ecal.MoveToEmail(email)

if moveisok = false then
	set ecal = nothing
	Response.Redirect "err.asp?" & getGRSN()
end if

if moveisok = true and Len(email) > 0 and Len(calid) > 0 and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ce_join = CLng(trim(request("ce_join")))

	ecal.ce_withGuest = 0
	if ce_join >= 0 then
		ce_withGuest = trim(request("ce_withGuest"))
		if IsNumeric(ce_withGuest) = true then
			ecal.ce_withGuest = CLng(ce_withGuest)
		end if
	end if

	ecal.ce_myname = trim(request("ce_myname"))
	ecal.ce_remark = trim(request("ce_remark"))

	if ce_join = -2 then
		ecal.ce_join = -1
		ecal.ce_askRemove = true
		ecal.ce_withGuest = 0
	else
		ecal.ce_askRemove = false
		ecal.ce_join = ce_join

		if ce_join = -1 then
			ecal.ce_withGuest = 0
		end if
	end if

	if LCase(ecal.ce_email) = LCase(Session("mail")) then
		ecal.ce_join = 1
		ecal.ce_askRemove = false
	end if

	isok = false
	if ecal.Set(email) = true then
		if ecal.Save() = true then
			isok = true
		end if
	end if

	set ecal = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl & "&" & purl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl & "&" & purl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
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
function window_onload() {
<%
if ecal.ce_askRemove = false then
	Response.Write "document.f1.ce_join.value = " & ecal.ce_join & ";" & Chr(13)
else
	Response.Write "document.f1.ce_join.value = -2;" & Chr(13)
end if
%>
}

function goback()
{
	if (document.f1.returl.value.length < 3)
		history.back();
	else
		location.href = document.f1.returl.value + "&" + document.f1.purl.value;
}

function gosub()
{
	document.f1.submit();
}

function godel()
{
	if (confirm("确实要删除吗?") == false)
		return ;

	location.href = "cal_del.asp?<%=getGRSN() %>&calmode=9&calid=<%=calid %>&email=<%=Server.URLEncode(email) %>&returl=<%=Server.URLEncode(returl) %>&purl=<%=Server.URLEncode(purl) %>";
}
//-->
</script>


<BODY LANGUAGE=javascript onload="return window_onload()">
<br>
<form method="post" action="cal_editguest.asp" name="f1">
<input type="hidden" name="returl" value="<%=returl %>">
<input type="hidden" name="purl" value="<%=purl %>">
<input type="hidden" name="calid" value="<%=calid %>">
<input type="hidden" name="email" value="<%=email %>">
<input type="hidden" name="calmode" value="8">
  <table width="90%" border="0" align="center" cellspacing="0" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="2%" height="28">&nbsp;</td>
      <td width="98%"><font class="s" color="<%=MY_COLOR_4 %>"><b>编辑客人回复</b></font></td>
      </td>
    </tr>
  </table><br>

  <table width="91%" border="0" align="center">
    <tr> 
      <td>
        <div align="center">
          <table width="100%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border-top:1px <%=MY_COLOR_1 %> solid;">
			<tr>
				<td valign=center width="18%" height=27 align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
		<b>姓名</b>:&nbsp;
				</td>
				<td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
		<input type="text" name="ce_myname" class='textbox' value="<%=ecal.ce_myname %>" size="40" maxlength="40">
				</td>
			</tr>
			<tr>
				<td valign=center height=27 align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
		<b>电子邮件</b>:&nbsp;
				</td>
				<td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><%=server.htmlencode(ecal.ce_email) %>&nbsp;</td>
			</tr>
			<tr>
				<td valign=center height=27 align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
		<b>帐号</b>:&nbsp;
				</td>
				<td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><%=server.htmlencode(ecal.ce_username) %>&nbsp;</td>
			</tr>
			<tr>
				<td valign=center height=27 align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
		<b>回复结果</b>:&nbsp;
				</td>
				<td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
		<select name="ce_join" class="drpdwn">
		<option value="1">是的 - 参加</option>
		<option value="0">也许 - 未决定的</option>
		<option value="-1">不 - 婉言谢绝</option>
		<option value="-2">不 - 不列入请柬</option>
		</select>
				</td>
			</tr>
			<tr>
				<td valign=center width="15%" height=27 align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
		<b>客人</b>:&nbsp;
				</td>
				<td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
		<input type="text" name="ce_withGuest" class='textbox' value="<%=ecal.ce_withGuest %>" size="3" maxlength="3">
				</td>
			</tr>
			<tr>
				<td valign=center width="15%" height=27 align=right style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
		<b>备注</b>:&nbsp;
				</td>
				<td align=left style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
		<input type="text" name="ce_remark" class='textbox' value="<%=ecal.ce_remark %>" size="50" maxlength="100">
				</td>
			</tr>
			<tr>
			<td colspan=2 height=50 valign=center align=right bgcolor="#ffffff"> 
      <input type="button" value="保存" style="WIDTH: 60px" onclick="javascript:gosub()" class="Bsbttn">&nbsp;&nbsp;
      <input type="button" value="删除" style="WIDTH: 60px" onclick="javascript:godel()" class="Bsbttn">&nbsp;&nbsp;
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
