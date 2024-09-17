<!--#include file="passinc.asp" --> 

<%
nm = trim(request("nm"))
vid = trim(request("vid"))
purl = trim(request("purl"))

is_domain_manager = false

if isadmin() = false and isAccountsAdmin() = false then
	is_domain_manager = true

	if LCase(nm) = LCase(Application("em_SystemAdmin")) then
		Response.Redirect "noadmin.asp"
	end if

	if Len(vid) > 0 then
		Response.Redirect "noadmin.asp"
	end if

	dim dm
	set dm = server.createobject("easymail.Domain")
	dm.Load
	if dm.GetUserManagerDomainCount(Session("wem")) < 1 then
		set dm = nothing
		Response.Redirect "noadmin.asp"
	end if

	dim curdomain

	if Len(nm) > 0 then
		set emusers = Application("em")

		emusers.GetUserByName nm, outname, outdomain, outcomment
		curdomain = outdomain

		set emusers = nothing

		outname = NULL
		outdomain = NULL
		outcomment = NULL
	end if


	allnum = dm.GetUserManagerDomainCount(Session("wem"))
	isok = false
	i = 0

	if Len(curdomain) > 0 then
		do while i < allnum
			if LCase(curdomain) = LCase(dm.GetUserManagerDomain(Session("wem"), i)) then
				isok = true
	            exit do
			end if

			i = i + 1
		loop
	end if

	set dm = nothing

	if isok = false then
		Response.Redirect "noadmin.asp"
	end if
end if


if isadmin() = false then
	if LCase(nm) = LCase(Application("em_SystemAdmin")) then
		Response.Redirect "noadmin.asp"
	end if
end if
%>

<%
dim ei
set ei = server.createobject("easymail.MoreRegInfo")

if Len(vid) > 30 then
	ei.LoadTempRegInfo vid
elseif Len(nm) > 0 then
	ei.LoadRegInfo nm
end if

allnum = ei.Count_RegInfo
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>

<SCRIPT LANGUAGE=javascript>
<!--
<%
if purl <> "" then
%>
parent.f1.document.leftval.purl.value = "<%=purl %>";
<%
end if
%>
function goback(){
	if (parent.f1.document.leftval.purl.value.length < 1)
<%
if is_domain_manager = false then
%>
		location.href = "showuser.asp?<%=getGRSN() %>";
<%
else
%>
		location.href = "showdomainusers.asp?<%=getGRSN() %>";
<%
end if
%>
	else
		location.href = parent.f1.document.leftval.purl.value;
}
//-->
</SCRIPT>
</HEAD>

<BODY>
<br><br>
<FORM ACTION="setreginfo.asp" METHOD="POST" NAME="f1">
<input type="hidden" name="allmsgs">
<%
if isadmin() = true then
%>
<table align="center" border="0" width="85%" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style="border:1px <%=MY_COLOR_1 %> solid;">
	<tr>
	<td align="center" width="33%" height="28" style="border-right:1px <%=MY_COLOR_1 %> solid;">
	<a href="sendloglist.asp?user=<%=Server.URLEncode(nm) %>&<%=getGRSN() %>&returl=<%=Server.URLEncode("viewreginfo.asp?nm=" & nm) %>">查看发信记录</a>
	</td>
	<td align="center" width="33%" height="28" style="border-right:1px <%=MY_COLOR_1 %> solid;">
	<a href="listmon.asp?user=<%=Server.URLEncode(nm) %>&<%=getGRSN() %>&inout=out&purl=<%=Server.URLEncode("viewreginfo.asp?nm=" & nm) %>">查看已监控到的发送邮件</a>
	</td>
	<td align="center">
	<a href="listmon.asp?user=<%=Server.URLEncode(nm) %>&<%=getGRSN() %>&inout=in&purl=<%=Server.URLEncode("viewreginfo.asp?nm=" & nm) %>">查看已监控到的接收邮件</a>
	</td>
	</tr>
</table>
<br><br>
<%
end if
%>
  <div align="center">
  <table align="center" border="0" width="85%" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
	<tr bgcolor="<%=MY_COLOR_2 %>"> 
	<td colspan="2" height="30" align="center" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<font class="s" color="<%=MY_COLOR_4 %>"><b>个人资料 - <%=nm %></b></font>
	</td>
	</tr>
	<tr><td colspan="2" align='left' width='40%' height='26' style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>&nbsp;<b>注释</b>:&nbsp;<%=server.htmlencode(ei.Comment) %></td></tr>
	<tr><td align='right' width='40%' height='26' style='border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;'>本次登录IP:&nbsp;</td><td width='60%' style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
	<%=server.htmlencode(ei.CurrentlyIP) %>&nbsp;</td></tr>
	<tr><td align='right' width='40%' height='26' style='border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;'>本次登录时间:&nbsp;</td><td width='60%' style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
	<%=server.htmlencode(ei.CurrentlyTime) %>&nbsp;</td></tr>
	<tr><td align='right' width='40%' height='26' style='border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;'>上次登录IP:&nbsp;</td><td width='60%' style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
	<%=server.htmlencode(ei.PreviousIP) %>&nbsp;</td></tr>
	<tr><td align='right' width='40%' height='26' style='border-bottom:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;'>上次登录时间:&nbsp;</td><td width='60%' style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
	<%=server.htmlencode(ei.PreviousTime) %>&nbsp;</td></tr>
<%
i = 0

do while i < allnum
	ei.Get_RegInfo i, s_name, s_sel, s_len, s_msg

	Response.Write "<tr><td nowrap align='right' width='40%' height='26' style='border-bottom:1px " & MY_COLOR_1 & " solid; border-right:1px " & MY_COLOR_1 & " solid;'>&nbsp;" & server.htmlencode(s_name) & ":&nbsp;</td><td width='60%' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>"
	Response.Write server.htmlencode(s_msg) & "&nbsp;</td></tr>" & Chr(13)

	s_name = NULL
	s_sel = NULL
	s_len = NULL
	s_msg = NULL

	i = i + 1
loop
%></td></tr>
    <tr> 
		<td colspan="2" height="40" bgcolor="#ffffff" align="right"><br>
          <input type="button" value=" 返回 " onclick="javascript:goback()" class="Bsbttn">
		</td>
    </tr>
  </table>
</form>
</div>
</FORM>
</BODY>
</HTML>

<%
set ei = nothing
%>
