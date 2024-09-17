<!--#include file="passinc.asp" -->

<%
dim cdomain
cdomain = trim(request("cdomain"))


sortby = 0
if IsNumeric(trim(request("sortby"))) = true then
	sortby = CInt(trim(request("sortby")))
end if


dim ed
set ed = server.createobject("easymail.domain")
ed.Load

dim wem_user
wem_user = Session("wem")

if ed.GetUserManagerDomainCount(wem_user) < 1 then
	set ed = nothing
	response.redirect "noadmin.asp"
end if


searchtext = trim(request("searchtext"))
sm = trim(request("sm"))

if request("page") = "" then
	page = 0
else
	page = CInt(request("page"))
end if

if page < 0 then
	page = 0
end if
%>

<%
dim mam
set mam = server.createobject("easymail.AdminManager")
mam.Load
Enable_Show_User_Memo = mam.Enable_Show_User_Memo
set mam = nothing


dim ei
set ei = Application("em")

'-----------------------------------------

dim sysinfo
set sysinfo = server.createobject("easymail.sysinfo")
sysinfo.Load
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<script type="text/javascript" src="images/sc_left.js"></script>

<SCRIPT LANGUAGE=javascript>
<!--
function mdel() {
	if (ischeck() == true)
	{
		if (confirm("确实要删除吗?") == false)
			return ;

		document.f1.mode.value = "del";
		document.f1.submit();
	}
}

function mforbid() {
	if (ischeck() == true)
	{
		if (confirm("确实要禁用吗?") == false)
			return ;

		document.f1.mode.value = "forbid";
		document.f1.submit();
	}
}

function mclear() {
	if (ischeck() == true)
	{
		document.f1.mode.value = "clear";
		document.f1.submit();
	}
}

function selectsortby_onchange() {
	location.href = "showdomainusers.asp?<%=getGRSN() %>&sortby=" + document.f1.sortby.value + "&cdomain=<%=cdomain %>&page=0&searchtext=<%=searchtext %>&sm=<%=sm %>";
}

function window_onload() {
<%
if sm = "1" then
	Response.Write "	document.f1.sm.value = '1';"
end if
%>
}
//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload()">
<br>
<FORM ACTION="savedomainusers.asp" METHOD="POST" NAME="f1">
  <table width="98%" nowrap align="center" cellspacing="0">
  <tr bgcolor="<%=MY_COLOR_3 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
      <td width="50%" colspan="4" height="28" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	&nbsp;
	<select id="sm" name="sm" class="drpdwn" size="1">
	<option value="0" selected>用户名</option>
	<option value="1">说明</option>
	</select>：<input type="text" name="searchtext" value="<%=searchtext %>" style="border:1px solid #555;">&nbsp;
	<input type="button" onclick="javascript:usersearch()" value=" 搜索 " class="sbttn">
	</td>
	<td align="right" width="50%" colspan="4" nowrap style="border-top:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	分类：<select name="sortby" class="drpdwn" size="1" LANGUAGE=javascript onchange="selectsortby_onchange()">
	<option value=""<% if sortby = 0 then Response.Write " selected" end if %>>---------所有用户----------</option>
	<option value="1"<% if sortby = 1 then Response.Write " selected" end if %>>被禁用户</option>
	<option value="2"<% if sortby = 2 then Response.Write " selected" end if %>>限制外发用户</option>
	<option value="3"<% if sortby = 3 then Response.Write " selected" end if %>>被域监控用户</option>
	<option value="4"<% if sortby = 4 then Response.Write " selected" end if %>>设置期满日期的用户</option>
	<option value="5"<% if sortby = 5 then Response.Write " selected" end if %>>含域名用户</option>
<%
anum = 0
do while anum < 7
	if sortby = anum + 6 then
		Response.Write "<option value=""" & anum + 6 & """ selected>" & getaccessmode(anum) & " 用户</option>"
	else
		Response.Write "<option value=""" & anum + 6 & """>" & getaccessmode(anum) & " 用户</option>"
	end if

	anum = anum + 1
loop
%>
	</select>
	</td>
	</tr>
</table><br>
   <input type="hidden" name="page" value="<%=page %>">
   <input type="hidden" name="mode">
  <table width="98%" border="0" align="center" cellspacing="0">
  <tr bgcolor="<%=MY_COLOR_2 %>">
      <td align="center" nowrap height="28" style="border-left:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><b>选择域名</b></font></td>
	</tr>
    <tr>
      <td bgcolor="<%=MY_COLOR_3 %>" height="22" align="center" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><%
i = 0
allnum = ed.GetUserManagerDomainCount(wem_user)


dim isok
isok = false

if cdomain = "" then
	cdomain = ed.GetUserManagerDomain(wem_user, 0)
end if


do while i < allnum
	cdomainstr = ed.GetUserManagerDomain(wem_user, i)

	if cdomainstr = cdomain then
		isok = true
	end if

	response.write "<a href=""showdomainusers.asp?" & getGRSN() & "&cdomain=" & cdomainstr & """>" & cdomainstr & "</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "

	cdomainstr = NULL

	i = i + 1
loop
%></div>
      </td>
    </tr>
</table><br>
<%
if isok = true then
	ed.GetControlMsgEx cdomain, dc_isshow, dc_maxuser, dc_dmanager, dc_maxsize, dc_allsize, dc_expires
%>
   <input type="hidden" name="cdomain" value="<%=cdomain %>">
  <table width="98%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
  <tr bgcolor="<%=MY_COLOR_2 %>">
      <td colspan="<%
if Enable_Show_User_Memo = false then
	Response.Write "12"
else
	Response.Write "13"
end if
%>" align="center" height="28" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<table width="100%" border="0"><tr><td align="center" height="22"><font class="s" color="<%=MY_COLOR_4 %>"><b><font class="s" color="#FF3333"><%=cdomain %></font>域用户管理</b>&nbsp;&nbsp;(<%=ed.GetUserNumberInDomain(cdomain) %>)</font>
	</td></tr>
<%
if dc_expires > 0 then
	Response.Write "<tr><td>本域期满日期: <font class='s' color='#FF3333'><b>" & getShowExpires(dc_expires) & "</b></font></td></tr>"
end if

Response.Write "<tr><td>本域允许创建的最大用户数: <font class='s' color='#FF3333'><b>" & dc_maxuser & "</b></font></td></tr><tr><td>本域可分配空间总数: <font class='s' color='#FF3333'><b>"

if dc_maxsize = 0 then
	Response.Write "不限</b></font></td></tr>"
else
	Response.Write dc_maxsize
	Response.Write "</b></font>(K)</td></tr>"
end if

if dc_allsize > 0 then
	Response.Write "<tr><td>本域已分配空间数: <font class='s' color='#FF3333'><b>" & dc_allsize & "</b></font>(K)</td></tr>"
end if
%></table>
      </td>
    </tr>
    <tr bgcolor="<%=MY_COLOR_2 %>">
      <td width="4%" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center"><input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()"></div>
      </td>
      <td width="4%" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">序号</div>
      </td>
      <td width="4%" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">修改</div>
      </td>
      <td width="14%" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">用户名</div>
      </td>
      <td width="15%" height="22" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">所在域名</div>
      </td>
      <td width="5%" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">是否<br>禁用</div>
      </td>
      <td width="13%" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">最后登录时间</div>
      </td>
<%
if Enable_Show_User_Memo = false then
%>
      <td width="19%" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;说明&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</div>
      </td>
<%
else
%>
      <td width="10%" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">&nbsp;&nbsp;&nbsp;说明&nbsp;&nbsp;&nbsp;</div>
      </td>
      <td width="9%" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">&nbsp;&nbsp;&nbsp;备注&nbsp;&nbsp;&nbsp;</div>
      </td>
<%
end if
%>
      <td width="12%" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">访问方式</div>
      </td>
      <td width="4%" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">限制<br>外发</div>
      </td>
<%
if sysinfo.enableDomainMonitor = true then
%>
      <td width="4%" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <div align="center">监控</div>
      </td>
<%
end if
%>
      <td width="4%" align="center" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
        <div align="center">期满<br>日期</div>
      </td>
    </tr>
    <%
dc_isshow = NULL
dc_maxuser = NULL
dc_dmanager = NULL
dc_maxsize = NULL
dc_allsize = NULL
dc_expires = NULL


dim isshow

allnum = ei.GetUsersCount
i = 0
showline = 0
startline = -1
li = 0
max_index = 0

do while i < allnum
	ei.GetUserByIndex3 i, name, domain, comment, forbid, lasttime, accessmode, limitout, expiresday, monitor

	isshow = false

	if domain = cdomain then
		if searchtext <> "" then
			if sm <> "1" then
				if InStr(name, searchtext) <> 0 then
					isshow = true
				end if
			else
				if InStr(comment, searchtext) <> 0 then
					isshow = true
				end if
			end if
		else
			isshow = true
		end if
	end if


	if isshow = true then
		if sortby = 1 then
			if forbid = false then
				isshow = false
			end if
		elseif sortby = 2 then
			if limitout = false then
				isshow = false
			end if
		elseif sortby = 3 then
			if monitor = false then
				isshow = false
			end if
		elseif sortby = 4 then
			if expiresday = "" then
				isshow = false
			end if
		elseif sortby = 5 then
			if InStr(name, "@") < 1 then
				isshow = false
			end if
		elseif sortby > 5 then
			if (accessmode + 6) <> sortby then
				isshow = false
			end if
		end if
	end if


	if isshow = true then
		if showline >= page * pageline and li < pageline then
			if startline = -1 then
				startline = i
			end if

			response.write "<tr><td align='center' height='24' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><input type='checkbox' name='check" & i & "'>"
			response.write "</td><td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & showline + 1
			response.write "</td><td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><a href='changepw.asp?" & getGRSN() & "&id=" & i & "&gourl=" & Server.URLEncode("showdomainusers.asp?" & getGRSN() & "&cdomain=" & cdomain & "&page=" & page & "&sm=" & sm & "&searchtext=" & searchtext) & "'><img src='images\edit.gif' border='0'></a>"
			response.write "</td><td align='center' nowrap style='border-bottom:1px " & MY_COLOR_1 & " solid;'><a href='viewreginfo.asp?nm=" & Server.URLEncode(name) & "&" & getGRSN() & "&purl=" & Server.URLEncode("showdomainusers.asp?" & getGRSN() & "&sortby=" & sortby & "&cdomain=" & cdomain & "&page=" & page & "&sm=" & sm & "&searchtext=" & searchtext) & "'>" & name & "</a>"
			response.write "&nbsp;</td><td align='center' nowrap style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & domain

			if forbid = TRUE then
				response.write "&nbsp;</td><td align='center' nowrap style='border-bottom:1px " & MY_COLOR_1 & " solid;'>禁用" 
			else
				response.write "&nbsp;</td><td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>&nbsp;"
			end if
			response.write "</td><td align='center' nowrap style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & lasttime

			response.write "&nbsp;</td><td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & comment

			if Enable_Show_User_Memo = true then
				Response.Write "&nbsp;</td><td align='left' style='border-left:1px " & MY_COLOR_1 & " solid; border-bottom:1px " & MY_COLOR_1 & " solid;'>" & ei.GetUserMemo(name)
				Response.Write "&nbsp;</td><td align='center' nowrap style='border-left:1px " & MY_COLOR_1 & " solid; border-bottom:1px " & MY_COLOR_1 & " solid;'>" & getaccessmode(accessmode)
			else
				Response.Write "&nbsp;</td><td align='center' nowrap style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & getaccessmode(accessmode)
			end if

			response.write "&nbsp;</td><td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>"
			if limitout = true then
				response.write "Yes"
			else
				response.write "&nbsp;"
			end if

			if sysinfo.enableDomainMonitor = true then
				response.write "&nbsp;</td><td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>"
				if monitor = true then
					response.write "Yes"
				else
					response.write "&nbsp;"
				end if
			end if

			response.write "</td><td align='center' nowrap style='border-bottom:1px " & MY_COLOR_1 & " solid;'>"
			if expiresday = "" then
				response.write "&nbsp;"
			else
				response.write "&nbsp;" & getShowExpires(expiresday) & "&nbsp;"
			end if

			response.write "</td></tr>" & Chr(13)

			max_index = i

			li = li + 1
		end if

		showline = showline + 1
	end if


	name = NULL
	domain = NULL
	comment = NULL
	forbid = NULL
	lasttime = NULL
	accessmode = NULL
	limitout = NULL
	expiresday = NULL
	monitor = NULL

	i = i + 1
loop
%> 
    <tr>
      <td colspan="<%
if Enable_Show_User_Memo = false then
	Response.Write "12"
else
	Response.Write "13"
end if
%>" height="40" align="left" bgcolor="#ffffff" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>&nbsp;
		<input type="button" value=" 添加 " onClick="javascript:location.href='domaincreate.asp?<%=getGRSN() %>'" class="sbttn">&nbsp;&nbsp;&nbsp;
		<input type="button" value=" 删除 " onclick="javascript:mdel()" class="sbttn">&nbsp;&nbsp;&nbsp;
		<input type="button" value=" 禁用 " onclick="javascript:mforbid()" class="sbttn">&nbsp;&nbsp;&nbsp;
		<input type="button" value=" 启用 " onclick="javascript:mclear()" class="sbttn">
      </td>
    </tr>
  </table>
<%
allnum = showline

allpage = CInt((allnum - (allnum mod pageline))/ pageline)

if allnum mod pageline <> 0 then
	allpage = allpage + 1
end if

if allpage = 0 then
	allpage = 1
end if



response.write "<table width='98%' border='0'><tr><td align='center' height='32'>(<font color='#FF3333'>" & page + 1 & "</font>/" & allpage & ")&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"

if page - 1 < 0 then
	response.write "<img src='images\gfirstp.gif' border='0' align='absmiddle'>&nbsp;"
	response.write "<img src='images\gprep.gif' border='0' align='absmiddle'>&nbsp;&nbsp;"
else
	response.write "<a href=""showdomainusers.asp?" & getGRSN() & "&sortby=" & sortby & "&cdomain=" & cdomain & "&page=" & 0 & "&sm=" & sm & "&searchtext=" & searchtext & """><img src='images\firstp.gif' border='0' align='absmiddle'></a>&nbsp;"
	response.write "<a href=""showdomainusers.asp?" & getGRSN() & "&sortby=" & sortby & "&cdomain=" & cdomain & "&page=" & page - 1 & "&sm=" & sm & "&searchtext=" & searchtext & """><img src='images\prep.gif' border='0' align='absmiddle'></a>&nbsp;&nbsp;"
end if


response.write "<select name='selectpage' class='drpdwn' size='1' LANGUAGE=javascript onchange='selectpage_onchange()'>"
i = 0

do while i < allpage
	if i <> page then
		response.write "<option value=""" & i & """>" & i + 1 & "</option>"
	else
		response.write "<option value=""" & i & """ selected>" & i + 1 & "</option>"
	end if
	i = i + 1
loop

response.write "</select>&nbsp;"


if ((page+1) * pageline) => allnum then
	response.write "<img src='images\gnextp.gif' border='0' align='absmiddle'>&nbsp;"
else
	response.write "<a href=""showdomainusers.asp?" & getGRSN() & "&sortby=" & sortby & "&cdomain=" & cdomain & "&page=" & page + 1 & "&sm=" & sm & "&searchtext=" & searchtext & """><img src='images\nextp.gif' border='0' align='absmiddle'></a>&nbsp;"
end if

if page + 1 = allpage then
	response.write "<img src='images\gendp.gif' border='0' align='absmiddle'>&nbsp;"
else
	response.write "<a href=""showdomainusers.asp?" & getGRSN() & "&sortby=" & sortby & "&cdomain=" & cdomain & "&page=" & allpage - 1 & "&sm=" & sm & "&searchtext=" & searchtext & """><img src='images\endp.gif' border='0' align='absmiddle'></a>"
end if

response.write "</td></tr></table>"


end if
%>
  </FORM>
<br>
</BODY>

<SCRIPT LANGUAGE=javascript>
<!--
function allcheck_onclick() {
	if (document.f1.allcheck.checked == true)
		checkall(true);
	else
		checkall(false);
}

function checkall(check) {
<%
if startline = -1 then
	Response.Write "var i = 0;" & Chr(13)
	Response.Write "var maxnum = " & max_index + 1 & ";" & Chr(13)
else
	Response.Write "var i = " & startline & ";" & Chr(13)
	Response.Write "var maxnum = " & max_index + 1 & ";" & Chr(13)
end if
%>
	var theObj;

	for(; i < maxnum; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			theObj.checked = check;
	}
}

function ischeck() {
<%
if startline = -1 then
	Response.Write "var i = 0;" & Chr(13)
	Response.Write "var maxnum = " & max_index + 1 & ";" & Chr(13)
else
	Response.Write "var i = " & startline & ";" & Chr(13)
	Response.Write "var maxnum = " & max_index + 1 & ";" & Chr(13)
end if
%>
	var theObj;

	for(; i < maxnum; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}

function usersearch(){
	location.href = "showdomainusers.asp?<%=getGRSN() %>&sortby=<%=sortby %>&cdomain=<%=cdomain %>&searchtext=" + document.f1.searchtext.value + "&sm=" + document.f1.sm.value;
}

function selectpage_onchange()
{
	location.href = "showdomainusers.asp?<%=getGRSN() & "&sortby=" & sortby & "&cdomain=" & cdomain & "&sm=" & sm & "&searchtext=" & searchtext %>&page=" + document.f1.selectpage.value;
}
//-->
</SCRIPT>

</HTML>

<%
set ei = nothing
set ed = nothing
set sysinfo = nothing


function getShowExpires(exday)
	if exday >= 19720101 then
		getShowExpires = Mid(Cstr(exday), 1, 4) & "-" & Mid(Cstr(exday), 5, 2) & "-" & Mid(Cstr(exday), 7, 2)
	end if
end function

function getaccessmode(amode)
	if amode = 0 then
		getaccessmode = "http/smtp/pop3,imap4"
	elseif amode = 1 then
		getaccessmode = "smtp/pop3,imap4"
	elseif amode = 2 then
		getaccessmode = "http/smtp"
	elseif amode = 3 then
		getaccessmode = "http/pop3,imap4"
	elseif amode = 4 then
		getaccessmode = "http"
	elseif amode = 5 then
		getaccessmode = "smtp"
	elseif amode = 6 then
		getaccessmode = "pop3,imap4"
	end if
end function
%>
