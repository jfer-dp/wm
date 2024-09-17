<!--#include file="passinc.asp" -->

<%
calid = trim(request("calid"))
sortby = trim(request("sortby"))

msgname = trim(request("msgname"))
if Len(msgname) < 1 then
	msgname = Session("wem")
end if

preturl = trim(request("preturl"))
ppreturl = trim(request("ppreturl"))

purl = "preturl=" & Server.URLEncode(preturl) & "&ppreturl=" & Server.URLEncode(ppreturl)

returl = "cal_showguest.asp?" & getGRSN() & "&calid=" & calid & "&msgname=" & Server.URLEncode(msgname)

if Len(sortby) > 0 then
	if IsNumeric(sortby) = true then
		sortby = CLng(sortby)

		if sortby < 0 then
			sortby = 0
		end if
	else
		sortby = 0
	end if
else
	sortby = 0
end if

sortmode = trim(request("sortmode"))
if sortmode <> "1" then
	sortmode = "0"
end if

dim ecal
set ecal = server.createobject("easymail.CalendarExtend")

ecal.SortListMode = sortby

ecal.Load Session("wem"), calid
%>

<html>
<head>
<TITLE>WinWebMail</TITLE>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<LINK href="images\hwem.css" rel=stylesheet>
<STYLE type=text/css>
<!--
a:hover {text-decoration:underline;}

.mjNoLine {
 	text-decoration: none; 
}
.mjRemove {
	text-decoration: line-through;
}
.mjLinkLeft {
	color: #447172;
 	text-decoration: none; 
	CURSOR: pointer;
}
.mjLink {
	color: #002f72;
 	text-decoration: none; 
	CURSOR: pointer;
}
.calendar_dayname {
	BORDER-TOP: #ffffc0 7px solid;
	BORDER-LEFT: #ffffc0 5px solid;
	BORDER-BOTTOM: #ffffc0 3px solid;
	FONT-WEIGHT: normal;
	color: #202020;
	BACKGROUND-COLOR: #ffffc0;
}
.mjEL {
	font-size: 9pt;
	color: #447172;
 	text-decoration: none; 
	CURSOR: pointer;
}
-->
</STYLE>
</head>

<script language="JavaScript">
<!--
function showtab4(p_page)
{
	if (p_page < 0)
		location.href = "<%=returl %>&page=" + document.getElementById("page").value + "&sortby=<%=sortby %>&sortmode=<%=sortmode %>&<%=purl %>";
	else
		location.href = "<%=returl %>&page=" + p_page + "&sortby=<%=sortby %>&sortmode=<%=sortmode %>&<%=purl %>";
}

function selectpage_onchange()
{
	showtab4(-1);
}

function myback()
{
<%
if Len(preturl) > 0 then
%>
	location.href = "<%=preturl %>&returl=<%=Server.URLEncode(ppreturl) %>";
<%
else
%>
	location.href = "cal_index.asp?<%=getGRSN() %>";
<%
end if
%>
}
//-->
</script>

<body>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="31" width="85%" align="left" bgcolor="#ffffff" style="border-left:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
&nbsp;&nbsp;<b>客人名单</b>&nbsp;&nbsp;[<a href="javascript:moreguest()" class=mjNoLine>邀请其他客人</a>]</font>
    </td>
    <td align="center" bgcolor="#ffffff" style="border-right:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
	<input type="button" value="返回" style="WIDTH: 50px" onclick="javascript:myback();" class="Bsbttn">
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="80%" style="border-left:3px #ffffff solid; border-right:3px #ffffff solid;">
<%
	vsd = trim(request("vsd"))
	if Len(vsd) < 1 then
		vsd = "0"
	end if

	if vsd = "1" then
		ecal.SelFuture
	elseif vsd = "2" then
		ecal.SelPast
	end if

	vmd = trim(request("vmd"))
	if Len(vmd) > 0 then
		ecal.Search "", false, false, CLng(vmd)
	end if

	allnum = ecal.Count
	allnb = ecal.Count

	if trim(request("page")) = "" then
		page = 0
	else
		page = CInt(trim(request("page")))
	end if

	if page < 0 then
		page = 0
	end if

	allpage = CInt((allnum - (allnum mod pageline))/ pageline)

	if allnum mod pageline <> 0 then
		allpage = allpage + 1
	end if

	if page >= allpage then
		page = allpage - 1
	end if

	if page < 0 then
		page = 0
	end if

	if allpage = 0 then
		allpage = 1
	end if

	add_bakurl = "&page=" & page & "&vsd=" & vsd & "&vmd=" & vmd & "&sortby=" & sortby & "&sortmode=" & sortmode
%>
<table width=100% border=0 cellspacing=0 cellpadding=0>
<tr><td class=calendar_dayname>
<input type="button" value="删除点选项目" style="WIDTH: 92px" onclick="javascript:delmulevent()" class="sbttn">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%
if page > 0 then
	Response.Write "<a href='javascript:showtab4(" & page - 1 & ")'><img src='images\prep.gif' border='0' align='absmiddle'></a>&nbsp;"
else
	Response.Write "<img src='images\gprep.gif' border='0' align='absmiddle'>&nbsp;"
end if
%>
<select id="page" name="page" class="drpdwn" size="1" LANGUAGE=javascript onchange="selectpage_onchange()">
<%
i = 0

do while i < allpage
	if i <> page then
		Response.Write "<option value=""" & i & """>" & i + 1 & "</option>"
	else
		Response.Write "<option value=""" & i & """ selected>" & i + 1 & "</option>"
	end if
	i = i + 1
loop
%></select>
<%
if page < allpage - 1 then
	Response.Write "&nbsp;<a href='javascript:showtab4(" & page + 1 & ")'><img src='images\nextp.gif' border='0' align='absmiddle'></a>"
else
	Response.Write "&nbsp;<img src='images\gnextp.gif' border='0' align='absmiddle'>"
end if
%>
</td></tr>
<tr><td>
<form method="post" action="cal_del.asp" name="f2">
<input type="hidden" name="returl" value="<%=returl & add_bakurl %>">
<input type="hidden" name="purl" value="<%=purl %>">
<input type="hidden" name="calid" value="<%=calid %>">
<input type="hidden" name="calmode" value="8">
<table width="100%" border="0" align="center" cellspacing="0" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
  <tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
    <td width="3%" align="center" height="25" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()"></td>
	<td width="12%" nowrap align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><%
if sortby = "3" then
	if sortmode = "0" then
		Response.Write "<a href=""javascript:setsort('3', '1')"">状态</a>&nbsp;<a href=""javascript:setsort('3', '1')""><img src='images\arrow_down.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "<a href=""javascript:setsort('3', '0')"">状态</a>&nbsp;<a href=""javascript:setsort('3', '0')""><img src='images\arrow_up.gif' border='0' align='absmiddle'></a>"
	end if
else
	Response.Write "<a href=""javascript:setsort('3', '0')"">状态</a>"
end if
%></font></td>
	<td width="25%" align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><%
if sortby = "0" or sortby = "2" then
	if sortby = "0" then
		if sortmode = "0" then
			Response.Write "<a href=""javascript:setsort('0', '1')"">帐号<img src='images\arrow_down.gif' border='0' align='absmiddle'></a>&nbsp;(<a href=""javascript:setsort('2', '0')"">姓名</a>)"
		else
			Response.Write "<a href=""javascript:setsort('0', '0')"">帐号<img src='images\arrow_up.gif' border='0' align='absmiddle'></a>&nbsp;(<a href=""javascript:setsort('2', '0')"">姓名</a>)"
		end if
	else
		if sortmode = "0" then
			Response.Write "<a href=""javascript:setsort('0', '1')"">帐号</a>&nbsp;(<a href=""javascript:setsort('2', '1')"">姓名<img src='images\arrow_down.gif' border='0' align='absmiddle'></a>)"
		else
			Response.Write "<a href=""javascript:setsort('0', '0')"">帐号</a>&nbsp;(<a href=""javascript:setsort('2', '0')"">姓名<img src='images\arrow_up.gif' border='0' align='absmiddle'></a>)"
		end if
	end if
else
	Response.Write "<a href=""javascript:setsort('0', '0')"">帐号</a>&nbsp;(<a href=""javascript:setsort('2', '0')"">姓名</a>)"
end if
%></font></td>
    <td width="20%" align="center" nowrap bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><%
if sortby = "1" then
	if sortmode = "0" then
		Response.Write "<a href=""javascript:setsort('1', '1')"">电子邮件</a>&nbsp;<a href=""javascript:setsort('1', '1')""><img src='images\arrow_down.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "<a href=""javascript:setsort('1', '0')"">电子邮件</a>&nbsp;<a href=""javascript:setsort('1', '0')""><img src='images\arrow_up.gif' border='0' align='absmiddle'></a>"
	end if
else
	Response.Write "<a href=""javascript:setsort('1', '0')"">电子邮件</a>"
end if
%></font></td>
    <td width="5%" align="center" nowrap bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>"><%
if sortby = "4" then
	if sortmode = "0" then
		Response.Write "<a href=""javascript:setsort('4', '1')"">客人</a>&nbsp;<a href=""javascript:setsort('4', '1')""><img src='images\arrow_down.gif' border='0' align='absmiddle'></a>"
	else
		Response.Write "<a href=""javascript:setsort('4', '0')"">客人</a>&nbsp;<a href=""javascript:setsort('4', '0')""><img src='images\arrow_up.gif' border='0' align='absmiddle'></a>"
	end if
else
	Response.Write "<a href=""javascript:setsort('4', '0')"">客人</a>"
end if
%></font></td>
	<td width="25%" align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>">备注</font></td>
	<td width="5%" nowrap align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>">编辑</font></td>
	<td width="5%" nowrap align="center" bgcolor="<%=MY_COLOR_2 %>" style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"><font class="s" color="<%=MY_COLOR_4 %>">删除</font></td>
  </tr>
<%
si = 0
i = 0

do while i < ((page + 1) * pageline) and i < allnum
	isHost = false

	if i >= page * pageline then
		if sortmode = "1" then
			showi = i
		else
			showi = allnum - i - 1
		end if

		ecal.MoveTo showi

		Response.Write "  <tr>"
		Response.Write "	<td align='center' height='22' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>"
		if LCase(ecal.ce_email) <> LCase(Session("mail")) then
			Response.Write "<input type='checkbox' name='check" & si & "' value='" & server.htmlencode(ecal.ce_email) & "'>"
		else
			isHost = true
			Response.Write "&nbsp;"
		end if
		Response.Write "</td>"

		Response.Write "    <td align='left' nowrap style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & get_show_joinmode(ecal.ce_join) & "</td>"
		Response.Write "    <td align='left' nowrap style='border-bottom:1px " & MY_COLOR_1 & " solid;'>"
		if ecal.ce_askRemove = true then
			Response.Write "<font class=mjRemove>"
		end if
		Response.Write server.htmlencode(ecal.ce_username)

		if Len(ecal.ce_myname) > 0 then
			Response.Write "&nbsp;" & server.htmlencode("(" & ecal.ce_myname & ")")
		end if
		Response.Write "</td>"

		Response.Write "    <td align='left' nowrap style='border-bottom:1px " & MY_COLOR_1 & " solid;'>"
		if ecal.ce_askRemove = true then
			Response.Write "<font class=mjRemove>"
		end if
		Response.Write server.htmlencode(ecal.ce_email) & "</td>"

		Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & ecal.ce_withGuest & "</td>"

		Response.Write "    <td align='left' nowrap width='1%' nowrap style='border-bottom:1px " & MY_COLOR_1 & " solid;'>" & server.htmlencode(ecal.ce_remark) & "&nbsp;</td>"

		Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><a href=""javascript:editguest('" & Server.URLEncode(ecal.ce_email) & "')""><img src='images\edit.gif' border='0' alt='编辑'></a></td>"

		if isHost = false then
			Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'><a href=""javascript:delguest('" & Server.URLEncode(ecal.ce_email) & "')""><img src='images\del.gif' border='0' alt='删除'></a></td>"
		else
			Response.Write "    <td align='center' style='border-bottom:1px " & MY_COLOR_1 & " solid;'>&nbsp;</td>"
		end if

		Response.Write "  </tr>" & chr(13)

		si = si + 1
	end if

	i = i + 1
loop
%>
</table>
</form>
</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
	<td>&nbsp;</td>
  </tr>
</table>
</body>


<SCRIPT language=javascript>
<!--
function showevent(evid)
{
	location.href = "cal_showinvite.asp?<%=getGRSN() %>&calid=" + evid + "&returl=<%=Server.URLEncode(returl & add_bakurl) %>";
}

function setsort(p_sortby, p_sortmode)
{
	location.href = "<%=returl %>&page=" + document.getElementById("page").value + "&sortby=" + p_sortby + "&sortmode=" + p_sortmode + "&<%=purl %>";
}

function allcheck_onclick() {
	if (document.f2.allcheck.checked == true)
		checkall(true);
	else
		checkall(false);
}

function checkall(check) {
	var i = 0;
	var theObj;

	for(; i<<%
if pageline < allnb then
	Response.Write pageline
else
	Response.Write allnb
end if
%>; i++)
	{
		theObj = eval("document.f2.check" + i);

		if (theObj != null)
			theObj.checked = check;
	}
}

function ischeck() {
	var i = 0;
	var theObj;

	for(; i<<%
if pageline < allnb then
	Response.Write pageline
else
	Response.Write allnb
end if
%>; i++)
	{
		theObj = eval("document.f2.check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}

function delmulevent()
{
	if (ischeck() == true)
	{
		if (confirm("确实要删除吗?") == false)
			return ;

		document.f2.submit();
	}
}

function editguest(email)
{
	location.href = "cal_editguest.asp?<%=getGRSN() %>&calid=<%=calid %>&email=" + email + "&returl=<%=Server.URLEncode(returl & add_bakurl) %>&<%=purl %>";
}

function delguest(email)
{
	if (confirm("确实要删除吗?") == false)
		return ;

	location.href = "cal_del.asp?<%=getGRSN() %>&calmode=9&calid=<%=calid %>&email=" + email + "&returl=<%=Server.URLEncode(returl & add_bakurl) %>&purl=<%=Server.URLEncode(purl) %>";
}

function moreguest()
{
	location.href = "cal_moreguest.asp?<%=getGRSN() %>&calid=<%=calid %>&msgname=<%=Server.URLEncode(msgname) %>&preturl=<%=Server.URLEncode(preturl) %>&ppreturl=<%=Server.URLEncode(ppreturl) %>";
}
//-->
</SCRIPT>
</html>

<%
set ecal = nothing


function get_show_joinmode(join_mode)
	if join_mode = -1 then
		get_show_joinmode = "<img src='images/cal/d.gif' border=0 alt='婉言谢绝'>&nbsp;婉言谢绝"
	elseif join_mode = 0 then
		get_show_joinmode = "<img src='images/cal/u.gif' border=0 alt='未决定的'>&nbsp;未决定"
	elseif join_mode = 1 then
		get_show_joinmode = "<img src='images/cal/a.gif' border=0 alt='参加'>&nbsp;参加"
	end if
end function
%>
