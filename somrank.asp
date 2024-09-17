<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

dim datestr
datestr = trim(request("date"))

dim skt
set skt = server.createobject("easymail.Stakeout")
skt.Load_SendOut_Statistic datestr

allnum = skt.SendOut_Statistic_Count

dim themax
if allnum > pageline then
	themax = pageline
else
	themax = allnum
end if

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim sysinfo
	if trim(request("startsom")) = "1" then
		set sysinfo = server.createobject("easymail.sysinfo")
		sysinfo.Load

		sysinfo.Enable_SendOutMonitor = true

		sysinfo.Save
		set sysinfo = nothing
		set skt = nothing

		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("somrank.asp?date=" & datestr)
	end if

	dim mt
	set mt = server.createobject("easymail.WMethod")

	set app_em = Application("em")

	dim ei
	set ei = server.createobject("easymail.UserWorkTimer")
	ei.Load_Templet

	i = 0
	do while i <= themax
		if trim(request("check" & i)) <> "" then
			set_user = trim(request("check" & i))
			ei.Set_Templet_To_User set_user

			if ei.is_update_disabled_user = true then
				if mt.Is_Disabled_User(set_user) = false then
					if ei.disabled_user_over = "1" or Len(ei.disabled_user_over) = 8 then
						app_em.ForbidUserByName set_user, true
					end if
				else
					if ei.disabled_user_over = "0" then
						app_em.ForbidUserByName set_user, false
					end if
				end if
			end if

			if ei.is_update_limitout = true then
				if mt.Is_Limitout_User(set_user) = false then
					if ei.limitout_over = "1" or Len(ei.limitout_over) = 8 then
						app_em.SetLimitOut set_user, true
					end if
				else
					if ei.limitout_over = "0" then
						app_em.SetLimitOut set_user, false
					end if
				end if
			end if
		end if 

	    i = i + 1
	loop

	set ei = nothing
	set app_em = nothing
	set mt = nothing
	set skt = nothing

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("somrank.asp?date=" & datestr & "&page=" & trim(request("page")))
end if

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

gourl = "somrank.asp?page=" & page & "&date=" & datestr
%>

<HTML>
<HEAD>
<%=s_lang_meta %>
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>

<STYLE type=text/css>
<!--
.EX_TITLE {
	border-left:1px <%=MY_COLOR_1 %> solid;
	border-right:1px <%=MY_COLOR_1 %> solid;
	border-bottom:1px <%=MY_COLOR_1 %> solid;
	BACKGROUND-COLOR: #F8F8D2;
}

.EX_TITLE_FONT {
	FONT-WEIGHT: bold;
	COLOR: #666666;
}
-->
</STYLE>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
<%
if skt.Get_Statistic_Date_Count > 0 and Len(datestr) > 0 then
	Response.Write "var isshow = true;"
else
	Response.Write "var isshow = false;"
end if
%>

function checkall(check) {
	var i = 0;
	var theObj;

	for(; i<<%=themax %>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			theObj.checked = check;
	}
}

function ischeck() {
	var i = 0;
	var theObj;

	for(; i<<%=themax %>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}

function allcheck_onclick() {
	if (document.f1.allcheck.checked == true)
		checkall(true);
	else
		checkall(false);
}

function selectpage_onchange()
{
	location.href = "somrank.asp?<%=getGRSN() & "&date=" & datestr %>&page=" + document.f1.page.value;
}

function headmessage() {
	var theObj;
	theObj = eval("document.all(\"headMessage\")");

	if (isshow == false)
	{
		var instr = "<%
if skt.Get_Statistic_Date_Count > 0 then
	i = skt.Get_Statistic_Date_Count - 1

	do while i >= 0
		Response.Write "<input type='checkbox' name='ckdate" & i & "' value='" & skt.Get_Statistic_Date(i) & "'"

		if InStr(datestr, skt.Get_Statistic_Date(i)) > 0 then
			Response.Write " checked"
		end if

		Response.Write ">" & Show_Som_Date(skt.Get_Statistic_Date(i)) & "&nbsp;&nbsp;&nbsp;"
		i = i - 1
	loop

	if skt.Get_Statistic_Date_Count > 0 then
		Response.Write "&nbsp;&nbsp;<input type='button' value='" & s_lang_0210 & "' onclick='javascript:msearch();' class='sbttn'>"
	end if
else
	set sysinfo = server.createobject("easymail.sysinfo")
	sysinfo.Load

	if sysinfo.Enable_SendOutMonitor = false then
		Response.Write "&nbsp;<input type='button' value='" & s_lang_0211 & "' onclick='javascript:start_som();' class='sbttn'>"
	else
		Response.Write "&nbsp;[" & s_lang_0212 & "]"
	end if

	set sysinfo = nothing
end if
%>";
		instr = "<table width='95%' align='center' border='0' bgcolor='<%=MY_COLOR_2 %>' cellspacing='0' style='border-left:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;<% if isMSIE = true then Response.Write " word-break: break-all;" %>'><tr><td align='center' bgcolor='<%=MY_COLOR_2 %>' style='border-bottom:1px <%=MY_COLOR_1 %> solid;'><a href='javascript:headmessage()'><%=s_lang_0213 %></a></td></tr><tr><td bgcolor='<%=MY_COLOR_3 %>'>" + instr + "</td></tr></table>"
		theObj.innerHTML = instr;
		isshow = true;
	}
	else
	{
		theObj.innerHTML = "<table width='95%' align='center' border='0' bgcolor='<%=MY_COLOR_2 %>' cellspacing='0' style='border-left:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;'><tr><td align='center' bgcolor='<%=MY_COLOR_2 %>'><a href='javascript:headmessage()'><%=s_lang_0213 %></a></td></tr></table>";
		isshow = false;
	}
}

function window_onload() {
	headmessage();
}

function msearch() {
	if (is_check_date() == true)
		location.href = "somrank.asp?<%=getGRSN() %>&date=" + get_check_date();
	else
		alert("<%=s_lang_0214 %>.");
}

function is_check_date() {
	var i = 0;
	var theObj;

	for(; i<<%=skt.Get_Statistic_Date_Count %>; i++)
	{
		theObj = eval("document.f1.ckdate" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}

function get_check_date() {
	var retstr = "";
	var i = 0;
	var theObj;

	for(; i<<%=skt.Get_Statistic_Date_Count %>; i++)
	{
		theObj = eval("document.f1.ckdate" + i);

		if (theObj != null)
		{
			if (theObj.checked == true)
			{
				if (retstr.length == 0)
					retstr = theObj.value;
				else
					retstr = retstr + "," + theObj.value;
			}
		}
	}

	return retstr;
}

function set_sys_tmp() {
	if (ischeck() == true)
		document.f1.submit();
}

function start_som() {
	document.f1.startsom.value = "1";
	document.f1.submit();
}

function findIt() {
	if (document.f1.searchstr.value != "")
		findInPage(document.f1.searchstr.value);
}

var DOM = (document.getElementById) ? 1 : 0;
var NS4 = (document.layers) ? 1 : 0;
var IE4 = 0;
if (document.all)
{
	IE4 = 1;
	DOM = 0;
}

var win = window;   
var n   = 0;

function findInPage(str) {
var txt, i, found;

if (str == "")
	return false;

if (DOM)
{
	win.find(str, false, true);
	return true;
}

if (NS4) {
	if (!win.find(str))
		while(win.find(str, false, true))
			n++;
	else
		n++;

	if (n == 0)
		alert("<%=s_lang_0215 %>.");
}

if (IE4) {
	txt = win.document.body.createTextRange();

	for (i = 0; i <= n && (found = txt.findText(str)) != false; i++) {
		txt.moveStart("character", 1);
		txt.moveEnd("textedit");
	}

if (found) {
	txt.moveStart("character", -1);
	txt.findText(str);
	txt.select();
	txt.scrollIntoView();
	n++;
}
else {
	if (n > 0) {
		n = 0;
		findInPage(str);
	}
	else
		alert("<%=s_lang_0215 %>.");
	}
}

return false;
}
//-->
</SCRIPT>


<BODY LANGUAGE=javascript onload="return window_onload()">
<FORM ACTION="somrank.asp" METHOD="POST" name="f1">
<input type="hidden" name="date" value="<%=datestr %>">
<input type="hidden" name="startsom">
<br>
<p id="headMessage"></p>
  <table width="95%" border="0" align="center">
    <tr>
      <td width="3%">&nbsp;</td>
      <td width="10%"><b><a href="uwt.asp?<%=getGRSN() %>&gourl=<%=Server.URLEncode(gourl) %>"><%=s_lang_0216 %></a></b></td>
      <td width="23%"><b><a href="javascript:set_sys_tmp();"><%=s_lang_0217 %></a></b></td>
      <td width="15%">
<%
if page > 0 then
	Response.Write "<a href=""somrank.asp?" & getGRSN() & "&date=" & datestr & "&page=" & page - 1 & """><img src='images\prep.gif' border='0' align='absmiddle'></a>&nbsp;"
else
	Response.Write "<img src='images\gprep.gif' border='0' align='absmiddle'>&nbsp;"
end if
%>
<select name="page" id="page" class="drpdwn" size="1" LANGUAGE=javascript onchange="selectpage_onchange()">
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
	Response.Write "&nbsp;<a href=""somrank.asp?" & getGRSN() &  "&date=" & datestr & "&page=" & page + 1 & """><img src='images\nextp.gif' border='0' align='absmiddle'></a>"
else
	Response.Write "&nbsp;<img src='images\gnextp.gif' border='0' align='absmiddle'>"
end if
%>
	<td width="49%" align="right"><input type="text" name="searchstr" class="textbox" size="10">&nbsp;<input type="button" value="<%=s_lang_0218 %>" onclick="javascript:findIt();" class="sbttn">&nbsp;&nbsp;&nbsp;&nbsp;
	<a href="right.asp?<%=getGRSN() %>"><b><%=s_lang_return %></b></a>&nbsp;&nbsp;&nbsp;&nbsp;
	<font class="s"><font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0203 %></b></font></font>&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
  </table>
<br>
  <table width="95%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
    <tr bgcolor="<%=MY_COLOR_2 %>" style='border:1px <%=MY_COLOR_1 %> solid;font-size: 9pt;'>
      <td width="5%" height="25" align="center" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()">
      </td>
      <td width="5%" align="center" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0219 %></b></font>
      </td>
      <td width="66%" align="center" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0220 %></b></font>
      </td>
      </td>
      <td width="7%" align="center" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0221 %></b></font>
      </td>
      <td width="10%" align="center" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0222 %></b></font>
      </td>
      <td width="7%" align="center" nowrap style="border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;"> 
        <font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0223 %></b></font>
      </td>
    </tr>
<%
minshowi = page * pageline
i = 0
li = 0

do while i < allnum
	if i >= minshowi and li < pageline then
		skt.Get_SendOut_Statistic_Info i, name, sendnum

		sline = i Mod 2
		if sline = 1 then
		    Response.Write "<tr bgcolor=" & MY_COLOR_3 & ">" & Chr(13)
    	else
		    Response.Write "<tr bgcolor=#f9f9f9>" & Chr(13)
		end if
%>
	<td height="22" align="center" style='border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;'><input type='checkbox' name='check<%=li %>' value='<%=server.htmlencode(name) %>'></td>
	<td align='center' style='border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;'><%=i+1 %></td>
	<td align="center" nowrap style='border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;<% if isMSIE = true then Response.Write " word-break: break-all;" %>'><%=server.htmlencode(name) %>&nbsp;</td>
	<td align="center" nowrap style='border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;'><%=server.htmlencode(sendnum) %></td>
	<td align="center" nowrap style='border-left:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;'><b><a href="listmon.asp?user=<%=Server.URLEncode(name) %>&<%=getGRSN() %>&inout=out&purl=<%=Server.URLEncode(gourl) %>"><%=s_lang_0224 %></a></b>&nbsp;&nbsp;<a href="listmon.asp?user=<%=Server.URLEncode(name) %>&<%=getGRSN() %>&inout=in&purl=<%=Server.URLEncode(gourl) %>"><%=s_lang_0225 %></a></td>
	<td align="center" nowrap style='border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;'><a href="uwtuser.asp?user=<%=Server.URLEncode(name) %>&<%=getGRSN() %>&gourl=<%=Server.URLEncode(gourl) %>"><%=s_lang_0226 %></a></td>
	</tr>
<%
	    li = li + 1
	end if

	name = NULL
	sendnum = NULL

	i = i + 1
loop
%>
</table>
<br><br>
</FORM>
</BODY>
</HTML>

<%
set skt = nothing

function Show_Som_Date(ostr)
	if Len(ostr) = 8 then
		Show_Som_Date = Mid(ostr, 1, 4) & "-" & Mid(ostr, 5, 2) & "-" & Mid(ostr, 7, 2)
	end if
end function
%>
