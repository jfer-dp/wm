<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

dim ei
set ei = server.createobject("easymail.CalOptions")
ei.Load Session("wem")

returl = trim(request("returl"))

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ei.StartWeekDay = CLng(trim(request("StartWeekDay")))

	if trim(request("ShowFeasts")) = "" then
		ei.ShowFeasts = false
	else
		ei.ShowFeasts = true
	end if

	if trim(request("ShowDayExt")) = "" then
		ei.ShowDayExt = false
	else
		ei.ShowDayExt = true
	end if

	if trim(request("ShowNLFeasts")) = "" then
		ei.ShowNLFeasts = false
	else
		ei.ShowNLFeasts = true
	end if

	if trim(request("ShowNL")) = "" then
		ei.ShowNL = false
	else
		ei.ShowNL = true
	end if

	if trim(request("Show24Hour")) = "" then
		ei.Show24Hour = false
	else
		ei.Show24Hour = true
	end if

	ei.MyCalendarViewState = CLng(trim(request("MyCalendarViewState")))
	ei.EventShareDefault = CLng(trim(request("EventShareDefault")))
	ei.TaskShareDefault = CLng(trim(request("TaskShareDefault")))


	ei.RemoveAllFriends

	dim msg
	msg = trim(request("allmsgs"))

	if Len(msg) > 0 then
		dim item
		dim ss
		dim se
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(9))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				ei.AddFriend item
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	isok = ei.Save()

	set ei = nothing

	if isok = true then
		if Len(returl) > 3 then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	else
		if Len(returl) > 3 then
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(returl)
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("cal_index.asp")
		end if
	end if
end if
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<STYLE type=text/css>
<!--
html {overflow: scroll; overflow-x: hidden; overflow-y: auto !important;}
body {font-family:<%=s_lang_font %>; font-size:9pt;color:#000000;margin-top:5px;margin-left:10px;margin-right:10px;margin-bottom:2px;background-color:#ffffff}
.tl {height:24px; text-align:right; border-bottom:1px #8CA5B5 solid;}
.trr {text-align:left; border-bottom:1px #8CA5B5 solid;}
.sbttn {font-family:<%=s_lang_font %>; font-size:9pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5;color:#000066;height:19px;text-decoration:none;cursor:pointer}
-->
</STYLE>
</head>

<script type="text/javascript">
<!--
function window_onload() {
<%
Response.Write "document.f1.StartWeekDay.value = """ & ei.StartWeekDay & """;" & Chr(13)
Response.Write "document.f1.MyCalendarViewState" & ei.MyCalendarViewState & ".checked = true;" & Chr(13)
Response.Write "document.f1.EventShareDefault" & ei.EventShareDefault & ".checked = true;" & Chr(13)
Response.Write "document.f1.TaskShareDefault" & ei.TaskShareDefault & ".checked = true;"
%>
}

function goback()
{
	if (document.f1.returl.value.length < 3)
		history.back();
	else
		location.href=document.f1.returl.value;
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
	document.f1.action = "cal_set.asp";
	document.f1.method = "POST";
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
			oOption.text = document.f1.addmsg.value;
			oOption.value = document.f1.addmsg.value;
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
}

function haveit()
{
	var tempstr = document.f1.addmsg.value;

	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].value == tempstr)
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
//-->
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
<br>
<form name="f1">
<input type="hidden" name="returl" value="<%=returl %>">
<input type="hidden" name="allmsgs">

<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td colspan="2" style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
效率手册选项
</td></tr>
<tr><td colspan=2 class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" cellspacing="0" cellpadding="0">
	<tr> 
	<td align="center">
	<table width="100%" border="0" align="center" cellspacing="0" style="border-top:1px #8CA5B5 solid;">
		<tr>
		<td colspan=2 height="24" valign=center align=left bgcolor="#DBEAF5" style='border-bottom:1px #8CA5B5 solid;'> 
		&nbsp;<b>常规选项</b>
		</td>
		</tr>

		<tr>
		<td valign=center width="40%" class="tl"> 
		每周开始的日期&nbsp;(按月查看模式)<%=s_lang_mh %>
		</td>
		<td class="trr">
		<select name="StartWeekDay" class="drpdwn">
<%
i = 0

do while i < 7
	Response.Write "<option value=""" & i & """>" & getWeekName3(i) & "</option>" & Chr(13)

	i = i + 1
loop
%>
		</select>
		</td>
		</tr>

		<tr>
		<td valign=center class="tl"> 
		显示日期扩展信息<%=s_lang_mh %>
		</td>
		<td class="trr">
		<input type="checkbox" name="ShowDayExt" value="checkbox" <% if ei.ShowDayExt = true then response.write "checked"%>>
		</td>
		</tr>

		<tr>
		<td valign=center class="tl"> 
		显示节日<%=s_lang_mh %>
		</td>
		<td class="trr">
		<input type="checkbox" name="ShowFeasts" value="checkbox" <% if ei.ShowFeasts = true then response.write "checked"%>>
		</td>
		</tr>

		<tr>
		<td valign=center class="tl"> 
		显示农历节日<%=s_lang_mh %>
		</td>
		<td class="trr">
		<input type="checkbox" name="ShowNLFeasts" value="checkbox" <% if ei.ShowNLFeasts = true then response.write "checked"%>>
		</td>
		</tr>

		<tr>
		<td valign=center class="tl"> 
		显示农历<%=s_lang_mh %>
		</td>
		<td class="trr">
		<input type="checkbox" name="ShowNL" value="checkbox" <% if ei.ShowNL = true then response.write "checked"%>>
		</td>
		</tr>

		<tr>
		<td valign=center class="tl"> 
		按24小时制显示时间<%=s_lang_mh %>
		</td>
		<td class="trr">
		<input type="checkbox" name="Show24Hour" value="checkbox" <% if ei.Show24Hour = true then response.write "checked"%>>
		</td>
		</tr>

		<tr>
		<td colspan=2 height="24" valign=center align=left bgcolor="<%=MY_COLOR_2 %>" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'> 
		&nbsp;<b>共享选项</b>
		</td>
		</tr>

		<tr>
		<td valign=center class="tl"> 
		效率手册查看<%=s_lang_mh %>
		</td>
		<td class="trr">
		<input type=radio value="0" name="MyCalendarViewState" id="MyCalendarViewState0">不允许他人查看我的效率手册
		<br>
		<input type=radio value="1" name="MyCalendarViewState" id="MyCalendarViewState1">我的朋友可以查看我的效率手册
		<br>
		<input type=radio value="2" name="MyCalendarViewState" id="MyCalendarViewState2">任何人都可以查看我的效率手册
		</td>
		</tr>

		<tr>
		<td height=22 valign=bottom colspan=2 align=center>
		<b>我的朋友帐号列表</b>&nbsp;<font color="#444444">(在您朋友名单列表的任何人都可以查看您的效率手册)</font><br>
		</td>
		</tr>

		<tr>
		<td colspan=2 valign=center align=left style="height:24px; border-bottom:1px #8CA5B5 solid;">
<table>
  <tr valign=top> 
	<td>
	&nbsp;<input maxlength=100 size=30 name="addmsg" class='n_textbox' onkeydown="goent()">
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
	<select class="drpdwn" style="WIDTH: 250px" multiple size=6 name=listall width="230">
<%
i = 0
allnum = ei.CountFriends

do while i < allnum
	tmsg = server.htmlencode(ei.GetFriend(i))
	Response.Write "<option value=""" & tmsg & """>" & tmsg & "</option>" & Chr(13)

	tmsg = NULL

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
		<td valign=center class="tl"> 
		活动共享(默认)<%=s_lang_mh %>
		</td>
		<td class="trr">
		<input type=radio value="0" name="EventShareDefault" id="EventShareDefault0"><b>私人</b>&nbsp;&nbsp;<font color="#444444">(其他人无法看到该活动)</font><br>
		<input type=radio value="1" name="EventShareDefault" id="EventShareDefault1"><b>显示繁忙状态</b>&nbsp;&nbsp;<font color="#444444">(他人可看到该活动的日期，但无法查阅详细内容)</font><br>
		<input type=radio value="2" name="EventShareDefault" id="EventShareDefault2"><b>公共</b>&nbsp;&nbsp;<font color="#444444">(他人可以看到活动安排的细节)</font>
		</td>
		</tr>

		<tr>
		<td valign=center class="tl"> 
		待办事项共享(默认)<%=s_lang_mh %>
		</td>
		<td class="trr">
		<input type=radio value="0" name="TaskShareDefault" id="TaskShareDefault0"><b>私人</b>&nbsp;&nbsp;<font color="#444444">(其他人无法看到该待办事项)</font><br>
		<input type=radio value="1" name="TaskShareDefault" id="TaskShareDefault1"><b>公共</b>&nbsp;&nbsp;<font color="#444444">(他人可以看到待办事项的细节)</font>
		</td>
		</tr>
	</table>
	</td></tr>
	</table>
</td></tr>

<tr><td colspan="2" align="left" style="background-color:white; padding-right:16px; padding-top:12px; padding-bottom:10px;">
<a class='wwm_btnDownload btn_blue' href="javascript:goback();"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();">保存</a>
</td></tr>
</table>
</form>
</body>
</html>

<%
set ei = nothing


function getWeekName3(wknum)
	temp_wk_str = ""

	if wknum = "0" then
		temp_wk_str = "星期日"
	elseif wknum = "1" then
		temp_wk_str = "星期一"
	elseif wknum = "2" then
		temp_wk_str = "星期二"
	elseif wknum = "3" then
		temp_wk_str = "星期三"
	elseif wknum = "4" then
		temp_wk_str = "星期四"
	elseif wknum = "5" then
		temp_wk_str = "星期五"
	elseif wknum = "6" then
		temp_wk_str = "星期六"
	end if

	getWeekName3 = temp_wk_str
end function
%>
