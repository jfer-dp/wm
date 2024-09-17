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
	ei.RemoveAllFavorites

	dim msg
	msg = trim(request("allmsgs"))

	if Len(msg) > 0 or Len(addname) > 0 then
		dim item
		dim ss
		dim se
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(9))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				ei.AddFavorite item
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
body {font-family:<%=s_lang_font %>; font-size:9pt;color:#000000;margin-top:5px;margin-left:10px;margin-right:10px;margin-bottom:2px;background-color:#ffffff}
.sbttn {font-family:<%=s_lang_font %>; font-size:9pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5;color:#000066;height:19px;text-decoration:none;cursor:pointer}
.textbox {BORDER:1px #555555 solid;}
-->
</STYLE>
</head>

<script type="text/javascript">
<!--
function sub()
{
	var tempstr = "";
	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		tempstr = tempstr + document.f1.listall[i].value + "\t";
	}

	document.f1.allmsgs.value = tempstr;
	document.f1.action = "cal_favorite.asp";
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

function goback()
{
	if (document.f1.returl.value.length < 3)
		history.back();
	else
		location.href=document.f1.returl.value;
}
//-->
</script>

<BODY>
<br>
<FORM NAME="f1">
<input type="hidden" name="returl" value="<%=returl %>">
<input type="hidden" name="allmsgs">

<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td colspan="2" style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
编辑我收藏的效率手册
</td></tr>
<tr><td colspan=2 class="block_top_td" style="height:12px; _height:14px;"></td></tr>
<tr><td align="center">

<table width="76%" align="center" border="0" cellspacing="0" cellpadding="0">
	<tr valign=bottom>
	<td width="30%" height="10" align="center">键入要添加的帐号<%=s_lang_mh %></td>
	<td width="30%"></td>
	<td width="40%" align="center">收藏列表(帐号)<%=s_lang_mh %></td>
	</tr>
	<tr valign=top> 
	<td>
	<input maxlength=120 size=30 name="addmsg" class='textbox' onkeydown="goent()">
	</td>
	<td align=middle align="center">
		<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0">
		<tr>
		<td align="center">
		<input class="sbttn" style="WIDTH: 90px" LANGUAGE=javascript onclick="add()" type=button value="添加 >>">
		</td>
		</tr>
		<tr>
		<td align="center"><br>
		<input class="sbttn" style="WIDTH: 90px" LANGUAGE=javascript onclick="delout()" type=button value="<< 删除">
		</td>
		</tr>
		</table>
	</td>
	<td>
	<select class="drpdwn" style="WIDTH: 230px" multiple size=10 name=listall width="230">
<%
i = 0
allnum = ei.CountFavorites

do while i < allnum
	tmsg = ei.GetFavorite(i)
	Response.Write "<option value=""" & server.htmlencode(tmsg) & """>" & server.htmlencode(tmsg) & "</option>"

	tmsg = NULL

	i = i + 1
loop
%>
	</select>
	</td>
	</tr>
</table>

</td></tr>

<tr><td colspan="3" class="block_top_td" style="height:10px;"></td></tr>

<tr><td colspan="3" align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="javascript:goback();"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:sub();">保存</a>
</td></tr>
</table>

<table width="90%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px #8CA5B5 solid; margin-top:60px; padding-bottom:10px;'>
	<tr>
	<td valign="top" style="padding:4px; width:30px;"><img src='images/remind.gif' border='0' align='absmiddle'></td>
	<td style="padding:4px; color:#444444;">您可以通过添加其他用户的帐号, 来查看该用户效率手册中公开的内容.<br>
	</td></tr>
</table>
</FORM>
</BODY>
</HTML>

<%
set ei = nothing
%>
