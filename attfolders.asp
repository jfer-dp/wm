<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
dim ischange
ischange = FALSE

if trim(request("mode")) <> "" then
	dim eusers
	set eusers = Application("em")
	MaxPerFolderNumber = eusers.GetMaxPerFolderNumber(Session("wem"))
	set eusers = nothing

	dim pf
	set pf = server.createobject("easymail.PerAttFolders")
	pf.Load Session("wem")

	if trim(request("mode")) = "del" then
		if pf.DeleteFolderByName(Mid(trim(request("pfname")), 2)) = FALSE then
			Response.Write "*" & a_lang_001 & "[" & server.htmlencode(Mid(trim(request("pfname")), 2)) & "]" & a_lang_002
		else
			ischange = TRUE
		end if
	elseif trim(request("mode")) = "add" then
		if MaxPerFolderNumber > pf.FolderCount then
			NewName = trim(request("NewName"))
			NewName = replace(NewName, Chr(9), " ")
			NewName = replace(NewName, """", "")
			NewName = replace(NewName, "'", "")
			pf.AddFolder NewName
			ischange = TRUE
		else
			Response.Write "*" & a_lang_003
		end if
	elseif trim(request("mode")) = "rename" then
		NewName = trim(request("NewName"))
		NewName = replace(NewName, """", "'")

		pf.RenameFolder Mid(trim(request("pfname")), 2), NewName
		ischange = TRUE
	end if

	if ischange = TRUE then
		pf.Save
	end if

	set pf = nothing
end if

noticemsg = trim(request("noticemsg"))

dim ei
set ei = server.createobject("easymail.InfoList")

ei.IsAttFolder = true
ei.LoadSizeInfo Session("wem")

dim all_mailnum
dim all_mailnewnum
dim all_mailsize

all_mailnum = 0
all_mailnewnum = 0
all_mailsize = 0
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
.table_min_width {width:660px; font-size:0pt; height:0px; width:0px; border:0px;}
.size_msg {padding:8px 12px 8px 12px; border-radius:4px; -webkit-border-radius:4px; text-align:left; border:1px #A5B6C8 solid;}
.wwm_alert_msg {padding:4px 8px 4px 8px; _padding-top:6px; line-height:18px; color:black; background:#FFF3C3; border-radius:4px; -webkit-border-radius:4px; text-align:left; border:1px #7E4F05 solid; word-break:break-all; word-wrap:break-word;}
.td_alert {padding-top:8px;}
.sbttn {font-family:<%=s_lang_font %>;font-size:9pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5;color:#000066;height:19px;text-decoration:none;cursor:pointer}
.b_sbttn {font-family:<%=s_lang_font %>;font-size:9pt; background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5;color:#000066;height:22px;text-decoration:none;cursor:pointer}
.title_tr {white-space:nowrap; background:#f2f4f6; height:24px;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.st_l, .st_r {height:26px; text-align:center; white-space:nowrap; border-left:1px solid #A5B6C8; border-top:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8; _padding-top:2px;}
.st_r {border-right:1px solid #A5B6C8;}
.cont_tr {background:white; height:24px; cursor:pointer;}
.cont_tr_np {background:white; height:24px;}
.cont_td {height:24px; white-space:nowrap; border-bottom:1px solid #A5B6C8; padding-left:6px; padding-right:6px;}
.cont_td_word {height:24px; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px; word-break:break-all; word-wrap:break-word;}
-->
</STYLE>
</HEAD>

<script type="text/javascript" src="images/sc_left.js"></script>

<script type="text/javascript">
<!--
function window_onload() {
<%
if ischange = TRUE then
	Response.Write "parent.f1.window.location.href = ""left.asp?" & getGRSN() & "&asp="" + escape(document.location.href);"
end if
%>

<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then
%>
	pf_onchange();
<%
end if
%>
}

function deletepf() {
	if (document.getElementById("pfName").value.charAt(0) == "0")
	{
		if (document.f1.pfName.selectedIndex != -1)
		{
			document.f1.mode.value = "del";
			document.f1.submit();
		}
	}
	else
		alert("<%=a_lang_349 %>");
}

function add() {
	if (document.f1.NewName.value.indexOf("\"") != -1 || document.f1.NewName.value.indexOf("'") != -1)
	{
		alert("<%=a_lang_004 %>");
		return;
	}

	if (document.f1.NewName.value != "")
	{
		document.f1.mode.value = "add";
		document.f1.submit();
	}
}

function rename() {
	if (document.f1.NewName.value.indexOf("\"") != -1 || document.f1.NewName.value.indexOf("'") != -1)
	{
		alert("<%=a_lang_004 %>");
		return;
	}

	if (document.f1.NewName.value != "")
	{
		document.f1.mode.value = "rename";
		document.f1.submit();
	}
}

function m_over(tag_obj) {
	tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj) {
	tag_obj.style.backgroundColor = "white";
}

function showmail(s_url) {
	location.href = s_url;
}

function pf_onchange() {
	if (document.getElementById("pfName").selectedIndex == 0)
	{
		document.getElementById("del_bk").style.display = "none";
		document.getElementById("create_bk").style.display = "inline";

		document.getElementById("bt_del").style.display = "none";
		document.getElementById("bt_cgname").style.display = "none";
		document.getElementById("bt_create").style.display = "inline";
	}
	else
	{
		document.getElementById("del_bk").style.display = "inline";
		document.getElementById("create_bk").style.display = "none";

		document.getElementById("bt_del").style.display = "inline";
		document.getElementById("bt_cgname").style.display = "inline";
		document.getElementById("bt_create").style.display = "none";
	}
}
//-->
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
<form name="f1" method="post" action="attfolders.asp?<%=getGRSN() %>">
<br>
<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr class="title_tr">
	<td width="55%" class="st_l"><%=a_lang_005 %></td>
	<td width="15%" class="st_l"><%=a_lang_006 %></td>
	<td width="15%" class="st_l"><%=a_lang_007 %></td>
	<td width="15%" class="st_r"><%=a_lang_008 %></td>
	</tr>

	<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);' onclick="javascript:showmail('listatt.asp?mb=att&<%=getGRSN() %>');">
	<td align="center" class='cont_td_word'>
	<%=a_lang_009 %>
<%
if Application("em_Enable_ShareFolder") = true then
	if isadmin() = true or Session("ReadOnlyUser") <> 1 then
%>
&nbsp;&nbsp;[<a href="ff_sharefolder.asp?foldername=att&<%=getGRSN() %>"><%=a_lang_010 %></a>]
<%
	end if
end if
%>
	</td>
	<td align="right" class='cont_td'><%= ei.attachmentsCount %></td>
	<td align="right" class='cont_td'>
<%
if ei.newAttachmentsCount = 0 then
	Response.Write ei.newAttachmentsCount
else
%>
	<font color="#901111"><b><%= ei.newAttachmentsCount %></b></font>
<%
end if
%>
	</td>
	<td align="right" class='cont_td'><%= CLng(ei.attachmentsSize/1000) %>K</td>
	</tr>
<%
all_mailnum = ei.attachmentsCount
all_mailnewnum = ei.newAttachmentsCount
all_mailsize = ei.attachmentsSize

allnum = ei.PerAttFolderCount

i = 0
do while i < allnum
	ei.GetPerAttFolderInfo i, pfname, pfmailcount, pfsize, pfnewmailcount

	all_mailnum = all_mailnum + pfmailcount
	all_mailnewnum = all_mailnewnum + pfnewmailcount
	all_mailsize = all_mailsize + pfsize

	Response.Write "<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);' onclick=""javascript:showmail('listatt.asp?mb=" & Server.URLEncode(pfname) & "&" & getGRSN() & "');"">"
	Response.Write "<td align='center' class='cont_td_word'>" & server.htmlencode(pfname)

	if Application("em_Enable_ShareFolder") = true then
		if isadmin() = true or Session("ReadOnlyUser") <> 1 then
			Response.Write "&nbsp;&nbsp;&nbsp;[<a href='ff_sharefolder.asp?mode=att&foldername=" & Server.URLEncode(pfname) & "&" & getGRSN() & "'>" & a_lang_010 & "</a>]"
		end if
	end if

	Response.Write "</td><td align='right' class='cont_td'>" & pfmailcount & "</td>"
	Response.Write "<td align='right' class='cont_td'>"

	if pfnewmailcount = 0 then
		Response.Write pfnewmailcount
	else
		Response.Write "<font color='#901111'><b>" & pfnewmailcount & "</b></font>"
	end if

	Response.Write "</td><td align='right' class='cont_td'>"
	Response.Write CLng(pfsize/1000) & "K</td></tr>" & Chr(13)

	pfname = NULL
	pfmailcount = NULL
	pfsize = NULL
	pfnewmailcount = NULL

	i = i + 1
loop
%>
	<tr class='cont_tr_np' onmouseover='m_over(this);' onmouseout='m_out(this);'>
	<td align="center" class='cont_td_word'><b><%=a_lang_011 %></b></td>
	<td align="right" class='cont_td'><b><%= all_mailnum %></b></td>
	<td align="right" class='cont_td'><b><%= all_mailnewnum %></b></td>
	<td align="right" class='cont_td'><b><%= CLng(all_mailsize/1000) %>K</b></td>
	</tr>

	<tr><td colspan="4" class="block_top_td" style="height:30px;"></td></tr>

<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then
%>
	<tr> 
	<td align="left" colspan="4">
<table width="100%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_012 %>
</td></tr>
<tr><td class="block_top_td" style="height:10px; _height:12px;"></td></tr>
<tr><td align="left" style="padding-left:8px; padding-right:8px;">

<table width="97%" border="0" align="left" cellspacing="0" bgcolor="white">
	<tr>
	<td align="left">
<select id="pfName" name="pfName" class="drpdwn" size="1" onchange="javascript:pf_onchange();">
<option value="">[<%=a_lang_013 %>]</option>
<%
allnum = ei.PerAttFolderCount
i = 0

do while i < allnum
	ei.GetPerAttFolderInfo i, pfname, pfmailcount, pfsize, pfnewmailcount

	if pfmailcount > 0 then
		Response.Write "<option value=""" & "1" & server.htmlencode(pfname) & """>" & server.htmlencode(pfname) & "</option>"
	else
		Response.Write "<option value=""" & "0" & server.htmlencode(pfname) & """>" & server.htmlencode(pfname) & "</option>"
	end if

	pfname = NULL
	pfmailcount = NULL
	pfsize = NULL
	pfnewmailcount = NULL

	i = i + 1
loop
%>
</select><span id="del_bk">&nbsp;</span><input id="bt_del" type="button" value="<%=a_lang_014 %>" onclick="javascript:deletepf()" class="b_sbttn">
<input type="text" name="NewName" size="20" maxlength="64" class='n_textbox'><span id="create_bk">&nbsp;</span><input id="bt_create" type="button" value="<%=a_lang_015 %>" name="button" onclick="javascript:add()" class="b_sbttn">
<input id="bt_cgname" type="button" value="<%=a_lang_016 %>" name="button" onclick="javascript:rename()" class="b_sbttn">
<input type="hidden" name="mode" value="add">
	</td>
	</tr>
</table>
	</td></tr>
</table>
<%
end if
%>
</td></tr>
</table>

</Form>
</BODY>
</HTML>

<%
set ei = nothing
%>
