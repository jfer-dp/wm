<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" -->

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

Session("svcal") = ""

dim ischange
ischange = FALSE

dim wmeth
set wmeth = server.createobject("easymail.WMethod")

if trim(request("revoke")) = "1" then
	wmeth.Revoke_Files Session("wem")
end if

dim uw
set uw = server.createobject("easymail.UserWeb")
uw.Load Session("wem")
useAutoClearTrashBox = uw.useAutoClearTrashBox

dim need_set_sender_name
need_set_sender_name = false

if Len(uw.MailName) < 1 then
	need_set_sender_name = true
end if

dim roupw_is_ok
roupw_is_ok = 0

if trim(request("mode")) <> "" then
	if trim(request("mode")) = "checkroupw" then
		if Session("ReadOnlyUser") = 1 then
			dim rou
			set rou = server.createobject("easymail.ReadOnlyUsers")
			rou.Load

			if rou.CheckPassword(Session("wem"), trim(request("roupw"))) = true then
				roupw_is_ok = 2
				Session("ReadOnlyUser") = 2
			else
				roupw_is_ok = 1
			end if

			set rou = nothing
		end if
	end if

	if trim(request("mode")) = "setclean" then
		uw.useAutoClearTrashBox = true
		uw.autoClearTrashBoxDays = 30
		uw.Save
		set uw = nothing
		set wmeth = nothing
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("viewmailbox.asp")
	end if

	dim pf
	set pf = server.createobject("easymail.PerFolders")
	pf.Load Session("wem")

	if trim(request("mode")) = "del" then
		if pf.DeleteFolderByName(Mid(trim(request("pfname")), 2)) = FALSE then
			Response.Write b_lang_305 & server.htmlencode(Mid(trim(request("pfname")), 2)) & b_lang_306
		else
			ischange = TRUE
		end if
	elseif trim(request("mode")) = "add" then
		if pf.MaxPerFolderNumber > pf.FolderCount then
			NewName = trim(request("NewName"))
			NewName = replace(NewName, Chr(9), " ")
			NewName = replace(NewName, """", "")
			NewName = replace(NewName, "'", "")
			pf.AddFolder NewName
			ischange = TRUE
		else
			Response.Write b_lang_307
		end if
	elseif trim(request("mode")) = "rename" then
		NewName = trim(request("NewName"))
		NewName = replace(NewName, """", "'")

		if pf.CanSetWithReceiveOutMail(NewName) = false and pf.Get_EnableRecOutMail(Mid(trim(request("pfname")), 2)) = true then
			Response.Write b_lang_308
		else
			pf.RenameFolder Mid(trim(request("pfname")), 2), NewName
			ischange = TRUE
		end if
	end if

	if ischange = TRUE then
		pf.Save
	end if

	set pf = nothing
end if


noticemsg = trim(request("noticemsg"))

set uw = nothing

dim am
set am = server.createobject("easymail.Attachments")
am.Load Session("wem"), Session("tid")

am.RemoveAll

set am = nothing


'-----------------------------------------
dim eusers
set eusers = Application("em")

maxsize = eusers.GetMailBoxSize(Session("wem"))

if maxsize < 0 then
	maxsize = 10000
end if


dim ei
set ei = server.createobject("easymail.InfoList")
'-----------------------------------------

ei.LoadSizeInfo Session("wem")

cursize = ei.allMailSize

dim bf

if maxsize > 0 then
	bf = CLng((100 * CLng(cursize / 1000)) / maxsize)

	if bf = 0 then
		if cursize > 0 then
			bf = 1
		else
			bf = 0
		end if
	end if
else
	bf = 100
end if

if bf > 100 then
	bf = 100
end if

if bf < 0 then
	bf = 0
end if


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
<link rel="stylesheet" type="text/css" href="images/popwin.css">

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
.revoke_c {clear:both;width:230px;text-align:center;font-size:12px;height:22px;line-height:14px;color:#666666;}
.revoke_c a:link,.revoke_c a:visited{color:#666666; text-decoration:underline;} 
.revoke_c a:hover{color:#901111;}
-->
</STYLE>
</HEAD>

<script type="text/javascript" src="images/sc_left.js"></script>

<script type="text/javascript">
<!--
function window_onload() {
try{
	if (parent.f1.document.getElementById("purl") != null)
		parent.f1.document.leftval.purl.value = "";
}catch(error){}
<%
if ischange = TRUE then
	Response.Write "parent.f1.window.location.href = ""left.asp?" & getGRSN() & """;" & Chr(13)
end if

dim mail_count_alert
mail_count_alert = 0

if trim(request("rla")) = "1" then
	dmn = trim(request("dmn"))

	if IsNull(dmn) = false and Len(dmn) > 0 then
		defaultMailsNumber = CLng(dmn)

		if defaultMailsNumber > 9 and defaultMailsNumber < 10000 then
			mc_bf = CLng((100 * CLng(ei.allMailCount)) / defaultMailsNumber)

			if mc_bf > 80 then 
				mail_count_alert = 1
			end if

			if mc_bf > 90 then 
				mail_count_alert = 2
			end if
		end if
	end if

	if bf > 90 or mail_count_alert = 2 then
		Response.Write "alert('" & b_lang_309 & "');"
	end if
end if
%>

<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then
%>
	pf_onchange();
<%
end if
%>

<%
if roupw_is_ok = 2 then
%>
	parent.f1.window.location.href = "left.asp?<%=getGRSN() %>";
<%
elseif roupw_is_ok = 1 then
%>
	alert("<%=b_lang_077 %>");
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
		alert("<%=b_lang_372 %>");
}

function add() {
	if (document.f1.NewName.value.indexOf("\"") != -1 || document.f1.NewName.value.indexOf("'") != -1)
	{
		alert("<%=b_lang_310 %>");
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
		alert("<%=b_lang_310 %>");
		return;
	}

	if (document.f1.NewName.value != "")
	{
		document.f1.mode.value = "rename";
		document.f1.submit();
	}
}

function emptyfolder(e) {
	stop_event(e);

	if (confirm("<%=b_lang_311 %>") == false)
		return ;

	location.href = "mulmail.asp?mode=cleanTrash&<%=getGRSN() %>&gourl=<%=Server.URLEncode("viewmailbox.asp?" & getGRSN()) %>";
}

function movetotrash(e) {
	stop_event(e);

	if (confirm("<%=b_lang_312 %>") == false)
		return ;

	location.href = "mulmail.asp?mode=mitt&<%=getGRSN() %>&gourl=<%=Server.URLEncode("viewmailbox.asp?" & getGRSN()) %>";
}

function stop_event(e) {
	if (!e) var e = window.event;
	e.cancelBubble = true;
	if (e.stopPropagation)
	e.stopPropagation();
}

function set_clean(e) {
	stop_event(e);

	document.f1.mode.value = "setclean";
	document.f1.submit();
}

function m_over(tag_obj) {
	tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj) {
	tag_obj.style.backgroundColor = "white";
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

function showmail(s_url) {
	location.href = s_url;
}

function check_roupw() {
	if (document.getElementById("roupw").value.length > 0)
	{
		document.f1.mode.value = "checkroupw";
		document.f1.submit();
	}
}

function stopent() {
<%
if isMSIE = true then
%>
	if (event.keyCode == 13)
		event.keyCode = 9;
<%
end if
%>
}

function ffstopent(event) {
<%
if isMSIE = false then
%>
	if (event.which == 13)
		return false;
<%
end if
%>
}
//-->
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
<form name="f1" method="post" action="viewmailbox.asp?<%=getGRSN() %>">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white" style="margin-top:6px;">
	<tr><td colspan="4" class="block_top_td"><div class="table_min_width"></div></td></tr>

	<tr><td colspan="4">
	<div class="size_msg">
		<table width="100%" border="0" align="center" cellspacing="0" bgcolor="white">
		<tr><td width="100%"> 
			<table style="border:1px #104A7B solid; padding:1px; margin-top:6px;" cellspacing=0 cellpadding=0 width="100%" border=0>
			<tr><td width="100%">
			<div style="font-size:3px; width:<%=bf %>%; height:6px; background-color:<%
if bf > 90 then
	Response.Write "#dd0000"
else
	Response.Write "#339933"
end if
%>"></div>
			</td></tr>
			</table>
		</td></tr>

		<tr><td colspan="3" class="block_top_td" style="height:4px;"></td></tr>

		<tr>
		<td colspan="3"><font color="#901111" style="font-size:18px;"><%=Session("mail") %></font> <%=b_lang_313 %><%=s_lang_mh %><%
if maxsize > 1000 then
	Response.Write CLng(maxsize / 1000) & b_lang_314
else
	Response.Write maxsize & "<b>KB</b>"
end if
%><%=b_lang_315 %><%=bf %><%=b_lang_316 %>
		</td></tr>
<%
if bf > 80 or mail_count_alert > 0 then
	Response.Write "<tr><td colspan='3' class='td_alert'>"
	Response.Write "<div class='wwm_alert_msg'><b>" & b_lang_317 & "</b>" & s_lang_mh & b_lang_318 & "</div>"
	Response.Write "</td></tr>"
end if

if noticemsg <> "" then
	Response.Write "<tr><td colspan='3' class='td_alert'>"
	Response.Write "<div class='wwm_alert_msg'><b>" & b_lang_317 & "</b>" & s_lang_mh & "<a href='logon.asp?forget=1&" & getGRSN() & "'>" & noticemsg & "</a></div>"
	Response.Write "</td></tr>"
end if
%>
		</table>
	</div>
	</td></tr>
<%
wl_allnum = eusers.Get_Other_IPs_Count(Session("wem"), Request.ServerVariables("REMOTE_ADDR"))
if wl_allnum > 0 then
%>
	<tr><td colspan="4" class="block_top_td" style="height:8px;"></td></tr>
	<tr><td align="left" colspan="4">
	<div class="size_msg">
<img src="images/shield.gif" align='absmiddle' border="0">&nbsp;&nbsp;<%=b_lang_319 %> <%
i = 0
do while i < wl_allnum
	wl_ip = eusers.Get_Other_IP(Session("wem"), Request.ServerVariables("REMOTE_ADDR"), i)

	if isadmin() = true or Session("ReadOnlyUser") <> 1 then
		if Len(wl_ip) > 0 then
			Response.Write "[<a href='kick.asp?" & getGRSN() & "&ip=" & wl_ip & "' style='text-decoration: none;' title='" & b_lang_320 & "'><b>" & wl_ip & "</b>&nbsp;<img src='images/kick.gif' align='absmiddle' border='0'></a>] "
		end if
	end if

	wl_ip = NULL

	i = i + 1
loop
%><%=b_lang_321 %>
	</div>
	</td></tr>
<%
end if
set eusers = nothing

if isadmin() = true or Session("ReadOnlyUser") <> 1 then

dim ecal
set ecal = server.createobject("easymail.CalendarNotice")
ecal.Load Session("wem")
calnt_count = ecal.Count
set ecal = nothing
	
if calnt_count > 0 then
%>
	<tr><td colspan="4" class="block_top_td" style="height:8px;"></td></tr>
	<tr><td align="left" colspan="4">
	<div class="size_msg">
<img src="images/remind.gif" align='absmiddle' border="0">&nbsp;&nbsp;<%=b_lang_322 %><a href="cal_listinvited.asp?<%=getGRSN() %>&fmeml=1"><%=calnt_count %><%=b_lang_323 %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="cal_listinvited.asp?<%=getGRSN() %>&fmeml=1"><%=b_lang_324 %></a>
	</div>
	</td></tr>
<%
end if

dim mht
set mht = server.createobject("easymail.Hint")
mht.GetRandHint h_isok, h_isDisabled, h_isUseGrsn, h_isPOP, h_expire, h_expProc, h_msg, h_url

if h_isok = true then
%>
	<tr><td colspan="4" class="block_top_td" style="height:8px;"></td></tr>
	<tr><td align="left" colspan="4">
	<div class="size_msg">
<img src="images/remind.gif" align='absmiddle' border="0">&nbsp;
<%
Response.Write server.htmlencode(h_msg)
Response.Write "&nbsp;&nbsp;&nbsp;"

if Len(h_url) > 0 then
	if h_isUseGrsn = true then
		if InStr(1, h_url, "?") = 0 Then
			Response.Write "<a href=""" & server.htmlencode(h_url & "?" & getGRSN()) & """"
		else
			Response.Write "<a href=""" & server.htmlencode(h_url & "&" & getGRSN()) & """"
		end if
	else
		Response.Write "<a href=""" & server.htmlencode(h_url) & """"
	end if

	if h_isPOP = true then
		 Response.Write " target=""_blank"">" & b_lang_325 & "</a>"
	else
		 Response.Write ">" & b_lang_325 & "</a>"
	end if
end if
%>
	</div>
	</td></tr>
<%
end if

h_isok = NULL
h_isDisabled = NULL
h_isUseGrsn = NULL
h_isPOP = NULL
h_expire = NULL
h_expProc = NULL
h_msg = NULL
h_url = NULL

set mht = nothing

end if
%>
	<tr><td colspan="4" class="block_top_td" style="height:12px;"></td></tr>

	<tr class="title_tr">
	<td width="55%" class="st_l"><%=b_lang_326 %></td>
	<td width="15%" class="st_l"><%=b_lang_327 %></td>
	<td width="15%" class="st_l"><%=b_lang_328 %></td>
	<td width="15%" class="st_r"><%=b_lang_329 %></td>
	</tr>

	<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);' onclick="javascript:showmail('listmail.asp?mode=in&<%=getGRSN() %>');">
	<td align="center" class='cont_td_word'>
	<%=b_lang_149 %>
<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then
	if ei.inboxMailCount > 20 then
		Response.Write "&nbsp;&nbsp;&nbsp;<a onclick=""javascript:movetotrash(event);""><img src='images/del.gif' align='absmiddle' border='0' title=""" & b_lang_330 & """></a>"
	end if
end if
%>
	</td>
	<td align="right" class='cont_td'><%= ei.inboxMailCount %></td>
	<td align="right" class='cont_td'>
<%
if ei.newInBoxMailCount = 0 then
	Response.Write ei.newInBoxMailCount
else
%>
<font color="#901111"><b><%= ei.newInBoxMailCount %></b></font>
<%
end if
%>
	</td>
	<td align="right" class='cont_td'><%= CLng(ei.inboxMailSize/1000) %>K</td>
	</tr>

	<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);' onclick="javascript:showmail('listmail.asp?mode=out&<%=getGRSN() %>');">
	<td align="center" class='cont_td_word'>
	<%=b_lang_150 %>
	</td>
	<td align="right" class='cont_td'><%= ei.outboxMailCount %></td>
	<td align="right" class='cont_td'>
<%
if ei.newOutBoxMailCount = 0 then
	Response.Write ei.newOutBoxMailCount
else
%>
<font color="#901111"><b><%= ei.newOutBoxMailCount %></b></font>
<%
end if
%>
	</td>
	<td align="right" class='cont_td'><%= CLng(ei.outboxMailSize/1000) %>K</td>
	</tr>

	<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);' onclick="javascript:showmail('listmail.asp?mode=sed&<%=getGRSN() %>');">
	<td align="center" class='cont_td_word'>
	<%=b_lang_151 %>
	</td>
	<td align="right" class='cont_td'><%= ei.sendboxMailCount %></td>
	<td align="right" class='cont_td'>
<%
if ei.newSendBoxMailCount = 0 then
	Response.Write ei.newSendBoxMailCount
else
%>
<font color="#901111"><b><%= ei.newSendBoxMailCount %></b></font>
<%
end if
%>
	</td>
	<td align="right" class='cont_td'><%= CLng(ei.sendboxMailSize/1000) %>K</td>
	</tr>

	<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);' onclick="javascript:showmail('listmail.asp?mode=del&<%=getGRSN() %>');">
	<td align="center" class='cont_td_word'>
<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then
%>
	<%=b_lang_331 %>&nbsp;&nbsp;&nbsp;<a onclick="javascript:emptyfolder(event);"><img src='images/del.gif' align='absmiddle' border='0' title="<%=b_lang_332 %>"></a>
<%
if useAutoClearTrashBox = false and ei.delboxMailCount > 99 then
%>
		&nbsp;&nbsp;&nbsp;<input type="button" value="<%=b_lang_333 %>" onclick="javascript:set_clean(event);" class="sbttn" style="<%=b_lang_334 %>">
<%
end if
%>
<%
else
%>
	<%=b_lang_331 %>
<%
end if
%>
	</td>
	<td align="right" class='cont_td'><%= ei.delboxMailCount %></td>
	<td align="right" class='cont_td'>
<%
if ei.newDelBoxMailCount = 0 then
	Response.Write ei.newDelBoxMailCount
else
%>
<font color="#901111"><b><%= ei.newDelBoxMailCount %></b></font>
<%
end if
%>
	</td>
	<td align="right" class='cont_td'><%= CLng(ei.delboxMailSize/1000) %>K</td>
	</tr>
<%
all_mailnum = ei.inboxMailCount + ei.outboxMailCount + ei.sendboxMailCount + ei.delboxMailCount
all_mailnewnum = ei.newInBoxMailCount + ei.newOutBoxMailCount + ei.newSendBoxMailCount + ei.newDelBoxMailCount
all_mailsize = ei.inboxMailSize + ei.outboxMailSize + ei.sendboxMailSize + ei.delboxMailSize

allnum = ei.PerFolderCount

i = 0
do while i < allnum
	ei.GetPerFolderInfo i, pfname, pfmailcount, pfsize, pfnewmailcount

	all_mailnum = all_mailnum + pfmailcount
	all_mailnewnum = all_mailnewnum + pfnewmailcount
	all_mailsize = all_mailsize + pfsize

	Response.Write "<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);' onclick=""javascript:showmail('listmail.asp?mode=" & Server.URLEncode(pfname) & "&" & getGRSN() & "');"">"
	Response.Write "<td align='center' class='cont_td_word'>" & server.htmlencode(pfname)

	if isadmin() = true or Session("ReadOnlyUser") <> 1 then
		if Application("em_Enable_ShareFolder") = true then
			Response.Write "&nbsp;&nbsp;[<a href='pf_setupfolder.asp?foldername=" & Server.URLEncode(pfname) & "&" & getGRSN() & "'>" & b_lang_335 & "</a>]"
			Response.Write "&nbsp;[<a href='ff_sharefolder.asp?foldername=" & Server.URLEncode(pfname) & "&" & getGRSN() & "'>" & b_lang_336 & "</a>]"
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

	<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);' onclick="javascript:showmail('attfolders.asp?<%=getGRSN() %>');">
	<td align="center" class='cont_td_word' style='color:#1e5494;'>
	<b><%=b_lang_337 %></b>
	</td>
	<td align="right" class='cont_td'><%= ei.allMailCount - all_mailnum %></td>
	<td align="right" class='cont_td'>
<%
if (ei.allNewMailCount - all_mailnewnum) = 0 then
	Response.Write "0"
else
%>
<font color="#901111"><b><%= ei.allNewMailCount - all_mailnewnum %></b></font>
<%
end if
%>
	</td>
	<td align="right" class='cont_td'><%= CLng((ei.allMailSize - all_mailsize)/1000) %>K</td>
	</tr>

	<tr class='cont_tr_np' onmouseover='m_over(this);' onmouseout='m_out(this);'>
	<td align="center" class='cont_td_word'><b><%=b_lang_338 %></b></td>
	<td align="right" class='cont_td'><b><%= ei.allMailCount %></b></td>
	<td align="right" class='cont_td'><b><%= ei.allNewMailCount %></b></td>
	<td align="right" class='cont_td'><b><%= CLng(ei.allMailSize/1000) %>K</b></td>
	</tr>

	<tr><td colspan="4" class="block_top_td" style="height:22px;"></td></tr>

<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then

wmeth.Get_Revoke_Info Session("wem"), files_number, wYear, wMonth, wDay

if files_number > 0 then
	dim d_show_str
	if Year(now()) = wYear and Month(now()) = wMonth and Day(now()) = wDay then
		d_show_str = b_lang_359
	else
		d_show_str = wYear & b_lang_360 & wMonth & b_lang_361 & wDay & b_lang_362
	end if
%>
	<tr> 
	<td align="right" colspan="4">
<div class="revoke_c"><a href="viewmailbox.asp?<%=getGRSN() %>&revoke=1"><%
Response.Write b_lang_363 & d_show_str & b_lang_364 & files_number & b_lang_365

files_number = NULL
wYear = NULL
wMonth = NULL
wDay = NULL
%></a></div><br>
	</td><tr>
<%
end if

end if
%>
	<tr>
	<td align="left" colspan="4">
<table width="100%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%
if isadmin() = false and Session("ReadOnlyUser") = 1 then
%>
<%=b_lang_366 %>
<%
else
%>
<%=b_lang_339 %>
<%
end if
%>
</td></tr>
<tr><td class="block_top_td" style="height:10px; _height:12px;"></td></tr>
<tr><td align="left" style="padding-left:8px; padding-right:8px;">

<%
if isadmin() = false and Session("ReadOnlyUser") = 1 then
%>
<input type="password" id="roupw" name="roupw" size="20" maxlength="30" class='n_textbox' onkeydown="stopent()" onkeypress="return ffstopent(event);">
<input type="button" value="<%=s_lang_ok %>" onclick="javascript:check_roupw()" class="b_sbttn">
<%
else
%>
<table width="97%" border="0" align="left" cellspacing="0" bgcolor="white">
	<tr>
	<td align="left">
<select id="pfName" name="pfName" class="drpdwn" size="1" onchange="javascript:pf_onchange();">
<option value=""><%=b_lang_340 %></option>
<%
allnum = ei.PerFolderCount
i = 0

do while i < allnum
	ei.GetPerFolderInfo i, pfname, pfmailcount, pfsize, pfnewmailcount

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
</select><span id="del_bk">&nbsp;</span><input id="bt_del" type="button" value="<%=b_lang_341 %>" onclick="javascript:deletepf()" class="b_sbttn">
<input type="text" name="NewName" size="20" maxlength="64" class='n_textbox'><span id="create_bk">&nbsp;</span><input id="bt_create" type="button" value="<%=b_lang_342 %>" name="button" onclick="javascript:add()" class="b_sbttn">
<input id="bt_cgname" type="button" value="<%=b_lang_344 %>" name="button" onclick="javascript:rename()" class="b_sbttn">
	</td>
	</tr>
</table>
<%
end if
%>
<input type="hidden" name="mode" value="add">
</td></tr>
</table>
	</td>
	</tr>

	<tr><td colspan="4" class="block_top_td" style="height:6px; border-bottom:1px #a7c5e2 solid;">&nbsp;</td></tr>
</table>

<table width="100%" border="0" align="center" cellspacing="0" bgcolor="white" style='margin-top:40px; padding-bottom:10px;'>
	<tr><td nowrap align="center">
<%=b_lang_343 %>
	</td></tr>
</table>

</Form>

<div id="pop_overlay">
</div>

<div id="pop_win" style="display:none; position:absolute;" class="mydiv">
	<div class="pop_base"><div class="pop_bd bd"><div class="title">
		<div class="title_left"><%=b_lang_357 %></div>
		<div class="title_right" title="<%=s_lang_close %>" id="pop_close_wind"><span>&nbsp;</span></div>
	</div>
	<div class="pop_content"><span id="pop_ctmsg"><%=b_lang_358 %></span><br>
	<input type="text" id="sender_name" size="36" maxlength="60" class='b_input'>
	</div>
	<div class="title_bottom">
	<div class="title_ok_cancel_div">
	<a id="pop_ok" class="wwm_btnDownload btn_gray" href="#" onclick="javascript:set_name();"><%=s_lang_ok %></a>&nbsp;
	<a class="wwm_btnDownload btn_gray" href="#" onclick="javascript:pop_close()"><%=s_lang_cancel %></a>
	</div></div></div></div>
</div>

<script type="text/javascript" src="images/bRoundCurve 1.0.js"></script>
<script type="text/javascript">
b_RoundCurve("revoke_c","#A0C044","#ffffff",1);

var doc = document.documentElement;
var body = document.body;
var oWin;
var oLay;
var oClose;

oWin = document.getElementById("pop_win");
oLay = document.getElementById("pop_overlay");
oClose = document.getElementById("pop_close_wind");

oClose.onclick = function ()
{
	pop_close();
}

<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then
	if need_set_sender_name = true then
		Response.Write "pop_show();"
		Response.Write "document.getElementById('sender_name').focus();"
	end if
else
	Response.Write "document.getElementById('roupw').focus();"
end if
%>

function pop_show()
{
	oLay.style.height = document.documentElement.scrollHeight + "px";
	oLay.style.width = document.documentElement.scrollWidth + "px";

	var ie_h = doc && doc.clientHeight  || body && body.clientHeight  || 0;
	var ie_w = doc && doc.clientWidth  || body && body.clientWidth  || 0;

	if (ie_h > document.documentElement.scrollHeight)
		oLay.style.height = ie_h + "px";

	if (ie_w > document.documentElement.scrollWidth)
		oLay.style.width = ie_w + "px";

	oLay.style.display = "block";
	oWin.style.display = "block"	
}

function pop_close()
{
	oLay.style.display = "none";
	oWin.style.display = "none"	
}

var g_newname;

function set_name()
{
	g_newname = document.getElementById("sender_name").value;
	g_newname = g_newname.replace(/\'/g,"");
	g_newname = g_newname.replace(/\"/g,"");

	if (g_newname.length < 1)
	{
		alert("<%=b_lang_193 %>.");
		document.getElementById("sender_name").value = "";
		document.getElementById('sender_name').focus();
		return ;
	}

	SendInfo();
}

var request = false;
try {
	request = new XMLHttpRequest();
} catch (trymicrosoft) {
try {
	request = new ActiveXObject("Msxml2.XMLHTTP");
} catch (othermicrosoft) {
try {
	request = new ActiveXObject("Microsoft.XMLHTTP");
} catch (failed) {
	request = false;
}}}

if (!request)
	alert("Error initializing XMLHttpRequest!");

function SendInfo()
{
	var url = "style.asp?mbsetname=" + escape(g_newname) + "&<%=getGRSN() %>";
	request.open("GET", url, true);
	request.onreadystatechange = updatePage;
	request.send(null);
}

function updatePage()
{
	if (request.readyState == 4)
	{
		if (request.status == 200)
			pop_close();
	}
}
</script>

</BODY>
</HTML>

<%
set ei = nothing
set wmeth = nothing
%>
