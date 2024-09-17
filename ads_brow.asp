<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" --> 

<%
dim is_domain_admin
is_domain_admin = false

if isadmin() = false then
	dim dm
	set dm = server.createobject("easymail.Domain")
	dm.Load

	allnum = dm.GetUserManagerDomainCount(Session("wem"))

	curdomain = Mid(Session("mail"), InStr(Session("mail"), "@") + 1)

	i = 0
	do while i < allnum
		domain = dm.GetUserManagerDomain(Session("wem"), i)

		if LCase(curdomain) = LCase(domain) then
			is_domain_admin = true
		end if

		domain = NULL

		i = i + 1
	loop

	set dm = nothing
else
	is_domain_admin = true
end if


dim ads
set ads = server.createobject("easymail.Addresses")
ads.Load Session("wem")

i = 0
allnum = ads.EmailCount
gourl = "ads_brow.asp?" & getGRSN

dim wemcert
set wemcert = server.createobject("easymail.WebEasyMailCert")

if allnum > 0 then
	wemcert.Load Session("wem"), Session("mail")
end if
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/slstyle.css">
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
<link rel="stylesheet" type="text/css" href="images/ads.css">
</HEAD>

<script type="text/javascript" src="images/sc_left.js"></script>
<script type="text/javascript" src="images/mglobal.js"></script>
<script type="text/javascript" src="images/jquery.min.js"></script>
<script type="text/javascript" src="images/jquery-powerFloat-min.js"></script>

<script language="JavaScript">
<!-- 
parent.f1.clean_ads();

if (ie == 6)
	document.execCommand("BackgroundImageCache", false, true); 

function addtoleft()
{
	var theObjTo;
	var theObjCc;
	var theObjBcc;

	var i;
	var to = "";
	var cc = "";
	var bcc = "";

	for (i = 0; i < <%=allnum %>; i++)
	{
		theObjTo = document.getElementById("checkto" + i);
		theObjCc = document.getElementById("checkcc" + i);
		theObjBcc = document.getElementById("checkbcc" + i);

		if (theObjTo != null && theObjTo.checked == true)
		{
			if (to != "")
				to = to + "," + theObjTo.value;
			else
				to = to + theObjTo.value;
		}
		else if (theObjCc != null && theObjCc.checked == true)
		{
			if (cc != "")
				cc = cc + "," + theObjCc.value;
			else
				cc = cc + theObjCc.value;
		}
		else if (theObjBcc != null && theObjBcc.checked == true)
		{
			if (bcc != "")
				bcc = bcc + "," + theObjBcc.value;
			else
				bcc = bcc + theObjBcc.value;
		}
	}

	parent.f1.document.leftval.to.value = to;
	parent.f1.document.leftval.cc.value = cc;
	parent.f1.document.leftval.bcc.value = bcc;
}

function gosend()
{
	addtoleft();
	location.href = "wframe.asp?<%=getGRSN() %>&gourl=<%=Server.URLEncode(gourl) %>";
}

function mdel()
{
	if (ischeck() == true)
	{
		if (confirm("<%=s_lang_0115 %>") == false)
			return ;

		conv_upinfo();
		document.form1.action = "ads_delete.asp?<%=getGRSN() %>&mdel=1&gourl=<%=Server.URLEncode(gourl) %>";
		document.form1.submit();
	}
}

function checkall(tgobj) {
	var theObj;
	for(var i = 0; i < <%=allnum %>; i++)
	{
		theObj = document.getElementById("checkdel" + i);
		if (theObj != null && theObj.disabled == false)
			theObj.checked = tgobj.checked;
	}
}

function ischeck() {
	var i = 0;
	var theObj;

	for(; i<<%=allnum %>; i++)
	{
		theObj = document.getElementById("checkdel" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}

function sub_up()
{
	var exname = fsa.upfile.value.substr(fsa.upfile.value.length - 4);
	exname.toLowerCase();

	if (fsa.upfile.value != "" && exname == ".csv")
	{
		fsa.action = "csvimport.asp";
		fsa.submit();
	}
	else
	{
		if (fsa.upfile.value != "" && exname != ".csv")
			alert("<%=s_lang_0356 %>.");
	}
}

function select_mto(msi) {
	if (msi == 1)
		location.href = "adg_brow.asp?<%=getGRSN() %>";
	else if (msi == 2)
		location.href = "ads_pubbrow.asp?<%=getGRSN() %>";
	else if (msi == 3)
		location.href = "ads_dm_pubbrow.asp?<%=getGRSN() %>";
}

<%
if isadmin() = true then
%>
function conv2pads()
{
	if (ischeck() == true)
	{
		if (confirm("<%=s_lang_0357 %>?") == false)
			return ;

		conv_upinfo();
		document.form1.action = "ads_conv.asp?cfrom=ads&cto=pads&<%=getGRSN() %>&gourl=<%=Server.URLEncode(gourl) %>";
		document.form1.submit();
	}
}
<%
end if

if is_domain_admin = true then
%>
function conv2dmpads()
{
	if (ischeck() == true)
	{
		if (confirm("<%=s_lang_0358 %>?") == false)
			return ;

		conv_upinfo();
		document.form1.action = "ads_conv.asp?cfrom=ads&cto=dmpads&<%=getGRSN() %>&gourl=<%=Server.URLEncode(gourl) %>";
		document.form1.submit();
	}
}
<%
end if
%>

function select_mopt(msi) {
	if (msi != 5)
		document.getElementById("csvimp").style.display = "none";

	if (msi == 2)
		location.href = "ads_add.asp?<%=getGRSN() %>&gourl=<%=Server.URLEncode(gourl) %>";
	else if (msi == 3)
		mdel();
	else if (msi == 4)
		gosend();
	else if (msi == 5)
		document.getElementById("csvimp").style.display = "inline";
	else if (msi == 6)
	{
		document.form1.action = "expcsv.asp";
		document.form1.submit();
	}
<%
if is_domain_admin = true then
%>
	else if (msi == 7)
		conv2dmpads();
<%
end if

if isadmin() = true then
%>
	else if (msi == 8)
		conv2pads();
<%
end if
%>
}

function conv_upinfo() {
	var i = 0;
	var theObj;
	var conv_str = "\t";

	for(; i<<%=allnum %>; i++)
	{
		theObj = document.getElementById("checkdel" + i);

		if (theObj != null)
		{
			if (theObj.checked == true)
				conv_str = conv_str + theObj.value + '\t';
		}
	}

	document.getElementById("upinfo").value = conv_str;
}
// -->
</script>

<BODY onload="return window_onload()">
<FORM ENCTYPE="multipart/form-data" METHOD=POST NAME="fsa">
<table width="100%" border="0" align="center" bgcolor="white" cellspacing="0" style="margin-top:4px;">
	<tr><td class="tool_top_td">

	<span class="st_span"><span id="pm_moveto" class="menu_pop" style='width:60px; +width:63px; _width:60px;'>
	<div class='attbg'><img style='margin: 6px 0pt 0pt;' src='images/popshow.gif'></div>
	<div class='menu_pop_text'><%=s_lang_0321 %></div>
	</span></span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span"><span id="pm_bj" class="menu_pop" style='+width:69px; _width:63px;'>
	<div class='attbg'><img style='margin: 6px 0pt 0pt;' src='images/popshow.gif'></div>
	<div class='menu_pop_text'><%=s_lang_0312 %>...</div>
	</span></span>
	<span style='float:left; width:8px;'>&nbsp;</span>

	<span id="csvimp" class="st_span" style='display:none;'>
	<a class="btn_addPic" href="javascript:void(0);"><span><em>+</em><%=s_lang_0359 %></span> <input class="filePrew" tabindex="3" id="upfile" name="upfile" size="3" type="file" onchange="javascript:sub_up()"></a>
	</span>

	<span class="st_right_span" style='padding-top:3px; _padding-top:2px;'>
	<input type="text" id="query" onkeyup="sorter.searchhtml('query', 4, 6)" class='n_textbox' size="10">
	</span>
	<span style='float:right; width:1px;'>&nbsp;</span>

	<span class="st_right_span" style='padding-top:9px;'><%=s_lang_0318 %><%=s_lang_mh %></span>

	</td></tr>
</talbe>
</FORM>

<form action="ads_brow.asp" method=post id=form1 name=form1>
	<div style="display:none;">
	<select id="columns" style="display:none;"><option value="4"></option></select>
	</div>

	<table align="center" id="table" class="tinytable" border="0" cellpadding="0" cellspacing="0">
		<thead>
			<tr>
    <th width="4%" class="nosort"><h3><input type="checkbox" onclick="checkall(this)" style="margin-top:1px; _margin:-1px -2px 2px -2px;"></h3></th>
    <th width="4%" class="nosort"><h3>To</h3></th>
    <th width="4%" class="nosort"><h3>Cc</h3></th>
    <th width="4%" class="nosort"><h3>Bcc</h3></th>
    <th width="31%" noWrap><h3><%=s_lang_0360 %></h3></th>
    <th width="35%" noWrap><h3><%=s_lang_0320 %></h3></th>
    <th width="6%" noWrap class="nosort"><h3><%=s_lang_0361 %></h3></th>
    <th width="6%" noWrap class="nosort"><h3><%=s_lang_0093 %></h3></th>
    <th width="6%" noWrap class="nosort"><h3><%=s_lang_del %></h3></th>
			</tr>
		</thead>
		<tbody>
<%
i = 0
do while i < allnum
	ads.MoveTo i
	show_mail_function = " onclick=""location.href='ads_edit.asp?id=" & i & addsortstr & "&gourl=" & Server.URLEncode(gourl) & "&" & getGRSN() & "'"""

	Response.Write "<tr>"
	Response.Write "<td align='center' style='height:20px; border-bottom:1px solid #8CA5B5;'><input type='checkbox' id='checkdel" & i & "' name='checkdel" & i & "' value=""" & ads.nickname & """></td>"

	Response.Write "<td align='center' style='border-bottom:1px solid #8CA5B5;'><input type='checkbox' id='checkto" & i & "' name='checkto" & i & "' value=""" & ads.email & """></td>"
	Response.Write "<td align='center' style='border-bottom:1px solid #8CA5B5;'><input type='checkbox' id='checkcc" & i & "' name='checkcc" & i & "' value=""" & ads.email & """></td>"
	Response.Write "<td align='center' style='border-bottom:1px solid #8CA5B5;'><input type='checkbox' id='checkbcc" & i & "' name='checkbcc" & i & "' value=""" & ads.email & """></td>"

	Response.Write "<td align='center' style='cursor:pointer; word-break:break-all; word-wrap:break-word; padding-bottom:4px; _padding-bottom:1px; border-bottom:1px solid #8CA5B5;'" & show_mail_function & ">" & server.htmlencode(ads.nickname) & "</td>"
	Response.Write "<td align='left' style='cursor:pointer; word-break:break-all; word-wrap:break-word; padding-bottom:4px; _padding-bottom:1px; border-bottom:1px solid #8CA5B5;'" & show_mail_function & ">" & server.htmlencode(ads.email) & "</td>"

if wemcert.EmailExist(ads.email) = false then
	Response.Write "<td align='center' style='border-bottom:1px solid #8CA5B5;'>&nbsp;</td>"
else
	Response.Write "<td align='center' style='border-bottom:1px solid #8CA5B5;'><a href='cert_showcert.asp?email=" & Server.URLEncode(ads.email) & "&" & getGRSN() & "&gourl=" & Server.URLEncode(gourl) & "'><img src='images/s0.gif' border='0' title='" & s_lang_0362 & "'></a></td>"
end if

	Response.Write "<td align='center' style='border-bottom:1px solid #8CA5B5;'><a href='ads_edit.asp?id=" & i & addsortstr & "&gourl=" & Server.URLEncode(gourl) & "&" & getGRSN() & "'><img src='images/edit.gif' border='0' title='" & s_lang_modify & "'></a></td>"
	Response.Write "<td align='center' style='border-bottom:1px solid #8CA5B5;'><a href='ads_delete.asp?id=" & Server.URLEncode(ads.nickname) & "&gourl=" & Server.URLEncode(gourl) & "&" & getGRSN() & "'><img src='images/del.gif' border='0' title='" & s_lang_del & "'></a></td>"
	Response.Write "</tr>" & Chr(13)

    i = i + 1
loop
%>
</tbody>
</table>
<input type="hidden" name="wemid" value="<%=wemid %>">
<input type="hidden" name="addresssend" value="">
<input type="hidden" name="sc" value="<%=sc %>">
<input type="hidden" name="fl" value="<%=fl %>">
<input type="hidden" id="upinfo" name="upinfo" value="">
</FORM>

<div id="pmc_moveto" class="qmpanel_shadow" style="display:none; position:absolute;">
	<div class="menu_base">
	<div id="md_moveto" class="menu_bd bd">
	<table width="100%">
		<div name="mi" onclick="select_mto(1);" class="menu_item"><%=s_lang_0317 %></div>
		<div name="mi" onclick="select_mto(2);" class="menu_item"><%=s_lang_0322 %></div>
		<div name="mi" onclick="select_mto(3);" class="menu_item"><%=s_lang_0323 %></div>
	</table>
	</div>
	</div>
</div>

<div id="pmc_bj" class="qmpanel_shadow" style="display:none; position:absolute;">
	<div class="menu_base">
	<div id="lb_bj" class="menu_bd bd">
	<table width="100%">
		<div name="mi" onclick="select_mopt(2);" class="menu_item"><%=s_lang_0363 %></div>
		<div name="mi" onclick="select_mopt(3);" class="menu_item"><%=s_lang_0364 %></div>
		<div name="mi" onclick="select_mopt(4);" class="menu_item"><%=s_lang_0365 %></div>
		<div class="menu_item_nofun"><div style="background:#ccc; padding-top:1px; margin-top: 5px;"></div></div>
		<div name="mi" onclick="select_mopt(5);" class="menu_item"><%=s_lang_0366 %></div>
		<div name="mi" onclick="select_mopt(6);" class="menu_item"><%=s_lang_0367 %></div>
<%
if is_domain_admin = true or isadmin() = true then
%>
		<div class="menu_item_nofun"><div style="background:#ccc; padding-top:1px; margin-top: 5px;"></div></div>
<%
end if

if is_domain_admin = true then
%>
		<div name="mi" onclick="select_mopt(7);" class="menu_item"><%=s_lang_0368 %></div>
<%
end if

if isadmin() = true then
%>
		<div name="mi" onclick="select_mopt(8);" class="menu_item"><%=s_lang_0369 %></div>
<%
end if
%>
	</table>
	</div>
	</div>
</div>

<script type="text/javascript">
var mil = document.getElementsByTagName("div");
for (var i=0; i<mil.length; i++) 
{
	if (mil[i].name == "mi")
	{
		mil[i].onmouseover = function(){this.style.backgroundColor='#3470cc';this.style.color='#ffffff';}
		mil[i].onmouseout = function(){this.style.backgroundColor='#ffffff';this.style.color='#000000';}
	}
}

var is_in_menu_moveto = false;
var is_menu_show_moveto = false;
var my_menu_time_moveto;
var is_in_menu_bj = false;
var is_menu_show_bj = false;
var my_menu_time_bj;

$(function() {
	$("#pm_moveto").powerFloat({
		width: 120,
		eventType: "click",
		target: "#pmc_moveto",
		showCall: function() {
			if (is_menu_show_moveto == true)
				$.powerFloat.hide();
			else
			{
				is_menu_show_bj = false;

				is_menu_show_moveto = true;
				clearTimeout(my_menu_time_moveto);
			}

			$("#pmc_moveto").mouseover(function() {
				is_in_menu_moveto = true;
				clearTimeout(my_menu_time_moveto);
			});

			$("#pmc_moveto").mouseout(function() {
				is_in_menu_moveto = false;
				my_menu_time_moveto = setTimeout("setTimeClose(1)", 1000);
			});

			$(".menu_item").click(function() {
				$.powerFloat.hide();
			});
		},
		hideCall: function() {
			setTimeout("set_menu_close(1)", 300);
		}
	});
});

$(function() {
	$("#pm_bj").powerFloat({
		width: 166,
		eventType: "click",
		target: "#pmc_bj",
		showCall: function() {
			if (is_menu_show_bj == true)
				$.powerFloat.hide();
			else
			{
				is_menu_show_moveto = false;

				is_menu_show_bj = true;
				clearTimeout(my_menu_time_bj);
			}

			$("#pmc_bj").mouseover(function() {
				is_in_menu_bj = true;
				clearTimeout(my_menu_time_bj);
			});

			$("#pmc_bj").mouseout(function() {
				is_in_menu_bj = false;
				my_menu_time_bj = setTimeout("setTimeClose(2)", 1000);
			});

			$(".menu_item").click(function() {
				$.powerFloat.hide();
			});
		},
		hideCall: function() {
			setTimeout("set_menu_close(2)", 300);
		}
	});
});

function set_menu_close(tgv)
{
	if (tgv == 1)
		is_menu_show_moveto = false;
	else if (tgv == 2)
		is_menu_show_bj = false;
}

function setTimeClose(tgv)
{
	if (is_menu_show_moveto == true && is_in_menu_moveto == false && tgv == 1)
		$.powerFloat.hide();

	if (is_menu_show_bj == true && is_in_menu_bj == false && tgv == 2)
		$.powerFloat.hide();
}
</script>

<script type="text/javascript" src="images/slscript.js"></script>

<script type="text/javascript">
o_s_asp = parent.f1.document.leftval.s_asp.value;
o_s_col = parent.f1.document.leftval.s_col.value;
o_s_mode = parent.f1.document.leftval.s_mode.value;
o_s_search = parent.f1.document.leftval.s_search.value;

	var sorter = new TINY.table.sorter('sorter','table',{
		headclass:'head',
		ascclass:'asc',
		descclass:'desc',
		evenclass:'evenrow',
		oddclass:'oddrow',
		evenselclass:'evenselected',
		oddselclass:'oddselected',
		paginate:true,
		size:9999,
		colddid:'columns',
		hoverid:'selectedrow',
		init:true
	});

if (o_s_asp == "ads_brow.asp")
{
	sorter.sort(o_s_col);
	if (o_s_mode == "1")
		sorter.sort(o_s_col);

	document.getElementById('query').value = o_s_search;
}
else
{
	parent.f1.document.leftval.s_asp.value = "";
	parent.f1.document.leftval.s_col.value = "";
	parent.f1.document.leftval.s_mode.value = "";
	parent.f1.document.leftval.s_search.value = "";
	sorter.sort(4);
}

function _save_mode(s_col, s_mode) {
	parent.f1.document.leftval.s_asp.value = "ads_brow.asp";
	parent.f1.document.leftval.s_col.value = s_col;
	parent.f1.document.leftval.s_mode.value = s_mode;
}

function _save_search(s_search) {
	parent.f1.document.leftval.s_asp.value = "ads_brow.asp";
	parent.f1.document.leftval.s_search.value = s_search;
}

function window_onload() {
	if (document.getElementById('query').value.length > 0)
		sorter.searchhtml('query', 4, 6);
}
</script>

</BODY>
</HTML>

<%
set ads = nothing
set wemcert = nothing
%>
