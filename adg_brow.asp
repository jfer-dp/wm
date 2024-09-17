<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" --> 

<%
dim ads
set ads = server.createobject("easymail.Addresses")
ads.Load Session("wem")

i = 0
allnum = ads.GroupCount
gourl = "adg_brow.asp?" & getGRSN
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

<script type="text/javascript">
<!-- 
if (ie == 6)
	document.execCommand("BackgroundImageCache", false, true); 

parent.f1.clean_ads();

function checkthis(gindex)
{
	location.href = "wframe.asp?gindex=" + gindex + "<%=addsortstr %>&<%=getGRSN() %>&gourl=<%=Server.URLEncode(gourl) %>";
}

function mdel()
{
	if (ischeck() == true)
	{
		if (confirm("<%=s_lang_0115 %>") == false)
			return ;

		conv_upinfo();
		document.form1.action = "adg_delete.asp?<%=getGRSN() %>&mdel=1&gourl=<%=Server.URLEncode(gourl) %>";
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

function select_mto(msi)
{
	if (msi == 1)
		location.href = "ads_brow.asp?<%=getGRSN() %>";
	else if (msi == 2)
		location.href = "ads_pubbrow.asp?<%=getGRSN() %>";
	else if (msi == 3)
		location.href = "ads_dm_pubbrow.asp?<%=getGRSN() %>";
}

function select_mopt(msi)
{
	if (msi == 1)
		location.href = "adg_add.asp?<%=getGRSN() %>&gourl=<%=Server.URLEncode(gourl) %>";
	else if (msi == 2)
		mdel();
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
<FORM NAME="fsa">
<table width="100%" border="0" align="center" bgcolor="white" cellspacing="0" style="margin-top:4px;">
	<tr><td class="tool_top_td">

	<span class="st_span"><span id="pm_moveto" class="menu_pop" style='width:60px; +width:63px; _width:60px;'>
	<div class='attbg'><img style='margin: 6px 0pt 0pt;' src='images/popshow.gif'></div>
	<div class='menu_pop_text'><%=s_lang_0317 %></div>
	</span></span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span"><span id="pm_bj" class="menu_pop" style='+width:69px; _width:63px;'>
	<div class='attbg'><img style='margin: 6px 0pt 0pt;' src='images/popshow.gif'></div>
	<div class='menu_pop_text'><%=s_lang_0312 %>...</div>
	</span></span>
	<span style='float:left; width:8px;'>&nbsp;</span>

	<span class="st_right_span" style='padding-top:3px; _padding-top:2px;'>
	<input type="text" id="query" onkeyup="sorter.searchhtml('query', 1, 3)" class='n_textbox' size="10">
	</span>
	<span style='float:right; width:1px;'>&nbsp;</span>

	<span class="st_right_span" style='padding-top:9px;'><%=s_lang_0318 %><%=s_lang_mh %></span>

	</td></tr>
</talbe>
</FORM>

<form action="adg_brow.asp" method=post id=form1 name=form1>
	<div style="display:none;">
	<select id="columns" style="display:none;"><option value="1"></option></select>
	</div>

	<table align="center" id="table" class="tinytable" border="0" cellpadding="0" cellspacing="0">
		<thead>
			<tr>
    <th width="4%" class="nosort"><h3><input type="checkbox" onclick="checkall(this)" style="margin-top:1px; _margin:-1px -2px 2px -2px;"></h3></th>
    <th width="29%" noWrap><h3><%=s_lang_0319 %></h3></th>
    <th width="55%" noWrap><h3><%=s_lang_0320 %></h3></th>
    <th width="6%" noWrap class="nosort"><h3><%=s_lang_0093 %></h3></th>
    <th width="6%" noWrap class="nosort"><h3><%=s_lang_del %></h3></th>
			</tr>
		</thead>
		<tbody>
<%
i = 0
do while i < allnum
	ads.GetGroupInfo i, nickname, emails
	show_mail_function = " onclick=""checkthis(" & i & ")"""

	Response.Write "<tr>"
	Response.Write "<td align='center' style='height:20px; border-bottom:1px solid #8CA5B5;'><input type='checkbox' id='checkdel" & i & "' name='checkdel" & i & "' value=""" & nickname & """></td>"

	Response.Write "<td align='center' style='cursor:pointer; word-break:break-all; word-wrap:break-word; padding-bottom:4px; _padding-bottom:1px; border-bottom:1px solid #8CA5B5;'" & show_mail_function & ">" & server.htmlencode(nickname) & "</td>"
	Response.Write "<td align='left' style='cursor:pointer; word-break:break-all; word-wrap:break-word; padding-bottom:4px; _padding-bottom:1px; border-bottom:1px solid #8CA5B5;'" & show_mail_function & ">" & replace(server.htmlencode(emails), ",", "<br>") & "</td>"

	Response.Write "<td align='center' style='border-bottom:1px solid #8CA5B5;'><a href='adg_edit.asp?id=" & i & "&gourl=" & Server.URLEncode(gourl) & "&" & getGRSN() & "'><img src='images/edit.gif' border='0' title='" & s_lang_modify & "'></a></td>"
	Response.Write "<td align='center' style='border-bottom:1px solid #8CA5B5;'><a href='adg_delete.asp?id=" & Server.URLEncode(nickname) & "&gourl=" & Server.URLEncode(gourl) & "&" & getGRSN() & "'><img src='images/del.gif' border='0' title='" & s_lang_del & "'></a></td>"
	Response.Write "</tr>" & Chr(13)

	nickname = NULL
	emails = NULL

    i = i + 1
loop
%>
</table>
<input type="hidden" name="wemid" value="<%=wemid %>">
<input type="hidden" name="addresssend" value="">
<input type="hidden" name="sc" value="<%=sc %>">
<input type="hidden" id="upinfo" name="upinfo" value="">
</FORM>

<div id="pmc_moveto" class="qmpanel_shadow" style="display:none; position:absolute;">
	<div class="menu_base">
	<div id="md_moveto" class="menu_bd bd">
	<table width="100%">
		<div name="mi" onclick="select_mto(1);" class="menu_item"><%=s_lang_0321 %></div>
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
		<div name="mi" onclick="select_mopt(1);" class="menu_item"><%=s_lang_0324 %></div>
		<div name="mi" onclick="select_mopt(2);" class="menu_item"><%=s_lang_0325 %></div>
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
		width: 110,
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

if (o_s_asp == "adg_brow.asp")
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
	sorter.sort(1);
}

function _save_mode(s_col, s_mode) {
	parent.f1.document.leftval.s_asp.value = "adg_brow.asp";
	parent.f1.document.leftval.s_col.value = s_col;
	parent.f1.document.leftval.s_mode.value = s_mode;
}

function _save_search(s_search) {
	parent.f1.document.leftval.s_asp.value = "adg_brow.asp";
	parent.f1.document.leftval.s_search.value = s_search;
}

function window_onload() {
	if (document.getElementById('query').value.length > 0)
		sorter.searchhtml('query', 1, 3);
}
</script>

</BODY>
</HTML>

<%
set ads = nothing
%>
