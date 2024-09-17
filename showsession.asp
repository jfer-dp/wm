<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" --> 

<%
gourl = trim(request("gourl"))
enc_gourl = Server.URLEncode(gourl)
msid = trim(request("msid"))
enc_rqs = Server.URLEncode(trim(Request.QueryString))

dim mlb
set mlb = server.createobject("easymail.Labels")
mlb.Load Session("wem")

dim ei
set ei = server.createobject("easymail.InfoList")
ei.LoadOneSession Session("wem"), msid

allnum = ei.getMailsCount

ei.getMailInfoEx allnum - 1, idname, isread, priority, sendMail, sendName, subject, size, etime, mstate
dim top_idname
top_idname = idname
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
<link rel="stylesheet" type="text/css" href="images/showsession.css">
</HEAD>

<script type="text/javascript" src="images/mglobal.js"></script>
<script type="text/javascript" src="images/jquery.min.js"></script>
<script type="text/javascript" src="images/jquery-powerFloat-min.js"></script>

<script type="text/javascript">
if (ie == 6)
	document.execCommand("BackgroundImageCache", false, true);

function window_onload() {
	if (parent.f1.document.leftval.temp.value.length < 1 || need_re_getnp() == true)
		ajax_getnp_info();
	else
		writepn();
}

function ajax_getnp_info() {
	var post_date = "mode=2&<%=getGRSN() %>" + parent.f1.document.leftval.sortinfo.value + "&fid=<%=msid %>";

$.ajax({
	type:"POST",
	url:"getnpinfo.asp",
	data:post_date,
	success:function(data){
		parent.f1.document.leftval.temp.value = data;
		writepn();
	},
	error:function(){
	}
});
}

function show_me(show_index, mid)
{
	if (document.getElementById("iframepage_" + show_index).src.length < 1)
	{
		document.body.style.cursor = "wait";
		document.getElementById("iframepage_" + show_index).src = "showmail.asp?inline=1&filename=" + mid + "&inlineid=" + show_index + "&<%=getGRSN() %>&gourl=<%=enc_gourl %>";
	}
	else
		document.body.style.cursor = "auto";

	var theObj = document.getElementById("msg_" + show_index);
	if (theObj.style.display == "inline")
		theObj.style.display = "none";
	else
		theObj.style.display = "inline";
}

function iFrameHeight(show_index) {
	var ifm = document.getElementById("iframepage_" + show_index);
	var subWeb = document.frames ? document.frames["iframepage_" + show_index].document : ifm.contentDocument;
	if(ifm != null && subWeb != null) {
		ifm.height = subWeb.body.scrollHeight;
	}
	document.body.style.cursor = "auto";
}

function showlabel(lb_id) {
	location.href = "listlabel.asp?<%=getGRSN() %>&lbid=" + lb_id;
}

function out_label(tglbid)
{
	var tgobj = document.getElementById("lbc_" + tglbid);
	if (tgobj != null)
	{
		tgobj.style.background = "";
		tgobj.style.width = "0px";
		tgobj.style.display = "";
	}
}

function up_label(tglbid)
{
	var tgobj = document.getElementById("lbc_" + tglbid);
	if (tgobj != null)
	{
		tgobj.style.background = "url('images/lbclose.gif')";
		tgobj.style.backgroundRepeat = "no-repeat";
		tgobj.style.backgroundPosition = "right center";
		tgobj.style.width = "19px";
		tgobj.style.display = "inline-block";
	}
}

function lbclose(tglbid, e)
{
	if (!e) var e = window.event;
	e.cancelBubble = true;
	if (e.stopPropagation)
	e.stopPropagation();

	var post_date = "setmode=2&<%=getGRSN() %>&lbid=" + tglbid + "&filename=<%=top_idname %>";

$.ajax({
	type:"POST",
	url:"showmail.asp",
	data:post_date,
	success:function(data){
		if (data != "1")
			document.location.reload(true);
		else
			$("#lba_" + tglbid).remove();
	},
	error:function(){
	}
});
}

function set_one_star(bj)
{
	var post_date = "setmode=1&<%=getGRSN() %>&bj=" + bj + "&filename=<%=top_idname %>";

$.ajax({
	type:"POST",
	url:"showmail.asp",
	data:post_date,
	success:function(data){
		if (data != "1")
			document.location.reload(true);
		else
		{
			var theObj = document.getElementById("icn_star");
			if (theObj != null)
			{
				if (bj == '1')
				{
					theObj.src = "images/star_yes.gif";
					theObj.title = "<%=b_lang_127 %>";
					theObj.onclick = new Function("set_one_star('0')");
				}
				else
				{
					theObj.src = "images/star_no.gif";
					theObj.title = "<%=b_lang_128 %>";
					theObj.onclick = new Function("set_one_star('1')");
				}
			}
		}
	},
	error:function(){
	}
});
}

function set_read(bj)
{
	var post_date = "setmode=5&<%=getGRSN() %>&bj=" + bj + "&filename=<%=top_idname %>";

$.ajax({
	type:"POST",
	url:"showmail.asp",
	data:post_date,
	success:function(data){
		alert_msg("<%=b_lang_104 %>");
	},
	error:function(){
	}
});
}

function get_lb_str(lb_id, lb_title, lb_color)
{
	return "<span id=\"lba_" + lb_id + "\"><span id=\"lb_" + lb_id + "\" class=\"wwm_lb_box\" style=\"background:#" + lb_color + ";\"\
 onclick=\"showlabel('" + lb_id + "')\" onmouseover=\"up_label('" + lb_id + "')\" onmouseout=\"out_label('" + lb_id + "')\"><span class=\"wwm_lb_text\">" + htmlEscape(lb_title) + "</span>\
<span id=\"lbc_" + lb_id + "\" class=\"wwm_lb_close\" onclick=\"lbclose('" + lb_id + "', event)\">&nbsp;</span></span>\r\n</span>";
}

function set_lb(tgid, tgcolor, tgtitle)
{
	var post_date = "setmode=6&<%=getGRSN() %>&lbid=" + tgid + "&filename=<%=top_idname %>";

$.ajax({
	type:"POST",
	url:"showmail.asp",
	data:post_date,
	success:function(data){
		if (data != "1")
			document.location.reload(true);
		else
		{
			var fdid = document.getElementById("lba_" + tgid);
			if (fdid == null)
				$("#labels_td").append(get_lb_str(tgid, tgtitle, tgcolor));
		}
	},
	error:function(){
	}
});
}

function delleftfilename(fname) {
	var s = parent.f1.document.leftval.temp.value;

	var sp = s.indexOf('\t' + fname);

	var newval;
	newval = s.substring(0, sp);
	newval = newval + s.substring(sp + fname.length + 1);

	parent.f1.document.leftval.temp.value = newval;
}

var pdeladd;
function prevnext(isnext){
	var s,ss;
	s = parent.f1.document.leftval.temp.value;
	ss = s.split("\t");

	var i;

	for(i = 0; i < ss.length; i++)
	{
		if (ss[i] == "<%=msid %>")
			break;
	}

	if (i < 0 || i >=ss.length)
		return;

	var mprev = "";
	var mnext = "";

	if (i > 0)
		mprev = ss[i - 1];

	if (i+1 < ss.length)
		mnext = ss[i + 1];

	if (isnext == '1' && mnext != "" && mnext != "|")
		location.href = "showsession.asp?msid=" + mnext + "&<%=getGRSN() %>&gourl=<%=enc_gourl %>";

	if (isnext == '0' && mprev != "" && mprev != "|")
		location.href = "showsession.asp?msid=" + mprev + "&<%=getGRSN() %>&gourl=<%=enc_gourl %>";

	if (mnext != "" && mnext != "|")
		pdeladd = "&gourl=<%=enc_gourl %>&nextfile=" + mnext;
	else
		pdeladd = "&gourl=<%=enc_gourl %>&nextfile=" + mprev;
}

function need_re_getnp() {
	var ss = parent.f1.document.leftval.temp.value.split("\t");

	if (ss.length > 0)
	{
		if (ss[0] == "<%=msid %>" || ss[ss.length - 1] == "<%=msid %>")
			return true;
	}

	return false;
}

function writepn() {
	var s,ss;
	s = parent.f1.document.leftval.temp.value;
	ss = s.split("\t");

	var i;

	for(i = 0; i < ss.length; i++)
	{
		if (ss[i] == "<%=msid %>")
			break;
	}

	if (i < 0 || i >=ss.length)
		return;

	var mprev = "";
	var mnext = "";

	if (i > 0)
		mprev = ss[i - 1];

	if (i+1 < ss.length)
		mnext = ss[i + 1];

	var writemprev = "";
	var writemnext = "";

	if (mprev != "" && mprev != "|")
		writemprev = "<a href=\"javascript:prevnext('0')\"><%=b_lang_129 %></a> ";
	else
		writemprev = "<font color='#a0a0a0'><%=b_lang_129 %></font> ";

	if (mnext != "" && mnext != "|")
		writemnext = "<a href=\"javascript:prevnext('1')\"><%=b_lang_130 %></a>";
	else
		writemnext = "<font color='#a0a0a0'><%=b_lang_130 %></font>";

	$("#top_pn").html(writemprev + writemnext);
	$("#bottom_pn").html(writemprev + writemnext);
}

function back() {
<% if gourl = "" then %>
	history.back();
<% else %>
	location_href("<%=gourl %>&<%=getGRSN() %>");
<% end if %>
}

function location_href(url) {
	location.href = url;
}

function delthis() {
	prevnext();
	delleftfilename("<%=msid %>");
<%
	if Session("delProc") = 0 then
%>
	location_href("delmail.asp?msid=<%=msid %>&<%=getGRSN() %>&gourl=<%=enc_gourl %>");
<%
	else
%>
	if (parent.f1.document.leftval.temp.value != "|\t|")
		location_href("delmail.asp?msid=<%=msid %>&<%=getGRSN() %>" + pdeladd);
	else
		location_href("delmail.asp?msid=<%=msid %>&<%=getGRSN() %>&gourl=<%=enc_gourl %>");
<%
	end if
%>
}

function realdelthis() {
	if (confirm("<%=b_lang_131 %>") == false)
		return ;

	prevnext();
	delleftfilename("<%=msid %>");
<%
	if Session("delProc") = 0 then
%>
	location_href("delmail.asp?realdel=1&msid=<%=msid %>&<%=getGRSN() %>&gourl=<%=enc_gourl %>");
<%
	else
%>
	if (parent.f1.document.leftval.temp.value != "|\t|")
		location_href("delmail.asp?realdel=1&msid=<%=msid %>&<%=getGRSN() %>" + pdeladd);
	else
		location_href("delmail.asp?realdel=1&msid=<%=msid %>&<%=getGRSN() %>&gourl=<%=enc_gourl %>");
<%
	end if
%>
}

function movemail(tgname) {
	prevnext();
	delleftfilename("<%=msid %>");
<%
if Session("delProc") = 0 then
%>
	location_href("delmail.asp?msid=<%=msid %>&mto=" + tgname + "&<%=getGRSN() %>&gourl=<%=enc_gourl %>");
<% else %>
	if (parent.f1.document.leftval.temp.value != "|\t|")
		location_href("delmail.asp?msid=<%=msid %>&<%=getGRSN() %>&mto=" + tgname + pdeladd);
	else
		location_href("delmail.asp?msid=<%=msid %>&mto=" + tgname + "&<%=getGRSN() %>&gourl=<%=enc_gourl %>");
<% end if %>
}

function alert_msg(amsg) {
	$("#top_show_msg").text(amsg);
	document.getElementById("top_show_msg").style.display = "inline";
	setTimeout("close_alert()", 3000);
}

function close_alert() {
	document.getElementById("top_show_msg").style.display = "none";
}
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
<a name="gotop" style="font-size:0pt; height:0px;"></a>
<table class="table_main" align="center" cellspacing="0" cellpadding="0">
	<tr><td class="block_top_td"><div class="table_min_width"></div></td></tr>
	<tr><td class="tool_top_td">

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:back()"><< <%=s_lang_return %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=reply&<%=getGRSN() %>&filename=<%=top_idname & "&backurl=showsession.asp?" & enc_rqs %>"><%=b_lang_132 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=replyall&<%=getGRSN() %>&filename=<%=top_idname & "&backurl=showsession.asp?" & enc_rqs %>"><%=b_lang_133 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

<%
dim emmail
set emmail = server.createobject("easymail.emmail")
emmail.LoadAll Session("wem"), top_idname

if emmail.IsSaveMail = false then
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=forward&<%=getGRSN() %>&filename=<%=top_idname & "&backurl=showsession.asp?" & enc_rqs %>"><%=b_lang_134 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
else
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=forward&<%=getGRSN() %>&filename=<%=top_idname & "&backurl=showsession.asp?" & enc_rqs %>"><%=b_lang_135 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
end if
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:delthis();"><%=s_lang_del %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:realdelthis();"><%=b_lang_136 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span"><span id="pm_moveto" class="menu_pop"<%=b_lang_156 %>>
	<div class='attbg'><img style='margin: 6px 0pt 0pt;' src='images/popshow.gif'></div>
	<div class='menu_pop_text'><%=b_lang_137 %>...</div>
	</span></span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span"><span id="pm_bj" class="menu_pop"<%=b_lang_157 %>>
	<div class='attbg'><img style='margin: 6px 0pt 0pt;' src='images/popshow.gif'></div>
	<div class='menu_pop_text'><%=b_lang_138 %>...</div>
	</span></span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span id="top_pn" class="st_right_span">
	</span>

	</td></tr>
<%
dim exMailInfo
set exMailInfo = server.createobject("easymail.ExMailInfo")

exMailInfo.Load Session("wem"), top_idname
%>
	<tr class="head_tr"><td class="head_subject_td" style="border-top:1px solid #aac1de;">
	<table width="100%" border="0" cellspacing="0" align="center"><tr><td id="labels_td">
	<span class="head_subject_span"><%=server.htmlencode(emmail.subject) %></span>
<%
	Response.Write " <span class='head_star_span'>"
	if exMailInfo.Have_Star = true then
		Response.Write "<img id='icn_star' src='images/star_yes.gif' border='0' style='cursor:pointer;' title='" & b_lang_127 & "' onclick=""set_one_star('0');""></a>"
	else
		Response.Write "<img id='icn_star' src='images/star_no.gif' border='0' style='cursor:pointer;' title='" & b_lang_128 & "' onclick=""set_one_star('1');""></a>"
	end if
	Response.Write "</span>"

if exMailInfo.LabelCount > 0 then
	lball = exMailInfo.LabelCount
	lbi = 0
	do while lbi < lball
		mlb.GetByID exMailInfo.GetLabel(lbi), ret_id, ret_title, ret_color

		Response.Write create_label_str(ret_id, ret_title, ret_color) & Chr(13)
		ret_id = NULL
		ret_title = NULL
		ret_color = NULL

		lbi = lbi + 1
	loop
end if
%>
	</td><td width="30" nowrap>
	<span class='wwm_color_in_head'><%=allnum %></span>
	</td></tr></table>
	</td></tr>
	<tr class="head_tr"><td class="head_td">
<%
	idname = NULL
	isread = NULL
	priority = NULL
	sendMail = NULL
	sendName = NULL
	subject = NULL
	size = NULL
	etime = NULL
	mstate = NULL

show_i = 0
i = 0
do while i < allnum
	ei.getMailInfoEx allnum - i - 1, idname, isread, priority, sendMail, sendName, subject, size, etime, mstate

	Response.Write "<div class='qmpanel_shadow'><div class='menu_base'><div class='menu_sesmsg bd'><div class='head_msg' onclick=""show_me('" & show_i & "', '" & idname & "')"">"
	Response.Write "<table align='center' cellspacing='0' cellpadding='0' width='100%'><tr style='cursor:pointer;'><td nowrap height='30' width='5%' align='left' style='padding-left:12px; padding-right:8px;'>"

	if mstate = 0 then
		Response.Write "<img id='icn_b_" & li & "' src='reply.gif' title='" & b_lang_139 & "' align='absmiddle' border='0'"
	elseif mstate = 1 then
		Response.Write "<img id='icn_b_" & li & "' src='forward.gif' title='" & b_lang_140 & "' align='absmiddle' border='0'"
	elseif mstate = 2 or mstate = 3 then
		if isread = true then
			Response.Write "<img id='icn_b_" & li & "' src='rsysmail.gif' title='" & b_lang_141 & "' align='absmiddle' border='0'"
		else
			Response.Write "<img id='icn_b_" & li & "' src='nsysmail.gif' title='" & b_lang_142 & "' align='absmiddle' border='0'"
		end if
	else
		if isread = true then
			Response.Write "<img id='icn_b_" & li & "' src='mail.gif' title='" & b_lang_143 & "' align='absmiddle' border='0'"
		else
			Response.Write "<img id='icn_b_" & li & "' src='newmail.gif' title='" & b_lang_144 & "' align='absmiddle' border='0'"
		end if
	end if

	if issign = true then
		if isread = true then
			Response.Write "><img id='icn_e_" & li & "' src='images/s0-1.gif' title='" & b_lang_145 & "' align='absmiddle' border='0'"
		else
			Response.Write "><img id='icn_e_" & li & "' src='images/s0.gif' title='" & b_lang_145 & "' align='absmiddle' border='0'"
		end if
	elseif isenc = true then
		if isread = true then
			Response.Write "><img id='icn_e_" & li & "' src='images/e0-1.gif' title='" & b_lang_146 & "' align='absmiddle' border='0'"
		else
			Response.Write "><img id='icn_e_" & li & "' src='images/e0.gif' title='" & b_lang_146 & "' align='absmiddle' border='0'"
		end if
	end if

	Response.Write "</td><td nowrap align='left' width='60%' style='padding-left:8px;'>"

	if Session("mail") = sendMail then
		Response.Write "<font color='#5fa207' style='font-weight:bold;'>" & b_lang_147 & "</font>"
	else
		Response.Write server.htmlencode(sendName & " <" & sendMail & ">")
	end if

	Response.Write "</td><td nowrap align='right' width='10%' style='padding-right:8px; color:#444;'>"
	Response.Write getShowSize(size)

	Response.Write "</td><td nowrap align='right' width='25%' style='padding-right:12px; color:#444;'>"
	Response.Write etime

	Response.Write "</td></tr></table>"

	Response.Write "</div><div id='msg_" & show_i & "' style='display:none; padding:0px; background:#efefef;'>"
	Response.Write "<iframe id='iframepage_" & show_i & "' name='iframepage_" & show_i & "' frameBorder=0 scrolling=no width='100%'></iframe>"
	Response.Write "</div></div></div></div><br>" & Chr(13)

	idname = NULL
	isread = NULL
	priority = NULL
	sendMail = NULL
	sendName = NULL
	subject = NULL
	size = NULL
	etime = NULL
	mstate = NULL

	show_i = show_i + 1
	i = i + 1
loop
%>
	</td></tr>

	<tr><td class="block_top_td" style="height:10px;"></td></tr>
	<tr><td class="tool_top_td">

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:back()"><< <%=s_lang_return %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=reply&<%=getGRSN() %>&filename=<%=top_idname & "&backurl=showsession.asp?" & enc_rqs %>"><%=b_lang_132 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=replyall&<%=getGRSN() %>&filename=<%=top_idname & "&backurl=showsession.asp?" & enc_rqs %>"><%=b_lang_133 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
if emmail.IsSaveMail = false then
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=forward&<%=getGRSN() %>&filename=<%=top_idname & "&backurl=showsession.asp?" & enc_rqs %>"><%=b_lang_134 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
else
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=forward&<%=getGRSN() %>&filename=<%=top_idname & "&backurl=showsession.asp?" & enc_rqs %>"><%=b_lang_135 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
end if
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:delthis();"><%=s_lang_del %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:realdelthis();"><%=b_lang_136 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span id="bottom_pn" class="st_right_span">
	</span>

	<tr><td class="block_top_td" style="height:16px;"></td></tr>
	<tr><td align="right">
	<span style="margin-right:16px;"><a href="#gotop"><img src='images\gotop.gif' border='0' title="<%=b_lang_125 %>"></a></span>
	</td></tr>
	</table>

<div id="top_show_msg" class="top_show_msg"></div>

<div id="pmc_moveto" class="qmpanel_shadow" style="display:none; position:absolute;">
	<div class="menu_base">
	<div id="md_moveto" class="menu_bd bd">
	<table width="100%">
		<div name="mi" onclick="movemail('del');" class="menu_item"><%=b_lang_148 %></div>
		<div name="mi" onclick="movemail('in');" class="menu_item"><%=b_lang_149 %></div>
		<div name="mi" onclick="movemail('out');" class="menu_item"><%=b_lang_150 %></div>
		<div name="mi" onclick="movemail('sed');" class="menu_item"><%=b_lang_151 %></div>
<%
dim pf
set pf = server.createobject("easymail.PerFolders")
pf.Load Session("wem")

pfNumber = pf.FolderCount

if pfNumber > 0 then
	Response.Write "<div class='menu_item_nofun'><div style='background:#ccc; padding-top:1px; margin-top: 5px;'></div></div>"
end if

dim moveto_set_max
moveto_set_max = false

if pfNumber > 6 then
	moveto_set_max = true
end if

dim moveto_max_len
moveto_max_len = 0
i = 0
do while i < pfNumber
	spfname = pf.GetFolderName(i)

	t_len = getLength(spfname)
	if t_len > moveto_max_len then
		moveto_max_len = t_len
	end if

	Response.Write "<div name='mi' onclick=""movemail('" & pf.GetFolderID(spfname) & "');"" class='menu_item'>" & server.htmlencode(spfname) & "</div>"
	spfname = NULL

	i = i + 1
loop
set pf = nothing
%>
	</table>
	</div>
	</div>
</div>

<div id="pmc_bj" class="qmpanel_shadow" style="display:none; position:absolute;">
	<div class="menu_base">
	<div id="lb_bj" class="menu_bd bd">
	<table width="100%">
		<div name="mi" onclick="set_read('1');" class="menu_item"><%=b_lang_152 %></div>
		<div name="mi" onclick="set_read('0');" class="menu_item"><%=b_lang_153 %></div>
		<div class="menu_item_nofun"><div style="background:#ccc; padding-top:1px; margin-top: 5px;"></div></div>
		<div name="mi" onclick="set_one_star('1');" class="menu_item"><%=b_lang_154 %></div>
		<div name="mi" onclick="set_one_star('0');" class="menu_item"><%=b_lang_155 %></div>
<%
dim bj_set_max
dim bj_lb_max_len_title
bj_set_max = false
bj_lb_max_len_title = 0

allnum = mlb.Count

if allnum > 6 then
	bj_set_max = true
end if

if allnum > 0 then
	Response.Write "		<div class='menu_item_nofun'><div style='background:#ccc; padding-top:1px; margin-top: 5px;'></div></div>"

	i = 0
	do while i < allnum
		mlb.GetByIndex i, ret_id, ret_title, ret_color

		t_len = getLength(ret_title)
		if t_len > bj_lb_max_len_title then
			bj_lb_max_len_title = t_len
		end if
		Response.Write "<div onclick=""set_lb('" & ret_id & "', '" & ret_color & "', '" & server.htmlencode(ret_title) & "');"" name='mi' class='menu_item'><span class='wwm_color_in_line' style='background:#" & ret_color & ";'>&nbsp;</span> " & server.htmlencode(ret_title) & "</div>" & Chr(13)

		ret_id = NULL
		ret_title = NULL
		ret_color = NULL

		i = i + 1
	loop
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
<%
if moveto_set_max = true then
	Response.Write "	$(""#md_moveto"").height(226);"
end if
%>
	$("#pm_moveto").powerFloat({
<%
if moveto_max_len > 10 then
	Response.Write "		width: " & (moveto_max_len * 6) + 38 & ","
else
	Response.Write "		width: 110,"
end if
%>
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
<%
if bj_set_max = true then
	Response.Write "	$(""#lb_bj"").height(232);"
end if
%>
	$("#pm_bj").powerFloat({
<%
if bj_lb_max_len_title > 10 then
	Response.Write "		width: " & (bj_lb_max_len_title * 6) + 60 & ","
else
	Response.Write "		width: 110,"
end if
%>
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

$(function(){
	$("a").each(function ()
	{
		var link = $(this);
		var href = link.attr("href");
		if(href && href[0] == "#")
		{
			var name = href.substring(1);
			$(this).click(function()
			{
				var nameElement = $("[name='"+name+"']");
				var idElement = $("#"+name);
				var element = null;
				if(nameElement.length > 0) {
					element = nameElement;
				} else if(idElement.length > 0) {
					element = idElement;
				}

				if(element)
				{
					var offset = element.offset();
					window.scrollTo(offset.left, offset.top);
				}

				return false;
			});
		}
	});
});
</script>
</BODY>
</HTML>

<%
top_idname = NULL

set exMailInfo = nothing
set emmail = nothing
set ei = nothing
set mlb = nothing


function getShowSize(bytesize)
	if bytesize < 1000 then
		getShowSize = "1K"
	else
		if bytesize < 1000000 then
			tmpSize = CDbl(bytesize/1000)
			tmpindex = InStr(1, tmpSize, ".")
			if tmpindex = 0 then
				getShowSize = tmpSize & "K"
			else
				getShowSize = CDbl(Left(tmpSize, tmpindex + 1)) & "K"
			end if
		else
			tmpSize = CStr(CDbl(bytesize/1000000))
			tmpindex = InStr(1, tmpSize, ".")
			if tmpindex = 0 then
				getShowSize = tmpSize & "M"
			else
				getShowSize = CDbl(Left(tmpSize, tmpindex + 2)) & "M"
			end if
		end if
	end if
end function

function create_label_str(nid, ret_title, ret_color)
	create_label_str = "<span id=""lba_" & nid & """> <span id=""lb_" & nid & """ class=""wwm_lb_box"" style=""background:#" & ret_color & ";"" onclick=""showlabel('" & nid & "')"" onmouseover=""up_label('" & nid & "')"" onmouseout=""out_label('" & nid & "')""><span class=""wwm_lb_text"">" & server.htmlencode(ret_title) & "</span><span id=""lbc_" & nid & """ class=""wwm_lb_close"" onclick=""lbclose('" & nid & "', event)"">&nbsp;</span></span>" & Chr(13) & Chr(10) & "</span>"
end function

function getLength(txt)
	txt=trim(txt)
	x = len(txt)
	y = 0
	for ii = 1 to x
		if asc(mid(txt,ii,1))<0 or asc(mid(txt,ii,1))>255 then
			y = y + 2
		else
			y = y + 1
		end if
	next
	getLength= y
end function
%>
