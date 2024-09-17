<!--#include file="passinc.asp" -->
<!--#include file="language.asp" -->

<%
Session("ac_ads_number") = ""

dim pf
set pf = server.createobject("easymail.PerFolders")
pf.Load Session("wem")

dim paf
set paf = server.createobject("easymail.PerAttFolders")
paf.Load Session("wem")

dim isdomainmanager
dim eu
set eu = Application("em")
isdomainmanager = eu.IsDomainManager(Session("wem"))
set eu = nothing

dim eads
set eads = server.createobject("easymail.EntAddress")
eads.Load
eads_allnum = eads.Count
set eads = nothing

dim mlb
set mlb = server.createobject("easymail.Labels")
mlb.Load Session("wem")
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<style type="text/css">
<!--
html	{overflow-x:hidden; overflow-y:hidden;}
body	{font-family:<%=s_lang_font %>; font-size:9pt; color:#1e5494; background-color:#89b5e9; margin-top:4px; margin-bottom:4px;}
a:hover {color:#1e5494; text-decoration:none}
a		{color:#1e5494; text-decoration:none}

.menu_item{line-height:22px;}
.menu_item:hover {background-color:#e0ecf9;} 
.bd{border-color:#aaa;}

.e_own {border:0px; background:url(images/expand_all.gif) no-repeat 0px 0px; width:9px; height:9px;}
.c_own {border:0px; background:url(images/collapse_all.gif) no-repeat 0px 0px; width:9px; height:9px;}

.menu_bd{overflow-x:hidden;overflow-y:auto;background:#fff;padding:4px 0;border:1px solid #6392c8;border-radius:5px; _padding-bottom:2px;}
.menu_item,.menu_item_high,.menu_item_nofun{padding:1px 12px 0 12px; white-space:nowrap; text-overflow:ellipsis;cursor:pointer; margin:0 4px -1px 4px; overflow:hidden; width:106px; text-overflow:ellipsis; -o-text-overflow:ellipsis;}
.menu_item_nofun {margin:0 -6px 0 -6px; width:118px;}
.menu_item_nofun_con {background:#ccc; padding-top:1px; margin:1px -10px 1px -2px;}
.menu_item_ec {float:left; margin-left:-11px; margin-top:0px; *margin-top:5px; width:11px;}
.menu_item_ec_text {float:left; margin-top:0px;}
.menu_item_nofun{color:#7b7b7b;cursor:pointer;}
.menu_item{background:#fff}
.menu_item_high{background:#3470cc;color:#fff;}
.menu_item_high a,.menu_item_high a:link,.menu_item_high a:visited,.menu_item_high a:active{display:block;color:#fff;text-decoration:none;}
.mailgroup_member .menu_item_nofun .bold{_padding-top:10px;}
.tips_maliciousLink .menu_bd{background-color:#FFE6E6;border-color:#E2AEAE;}
.tips_unknowLink .menu_bd{background-color:#fff9e3;border-color:#e9c968;box-shadow:0 1px 2px rgba(0,0,0,0.2);}
.tips_maliciousLink .icon_caution_s,.tips_unknowLink .icon_info_s{display:inline-block;vertical-align:middle;margin:-1px 5px 0 -4px;_vertical-align:text-bottom;_margin:5px 5px 0 -4px;}
.tips_maliciousLink .menu_item_nofun,.tips_unknowLink .menu_item_nofun{color:#000000;}
.menu_item .arrow_meunico{width:12px;height:12px;margin:0 -5px 0 0;*margin:5px -5px 0 0;vertical-align:middle;}
.wwm_color_in_line{padding:1px; border-radius:3px; -webkit-border-radius:3px; display:-moz-inline-box; display:inline-block; padding-left:2px; padding-right:2px; height:8px; width:6px; font-size:0pt;}
-->
</style>
</HEAD>

<script type="text/javascript" src="images/mglobal.js"></script>

<script language="JavaScript">
<!--
var asp2id_array = new Array(["labels.asp", "p_labels"], ["listsesmail.asp", "listmail_in"], ["writemail.asp", "id_wframe"], ["viewmailbox.asp", "p_listmail"], ["attfolders.asp", "p_attfolders"]
, ["adg_brow.asp", "id_ads_brow"], ["ads_pubbrow.asp", "id_ads_brow"], ["ads_dm_pubbrow.asp", "id_ads_brow"], ["cert_share.asp", "id_cert_index"], ["searchlistmail.asp", "id_findmail"]);

function find_id(aspfile) {
	for (var i = 0; i < asp2id_array.length; i++)
	{
		if (asp2id_array[i][0] == aspfile.toLowerCase())
			return asp2id_array[i][1];
	}

	return "";
}

function window_onload() {
<%
asp = trim(request("asp"))

if asp = "" then
%>
	select_id_cg_color("p_listmail");
<%
else
%>
	select_one("<%=asp %>", "");
<%
end if
%>
}

var fed_scrollTop = 0;
function time_scroll() {
	if (ie != 6 || fed_scrollTop < 2)
		return ;

	document.getElementById("full_ec_div").scrollTop = fed_scrollTop;
	fed_scrollTop = 0;
}

function theright(rurl, onlyone, tgobj) {
	if (ie == 6)
		fed_scrollTop = document.getElementById("full_ec_div").scrollTop;

	select_cg_color(tgobj);

	var theObj
	if (rurl == "viewmailbox.asp")
		show_folder(1, -1, -1, -1);
	else if (rurl == "labels.asp")
		show_folder(-1, 1, -1, -1);
	else if (rurl == "attfolders.asp")
		show_folder(-1, -1, 1, -1);

	check_div_height();
	var mrstr = String(Math.random());

	if (onlyone == true)
		parent.f2.window.location.href = rurl + "?GRSN=" + mrstr.substring(2, 10);
	else
		parent.f2.window.location.href = rurl + "&GRSN=" + mrstr.substring(2, 10);

	if (ie == 6)
		setTimeout("time_scroll()", 1);
}

function show_folder(s_listmail, s_labels, s_attfolders, s_other)
{
	if (s_listmail > -1)
		show_folder_one(s_listmail, document.getElementById("chd_listmail"), document.getElementById("ec_listmail"));

	if (s_labels > -1)
		show_folder_one(s_labels, document.getElementById("chd_labels"), document.getElementById("ec_labels"));

	if (s_attfolders > -1)
		show_folder_one(s_attfolders, document.getElementById("chd_attfolders"), document.getElementById("ec_attfolders"));

	if (s_other > -1)
		show_folder_one(s_other, document.getElementById("chd_other"), document.getElementById("ec_other"));
}

function show_folder_one(s_f, chd_obj, ec_obj)
{
	if (chd_obj != null && ec_obj != null)
	{
		if (s_f == 0)
		{
			ec_obj.innerHTML = "<img src='images/null.gif' class='e_own'>";
			chd_obj.style.display = "none";
		}
		else if (s_f == 1)
		{
			ec_obj.innerHTML = "<img src='images/null.gif' class='c_own'>";
			chd_obj.style.display = "inline";
		}
		else if (s_f == 2)
		{
			if (chd_obj.style.display == "inline")
			{
				ec_obj.innerHTML = "<img src='images/null.gif' class='e_own'>";
				chd_obj.style.display = "none";
			}
			else
			{
				ec_obj.innerHTML = "<img src='images/null.gif' class='c_own'>";
				chd_obj.style.display = "inline";
			}
		}
	}
}

function stop_event(e) {
	if (!e) var e = window.event;
	e.cancelBubble = true;
	if (e.stopPropagation)
	e.stopPropagation();
}

function show_listmail(e) {
	stop_event(e);
	show_folder(2, -1, -1, -1);

	clean_ie6_mout();
	check_div_height();
}

function show_labels(e) {
	if (ie == 6)
		fed_scrollTop = document.getElementById("full_ec_div").scrollTop;

	stop_event(e);
	show_folder(-1, 2, -1, -1);

	clean_ie6_mout();
	check_div_height();

	if (ie == 6)
		setTimeout("time_scroll()", 1);
}

function show_attfolders(e) {
	if (ie == 6)
		fed_scrollTop = document.getElementById("full_ec_div").scrollTop;

	stop_event(e);
	show_folder(-1, -1, 2, -1);

	clean_ie6_mout();
	check_div_height();

	if (ie == 6)
		setTimeout("time_scroll()", 1);
}

function show_other(e) {
	if (ie == 6)
		fed_scrollTop = document.getElementById("full_ec_div").scrollTop;

	stop_event(e);
	show_folder(-1, -1, -1, 2);

	check_div_height();

	if (ie == 6)
		setTimeout("time_scroll()", 1);
}

function get_div_height() {
	var ret_height = document.documentElement.clientHeight - document.getElementById("dc_top").offsetHeight - 144;

	var theObj = document.getElementById("id_ea_brow");
	if (theObj != null)
		ret_height = ret_height - 22;

	return ret_height;
}

function check_div_height() {
	var auto_height = get_div_height();

	var theObj = document.getElementById("dc_btm_1");
	var theObj2 = document.getElementById("dc_btm_2");

	if (theObj != null || theObj2 != null)
		auto_height = auto_height - 100;

	theObj = document.getElementById("full_ec_div");

	if (ie == 6)
	{
		if (theObj.scrollHeight > auto_height)
		{
			if (theObj.style.overflowY != "scroll")
				theObj.style.width = "138px";
			else
				theObj.style.height = "auto";
		}
		else
			theObj.style.width = "auto";
	}

	if (auto_height < 90)
		auto_height = 90;

	if (theObj.scrollHeight > auto_height)
	{

		theObj.style.height = auto_height + "px";
		theObj.style.overflowY = "scroll";
	}
	else
	{
		theObj.style.height = "auto";
		theObj.style.overflowY = "hidden";
	}
}

function clean_ie6_mout() {
	if (ie != 6)
		return ;

	if (document.getElementById("p_listmail").style.color.colorHex() != '#ffffff')
	{
		document.getElementById("p_listmail").style.backgroundColor = '#ffffff';
		document.getElementById("p_listmail").style.color = '#1e5494';
	}

	if (document.getElementById("p_attfolders").style.color.colorHex() != '#ffffff')
	{
		document.getElementById("p_attfolders").style.backgroundColor = '#ffffff';
		document.getElementById("p_attfolders").style.color = '#1e5494';
	}

<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then
%>
	if (document.getElementById("p_labels").style.color.colorHex() != '#ffffff')
	{
		document.getElementById("p_labels").style.backgroundColor = '#ffffff';
		document.getElementById("p_labels").style.color = '#1e5494';
	}

	document.getElementById("p_other").style.backgroundColor = '#ffffff';
	document.getElementById("p_other").style.color = '#1e5494';
<%
end if
%>
}

function select_one(right_url, tg_name) {
	var sp = right_url.lastIndexOf('/');
	var u_asp = "";
	var u_tg = "";
	var mil;
	var zk_listmail = false;
	var zk_labels = false;
	var zk_attfolders = false;
	var zk_other = false;

	if (sp > -1)
	{
		var ep = right_url.indexOf(".asp", sp);
		if (ep > -1)
		{
			u_asp = right_url.substr(sp + 1, ep - sp + 3).toLowerCase();
			u_tg = right_url.substring(ep + 4);
		}
	}

	if (u_asp == "listmail.asp")
	{
		if (tg_name == "in")
			select_id_cg_color("listmail_in");
		else if (tg_name == "out")
		{
			select_id_cg_color("listmail_out");
			zk_listmail = true;
		}
		else if (tg_name == "sed")
		{
			select_id_cg_color("listmail_sed");
			zk_listmail = true;
		}
		else if (tg_name == "del")
		{
			select_id_cg_color("listmail_del");
			zk_listmail = true;
		}
		else
		{
			mil = document.getElementsByTagName("div");
			for(var i=0; i<mil.length; i++) 
			{
				if (mil[i].innerHTML == "&nbsp;&nbsp;" + tg_name)
				{
					if (mil[i].id.length < 1 && mil[i].parentNode.id == "chd_listmail")
					{
						select_cg_color(mil[i]);
						zk_listmail = true;
						break;
					}
				}
			}
		}
	}
	else if (u_asp == "listatt.asp")
	{
		if (tg_name == "att")
		{
			select_id_cg_color("attfolders_att");
			zk_attfolders = true;
		}
		else
		{
			mil = document.getElementsByTagName("div");
			for(var i=0; i<mil.length; i++) 
			{
				if (mil[i].innerHTML == "&nbsp;&nbsp;" + tg_name)
				{
					if (mil[i].id.length < 1 && mil[i].parentNode.id == "chd_attfolders")
					{
						select_cg_color(mil[i]);
						zk_attfolders = true;
						break;
					}
				}
			}
		}
	}
	else if (u_asp == "listlabel.asp")
	{
		sp = u_tg.indexOf("lbid=");

		if (sp > -1)
		{
			if (u_tg.substr(sp + 5, 16).toLowerCase() == "%2d%2dstar%2d%2d")
				select_id_cg_color("id_listlabel");
			else
			{
				ep = u_tg.substr(sp + 5, 8);
				if (ep.length == 8)
				{
					select_id_cg_color("labels_" + ep);
					zk_labels = true;
				}
			}
		}
	}
	else
	{
		if (select_id_cg_color("id_" + u_asp.substr(0, u_asp.length - 4)) == false)
		{
			var aspid = find_id(u_asp);
			if (aspid.length > 0)
			{
				if (select_id_cg_color(aspid) == true)
				{
					var aspid_pid = document.getElementById(aspid).parentNode.id;
					if (aspid_pid == "chd_listmail")
						zk_listmail = true;
					else if (aspid_pid == "chd_labels")
						zk_labels = true;
					else if (aspid_pid == "chd_attfolders")
						zk_attfolders = true;
					else if (aspid_pid == "chd_other")
						zk_other = true;
				}
			}
		}

		if (u_asp == "cert_index.asp" || u_asp == "cal_index.asp" || u_asp == "ff_showall.asp" || u_asp == "poll_showall.asp" || u_asp == "showallpf.asp")
			zk_other = true;
	}

	if (zk_listmail == true)
		show_folder(1, -1, -1, -1);

	if (zk_labels == true)
		show_folder(-1, 1, -1, -1);

	if (zk_attfolders == true)
		show_folder(-1, -1, 1, -1);

	if (zk_other == true)
		show_folder(-1, -1, -1, 1);
}

var old_tgobj;
function select_cg_color(tgobj) {
	var isok = false;
	if (old_tgobj != null)
	{
		old_tgobj.style.backgroundColor = '#ffffff';
		old_tgobj.style.color = '#1e5494';
		set_mouse(old_tgobj);
	}

	if (tgobj != null)
	{
		tgobj.style.backgroundColor = '#5991cf';
		tgobj.style.color = '#ffffff';
		old_tgobj = tgobj;
		isok = true;
	}

	return isok;
}

function select_id_cg_color(tg_id) {
	return select_cg_color(document.getElementById(tg_id));
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

var left_entads_content_div = "";
var left_entfolder_count = 0;
var left_entads_count = 0;
var left_array_ent_ads = [];
var array_ads = new Array();
var ar_index = 0;
var ar_is_request = false;

function SendInfo() {
	if (ar_is_request == true)
		return ;

	clean_ads();

	var url = "ajadsadg.asp?" + getJsGrsn();
	request.open("GET", url, true);
	request.onreadystatechange = updatePage;
	request.send(null);
}

function updatePage() {
	if (request.readyState == 4)
	{
		if (request.status == 200)
		{
			fjarray(request.responseText);
			ar_is_request = true;
		}
	}
}

function clean_ads() {
	array_ads.splice(0, array_ads.length);
	ar_index = 0;
	ar_is_request = false;
}

function fjarray(svpstr) {
	var s,ss,nickname,email;
	s = svpstr;
	ss = s.split("\t");

	var i;
	for(i = 0; i < ss.length; i++)
	{
		if (ss[i].length > 1)
		{
			if (ss[i].charAt(0) == '\f')
			{
				array_ads[ar_index] = new Array(ss[i].substr(1), '', 0, 0);
				ar_index++;
			}
			else
			{
				var fs = ss[i].indexOf('\f');
				if (fs > 0)
				{
					array_ads[ar_index] = new Array(ss[i].substr(fs + 1), ss[i].substring(0, fs), 0, 0);
					ar_index++;
				}
			}
		}
	}
}
//-->
</script>

<body onload="return window_onload()">
<div id="dc_top" class="menu_bd bd">
<table width="100%">
<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then
%>
	<div id="id_wframe" name="mi" onclick="theright('wframe.asp', true, this)" class="menu_item"><%=s_lang_0326 %></div>
	<div class="menu_item_nofun"><div class="menu_item_nofun_con"></div></div>
<%
end if
%>
	<div id="listmail_in" name="mi" onclick="theright('listmail.asp?mode=in', false, this)" class="menu_item"><%=s_lang_0327 %></div>
<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then
%>
<%
Session("SH_Admin") = false

dim shm
set shm = server.createobject("easymail.SH_Manager")
shm.Load

if shm.isEnabled = true and (isadmin() = true or shm.isAdmin(Session("wem")) = true) then
	Session("SH_Admin") = true
%>
	<div id="id_mailsh" name="mi" onclick="theright('listsh.asp?fin=1', false, this)" class="menu_item"><%=s_lang_0588 %></div>
<%
end if

set shm = nothing
%>
	<div id="id_sendloglist" name="mi" onclick="theright('sendloglist.asp', true, this)" class="menu_item"><%=s_lang_0349 %></div>
<% if IsEnterpriseVersion = true and (Application("em_Enable_MailRecall") = true or Application("em_Enable_MailRecall") = "") then %>
	<div id="id_recalllist" name="mi" onclick="theright('recalllist.asp', true, this)" class="menu_item"><%=s_lang_0123 %></div>
<% end if %>
<%
end if
%>
	<div id="id_findmail" name="mi" onclick="theright('findmail.asp', true, this)" class="menu_item"><%=s_lang_0328 %></div>
</table>
</div>

<div style="font-size:0pt; height:4px;">&nbsp;</div>

<div class="menu_bd bd">
<table width="100%">
<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then
%>
	<div id="id_listlabel" name="mi" onclick="theright('listlabel.asp?lbid=%2D%2Dstar%2D%2D', false, this)" class="menu_item"><%=s_lang_0329 %></div>
	<div id="id_ads_brow" name="mi" onclick="theright('ads_brow.asp', true, this)" class="menu_item"><%=s_lang_0350 %></div>
<%
if isadmin() = true or (Application("em_EnableEntAddress") = true and eads_allnum > 0) then
%>
	<div id="id_ea_brow" name="mi" onclick="theright('ea_brow.asp', true, this)" class="menu_item"><%=s_lang_0330 %></div>
<%
end if
%>

	<div class="menu_item_nofun"><div class="menu_item_nofun_con"></div></div>
<%
end if
%>

<div id="full_ec_div" style="overflow-x:hidden; overflow-y:hidden;">
	<div id="p_listmail" name="mi" onclick="theright('viewmailbox.asp', true, this)" class="menu_item"><span id="ec_listmail" class="menu_item_ec" onclick="show_listmail(event);"><img src='images/null.gif' class='e_own'></span><span class="menu_item_ec_text"><%=s_lang_0331 %></span></div>
	<div id="chd_listmail" style="display:none;">
		<div id="listmail_out" name="mi" onclick="theright('listmail.asp?mode=out', false, this)" class="menu_item">&nbsp;&nbsp;<%=s_lang_0332 %></div>
		<div id="listmail_sed" name="mi" onclick="theright('listmail.asp?mode=sed', false, this)" class="menu_item">&nbsp;&nbsp;<%=s_lang_0333 %></div>
		<div id="listmail_del" name="mi" onclick="theright('listmail.asp?mode=del', false, this)" class="menu_item">&nbsp;&nbsp;<%=s_lang_0334 %></div>
<%
i = 0
pfNumber = pf.FolderCount

do while i < pfNumber
	spfname = pf.GetFolderName(i)

	Response.Write "<div name='mi_listmail' onclick=""theright('listmail.asp?mode=" & Server.URLEncode(spfname) & "', false, this)"" class='menu_item'>&nbsp;&nbsp;" & server.htmlencode(spfname) & "</div>" & Chr(13)

	spfname = NULL
	i = i + 1
loop
%>
	</div>

<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then
%>
<%
i = 0
allnum = mlb.Count

if allnum < 1 then
%>
	<div id="p_labels" name="mi" onclick="theright('labels.asp', true, this)" class="menu_item"><%=s_lang_0335 %></div>
<%
else
%>
	<div id="p_labels" name="mi" onclick="theright('labels.asp', true, this)" class="menu_item"><span id="ec_labels" class="menu_item_ec" onclick="show_labels(event);"><img src='images/null.gif' class='e_own'></span><span class="menu_item_ec_text"><%=s_lang_0335 %></span></div>
	<div id="chd_labels" style="display:none;">
<%
do while i < allnum
	mlb.GetByIndex i, ret_id, ret_title, ret_color

	Response.Write "<div id='labels_" & ret_id & "' name='mi' onclick=""theright('listlabel.asp?lbid=" & Server.URLEncode(ret_id) & "', false, this)"" class='menu_item' style='height:22px; _padding-top:6px; _margin-bottom:-6px;'><span class='wwm_color_in_line' style='background:#" & ret_color & ";'></span>&nbsp;" & server.htmlencode(ret_title) & "</div>" & Chr(13)

	ret_id = NULL
	ret_title = NULL
	ret_color = NULL

	i = i + 1
loop

	Response.Write "</div>"
end if
%>
<%
end if
%>

	<div id="p_attfolders" name="mi" onclick="theright('attfolders.asp', true, this)" class="menu_item"><span id="ec_attfolders" class="menu_item_ec" onclick="show_attfolders(event);"><img src='images/null.gif' class='e_own'></span><span class="menu_item_ec_text"><%=s_lang_0336 %></span></div>
	<div id="chd_attfolders" style="display:none;">
		<div id="attfolders_att" name="mi" onclick="theright('listatt.asp?mb=att', false, this)" class="menu_item">&nbsp;&nbsp;<%=s_lang_0296 %></div>
<%
i = 0
pafNumber = paf.FolderCount

do while i < pafNumber
	spfname = paf.GetFolderName(i)

	Response.Write "<div name='mi_attfolders' onclick=""theright('listatt.asp?mb=" & Server.URLEncode(spfname) & "', false, this)"" class='menu_item'>&nbsp;&nbsp;" & server.htmlencode(spfname) & "</div>" & Chr(13)

	spfname = NULL
	i = i + 1
loop
%>
	</div>

<%
if isadmin() = true or Session("ReadOnlyUser") <> 1 then
%>
	<div id="p_other" name="mi" onclick="show_other(event);" class="menu_item"><span id="ec_other" class="menu_item_ec"><img src='images/null.gif' class='e_own'></span><span class="menu_item_ec_text"><%=s_lang_0337 %></span></div>
	<div id="chd_other" style="display:none;">
<%
if IsEnterpriseVersion = true then
%>
		<div id="id_archive" name="mi" onclick="theright('showarchive.asp', true, this)" class="menu_item">&nbsp;&nbsp;<%=s_lang_0570 %></div>
<%
end if
%>
		<div id="id_cert_index" name="mi" onclick="theright('cert_index.asp', true, this)" class="menu_item">&nbsp;&nbsp;<%=s_lang_0338 %></div>
		<div id="id_cal_index" name="mi" onclick="theright('cal_index.asp', true, this)" class="menu_item">&nbsp;&nbsp;<%=s_lang_0339 %></div>
<%
if Application("em_Enable_ShareFolder") = true then
%>
		<div id="id_ff_showall" name="mi" onclick="theright('ff_showall.asp', true, this)" class="menu_item">&nbsp;&nbsp;<%=s_lang_0340 %></div>
<%
end if
%>
		<div id="id_poll_showall" name="mi" onclick="theright('poll_showall.asp', true, this)" class="menu_item">&nbsp;&nbsp;<%=s_lang_0341 %></div>
<%
if Application("em_EnableBBS") = true then
%>
		<div id="id_showallpf" name="mi" onclick="theright('showallpf.asp', true, this)" class="menu_item">&nbsp;&nbsp;<%=s_lang_0342 %></div>
<%
end if
%>
	</div>
</div>

	<div class="menu_item_nofun"><div class="menu_item_nofun_con"></div></div>

	<div id="id_user_right" name="mi" onclick="theright('user_right.asp', true, this)" class="menu_item"><%=s_lang_0343 %></div>
	<div id="id_nb_brow" name="mi" onclick="theright('nb_brow.asp', true, this)" class="menu_item"><%=s_lang_0344 %></div>

<%
end if
%>
	<div class="menu_item_nofun"><div class="menu_item_nofun_con"></div></div>

	<div name="mi" onclick="theright('default.asp?logout=true', false, this)" class="menu_item"><%=s_lang_0395 %></div>
</table>
</div>


<% if isadmin() = false and Application("em_SpamAdmin") = LCase(Session("wem")) then %>
<div style="font-size:0pt; height:4px;">&nbsp;</div>

<div id="dc_btm_1" class="menu_bd bd">
<table width="100%">
	<div id="id_right" name="mi" onclick="theright('right.asp', true, this)" class="menu_item"><%=s_lang_0345 %></div>
</table>
</div>
<% end if %>

<% if isadmin() = true then%>
<div style="font-size:0pt; height:4px;">&nbsp;</div>

<div id="dc_btm_2" class="menu_bd bd">
<table width="100%">
	<div id="id_right" name="mi" onclick="theright('right.asp', true, this)" class="menu_item"><%=s_lang_0345 %></div>
	<div id="id_showuser" name="mi" onclick="theright('showuser.asp', true, this)" class="menu_item"><%=s_lang_0346 %></div>
</table>
</div>
<%
else
	if isdomainmanager = true then
%>
<div style="font-size:0pt; height:4px;">&nbsp;</div>

<div id="dc_btm_2" class="menu_bd bd">
<table width="100%">
	<div id="id_domainright" name="mi" onclick="theright('domainright.asp', true, this)" class="menu_item"><%=s_lang_0347 %></div>
<%
		if isAccountsAdmin() = false then
%>
	<div id="id_showdomainusers" name="mi" onclick="theright('showdomainusers.asp', true, this)" class="menu_item"><%=s_lang_0348 %></div>
<%
		else
%>
	<div id="id_showuser" name="mi" onclick="theright('showuser.asp', true, this)" class="menu_item"><%=s_lang_0346 %></div>
<%
		end if
%>
</table>
</div>
<%
	else
		if isAccountsAdmin() = true then
%>
<div style="font-size:0pt; height:4px;">&nbsp;</div>

<div id="dc_btm_2" class="menu_bd bd">
<table width="100%">
	<div id="id_showuser" name="mi" onclick="theright('showuser.asp', true, this)" class="menu_item"><%=s_lang_0346 %></div>
</table>
</div>
<%
		end if
	end if
end if
%>

<div style="position:absolute; display:none;">
<form id="leftval" name="leftval">
<INPUT NAME="tgname" TYPE="hidden">
<INPUT NAME="sortinfo" TYPE="hidden">
<INPUT NAME="temp" TYPE="hidden">
<INPUT NAME="purl" TYPE="hidden">
<INPUT NAME="to" TYPE="hidden">
<INPUT NAME="cc" TYPE="hidden">
<INPUT NAME="bcc" TYPE="hidden">
<select name="ads" id="ads" style="width:10px; visibility:hidden">
</select>
<INPUT NAME="s_asp" TYPE="hidden">
<INPUT NAME="s_col" TYPE="hidden">
<INPUT NAME="s_mode" TYPE="hidden">
<INPUT NAME="s_search" TYPE="hidden">
</form>
</div>

<script type="text/javascript">
if (ie == 6)
{
	var mil = document.getElementsByTagName("div");
	for (var i=0; i<mil.length; i++) 
	{
		if (mil[i].name == "mi" || mil[i].name == "mi_listmail" || mil[i].name == "mi_attfolders")
			set_mouse(mil[i]);
	}
}

function set_mouse(tgobj) {
	if (tgobj != null)
	{
		tgobj.onmouseover = function(){
			if (this.style.color.colorHex() != '#ffffff')
			{
				this.style.backgroundColor='#e0ecf9';
				this.style.color='#1e5494';
			}
		}
		tgobj.onmouseout = function(){
			if (this.style.color.colorHex() != '#ffffff')
			{
				this.style.backgroundColor='#ffffff';
				this.style.color='#1e5494';
			}
		}
	}
}

if (ie == 6)
	document.execCommand("BackgroundImageCache", false, true); 
</script>

</body>
</html>

<%
set mlb = nothing
set paf = nothing
set pf = nothing
%>
