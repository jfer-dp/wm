<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
sname = trim(request("dsname"))
sfname = trim(request("dsfname"))

if Request.ServerVariables("REQUEST_METHOD") = "GET" then
	rqs = trim(Request.QueryString)

	ss = InStr(rqs, "&backurl=")
	goback_url = Mid(rqs, ss + 9)

	ss = InStr(goback_url, "&gourl=")

	sgourl = ""
	pgourl = ""

	if ss <> 0 then
		sgourl = Mid(goback_url, 1, ss - 1)
		pgourl = Mid(goback_url, ss + 7)
	end if
end if

if trim(request("to")) <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	Set Obj=Server.CreateObject("EasyMail.EMMail")

if sname <> "" and sfname <> "" then
	openresult = Obj.OpenFriendFolder(Session("wem"), sname, sfname, false)

	if openresult = -1 then
		set Obj = nothing
		Response.Redirect "err.asp?errstr=" & a_lang_158
	elseif  openresult = 1 then
		set Obj = nothing
		Response.Redirect "err.asp?errstr=" & a_lang_119
	elseif  openresult = 2 then
		set Obj = nothing
		Response.Redirect "err.asp?errstr=" & a_lang_159
	end if
end if

	sgourl = trim(request("sgourl"))
	pgourl = trim(request("pgourl"))

	if Obj.DeliverMail(Session("wem"), trim(request("filename")), trim(request("to"))) = false then
		Set Obj=nothing
		Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("showmail.asp?" & trim(request("sgourl"))) & "&pgourl=" & pgourl
	else
		Set Obj=nothing
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("showmail.asp?" & trim(request("sgourl"))) & "&pgourl=" & pgourl
	end if
end if

dim ads
set ads = server.createobject("easymail.Addresses")
ads.Load Session("wem")
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
<link rel="stylesheet" type="text/css" href="images/selads.css">

<script type="text/javascript" src="images/mglobal.js"></script>
<script type="text/javascript" src="images/selads.js"></script>

<STYLE type=text/css>
<!--
.textbox_wwm {padding:2px 8px 0pt 3px; border:1px solid #999;background-color:#FFFFEE; font-size:13px; width:420px;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
-->
</STYLE>
</HEAD>

<script LANGUAGE=javascript>
<!--
if (ie == 6)
	document.execCommand("BackgroundImageCache", false, true); 

qF = "to";

var new_ads_adg_number = <%=ads.EmailCount + ads.GroupCount %>;
if (parent.f1.ar_is_request == true && parent.f1.array_ads.length != new_ads_adg_number)
	parent.f1.clean_ads();

function window_onload() {
try{
	if (parent.f1.document.getElementById("leftval") != null)
	{
		if (parent.f1.array_ads.length < 1)
		{
			parent.f1.SendInfo();
			setTimeout("wait_left_send_for_deliver()", 10);
		}
		else
		{
			array_ads = parent.f1.array_ads;
			main_write_ads(document.getElementById('main_dsearch').value.toLowerCase());
			main_check_search_str();
		}
	}
}catch(error){}

	ar_max_rq = 0;
}

function sendit() {
	if (check_sendto_number() == false)
	{
		alert("<%=a_lang_160 %>");
		document.f1.to.focus();
		return ;
	}

	if (document.f1.to.value != "")
	{
		parent.f1.document.leftval.purl.value = "showmail.asp?<%=trim(request("backurl")) %>";
		document.f1.submit();
	}
}

function check_sendto_number() {
	var all_sendto_num = 0;

	all_sendto_num = get_char_number(document.f1.to.value, ",") + 1;

	if (<%
set sysinfo = server.createobject("easymail.sysinfo")
sysinfo.Load

Response.Write sysinfo.Web_Max_Recipients

set sysinfo = nothing
%> < all_sendto_num)
		return false;

	return true;
}

function goback() {
	location.href = "showmail.asp?<%=goback_url %>";
}
//-->
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
<form name="f1" method="post" action="delivermail.asp">
<input type="hidden" name="dsname" value="<%=sname %>">
<input type="hidden" name="dsfname" value="<%=sfname %>">
<input type="hidden" name="filename" value="<%=trim(request("filename")) %>">
<input type="hidden" name="sgourl" value="<%=sgourl %>">
<input type="hidden" name="pgourl" value="<%=pgourl %>">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_161 %>
</td></tr>
<tr><td class="block_top_td" style="height:6px; _height:8px;"></td></tr>
<tr><td align="left" style="padding-left:20px; padding-top:14px;">

<table border="0" width="97%" cellpadding=0 cellspacing=0 align="left" bgcolor="white">
	<tr><td noWrap width="20%" height="35" align="right" valign="top" style="padding-top:8px; padding-left:10px; color:#444444;">
	<a href="javascript:selectAdd('Deliver')"><%=a_lang_162 %></a><%=s_lang_mh %>
	</td><td noWrap width="50%" align="left" valign="top">
	<textarea name="to" id="to" size="40" class='textbox_wwm' cols="40" rows="4"></textarea>
	</td>
<%
if Application("em_EnableEntAddress") = true then
%>
	<td noWrap align="left" valign="top" style="padding-left:6px; padding-top:24px;">
	<a href="javascript:eapop('Deliver')" title="<%=a_lang_163 %>"><img src="images/entads.gif" border="0" align="absmiddle"></a>
	</td>
<%
end if
%>
	<td noWrap width="30%" align="left" valign="top" style="padding-left:10px; padding-right:10px;">
<div id="main_ads" style="width:190px; height:130px; border:1px solid #999999;">

<div style="padding:4px 4px 4px 4px; border-bottom:1px solid #d3e1f0;">
<div style="border:1px solid #888888; width:180px; display:inline-block;">
<input type="text" id="main_dsearch" onkeyup="main_ds_keyup();" style="font-size:12px; border:0px; width:148px; height:18px; padding-left:3px; _margin:2px 0px -2px 1px;">
<span id="main_sicon" style="background-image:url(images/ok_search.gif); background-repeat:no-repeat; border:0px; width:15px; height:15px; font-size:10px; cursor:pointer; display:inline-block; margin:2px 4px -2px 0px; _margin:-1px 4px 1px 0px;" onclick="main_icon_click();"></span>
</div>
</div>

<div id="main_ads_in" style="width:190px; height:95px; overflow-x:hidden; overflow-y:auto;">
<div id="main_ads_name_id" class="s_ads_name"><%=a_lang_335 %></div>
<div id="main_left_ads_div"></div>
<div id="main_adg_name_id" class="s_ads_name" style="border-top:1px solid #d3e1f0; display:none;"><%=a_lang_336 %></div>
<div id="main_left_adg_div" style="display:none;"></div>
</div>

</div>
</td></tr>
</table>
</td></tr>
<tr><td class="block_top_td" style="height:12px;"></td></tr>

<tr><td align="left" style="background-color:white; padding-top:16px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
	<a class="wwm_btnDownload btn_blue" href="javascript:sendit();"><%=a_lang_166 %></a>&nbsp;
	<a class="wwm_btnDownload btn_blue" href="javascript:goback();"><%=s_lang_cancel %></a>
</td></tr>
</table>
</form>

<div id="pop_ads_div" class="mydiv" style="display:none;">
	<div class="pop_base"><div class="pop_bd bd"><div class="title">
		<div class="title_left"><%=a_lang_337 %></div>
		<div class="title_right" title="<%=s_lang_close %>" onclick="javascript:close_ads(0);"><span>&nbsp;</span></div>
	</div>
	<div class="pop_content">
<table width="420" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td width="200">
<div style="width:190px; height:292px; border:1px solid #4e86c4;">

<div style="padding:4px 4px 4px 4px; border-bottom:1px solid #d3e1f0;">
<div style="border:1px solid #888888; width:180px; display:inline-block;">
<input type="text" id="dsearch" onkeyup="ds_keyup();" style="font-size:12px; border:0px; width:147px; height:18px; padding-left:3px; _margin:2px 0px -2px 1px;">
<span id="sicon" style="background-image:url(images/ok_search.gif); background-repeat:no-repeat; border:0px; width:15px; height:15px; font-size:10px; cursor:pointer; display:inline-block; margin:2px 4px -2px 0px; _margin:-1px 4px 1px 0px;" onclick="icon_click();"></span>
</div>
</div>

<div style="width:190px; height:258px; overflow-x:hidden; overflow-y:auto;">
<div id="ads_name_id" class="s_ads_name"><%=a_lang_335 %></div>
<div id="left_ads_div"></div>
<div id="adg_name_id" class="s_ads_name" style="border-top:1px solid #d3e1f0; display:none;"><%=a_lang_336 %></div>
<div id="left_adg_div" style="display:none;"></div>
</div>

</div>
</td>
<td width="20">
<img src="images/adsright.gif" border="0">
</td>
<td width="200">
<div id="right_ads_div" style="width:190px; height:292px; border:1px solid #4e86c4; overflow-x:hidden; overflow-y:auto;">
</div>
</td></tr>
</table>
	</div>
	<div class="title_bottom">
	<div class="title_ok_cancel_div">
	<a id="pop_ok" class="wwm_btnDownload btn_gray" href="#" onclick="javascript:close_ads(1);"><%=s_lang_ok %></a>&nbsp;
	<a class="wwm_btnDownload btn_gray" href="#" onclick="javascript:close_ads(0);"><%=s_lang_cancel %></a>
	</div></div></div></div>
</div>

<%
if Application("em_EnableEntAddress") = true then
%>
<div id="pop_entads_div" class="mydiv" style="display:none;">
	<div class="pop_base"><div class="pop_bd bd"><div class="title">
		<div class="title_left"><%=a_lang_338 %></div>
		<div class="title_right" title="<%=s_lang_close %>" onclick="javascript:close_ent_ads(0);"><span>&nbsp;</span></div>
	</div>
	<div id="entads_content_div" class="pop_content" style="text-align:left; overflow-x:auto; overflow-y:auto;">
	</div>
	<div id="entads_find_div" class="pop_content" style="display:none;">
<table width="420" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td width="200">
<div style="width:190px; height:292px; border:1px solid #4e86c4;">

<div style="padding:4px 4px 4px 4px; border-bottom:1px solid #d3e1f0;">
<div style="border:1px solid #888888; width:180px; display:inline-block;">
<input type="text" id="ent_dsearch" onkeyup="ent_ds_keyup();" style="font-size:12px; border:0px; width:147px; height:18px; padding-left:3px; _margin:2px 0px -2px 1px;">
<span id="ent_sicon" style="background-image:url(images/ok_search.gif); background-repeat:no-repeat; border:0px; width:15px; height:15px; font-size:10px; cursor:pointer; display:inline-block; margin:2px 4px -2px 0px; _margin:-1px 4px 1px 0px;" onclick="ent_icon_click();"></span>
</div>
</div>

<div style="width:190px; height:258px; overflow-x:hidden; overflow-y:auto;">
<div id="ent_ads_name_id" class="s_ads_name"><%=a_lang_339 %></div>
<div id="ent_left_ads_div"></div>
</div>

</div>
</td>
<td width="20">
<img src="images/adsright.gif" border="0">
</td>
<td width="200">
<div id="ent_right_ads_div" style="width:190px; height:292px; border:1px solid #4e86c4; overflow-x:hidden; overflow-y:auto;">
</div>
</td></tr>
</table>
	</div>
	<div class="title_bottom">
	<div class="title_ok_cancel_div">
	<a id="btex_id" class="wwm_btnDownload btn_gray" href="#" onclick="javascript:exall();"><span id="btex"><%=s_lang_ex %></span></a>&nbsp;
	<a class="wwm_btnDownload btn_gray" href="#" onclick="javascript:show_ent_find();"><span id="entf_bt"><%=a_lang_340 %></span></a>&nbsp;
	<a id="pop_ok" class="wwm_btnDownload btn_gray" href="#" onclick="javascript:close_ent_ads(1);"><%=s_lang_ok %></a>&nbsp;
	<a class="wwm_btnDownload btn_gray" href="#" onclick="javascript:close_ent_ads(0);"><%=s_lang_cancel %></a>
	</div></div></div></div>
</div>
<%
end if
%>

<div id="bg" class="bg" style="display:none;"></div>
<iframe id='popIframe' class='popIframe' frameborder='0'></iframe>

</BODY>
</HTML>

<%
set ads = nothing
%>
