<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
dim ei
set ei = server.createobject("easymail.emmail")

inlineid = trim(request("inlineid"))
is_inline = false
if trim(request("inline")) = "1" then
	is_inline = true
end if

gourl = trim(request("gourl"))
enc_gourl = Server.URLEncode(gourl)
sname = trim(request("sname"))
sfname = trim(request("sfname"))
enc_rqs = Server.URLEncode(trim(Request.QueryString))

if sname <> "" and sfname <> "" then
	openresult = ei.OpenFriendFolder(Session("wem"), sname, sfname, false)

	if openresult = -1 then
		set ei = nothing
		Response.Redirect "err.asp?errstr=" & s_lang_0396
	elseif  openresult = 1 then
		set ei = nothing
		Response.Redirect "err.asp?errstr=" & s_lang_0397
	elseif  openresult = 2 then
		set ei = nothing
		Response.Redirect "err.asp?errstr=" & s_lang_0398
	end if
end if

if Request.ServerVariables("REQUEST_METHOD") = "POST" and (sname = "" or sfname = "") then
	set_mode = trim(Request.Form("setmode"))
	dim is_ok
	is_ok = true

	if set_mode = "3" then
		cname = UnEscape(trim(Request.Form("cname")))
		email = UnEscape(trim(Request.Form("email")))

		if cname = "" then
			cname = email
		end if

		dim ads
		set ads = server.createobject("easymail.Addresses")
		ads.Load Session("wem")

		if ads.Simple_Add_Email(cname, email) = false then
			is_ok = false
		else
			ads.Save
		end if
		set ads = nothing
	elseif set_mode = "4" then
		killem = UnEscape(trim(Request.Form("kill")))

		dim umg
		set umg = server.createobject("easymail.usermessages")
		umg.Load Session("wem")

		umg.AddRejectEMail killem

		umg.SaveReject
		set umg = nothing
	end if

	set ei = nothing

	if is_ok = true then
		Response.Write "1"
	else
		Response.Write "0"
	end if
	Response.End
end if

if sname = "" or sfname = "" then
	dim pf
	set pf = server.createobject("easymail.PerFolders")
	pf.Load Session("wem")
end if

'-----------------------------------------
filename = trim(request("filename"))

pt = trim(request("pt"))
subismessage = false

if pt <> "" then
	bd = trim(request("bd"))

	if bd <> "" then
		ei.LoadAll2 Session("wem"), filename, CDbl(pt), bd
	else
		ei.LoadAll1 Session("wem"), filename, CDbl(pt)
	end if

	subismessage = true
else
	ei.LoadAll Session("wem"), filename
end if

charset = UCase(ei.Text_CharSet)
if charset = "" or charset = "DEFAULT_CHARSET" then
	charset = s_lang_0553
end if

invitationID = ei.InvitationID
remindID = ei.RemindID

dim wemcert
set wemcert = server.createobject("easymail.WebEasyMailCert")

issign = false
isenc = false

if ei.IsSignature = true then
	issign = true
end if

if ei.IsSignAndEncrypt = true then
	isenc = true
end if

if ei.IsSignature = true or ei.IsSignAndEncrypt = true then
	if trim(request("scpw")) <> "" then
		Session("scpw") = trim(request("scpw"))
	end if

	if ei.DecryptAndVerify(Session("wem"), Session("mail"), Session("scpw")) = false then
		if Session("scpw") <> "" then
			wemcert.Load Session("wem"), Session("mail")

			if wemcert.CheckPassIsGood(Session("scpw"), -2) = false then
				Session("scpw") = ""
			end if
		end if
	end if
end if


Response.Cookies("name") = Session("wem")

dim userweb
set userweb = server.createobject("easymail.UserWeb")
userweb.Load Session("wem")

enableAutoAdaptCharSet = userweb.enableAutoAdaptCharSet
EnableShowHtmlMail = userweb.EnableShowHtmlMail

set userweb = nothing

dim allnum
allnum = 0
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html<%
if enableAutoAdaptCharSet = true then
	Response.Write "; charset=" & charset
else
	Response.Write "; charset=" & s_lang_0553
end if
%>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
<link rel="stylesheet" type="text/css" href="images/showmail.css">

<STYLE type=text/css>
<!--
.sbttn {font-family:<%=s_lang_font %>; font-size:9pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5;color:#000066;height:19px;text-decoration:none;cursor:pointer;}
.Bsbttn {font-family:<%=s_lang_font %>; font-size:10pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5; color:#000066;text-decoration:none;cursor:pointer;}
<%
if is_inline = false then
%>
.table_main {width:98%; border:0px; margin-top:8px;}
.bt_tool_table {width:98%; border:0px;}
<%
else
%>
body {margin: 0px 0px 0px 0px; padding: 0px 0px 0px 0px;}
.table_main {width:100%; border:0px;}
.bt_tool_table {width:100%; border:0px;}
<%
end if
%>
-->
</STYLE>
</HEAD>

<script type="text/javascript" src="images/mglobal.js"></script>
<script type="text/javascript" src="images/jquery.min.js"></script>
<script type="text/javascript" src="images/jquery-powerFloat-min.js"></script>

<script language="JavaScript">
<!-- 
if (ie == 6)
	document.execCommand("BackgroundImageCache", false, true); 

function window_onload() {
<%
if is_inline = false then
%>
if (ie != false)
	document.body.focus();
<%
end if
%>
	hide_rads(1);
	cg_p_height();
}

function back() {
<% if gourl = "" then %>
	history.back();
<% else %>
	location_href("<%=gourl %>&<%=getGRSN() %>");
<% end if %>
}

<%
if sname = "" or sfname = "" then
%>
function movemail(tgname) {
	location_href("movemail.asp?filename=<%=filename %>&mto=" + tgname + "&<%=getGRSN() %>&gourl=<%=enc_gourl %>");
}

function delthis() {
	if (confirm("<%=s_lang_0399 %>") == false)
		return ;

	location_href("delmail.asp?filename=<%=filename %>&<%=getGRSN() %>&gourl=<%=enc_gourl %>");
}
<%
end if
%>
function add2ads(vname, vemail) {
	var post_date = "setmode=3&<%=getGRSN() %>&cname=" + escape(vname) + "&email=" + escape(vemail);

$.ajax({
	type:"POST",
	url:"showarc.asp",
	data:post_date,
	success:function(data){
		alert_msg("<%=s_lang_0453 %>");
	},
	error:function(){
		alert_msg("<%=s_lang_0454 %>");
	}
});
}

function add2kill(vemail) {
	var post_date = "setmode=4&<%=getGRSN() %>&kill=" + escape(vemail);

$.ajax({
	type:"POST",
	url:"showarc.asp",
	data:post_date,
	success:function(data){
		alert_msg("<%=s_lang_0453 %>");
	},
	error:function(){
		alert_msg("<%=s_lang_0454 %>");
	}
});
}

function show_head_message()
{
	var theObj = document.getElementById("headmsg_span");
	if (theObj.style.display == "inline")
		theObj.style.display = "none";
	else
		theObj.style.display = "inline";

	cg_p_height();
}

function saveatt(anum) {
	location_href("mail2att.asp?<%=getGRSN() %>&filename=<%=filename %>&sname=<%=sname %>&sfname=<%=sfname %><%
if pt <> "" then
	response.write "&pt=" & pt
end if

if bd <> "" then
	response.write "&bd=" & Server.URLEncode(bd)
end if
%>&attnum=" + anum);
}

function doZoom(size){
<%
if ei.IsHtmlMail = false or EnableShowHtmlMail = false then
%>
	document.getElementById('zoom').style.fontSize=size+'px'
	cg_p_height();
<%
end if
%>
}

function setpw(){
	if (document.f1.sc_pw.value.length < 8)
		alert("<%=s_lang_0191 %>!");
	else
<%
if is_inline = false then
%>
		location.href = "showarc.asp?filename=<%=filename %>&scpw=" + document.f1.sc_pw.value + "&gourl=<%=enc_gourl %>";
<%
else
%>
		location.href = "showarc.asp?filename=<%=filename %>&inlineid=<%=inlineid %>&inline=<%=trim(request("inline")) %>&scpw=" + document.f1.sc_pw.value + "&gourl=<%=enc_gourl %>";
<%
end if
%>
}

function hide_rads(hr_set_none)
{
	if (hr_set_none == 0)
	{
		if (document.getElementById("ex_info_div").style.display == "none")
		{
			$("#rads_showstr").html("<%=s_lang_0457 %>");
			document.getElementById("ex_info_div").style.display = "inline";
		}
		else
		{
			$("#rads_showstr").html("<%=s_lang_0458 %>");
			document.getElementById("ex_info_div").style.display = "none";
		}
	}
	else
	{
		$("#rads_showstr").html("<%=s_lang_0458 %>");
		document.getElementById("ex_info_div").style.display = "none";
	}

	cg_p_height();
}

function show_it(name) {
	var show_span = document.getElementById(name + "_span")

	if (show_span.style.display == "none")
		show_span.style.display = "inline";
	else
		show_span.style.display = "none";

	cg_p_height();
}

function m_over(tag_obj) {
	tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj) {
	tag_obj.style.backgroundColor = "white";
}

function iFrameHeight() {
	var ifm= document.getElementById("iframepage");
	var subWeb = document.frames ? document.frames["iframepage"].document : ifm.contentDocument;
	if(ifm != null && subWeb != null) {
		ifm.height = subWeb.body.scrollHeight;
	}
}

function cg_p_height() {
<%
if is_inline = true then
	Response.Write "parent.iFrameHeight('" & inlineid & "');"
end if
%>
}

function location_href(url) {
<%
if is_inline = true then
	Response.Write "parent.location.href = url;"
else
	Response.Write "location.href = url;"
end if
%>
}

function parent_href(url) {
	parent.location.href = url;
}

function alert_msg(amsg) {
	$("#top_show_msg").text(amsg);
	document.getElementById("top_show_msg").style.display = "inline";
	setTimeout("close_alert()", 3000);
}

function close_alert() {
	document.getElementById("top_show_msg").style.display = "none";
}
// -->
</script>

<BODY id="onfc" LANGUAGE=javascript onload="return window_onload()">
<a name="gotop" style="font-size:0pt; height:0px;"></a>
<form name="f1">
<table id="table_main" class="table_main" align="center" cellspacing="0" cellpadding="0">
	<tr><td class="block_top_td"><div class="table_min_width"></div></td></tr>
<%
if is_inline = false then
	if subismessage = false then
%>
	<tr><td class="tool_top_td">

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:back()"><< <%=s_lang_return %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=reply&<%=getGRSN() %>&filename=<%=filename & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&backurl=showarc.asp?" & enc_rqs %>"><%=s_lang_0460 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=replyall&<%=getGRSN() %>&filename=<%=filename & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&backurl=showarc.asp?" & enc_rqs %>"><%=s_lang_0461 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

<%
if ei.IsSaveMail = false then
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=forward&<%=getGRSN() %>&filename=<%=filename & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&backurl=showarc.asp?" & enc_rqs %>"><%=s_lang_0462 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
else
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=forward&<%=getGRSN() %>&filename=<%=filename & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&backurl=showarc.asp?" & enc_rqs %>"><%=s_lang_0463 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
end if

if sname = "" or sfname = "" then
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:delthis();"><%=s_lang_del %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span"><span id="pm_moveto" class="menu_pop"<%=s_lang_0551 %>>
	<div class='attbg'><img style='margin: 6px 0pt 0pt;' src='images/popshow.gif'></div>
	<div class='menu_pop_text'><%=s_lang_0404 %>...</div>
	</span></span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
end if
%>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="<%
	Response.Write "showeml.asp?outmode=down&filename=" & filename & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN()
%>" target="_blank"><%=s_lang_0373 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span id="top_pn" class="st_right_span">
	</span>

	</td></tr>
<%
	end if
end if
%>
	<tr class="head_tr"><td id="subject_td" class="head_subject_td" style="border-top:1px solid #aac1de;">
	<span class="head_subject_span"><%=server.htmlencode(ei.subject) %></span>
	</td></tr>
<%
if issign = true or isenc = true then
%>
	<tr class="head_tr"><td class="head_td">
<%
if ei.NeedPass = false then
	Response.Write "<span style='float:left;'>"

	if ei.UnknownSigningKey = true then
		if issign = true then
			Response.Write "<img src='images/s3.gif' border='0' align='absmiddle' title='" & s_lang_0464 & "'>"
		elseif isenc = true then
			Response.Write "<img src='images/e3.gif' border='0' align='absmiddle' title='" & s_lang_0465 & "'>"
		end if
	else
		if ei.isVerified = true then
			if issign = true then
				Response.Write "<img src='images/s1.gif' border='0' align='absmiddle' title='" & s_lang_0466 & "'>"
			elseif isenc = true then
				Response.Write "<img src='images/e1.gif' border='0' align='absmiddle' title='" & s_lang_0467 & "'>"
			end if
		else
			if issign = true then
				Response.Write "<img src='images/s2.gif' border='0' align='absmiddle' title='" & s_lang_0468 & "'>"
			elseif isenc = true then
				Response.Write "<img src='images/e2.gif' border='0' align='absmiddle' title='" & s_lang_0469 & "'>"
			end if
		end if
	end if
else
	Response.Write "<span style='float:left;'>"
	Response.Write s_lang_0470 & s_lang_mh
end if

if ei.NeedPass = false then
%>
	</span>
	<span class='ca_span'>
	<%=s_lang_0471 %><%=s_lang_mh %><%
if ei.Signer <> "" then
	Response.Write "<font color='black'>" & server.htmlencode(ei.Signer) & "</font>"
else
	Response.Write "<font color='#901111'>" & s_lang_0472 & "</font>"
end if
%><br>
	<%=s_lang_0473 %><%=s_lang_mh %><%=ei.SignedTime %>
<%
else
	if ei.isRecipient = true then
		if wemcert.LightHasSecCert(Session("wem")) = false then
%>[<font color='#901111'><%=s_lang_0474 %><a href="cert_index.asp?<%=getGRSN() %>"><%=s_lang_0475 %></a>
<%
		else
			if Session("scpw") = "" and trim(request("scpw")) <> "" then
%>[<font color='#901111'><%=s_lang_0476 %></font>] <font color='#901111'><%=s_lang_0397 %>!</font> <font color='black'><%=s_lang_0477 %><%=s_lang_mh %></font><input type="password" name="sc_pw" class='pw_input' size="11">&nbsp;
<%
			else
%>[<font color='#901111'><%=s_lang_0476 %></font>] <font color='black'><%=s_lang_0477 %><%=s_lang_mh %></font><input type="password" name="sc_pw" class='pw_input' size="11">&nbsp;
<%
			end if
%>
<input type="button" value="<%=s_lang_0313 %>" onclick="javascript:setpw()" class="sbttn">
<%
		end if
	else
%>[<font color='#901111'><%=s_lang_0476 %></font>] <font color='black'><%=s_lang_0478 %></font>
<%
	end if
end if
%>
	</span>
	</td></tr>
<%
end if
%>

	<tr class="head_tr"><td class="head_td">
	<span style="float:left;"><%=s_lang_0479 %><%=s_lang_mh %></span>
	<span style="float:left;"><%
if ei.FromName = ei.FromMail then
	receiver = "<font color='#5fa207' style='font-weight:bold;'>" & server.htmlencode(ei.FromMail) & "</font>"
else
	receiver = "<font color='#5fa207' style='font-weight:bold;'>" & server.htmlencode(ei.FromName) & "</font>" & server.htmlencode(" <" & ei.FromMail & ">")
end if

dim item
item = server.htmlencode(ei.FromName)
item = replace(item, "'", "")
item = replace(item, """", "")

rec_email = ei.FromMail
rec_email = replace(rec_email, "'", "")
rec_email = replace(rec_email, """", "")

is_own = false
if LCase(ei.FromMail) = LCase(Session("mail")) then
	Response.Write "<img src='images/null.gif' class='head_own'>&nbsp;"

	is_own = true
end if

Response.Write receiver

if InStr(ei.FromMail, "@") then
	if is_own = false then
		Response.Write "&nbsp;&nbsp;&nbsp;[<a href='javascript:add2ads(""" & item & """,""" & rec_email & """)'>" & s_lang_0480 & "</a>&nbsp;&nbsp;<a href='javascript:add2kill(""" & rec_email & """)'>" & s_lang_0481 & "</a>]"
	end if
end if
%></span>
	</td></tr>

	<tr class="head_tr"><td class="head_td">
	<span style="float:left;"><%=s_lang_0482 %><%=s_lang_mh %><%=ei.Time %></span>
	<span style="float:right;">[<a href="javascript:hide_rads(0)"><span id="rads_showstr"></span></a>]&nbsp;&nbsp;<%=s_lang_0483 %><%=s_lang_mh %>[<a href="javascript:doZoom(16)"><%=s_lang_0484 %></a> <a href="javascript:doZoom(14)"><%=s_lang_0485 %></a> <a href="javascript:doZoom(12)"><%=s_lang_0486%></a>]</span>
	</td></tr>

	<tr class="head_tr"><td class="head_td">
	<span style="float:left;"><%=s_lang_0487 %><%=s_lang_mh %><font color="black"><%=getShowSize(ei.Size) %></font></span>
	<span id="go_att"></span>
	</td></tr>
<%
if ei.IsHtmlMail = true then
%>
	<tr class="head_tr"><td class="head_td">
	<span style="float:left;"><%=s_lang_0495 %><%=s_lang_mh %><%
	Response.Write "<a href=""showatt.asp?ishtml=1&filename=" & filename & "&count=0&pt=" & pt & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN() & """ target='_blank'>" & s_lang_0496 & "</a>"
%></span>
	</td></tr>
<%
end if
%>

	<tr><td>
<div id="ex_info_div" style="display:none">
	<table width="100%" border="0" cellspacing="0">
	<tr class="head_tr"><td class="head_td">
	<span style="float:left;"><span id="cal_mode"><%=s_lang_0497 %></span><%=s_lang_mh %></span>
	<span style="float:left;"><span id="cal_msg"><%
allnum = ei.to_count
i = 0
first_show = true
is_own = false

do while i < allnum
	ei.GetToAds i, ret_name, ret_email

	if InStr(ret_email, "@") then
		if first_show = false then
			Response.Write "<br>"
		else
			first_show = false
		end if

		is_own = false
		if LCase(ret_email) = LCase(Session("mail")) then
			Response.Write "<img src='images/null.gif' class='head_own'>&nbsp;"
			is_own = true
		end if

		if ret_name = ret_email or Len(ret_name) < 1 then
			Response.Write "<font color='black'>" & server.htmlencode(ret_email) & "</font>"
		else
			Response.Write "<font color='black'>" & server.htmlencode(ret_name) & "</font>" & server.htmlencode(" <" & ret_email & ">")
		end if

		if is_own = false then
			Response.Write "&nbsp;&nbsp;&nbsp;[<a href='javascript:add2ads(""" & server.htmlencode(ret_name) & """, """ & server.htmlencode(ret_email) & """)'>" & s_lang_0480 & "</a>]" & Chr(13)
		end if
	end if

	ret_name = NULL
	ret_email = NULL

	i = i + 1
loop
%></span></span>
	</td></tr>

	<tr class="head_tr"><td class="head_td">
	<span style="float:left;"><%=s_lang_0501 %><%=s_lang_mh %><%
xmsp = ei.XMSMailPriority

if xmsp = "High" then
	Response.Write s_lang_0130
elseif xmsp = "Low" then
	Response.Write s_lang_0131
else
	Response.Write s_lang_0146
end if
%></span>
	</td></tr>
<%
allnum = ei.cc_count
if allnum > 0 then
%>
	<tr class="head_tr"><td class="head_td">
	<span style="float:left;"><%=s_lang_0502 %><%=s_lang_mh %></span>
	<span style="float:left;"><%
i = 0
first_show = true
is_own = false

do while i < allnum
	ei.GetCcAds i, ret_name, ret_email

	if InStr(ret_email, "@") then
		if first_show = false then
			Response.Write "<br>"
		else
			first_show = false
		end if

		is_own = false
		if LCase(ret_email) = LCase(Session("mail")) then
			Response.Write "<img src='images/null.gif' class='head_own'>&nbsp;"
			is_own = true
		end if

		if ret_name = ret_email or Len(ret_name) < 1 then
			Response.Write "<font color='black'>" & server.htmlencode(ret_email) & "</font>"
		else
			Response.Write "<font color='black'>" & server.htmlencode(ret_name) & "</font>" & server.htmlencode(" <" & ret_email & ">")
		end if

		if is_own = false then
			Response.Write "&nbsp;&nbsp;&nbsp;[<a href='javascript:add2ads(""" & server.htmlencode(ret_name) & """, """ & server.htmlencode(ret_email) & """)'>" & s_lang_0480 & "</a>]" & Chr(13)
		end if
	end if

	ret_name = NULL
	ret_email = NULL

	i = i + 1
loop
%></span>
	</td></tr>
<%
end if
%>

	<tr class="head_tr"><td class="head_td">
	<span style="float:left;"><%=s_lang_0503 %><%=s_lang_mh %><a href="javascript:show_head_message()"><%=s_lang_0504 %></a></span><br>
	<span id="headmsg_span" style="display:none;">
<table width="100%" border="0" cellspacing="0" align="center" bgcolor='white' class="table_in">
	<tr><td height='20' class='headmsg_td'>
<%
ht = server.htmlencode(ei.HeadMessage)
ht = replace(ht, Chr(13), "<br>")
ht = replace(ht, Chr(32), "&nbsp;")
ht = replace(ht, Chr(9), "&nbsp;&nbsp;&nbsp;&nbsp;")
Response.Write ht
%>
	</td></tr>
</table>
	</span>
	</td></tr>

	</table>
	</div>
	</td></tr>

<%
if ei.IsHtmlMail = true then
	if EnableShowHtmlMail = true then
%>
	<tr bgcolor="white"><td class="iframe_td">
<iframe src="<%="showatt.asp?ishtml=1&filename=" & filename & "&count=0&pt=" & pt & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN() %>" id="iframepage" name="iframepage" frameBorder=0 scrolling=no width="100%" onLoad="iFrameHeight()"></iframe>
	</td></tr>
<%
	end if
end if

if ei.IsHtmlMail = false or EnableShowHtmlMail = false then
%>
	<tr bgcolor="white">
    <td id="zoom" class="zoom_td">
<%
end if

dim ecalnt
dim ecal
isok = true

if Len(invitationID) > 10 then
	set ecalnt = server.createobject("easymail.CalendarNotice")
	isok = ecalnt.Load(Session("wem"))

	if isok = true then
		isok = ecalnt.MoveToID(invitationID)
	end if

	set ecal = server.createobject("easymail.Calendar")
	if isok = true then
		isok = ecal.Load(ecalnt.bi_host_account)
	end if

	if isok = true then
		isok = ecal.MoveToID(invitationID)
	end if
end if

if Len(remindID) > 10 then
	set ecal = server.createobject("easymail.Calendar")
	if isok = true then
		isok = ecal.Load(Session("wem"))
	end if

	if isok = true then
		isok = ecal.MoveToID(remindID)
	end if
end if


if (Len(invitationID) < 10 and Len(remindID) < 10) or isok = false then

if ei.ContentType = "text/html" then
	if charset = "UTF-8" then
		utf_pos = InStr(ei.Text, "charset=UTF-8")

		if utf_pos > 0 then
			t = Mid(ei.Text, 1, utf_pos - 1)
			t = t & Mid(ei.Text, utf_pos + 13)
		else
			t = ei.Text
		end if
	else
		t = ei.Text
	end if
else
	if (issign = true or isenc = true) and ei.DecryptOrVerifyStr <> "" then
		t = server.htmlencode(ei.DecryptOrVerifyStr)
	else
		t = server.htmlencode(ei.Text)
	end if

	if Len(t) < 100000 then
		t = ei.ConvText2Html(t)
	end if

	t = replace(RemoveEndRN(t), Chr(10), "<br>")
	t = replace(t, Chr(32) & Chr(32), "&nbsp;&nbsp;")
	t = replace(t, Chr(9), "&nbsp;&nbsp;&nbsp;&nbsp;")
end if

if ei.IsHtmlMail = false or EnableShowHtmlMail = false then
	Response.Write t
%>&nbsp;</td>
  </tr>
<%
end if

allnum = ei.AttachmentCount

if allnum = 1 and ei.IsHtmlMail = true then
	allnum = 0
end if

i = 0
if allnum > 0 then
%>
	<tr><td>
<table width="<%
if is_inline = false then
	Response.Write "100%"
else
	Response.Write "97%"
end if
%>" border="0" cellspacing="0" align="center" class='table_att_in'>
<tr><td class='att_in_td'><img src='images/attach.gif' border='0' align='absmiddle' style='padding-right:8px;'><%=s_lang_0505 %></td></tr>
<%
if ei.IsHtmlMail = true then
	i = 1
	allnum = ei.AllAttachmentCount
	show_i = 1

	do while i < allnum
		if ei.AttachmentCanShow(i) = true then
			Response.Write "<tr style='cursor:default;' onmouseover='m_over(this);' onmouseout='m_out(this);'><td class='att_td'><span class='att_span'>"
			if ei.GetAttachmentName(show_i) = "" then
			    Response.Write "<a class='att' href=""showatt.asp?filename=" & filename & "&count=" & i & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN() & """ target='_blank'>" & "html" & "</a>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<a href=""showatt.asp?isdown=1&filename=" & filename & "&count=" & i & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN() & """ target='_blank'>" & s_lang_0373 & "</a>]"
				Response.Write "&nbsp;&nbsp;[<a href=""JavaScript:saveatt(" & show_i & ")"")>" & s_lang_0506 & "</a>]"
			else
				if ei.AttachmentIsMessage(show_i) = false then
			    	Response.Write "<a class='att' href=""showatt.asp?filename=" & filename & "&count=" & i & "&pt=" & pt & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN() & """ target='_blank'>" & server.htmlencode(ei.GetAttachmentName(show_i)) & "</a>"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<a href=""showatt.asp?isdown=1&filename=" & filename & "&count=" & i & "&pt=" & pt & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN() & """ target='_blank'>" & s_lang_0373 & "</a>]"
					Response.Write "&nbsp;&nbsp;[<a href=""JavaScript:saveatt(" & show_i & ")"")>" & s_lang_0506 & "</a>]"
				else
			    	Response.Write "<a class='att' href=""showarc.asp?filename=" & filename & "&count=" & i & "&pt=" & ei.GetAttachmentPT(show_i) & "&bd=" & Server.URLEncode(ei.GetEmlAttachmentBD(i)) & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN() & """ target='_blank'>" & server.htmlencode(ei.GetAttachmentName(show_i)) & "</a>"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<a href=""showatt.asp?isdown=1&filename=" & filename & "&count=" & i & "&pt=" & pt & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN() & """ target='_blank'>" & s_lang_0373 & "</a>]"
					Response.Write "&nbsp;&nbsp;[<a href=""JavaScript:saveatt(" & show_i & ")"")>" & s_lang_0506 & "</a>]"
				end if
			end if
			Response.Write "<span></td></tr>" & Chr(13)

			show_i = show_i + 1
		end if

	    i = i + 1
	loop
else
	do while i < allnum
		Response.Write "<tr style='cursor:default;' onmouseover='m_over(this);' onmouseout='m_out(this);'><td class='att_td'><span class='att_span'>"
		if ei.GetAttachmentName(i) = "" then
		    Response.Write "<a class='att' href=""showatt.asp?filename=" & filename & "&count=" & i & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN() & """ target='_blank'>" & "html" & "</a>"
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<a href=""showatt.asp?isdown=1&filename=" & filename & "&count=" & i & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN() & """ target='_blank'>" & s_lang_0373 & "</a>]"
			Response.Write "&nbsp;&nbsp;[<a href=""JavaScript:saveatt(" & i & ")"")>" & s_lang_0506 & "</a>]"
		else
			if ei.AttachmentIsMessage(i) = false then
		    	Response.Write "<a class='att' href=""showatt.asp?filename=" & filename & "&count=" & i & "&pt=" & pt & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN() & """ target='_blank'>" & server.htmlencode(ei.GetAttachmentName(i)) & "</a>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<a href=""showatt.asp?isdown=1&filename=" & filename & "&count=" & i & "&pt=" & pt & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN() & """ target='_blank'>" & s_lang_0373 & "</a>]"
				Response.Write "&nbsp;&nbsp;[<a href=""JavaScript:saveatt(" & i & ")"")>" & s_lang_0506 & "</a>]"
			else
		    	Response.Write "<a class='att' href=""showarc.asp?filename=" & filename & "&count=" & i & "&pt=" & ei.GetAttachmentPT(i) & "&bd=" & Server.URLEncode(ei.GetEmlAttachmentBD(i)) & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN() & """ target='_blank'>" & server.htmlencode(ei.GetAttachmentName(i)) & "</a>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<a href=""showatt.asp?isdown=1&filename=" & filename & "&count=" & i & "&pt=" & pt & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN() & """ target='_blank'>" & s_lang_0373 & "</a>]"
				Response.Write "&nbsp;&nbsp;[<a href=""JavaScript:saveatt(" & i & ")"")>" & s_lang_0506 & "</a>]"
			end if
		end if
		Response.Write "<span></td></tr>" & Chr(13)

	    i = i + 1
	loop
end if
%>
</table>
	</td></tr>
<%
end if

else

	dim ecalset
	set ecalset = server.createobject("easymail.CalOptions")
	ecalset.Load Session("wem")

	show_APM = false
	if ecalset.Show24Hour = false then
		show_APM = true
	end if

	set ecalset = nothing

	ecal.get_bi_start_date b_year, b_month, b_day, b_hour, b_minute
%>
<script language="JavaScript">
<!--
var Stag = document.getElementById("cal_mode");
Stag.innerHTML = "<%=s_lang_0507 %>";

Stag = document.getElementById("cal_msg");
Stag.innerHTML = "<font color='#BC131A'><b><%

if Len(invitationID) > 10 then
	Response.Write s_lang_0508
elseif Len(remindID) > 10 then
	Response.Write s_lang_0509
end if
%></b></font>";

var show_APM = <%=LCase(CStr(show_APM)) %>;
<%=s_lang_0549 %>

function getShowIconStr(bmode, bremind, brp)
{
	var s_str = "";

	if (bmode == 3)
		s_str = s_str + "<img src='images/cal/bdc.gif' border=0 align='absmiddle' title='<%=s_lang_0510 %>'>";

	if (bremind == true)
		s_str = s_str + "<img src='images/cal/bell.gif' border=0 align='absmiddle' title='<%=s_lang_0511 %>'>";

	if (brp == true)
		s_str = s_str + "<img src='images/cal/repeat.gif' border=0 align='absmiddle' title='<%=s_lang_0512 %>'>";

	if (s_str.length > 0)
		s_str = s_str + "&nbsp;";

	return s_str;
}

function write_getShowIconStr(bmode, bremind, brp)
{
	document.write(getShowIconStr(bmode, bremind, brp));
}

function write_ShowDateUrl(by, bm, bd)
{
	var showS_Str = "";
	currentDate = new Date(by, bm - 1, bd);

	showS_Str = "<a href=\"cal_index.asp?<%=getGRSN() %>&tsn=0&sy=" + by + "&sm=" + bm + "&sd=" + bd + "\">";
	showS_Str = showS_Str + <%=s_lang_0513 %>;
	showS_Str = showS_Str + "</a>";

	document.write(showS_Str);
}

function get_APM(vsh, vsm)
{
	var t_str = ""
	if (show_APM == false)
	{
		t_str = vsh + ":";

		if (vsm < 10)
			t_str = t_str + "0";
		t_str = t_str + vsm;
	}
	else
	{
		if (vsm < 10)
			t_str = "0";
		t_str = t_str + vsm;

		if (vsh == 0)
			t_str = "12:" + t_str + "AM";
		else if (vsh == 12)
			t_str = "12:" + t_str + "PM";
		else if (vsh < 12)
			t_str = vsh + ":" + t_str + "AM";
		else
			t_str = vsh + ":" + t_str + "PM";
	}

	return t_str;
}

function convWeeekName(wnum) {
	if (wnum > 6)
		wnum = wnum - 7;

	if (wnum == 0)
		return "<%=s_lang_0514 %>";
	else if (wnum == 1)
		return "<%=s_lang_0515 %>";
	else if (wnum == 2)
		return "<%=s_lang_0516 %>";
	else if (wnum == 3)
		return "<%=s_lang_0517 %>";
	else if (wnum == 4)
		return "<%=s_lang_0518 %>";
	else if (wnum == 5)
		return "<%=s_lang_0519 %>";
	else if (wnum == 6)
		return "<%=s_lang_0520 %>";
}

function write_ShowEventTime(vy, vm, vd, vh, vmin, vnt)
{
	document.write(getShowStartStr(vy, vm, vd, vh, vmin, vnt));
}

function getShowStartStr(vy, vm, vd, vh, vmin, vnt)
{
	if (vnt == 0)
		return "<%=s_lang_0521 %>"

	var s_str = "";
	currentDate = new Date(vy, vm - 1, vd, vh, vmin);

	s_str = get_APM(currentDate.getHours(), currentDate.getMinutes()) + "-";
	currentDate.setTime(currentDate.getTime() + (vnt * 1000));
	s_str = s_str + get_APM(currentDate.getHours(), currentDate.getMinutes());

	return s_str;
}

function showevent()
{
<%
if Len(invitationID) > 10 then
%>
	location_href("cal_showinvite.asp?<%=getGRSN() %>&calid=<%=invitationID %>&fmeml=1");
<%
elseif Len(remindID) > 10 then
%>
	location_href("cal_showinvite.asp?<%=getGRSN() %>&calid=<%=remindID %>&fmeml=1&fmcal=1");
<%
end if
%>
}
<%
if Len(invitationID) > 10 then
%>
function showntlist()
{
	location_href("cal_listinvited.asp?<%=getGRSN() %>&fmeml=1");
}

function showcal()
{
	location_href("cal_index.asp?<%=getGRSN() %>&tsn=0&sy=<%=b_year %>&sm=<%=b_month %>&sd=<%=b_day %>");
}
<%
else
%>
function write_clockstr()
{
	var eventDate = new Date(<%=b_year %>, <%=b_month %> - 1, <%=b_day %>, <%=b_hour %>, <%=b_minute %>);
	var nowdate = new Date();

	if (eventDate.getTime() >= nowdate.getTime())
	{
		var num_min = parseInt((eventDate.getTime() - nowdate.getTime())/(1000 * 60));
		var num_hour = parseInt(num_min/60)

		if (num_hour > 0)
			num_min = num_min - num_hour*60;

		if (num_hour > 0)
			document.write("<%=s_lang_0522 %>" + num_hour.toString() + "<%=s_lang_0523 %>" + num_min.toString() + "<%=s_lang_0524 %>");
		else
			document.write("<%=s_lang_0522 %>" + num_min.toString() + "<%=s_lang_0524 %>");
	}
	else
	{
<%
if ecal.bi_isRepeat = false then
%>
		if (eventDate.getDate() == nowdate.getDate())
			document.write("<%=s_lang_0525 %>");
		else
			document.write("<%=s_lang_0526 %>");
<%
else
%>
		document.write("<%=s_lang_0527 %>");
<%
end if
%>
	}
}
<%
end if
%>
//-->
</script>
<table width=100% border=0 cellspacing=0 cellpadding=0>
<tr>
<td width="50%" style="border:5px #ffffff solid;" valign="top">
	<table width=100% border=0 cellspacing=0 cellpadding=0 bgcolor="<%=MY_COLOR_3 %>" style="border:1px solid #8CA5B5;">
	<tr>
	<td height="23" align="center" style="border-bottom:1px solid #8CA5B5;">
<font class=s color="#104A7B"><b><%
if Len(invitationID) > 10 then
	Response.Write s_lang_0508
elseif Len(remindID) > 10 then
	Response.Write s_lang_0509
end if
%></b></font>
	</td>
	</tr>
	<tr>
	<td height="10"></td>
	</tr>
<%
if Len(invitationID) > 10 then
%>
	<tr>
	<td height="33" valign="top">
&nbsp;&nbsp;<input type="button" value="<%=s_lang_0528 %>" class="Bsbttn" LANGUAGE=javascript onclick="showevent()">
	</td>
	</tr>
	<tr>
	<td height="33" valign="top">
&nbsp;&nbsp;<input type="button" value="<%=s_lang_0529 %>" class="Bsbttn" LANGUAGE=javascript onclick="showntlist()">
	</td>
	</tr>
	<tr>
	<td height="33" valign="top">
&nbsp;&nbsp;<input type="button" value="<%=s_lang_0530 %>" class="Bsbttn" LANGUAGE=javascript onclick="showcal()">
	</td>
	</tr>
<%
elseif Len(remindID) > 10 then
%>
	<tr>
	<td height="55" valign="top">
&nbsp;&nbsp;<img src='images/cal/clock.gif' border='0' title="<%=s_lang_0531 %>">&nbsp;&nbsp;[<font color='#BC131A'><b><script>write_clockstr();</script></b></font>]
	</td>
	</tr>
	<tr>
	<td height="33" valign="top" align="center">
<input type="button" value="<%=s_lang_0532 %>" class="Bsbttn" LANGUAGE=javascript onclick="showevent()">
	</td>
	</tr>
<%
end if
%>
	</table>
</td>
<td valign="top">
<table width=100% border=0 cellspacing=0 cellpadding=0>
<tr><td width="100%" style="border:5px #ffffff solid;">
	<table width=100% border=0 cellspacing=0 cellpadding=0 bgcolor="#ffffff" style="border-top:1px #b0b0b0 solid;">
	<tr>
	<td height="23" colspan="2" bgcolor="#eeeeee" style="border-bottom:1px #b0b0b0 solid;">
&nbsp;<b><%=s_lang_0533 %></b>
</td>
	</tr>
<tr><td height="3" colspan="2"></td></tr>
<tr><td height="6" colspan="2" bgcolor="#ffffc0"></td></tr>
	<tr>
	<td colspan="2" bgcolor="#ffffc0" style="border-left:7px #ffffc0 solid; border-right:7px #ffffc0 solid;">
<font class=s color="#104A7B"><script>write_getShowIconStr(<%=ecal.bi_mode %>,<%=LCase(CStr(ecal.bi_remind)) %>,<%=LCase(CStr(ecal.bi_isRepeat)) %>)</script><b><%=server.htmlencode(ecal.bi_name) %></b></font><br>
<font class=s color="#104A7B"><%
ht = server.htmlencode(ecal.bi_note)
ht = replace(ht, Chr(13), "<br>")
ht = replace(ht, Chr(32), "&nbsp;")
ht = replace(ht, Chr(9), "&nbsp;&nbsp;&nbsp;&nbsp;")
Response.Write ht
%></font>
	</td>
	</tr>
<tr><td height="3" colspan="2" bgcolor="#ffffc0"></td></tr>
<tr><td height="4" colspan="2"></td></tr>
	<tr>
	<td width="30%">
&nbsp;<b><%=s_lang_0534 %></b>:
	</td>
	<td>
<%
if Len(invitationID) > 10 then
	Response.Write server.htmlencode(ecalnt.bi_host_account)

	if LCase(ecal.bi_notice_name) <> LCase(ecalnt.bi_host_account) then
		Response.Write "&nbsp;(" & server.htmlencode(ecal.bi_notice_name) & ")"
	end if
elseif Len(remindID) > 10 then
	Response.Write server.htmlencode(ecal.bi_notice_name)
end if
%>
	</td>
	</tr>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b><%=s_lang_0535 %></b>:
	</td>
	<td>
<%
Response.Write "<a href=""mailto:" & server.htmlencode(ecal.bi_notice_email) & "?subject=" & server.htmlencode(ecal.bi_name) & """>" & server.htmlencode(ecal.bi_notice_email) & "</a>"
%>
	</td>
	</tr>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b><%=s_lang_0536 %></b>:
	</td>
	<td>
<%
Response.Write "<script>write_ShowDateUrl(" & b_year & "," & b_month & "," & b_day & ")</script>"
%>
	</td>
	</tr>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b><%=s_lang_0537 %></b>:
	</td>
	<td>
<%
Response.Write "<script>write_ShowEventTime(" & b_year & "," & b_month & "," & b_day & "," & b_hour & "," & b_minute & "," & ecal.bi_needtime & ")</script>"
%>
	</td>
	</tr>
<%
if ecal.bi_isRepeat = true then
%>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b><%=s_lang_0538 %></b>:
	</td>
	<td>
<%=s_lang_0539 %>
	</td>
	</tr>
<%
end if

if Len(ecal.bi_place) > 0 then
%>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b><%=s_lang_0540 %></b>:
	</td>
	<td>
<%=server.htmlencode(ecal.bi_place) %>
	</td>
	</tr>
<%
end if

if Len(ecal.bi_city) > 0 then
%>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b><%=s_lang_0541 %></b>:
	</td>
	<td>
<%=server.htmlencode(ecal.bi_city) %>
	</td>
	</tr>
<%
end if

if Len(ecal.bi_address) > 0 then
%>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b><%=s_lang_0542 %></b>:
	</td>
	<td>
<%=server.htmlencode(ecal.bi_address) %>
	</td>
	</tr>
<%
end if

if Len(ecal.bi_phone) > 0 then
%>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b><%=s_lang_0543 %></b>:
	</td>
	<td>
<%=server.htmlencode(ecal.bi_phone) %>
	</td>
	</tr>
<%
end if

if Len(ecal.bi_other_phone) > 0 then
%>
<tr><td height="3" colspan="2"></td></tr>
	<tr>
	<td>
&nbsp;<b><%=s_lang_0544 %></b>:
	</td>
	<td>
<%=server.htmlencode(ecal.bi_other_phone) %>
	</td>
	</tr>
<%
end if

if Len(invitationID) > 10 then
	show_bi_notice_datetime = ecalnt.show_bi_notice_datetime
elseif Len(remindID) > 10 then
	show_bi_notice_datetime = ecal.show_bi_notice_datetime
end if

if Len(show_bi_notice_datetime) > 0 then
%>
<tr><td height="6" colspan="2"></td></tr>
	<tr>
	<td colspan="2">
&nbsp;<b><%=s_lang_0545 %>&nbsp;<%=server.htmlencode(show_bi_notice_datetime) %></b>
	</td>
	</tr>
<%
end if
%>
</table>
</td>
</tr>
</table>
</td>
</tr>
</table>
<%
	b_year = NULL
	b_month = NULL
	b_day = NULL
	b_hour = NULL
	b_minute = NULL
end if
%>
</td></tr>
</table>

<table align="center" cellspacing="0" cellpadding="0" class="bt_tool_table">
<%
if is_inline = false then
	if subismessage = false then
%>
	<tr><td class="block_top_td" style="height:12px;"><div class="table_min_width"></div></td></tr>
	<tr><td class="tool_top_td">

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:back()"><< <%=s_lang_return %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=reply&<%=getGRSN() %>&filename=<%=filename & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&backurl=showarc.asp?" & enc_rqs %>"><%=s_lang_0460 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=replyall&<%=getGRSN() %>&filename=<%=filename & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&backurl=showarc.asp?" & enc_rqs %>"><%=s_lang_0461 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

<%
if ei.IsSaveMail = false then
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=forward&<%=getGRSN() %>&filename=<%=filename & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&backurl=showarc.asp?" & enc_rqs %>"><%=s_lang_0462 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
else
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=forward&<%=getGRSN() %>&filename=<%=filename & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&backurl=showarc.asp?" & enc_rqs %>"><%=s_lang_0463 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
end if

if sname = "" or sfname = "" then
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:delthis();"><%=s_lang_del %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

<%
end if
%>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="<%
	Response.Write "showeml.asp?outmode=down&filename=" & filename & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN()
%>" target="_blank"><%=s_lang_0373 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span id="bottom_pn" class="st_right_span">
	</span>

	</td></tr>
	<tr><td class="block_top_td" style="height:16px;"></td></tr>
	<tr><td align="right">
	<span style="margin-right:16px;"><a href="#gotop"><img src='images/gotop.gif' border='0' title="<%=s_lang_0152 %>"></a></span>
<%
	end if
else
%>
	<tr><td class="block_top_td" style="height:12px;"><div class="table_min_width"></div></td></tr>
	<tr><td class="il_tool_top_td">
	<span style='float:left; width:10px;'>&nbsp;</span>

	<span class="st_span">
	<a href="javascript:location_href('wframe.asp?mode=reply&<%=getGRSN() %>&filename=<%=filename & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&backurl=goback" %>')" class="bg_href"><%=s_lang_0460 %></a>
	</span>
	<span style='float:left; width:30px;'>&nbsp;</span>

	<span class="st_span">
	<a href="javascript:location_href('wframe.asp?mode=replyall&<%=getGRSN() %>&filename=<%=filename & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&backurl=goback" %>')" class="bg_href"><%=s_lang_0461 %></a>
	</span>
	<span style='float:left; width:30px;'>&nbsp;</span>

<%
if ei.IsSaveMail = false then
%>
	<span class="st_span">
	<a href="javascript:location_href('wframe.asp?mode=forward&<%=getGRSN() %>&filename=<%=filename & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&backurl=goback" %>')" class="bg_href"><%=s_lang_0462 %></a>
	</span>
	<span style='float:left; width:30px;'>&nbsp;</span>
<%
else
%>
	<span class="st_span">
	<a href="javascript:location_href('wframe.asp?mode=forward&<%=getGRSN() %>&filename=<%=filename & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&backurl=goback" %>')" class="bg_href"><%=s_lang_0463 %></a>
	</span>
	<span style='float:left; width:30px;'>&nbsp;</span>
<%
end if

if sname = "" or sfname = "" then
%>
	<span class="st_span">
	<a href="javascript:delthis();" class="bg_href"><%=s_lang_del %></a>
	</span>
	<span style='float:left; width:30px;'>&nbsp;</span>

<%
end if
%>

	<span class="st_span">
	<a href="<%
	Response.Write "showeml.asp?outmode=down&filename=" & filename & "&sname=" & Server.URLEncode(sname) & "&sfname=" & Server.URLEncode(sfname) & "&" & getGRSN()
%>" target="_blank" class="bg_href"><%=s_lang_0373 %></a>
	</span>
	<span style='float:left; width:30px;'>&nbsp;</span>

	<span class="st_span">
	<a href="javascript:parent_href('showarc.asp?filename=<%=filename %>&gourl=<%=enc_gourl %>')" class="bg_href"><%=s_lang_0547 %></a>
	</span>
<%
end if
%>
	</td></tr>
</table>
<%
if is_inline = false and allnum > 0 then
%>
<a name="goatt">&nbsp;</a>
<script type="text/javascript">
<!--
$("#go_att").append("&nbsp;<a href='#goatt'><img src='images/atta.gif' border='0' title='<%=s_lang_0548 %>' align='absmiddle'></a>");
//-->
</script>
<%
end if

if is_inline = false then
	Response.Write "<br>"
end if
%>
</form>

<div id="top_show_msg" class="top_show_msg"></div>

<div id="pmc_moveto" class="qmpanel_shadow" style="display:none; position:absolute;">
	<div class="menu_base">
	<div id="md_moveto" class="menu_bd bd">
	<table width="100%">
		<div name="mi" onclick="movemail('del');" class="menu_item"><%=s_lang_0334 %></div>
		<div name="mi" onclick="movemail('in');" class="menu_item"><%=s_lang_0327 %></div>
		<div name="mi" onclick="movemail('out');" class="menu_item"><%=s_lang_0332 %></div>
		<div name="mi" onclick="movemail('sed');" class="menu_item"><%=s_lang_0430 %></div>
<%
dim moveto_set_max
moveto_set_max = false

if is_inline = false then
	if sname = "" or sfname = "" then

pfNumber = pf.FolderCount

if pfNumber > 0 then
	Response.Write "<div class='menu_item_nofun'><div style='background:#ccc; padding-top:1px; margin-top: 5px;'></div></div>"
end if

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

	Response.Write "<div name='mi' onclick=""movemail('" & server.htmlencode(spfname) & "');"" class='menu_item'>" & server.htmlencode(spfname) & "</div>" & Chr(13)
	spfname = NULL

	i = i + 1
loop

	end if
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
if sname = "" or sfname = "" then
	set pf = nothing
end if


set ecal = nothing
set ecalnt = nothing

set wemcert = nothing
set ei = nothing


function RemoveEndRN(ostr)
	dim rern_haveRN
	dim rern_len
	dim rern_char

	rern_haveRN = false
	rern_len = Len(ostr)

	do while rern_len > 1
		rern_char = Mid(ostr, rern_len, 1)

		if rern_char <> Chr(13) and rern_char <> Chr(10) then
			Exit Do
		else
			rern_haveRN = true
		end if

		rern_len = rern_len - 1
	loop

	if rern_haveRN = true and rern_len > 0 then
		RemoveEndRN = Mid(ostr, 1, rern_len)
	else
		RemoveEndRN = ostr
	end if
end function

function get_date_showstr(show_date_str)
	if Len(show_date_str) = 14 or Len(show_date_str) = 12 then
		tmp_month = Mid(show_date_str, 5, 2)
		if Mid(tmp_month, 1, 1) = "0" then
			tmp_month = Mid(tmp_month, 2, 1)
		end if

		tmp_day = Mid(show_date_str, 7, 2)
		if Mid(tmp_day, 1, 1) = "0" then
			tmp_day = Mid(tmp_day, 2, 1)
		end if

		get_date_showstr = Mid(show_date_str, 1, 4) & s_lang_0139 & tmp_month & s_lang_0140 & tmp_day & s_lang_0141
	else
		get_date_showstr = ""
	end if
end function

function get_time_showstr(show_date_str)
	if Len(show_date_str) = 14 or Len(show_date_str) = 12 then
		t_time_hour = CInt(Mid(show_date_str, 9, 2))

		dim t_hour_name
		if t_time_hour >= 0 and t_time_hour < 6 then
			t_hour_name = s_lang_0268
		elseif t_time_hour >= 6 and t_time_hour < 12 then
			t_hour_name = s_lang_0269
		elseif t_time_hour >= 12 and t_time_hour < 14 then
			t_hour_name = s_lang_0283
			if t_time_hour > 12 then
				t_time_hour = t_time_hour - 12
			end if
		elseif t_time_hour >= 14 and t_time_hour < 18 then
			t_hour_name = s_lang_0270
			t_time_hour = t_time_hour - 12
		elseif t_time_hour >= 18 then
			t_hour_name = s_lang_0271
			t_time_hour = t_time_hour - 12
		end if

		ts_time_hour = t_time_hour

		get_time_showstr = t_hour_name & ts_time_hour& ":" & Mid(show_date_str, 11, 2)
	else
		get_time_showstr = ""
	end if
end function

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

Function UnEscape(str)
    dim i,s,c
    s=""
    For i=1 to Len(str)
        c=Mid(str,i,1)
        If Mid(str,i,2)="%u" and i<=Len(str)-5 Then
            If IsNumeric("&H" & Mid(str,i+2,4)) Then
                s = s & CHRW(CInt("&H" & Mid(str,i+2,4)))
                i = i+5
            Else
                s = s & c
            End If
        ElseIf c="%" and i<=Len(str)-2 Then
            If IsNumeric("&H" & Mid(str,i+1,2)) Then
                s = s & CHRW(CInt("&H" & Mid(str,i+1,2)))
                i = i+2
            Else
                s = s & c
            End If
        Else
            s = s & c
        End If
    Next
    UnEscape = s
End Function

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
