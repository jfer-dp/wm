<!--#include file="passinc.asp" --> 

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

isMobile = false
dim http_user_agent
http_user_agent = LCase(Request.ServerVariables("HTTP_User-Agent"))
if InStr(http_user_agent, "applewebkit") > 0 or InStr(http_user_agent, "mobile") > 0 then
	if InStr(http_user_agent, "iphone") > 0 or InStr(http_user_agent, "ipod") > 0 or InStr(http_user_agent, "android") > 0 or InStr(http_user_agent, "ios") > 0 or InStr(http_user_agent, "ipad") > 0 then
		isMobile = true
	end if
end if

gindex = trim(request("gindex"))
gourl = trim(request("gourl"))

dim userweb
set userweb = server.createobject("easymail.UserWeb")
userweb.Load Session("wem")

if userweb.useRichEditer = false or isMobile = true then
	useRichEditer = false
else
	useRichEditer = true
end if

MailName = userweb.MailName

if Len(MailName) < 1 then
	MailName = Session("wem")
end if


dim ads
set ads = server.createobject("easymail.Addresses")

if gindex <> "" and IsNumeric(gindex) = true then
	sortstr = request("sortstr")
	sortmode = request("sortmode")

	if sortstr <> "" then
		if sortmode = "1" then
			sortmode = true
	
			ads.SetSort sortstr, sortmode
		elseif sortmode = "0" then
			sortmode = false
	
			ads.SetSort sortstr, sortmode
		end if
	end if
end if


ads.Load Session("wem")


dim wemcert
set wemcert = server.createobject("easymail.WebEasyMailCert")
wemcert.Load Session("wem"), Session("mail")

if Session("scpw") <> "" then
	if wemcert.CheckPassIsGood(Session("scpw"), -2) = false then
		Session("scpw") = ""
	end if
end if
%>

<!DOCTYPE html>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
<link rel="stylesheet" type="text/css" href="images/selads.css">
<link rel="stylesheet" href="images/autocompleter.css" type="text/css" media="screen">

<script type="text/javascript" src="images/sc_left.js"></script>
<script language="JavaScript" type="text/javascript" src="rte/mootools.v1.11.js"></script>
<script language="JavaScript" type="text/javascript" src="rte/observer.js"></script>
<script language="JavaScript" type="text/javascript" src="rte/autocompleter.js"></script>
<script language="JavaScript" type="text/javascript" src="rte/wrte1.js"></script>
<script language="JavaScript" type="text/javascript" src="rte/wrte2.js"></script>
<script type="text/javascript" src="images/mglobal.js"></script>
<script type="text/javascript" src="images/selads.js"></script>

<style type="text/css">
html {overflow: scroll; overflow-x: hidden; overflow-y: auto !important;}
a:hover {color:#1e5494; text-decoration:underline}
a		{color:#1e5494; text-decoration:none}
body {margin:5px 10px 5px 10px;}
.textbox_wwm {padding:2px 8px 0pt 3px; border:1px solid #999;background-color:#FFFFEE; font-size:13px; width:490px; _width:500px;}
.textarea_wwm {padding:2px 8px 0pt 3px; border:1px solid #999; font-size:14px;}
.wwm_msg {padding:8px; margin:-6px 0 14px 0; color:#7E4F05; line-height:18px; background:#FFF3C3;border-radius:4px; -webkit-border-radius:4px;padding-left:20px;padding-right:20px;text-align:left;border: #7E4F05 1px solid;}
.wwm_ar_msg {padding:3px; margin:-6px 0 14px 0; color:#104A7B; line-height:18px; background:#e0ecf9;border-radius:4px; -webkit-border-radius:4px;padding-left:8px;padding-right:8px;text-align:left;border: #8db8e7 1px solid;}
.subject_input {font-size:14px; padding:5px 5px 3px 5px; outline:none; border:1px solid #999999; background-color:#FFFFEE; width:491px; _width:501px;}
.tr_st {height:26px; border-bottom: #A5B6C8 1px solid;}
.sbttn {font-family:宋体,MS SONG,SimSun,tahoma,sans-serif; font-size:9pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5;color:#000066;height:19px;text-decoration:none;cursor:pointer}

.box{border:1px solid <%
if isMSIE = true then
	Response.Write "#eff5fb"
else
	Response.Write "#999"
end if
%>;width:186px;overflow:hidden;<%
if isMSIE = false then
	Response.Write "height:72px;"
end if
%>}
.box2{border:1px solid #eff5fb;width:186px;overflow:hidden;}
.select_st{position:relative;left:-2px;top:-2px;font-size:12px;width:186px;line-height:14px;border:0px;color:#222222;background-color:#FFFFEE;<%
if isMSIE = false then
	Response.Write "height:72px;"
end if
%>}
</style>
</HEAD>

<script LANGUAGE=javascript>
<!--
if (ie == 6)
	document.execCommand("BackgroundImageCache", false, true); 

<%
if useRichEditer = true then
%>
initRTE("./rte/images/", "./rte/", "", false);
<%
end if
%>

var TIMEOUT_ALERT_STR = "您已经超时, 立即保存信息";

var allminutes = <%=Session.TimeOut %>;
setTimeout("Time()", 60000);

window.addEvent('domready', function(){
	var el = $ES('textarea');
	var tokens = new Array();
	var i = 0;
	var ads_allnum = parent.parent.f1.document.leftval.ads.length;

	for (i; i < ads_allnum; i++)
	{
		tokens[i] = new Array();
		tokens[i][0] = parent.parent.f1.document.leftval.ads[i].text + " ";
		tokens[i][1] = parent.parent.f1.document.leftval.ads[i].value;
	}

	var completer1 = new Autocompleter.Local(el[0], tokens, {
		'delay': 100,
		'filterTokens': function() {
			var regex = new RegExp('^' + this.queryValue.escapeRegExp(), 'i');
			return this.tokens.filter(function(token){
				return (regex.test(token[0]) || regex.test(token[1]));
			});
		},
		'injectChoice': function(choice) {
			var el = new Element('li')
				.setHTML(this.markQueryValue(choice[0]))
				.adopt(new Element('span', {'class': 'example-info'}).setHTML(this.markQueryValue(choice[1])));
			el.inputValue = choice[0];
			this.addChoiceEvents(el).injectInside(this.choices);
		}
	});

	var completer2 = new Autocompleter.Local(el[1], tokens, {
		'delay': 100,
		'filterTokens': function() {
			var regex = new RegExp('^' + this.queryValue.escapeRegExp(), 'i');
			return this.tokens.filter(function(token){
				return (regex.test(token[0]) || regex.test(token[1]));
			});
		},
		'injectChoice': function(choice) {
			var el = new Element('li')
				.setHTML(this.markQueryValue(choice[0]))
				.adopt(new Element('span', {'class': 'example-info'}).setHTML(this.markQueryValue(choice[1])));
			el.inputValue = choice[0];
			this.addChoiceEvents(el).injectInside(this.choices);
		}
	});

	var completer3 = new Autocompleter.Local(el[2], tokens, {
		'delay': 100,
		'filterTokens': function() {
			var regex = new RegExp('^' + this.queryValue.escapeRegExp(), 'i');
			return this.tokens.filter(function(token){
				return (regex.test(token[0]) || regex.test(token[1]));
			});
		},
		'injectChoice': function(choice) {
			var el = new Element('li')
				.setHTML(this.markQueryValue(choice[0]))
				.adopt(new Element('span', {'class': 'example-info'}).setHTML(this.markQueryValue(choice[1])));
			el.inputValue = choice[0];
			this.addChoiceEvents(el).injectInside(this.choices);
		}
	});
});

var new_ads_adg_number = <%=ads.EmailCount + ads.GroupCount %>;
if (parent.parent.f1.ar_is_request == true && parent.parent.f1.array_ads.length != new_ads_adg_number)
	parent.parent.f1.clean_ads();

function sub(smode){
	if (allminutes < 1)
	{
		alert(TIMEOUT_ALERT_STR);
		return ;
	}

	if (parent.f4.document.fsa.upfile.value != "")
	{
		var last_pt = parent.f4.document.fsa.upfile.value.lastIndexOf('\\');

		if (last_pt > -1)
		{
			if (confirm("您的附件 \"" + parent.f4.document.fsa.upfile.value.substring(last_pt + 1) + "\" 可能忘记上传, 是否继续?") == false)
				return ;
		}
		else
		{
			if (confirm("您的附件可能忘记上传, 是否继续?") == false)
				return ;
		}
	}

	if (check_sendto_number() == false)
	{
		alert("邮件接收地址过多.");
		document.f1.EasyMail_To.focus();
		return ;
	}

	if (document.f1.needCheckCertPassword.value == "1")
	{
		document.f1.EasyMail_CertPW.value = document.getElementById("CertPW").value;
		if (document.f1.EasyMail_CertPW.value.length < 8)
		{
			alert("密码输入错误!");
			document.getElementById("CertPW").focus();
			return ;
		}
	}

	if (document.f1.EasyMail_To.value != "")
	{
		if (smode == 1)
			document.getElementById("sending").style.display = "inline";
		else
			document.getElementById("esave").style.display = "inline";
<%
if useRichEditer = false then
%>
		dec_EasyMail_Text();
<%
else
%>
	updateRTE('richedit');

	document.f1.RichEdit_Text.value = getText(document.f1.richedit.value);
	document.f1.RichEdit_Html.value = document.f1.richedit.value;

	dec_RichEdit_Text();
	dec_RichEdit_Html();
<%
end if
%>

document.f1.AddFromAttFileString.value = "";

var i = 0;
for (i; i < document.f1.NetAtts.length; i++)
{
	document.f1.AddFromAttFileString.value = document.f1.AddFromAttFileString.value + document.f1.NetAtts[i].value + "\t";
}

		conv_zAttFileString();
		document.f1.submit();
	}
	else
	{
		alert("请输入收件人地址");
		document.f1.EasyMail_To.focus();
	}
}

function window_onload() {
flash_att_div();

<%
if isMSIE = false then
%>
if (window.screen.width < 900)
	parent.document.getElementById("f3").style.overflow = "scroll";
<%
end if
%>

var obj_textarea = document.getElementById("EasyMail_Text");
if (obj_textarea != null)
{
	obj_textarea.style.height = document.getElementById("tdtt").offsetHeight + "px";
	obj_textarea.style.width = document.getElementById("tdtt").offsetWidth + "px";
<%
if isMobile = false then
%>
	if (window.screen.width < 900)
		obj_textarea.style.width = "570px";
<%
end if
%>
}
else
{
	if (window.screen.width < 900)
	{
		document.getElementById("Buttons1_richedit").width = "570";
		document.getElementById("Buttons2_richedit").width = "570";
		document.getElementById("richedit").width = "570";
	}
}

<%
if gindex <> "" and IsNumeric(gindex) = true then
	ads.GetGroupInfo CInt(gindex), nickname, emails
	response.write "document.f1.EasyMail_To.value = """ & emails & """;"

	nickname = NULL
	emails = NULL
else
%>
	document.f1.EasyMail_To.value = parent.parent.f1.document.leftval.to.value;
	document.f1.EasyMail_Cc.value = parent.parent.f1.document.leftval.cc.value;
	document.f1.EasyMail_Bcc.value = parent.parent.f1.document.leftval.bcc.value;

	parent.parent.f1.document.leftval.to.value = "";
	parent.parent.f1.document.leftval.cc.value = "";
	parent.parent.f1.document.leftval.bcc.value = "";
<%
end if
%>

<%
if isMSIE <> true then
%>
	document.f1.EasyMail_To.rows = 3;
	document.f1.EasyMail_Cc.rows = 1;
	document.f1.EasyMail_Bcc.rows = 1;
<%
end if
%>

	hide_cc();
	hide_bcc();
	hide_ex();

	if (document.f1.EasyMail_Cc.value.length > 0)
		hide_cc();

	if (document.f1.EasyMail_Bcc.value.length > 0)
		hide_bcc();

try{
	if (parent.parent.f1.document.getElementById("leftval") != null)
	{
		if (parent.parent.f1.array_ads.length < 1)
		{
			parent.parent.f1.SendInfo();
			setTimeout("wait_left_send()", 10);
		}
		else
		{
			array_ads = parent.parent.f1.array_ads;
			main_write_ads(document.getElementById('main_dsearch').value.toLowerCase());
			main_check_search_str();
		}
	}
}catch(error){}

	ar_max_rq = 0;
	change_main_ads_height();

	var tmp_wd = document.getElementById("tcb_td").clientWidth;
	document.getElementById("EasyMail_To").style.width = (tmp_wd - 10) + "px";
	document.getElementById("EasyMail_Cc").style.width = (tmp_wd - 10) + "px";
	document.getElementById("EasyMail_Bcc").style.width = (tmp_wd - 10) + "px";
	document.getElementById("EasyMail_Subject").style.width = (tmp_wd - 10) + "px";
}

function addNetAtt(){
	if (document.f1.NetSaveAtts.selectedIndex != -1)
	{
		var oOption = document.createElement("OPTION");
		oOption.text = document.f1.NetSaveAtts[document.f1.NetSaveAtts.selectedIndex].text;
		oOption.value = document.f1.NetSaveAtts[document.f1.NetSaveAtts.selectedIndex].value;
<%
if isMSIE = true then
%>
		document.f1.NetAtts.add(oOption);
<%
else
%>
		document.f1.NetAtts.appendChild(oOption);
<%
end if
%>
		document.f1.NetAtts.selectedIndex = document.f1.NetAtts.length - 1;

		if (document.f1.NetAtts.length > 0)
			document.f1.NetAtts.size = document.f1.NetAtts.length;
		else
			document.f1.NetAtts.size = 1;
	}
}

function sec_onchange(){
<%
if wemcert.LightHasSecCert(Session("wem")) = false then
%>
	document.f1.EasyMail_CertServer.value = 0;
	alert("请先上传您的私人数字证书");
<%
else
%>
	if (document.f1.EasyMail_CertServer.value == 0)
	{
		document.f1.needCheckCertPassword.value = "0";
		document.getElementById("cert_check_1").style.display = "none";
		document.getElementById("cert_check_2").style.display = "none";
	}
	else if (document.f1.EasyMail_CertServer.value == 1)
	{
		document.getElementById("cert_check_1").style.display = "inline";
		document.getElementById("cert_check_2").style.display = "none";
	}
	else if (document.f1.EasyMail_CertServer.value == 2)
	{
		document.getElementById("cert_check_1").style.display = "none";
		document.getElementById("cert_check_2").style.display = "inline";
	}
<%
	if Session("scpw") = "" and wemcert.IsNeedSecCertPassword = true then
%>
	if (document.f1.EasyMail_CertServer.value > 0)
		document.f1.needCheckCertPassword.value = "1";
<%
	end if
end if
%>}

function check_sendto_number() {
	var all_sendto_num = 0;

	if (document.f1.EasyMail_To.value != "")
		all_sendto_num = all_sendto_num + get_char_number(document.f1.EasyMail_To.value, ",") + 1;

	if (document.f1.EasyMail_Cc.value != "")
		all_sendto_num = all_sendto_num + get_char_number(document.f1.EasyMail_Cc.value, ",") + 1;

	if (document.f1.EasyMail_Bcc.value != "")
		all_sendto_num = all_sendto_num + get_char_number(document.f1.EasyMail_Bcc.value, ",") + 1;

	if (<%
set sysinfo = server.createobject("easymail.sysinfo")
sysinfo.Load

Application("em_Enable_MailRecall") = sysinfo.Enable_MailRecall

Response.Write sysinfo.Web_Max_Recipients

set sysinfo = nothing
%> < all_sendto_num)
		return false;

	return true;
}

function checkTime()
{
	var err = "日期错误";
	var nowdate = new Date(<%=Year(now()) & "," & Month(now()) - 1 & "," & Day(now()) & "," & Hour(now()) & "," & Minute(now()) %>);
	var mydate = new Date(document.f1.t_year.value, document.f1.t_month.value - 1, document.f1.t_day.value, document.f1.t_hour.value, 1);

	var nmonth = document.f1.t_month.value;
	var nday = document.f1.t_day.value;
	var nhour = document.f1.t_hour.value;

	document.getElementById("ex_showstr").innerHTML = "隐藏扩展功能";
	document.getElementById("ex_function_div").style.display = "inline";
	exfunc_is_show = false;

	if (document.f1.t_year.value == "" || document.f1.t_year.value > 9999 || document.f1.t_year.value < <%=Year(now()) %>)
	{
		alert(err);
		document.f1.t_year.focus();
		return false;
	}

	if (nmonth == "" || nmonth > 12 || nmonth < 1)
	{
		alert(err);
		document.f1.t_month.focus();
		return false;
	}

	if (nday == "" || nday > 31 || nday < 1)
	{
		alert(err);
		document.f1.t_day.focus();
		return false;
	}

	if (nhour == "" || nhour > 23 || nhour < 0)
	{
		alert(err);
		document.f1.t_hour.focus();
		return false;
	}

	if (mydate > nowdate)
	{
		if (document.f1.t_month.value < 10)
			nmonth = "0" + cutz(document.f1.t_month.value);

		if (document.f1.t_day.value < 10)
			nday = "0" + cutz(document.f1.t_day.value);

		if (document.f1.t_hour.value < 10)
			nhour = "0" + cutz(document.f1.t_hour.value);

		if (nhour == "0")
			nhour = "00"

		document.f1.EasyMail_TimerSend.value = document.f1.t_year.value + nmonth + nday + nhour;
	}
	else
	{
		alert("定时发送的日期应该比现在迟.");
		document.f1.t_hour.focus();
		return false;
	}

	return true;
}
//-->
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
<br>
<div style="position:absolute; left:12px; top:10px;">
<a href="help.asp#writemail" target="_blank"><img src="images/help.gif" border="0" title="帮助"></a></div>
   <FORM ACTION="sendmail.asp?<%=getGRSN() %>" METHOD=POST NAME="f1" target="_parent">
    <INPUT NAME="EasyMail_CharSet" TYPE="hidden" Value="<%=userweb.CharSet %>">
    <INPUT NAME="useRichEditer" TYPE="hidden" Value="<%
if useRichEditer = true then
	Response.Write "true"
else
	Response.Write "false"
end if
%>">
    <INPUT NAME="needCheckCertPassword" TYPE="hidden" value="0">
    <INPUT NAME="EasyMail_CertPW" TYPE="hidden">
    <INPUT NAME="AddFromAttFileString" TYPE="hidden">
    <input id="zAttFileString" name="zAttFileString" type="hidden">
    <INPUT NAME="MailName" TYPE="hidden" Value="<%=MailName %>">
    <INPUT NAME="SendMode" TYPE="hidden" Value="send">
    <INPUT NAME="EasyMail_From" TYPE="hidden" value="<%= Session("wem")%>" maxlength="64">
    <INPUT NAME="EasyMail_TimerSend" TYPE="hidden" maxlength="10">
  <table id="title_table" width="96%" border="0" bgColor="white" align="center" cellspacing="0">
    <tr><td colspan="2" width="100%" align="left" noWrap>
	<table width="100%" cellspacing=0 cellpadding=0>
	<tr>
	<td height="30" align="left" width="30%" style="padding-left:14px;">
[<a href="javascript:hide_cc()"><span id="cc_showstr">显示抄送地址</span></a>&nbsp;|&nbsp;<a href="javascript:hide_bcc()"><span id="bcc_showstr">显示暗送地址</span></a>]
	</td>
	<td height="30" align="right" width="70%" style="padding-right:20px;">
		<a class="wwm_btnDownload btn_blue" style="width: 60px" href="#" onclick="javascript:showSending()">发送</a>&nbsp;
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:showSave()">保存</a>&nbsp;
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:timerSending()">定时发送</a>&nbsp;
<% if gourl = "" then %>
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:parent.location.href='viewmailbox.asp?<%=getGRSN() %>'">取消</a>
<%
else
	if InStr(gourl, "?") > 0 then
%>
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:parent.location.href='<%=gourl %>&<%=getGRSN() %>'">取消</a>
<% else %>
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:parent.location.href='<%=gourl %>?<%=getGRSN() %>'">取消</a>
<%
	end if
end if
%>
	</td></tr>
	</table>
      </td>
    </tr>
    <tr> 
      <td colspan="2" height="15" align="center">
	<hr size="1" Color="#A5B6C8">
      </td>
    </tr>
    <tr> 
      <td width="100%">
<table width="90%" cellspacing=0 cellpadding=0 align="left" border="0">
	<tr>
	<td noWrap align="right" width="70" style="padding-left:2px;"><a class="wwm_btnDownload btn_gray" style="width:38px; padding-left:4px; padding-right:4px;" href="javascript:selectAdd('To')">收件人</a>：&nbsp;</td>
	<td id="tcb_td" noWrap align="left">
<textarea autocomplete="off" name="EasyMail_To" id="EasyMail_To" size="40" class='textbox_wwm' onfocus="setIt('EasyMail_To')" cols="40" rows="4"><%
	if trim(request("addresssend")) <> "" then
		Response.Write mto
	end if
%></textarea>
	</td>
	<td noWrap align="left" width="4" style="padding-left:1px; padding-right:9px;">
<%
if Application("em_EnableEntAddress") = true then
%><a href="javascript:eapop('To')" title="企业地址簿"><img src="images/entads.gif" border="0" align="absmiddle"></a><%
end if
%></td>
	</tr>
</table>
	</td>
	<td valign=top align="left" rowspan="5">
<div id="main_ads" style="width:190px; height:130px; border:1px solid #999999;">

<div style="padding:4px 4px 4px 4px; border-bottom:1px solid #d3e1f0;">
<div style="border:1px solid #888888; width:180px; display:inline-block;">
<input type="text" id="main_dsearch" onkeyup="main_ds_keyup();" style="font-size:12px; border:0px; width:148px; height:18px; padding-left:3px; _margin:2px 0px -2px 1px;">
<span id="main_sicon" style="background-image:url(images/ok_search.gif); background-repeat:no-repeat; border:0px; width:15px; height:15px; font-size:10px; cursor:pointer; display:inline-block; margin:2px 4px -2px 0px; _margin:-1px 4px 1px 0px;" onclick="main_icon_click();"></span>
</div>
</div>

<div id="main_ads_in" style="width:190px; height:95px; overflow-x:hidden; overflow-y:auto;">
<div id="main_ads_name_id" class="s_ads_name">地址簿</div>
<div id="main_left_ads_div"></div>
<div id="main_adg_name_id" class="s_ads_name" style="border-top:1px solid #d3e1f0; display:none;">通迅组</div>
<div id="main_left_adg_div" style="display:none;"></div>
</div>

</div>
	</td>
  </tr>
  <tr><td width="100%">
<div id="cc_div" style="display:none;">
<table cellspacing=0 cellpadding=0 width="90%" align="left" border="0"><tr>
	<td noWrap align="right" width="70" style="padding-left:2px; padding-bottom:4px; *padding-bottom:0px;"><a class="wwm_btnDownload btn_gray" style="width:38px; padding-left:4px; padding-right:4px;" href="javascript:selectAdd('Cc')">抄送</a>：&nbsp;</td>
	<td noWrap align="left">
<textarea autocomplete="off" name="EasyMail_Cc" id="EasyMail_Cc" size="48" class='textbox_wwm' onfocus="setIt('EasyMail_Cc')" cols="40" rows="2" style="height:32px;"><%
	if trim(request("addresssend")) <> "" then
		response.write " value=""" & mcc & """"
	end if
%></textarea>
	</td>
	<td noWrap align="left" width="4" style="padding-left:1px; padding-right:9px;">
<%
if Application("em_EnableEntAddress") = true then
%><a href="javascript:eapop('Cc')" title="企业地址簿"><img src="images/entads.gif" border="0" align="absmiddle"></a><%
end if
%></td></tr></table>
</div>
</td></tr>

  <tr><td width="100%">
<div id="bcc_div" style="display:none;">
<table cellspacing=0 cellpadding=0 width="90%" align="left" border="0"><tr>
	<td noWrap align="right" width="70" style="padding-left:2px; padding-bottom:4px; *padding-bottom:0px;"><a class="wwm_btnDownload btn_gray" style="width:38px; padding-left:4px; padding-right:4px;" href="javascript:selectAdd('Bcc')">暗送</a>：&nbsp;</td>
	<td noWrap align="left">
<textarea autocomplete="off" name="EasyMail_Bcc" id="EasyMail_Bcc" size="48" class='textbox_wwm' onfocus="setIt('EasyMail_Bcc')" cols="40" rows="2" style="height:32px;"><%
	if trim(request("addresssend")) <> "" then
		response.write " value=""" & mbcc & """"
	end if
%></textarea>
	</td>
	<td noWrap align="left" width="4" style="padding-left:1px; padding-right:9px;">
<%
if Application("em_EnableEntAddress") = true then
%><a href="javascript:eapop('Bcc')" title="企业地址簿"><img src="images/entads.gif" border="0" align="absmiddle"></a><%
end if
%></td></tr></table>
</div>
</td></tr>
  <tr><td width="100%">
	<table cellspacing=0 cellpadding=0 width="100%" align="left" border="0" style="padding-bottom:4px; _padding-bottom:2px;"><tr>
	<td noWrap align=right style="width:68px; _width:70px;">主题：&nbsp;</td>
	<td align="left"><input name="EasyMail_Subject" id="EasyMail_Subject" type="text" maxlength="512" class='subject_input'>
	</td></tr></table>
	</td></tr>

	<tr><td noWrap width="60%" align="left" style="padding-left:3px; padding-top:4px; _padding-top:2px;">
<span id="all_att_bt_div"<%
if isMobile = true then
	Response.Write " style='display:none;'"
end if
%>>
<%
if trim(request.Cookies("cookie_ZATT_Is_Enable")) = "False" then
%>
<a class="wwm_btnDownload btn_gray" href="javascript:addatt()">添加附件</a>
<%
else
%>
<a class="wwm_btnDownload btn_gray" href="javascript:addatt()">添加附件</a>
<a class="wwm_btnDownload btn_gray" href="javascript:addzatt()">添加链接式附件</a>
<a class="wwm_btnDownload btn_gray" href="javascript:addzatt_fts()">从中转站添加链接式附件</a>
<%
end if
%>
</span>
&nbsp;&nbsp;<span class="wwm_ar_msg">提醒:无动作超时时间还剩<font color="#FF3333"><b><span id="minpt"><%=Session.TimeOut %></span></b></font>分钟</span>
	</td>
	</tr>
</table>
	<div id="uploading" class="c_attdiv" style="display:none; background-image:url(images/load.gif); background-position:left; background-repeat:no-repeat;"><span style="padding-left:22px;"></span>上传附件<span id="upname"></span>&nbsp;<a href="#" onclick="hide_up_att()"><img src="images\filter.gif" border="0" align="absmiddle"></a></div><div id="pattdiv"></div><div id="pzattdiv"></div>
<table align="center" cellspacing=0 cellpadding=0 width="98%" style="padding-top:8px; _padding-top:4px;">
	<tr>
	<td colspan="2" align="center">
<%
if useRichEditer = false then
%>
<table width="97%" align="center" cellspacing=0 cellpadding=0><tr><td height="315" width="100%" name="tdtt" id="tdtt">
	<textarea name="EasyMail_Text" id="EasyMail_Text" cols="76" rows="12" class='textarea_wwm'></textarea>
</td></tr></table>
<%
else
%>
<table width="98%" align="center" cellspacing=0 cellpadding=0><tr><td>
<script language="JavaScript" type="text/javascript">
<!--
writeRichText('richedit', '', "100%", <%
if isMSIE = true then
	Response.Write "270"
else
	Response.Write "253"
end if
%>, true, false);
//-->
</script>
</td></tr></table>
<%
end if
%>
      </td>
    </tr>
	<tr>
	<td height="30" colspan="2" align="left" noWrap>
	<table width="99%" cellspacing=0 cellpadding=0 border="0">
	<tr><td width="20%" style="padding-left:20px;">
	[<a href="javascript:hide_ex()"><span id="ex_showstr"></span></a>]
	</td>
	<td height="30" align="right" style="padding-right:20px;">
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:showSending()">发送</a>&nbsp;
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:showSave()">保存</a>&nbsp;
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:timerSending()">定时发送</a>&nbsp;
<% if gourl = "" then %>
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:parent.location.href='viewmailbox.asp?<%=getGRSN() %>'">取消</a>
<%
else
	if InStr(gourl, "?") > 0 then
%>
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:parent.location.href='<%=gourl %>&<%=getGRSN() %>'">取消</a>
<% else %>
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:parent.location.href='<%=gourl %>?<%=getGRSN() %>'">取消</a>
<%
	end if
end if
%>
	</td></tr>
	</table>
	</td>
	</tr>

    <tr>
    <td noWrap colspan="2">
<div id="ex_function_div">
	<table width="96%" align="center" border="0" cellspacing="0">
	<tr><td width="20%" align="right" noWrap class="tr_st" style="BORDER-TOP: #A5B6C8 1px solid;">保存副本到发件箱：</td>
	<td width="80%" noWrap align="left" class="tr_st" style="BORDER-TOP: #A5B6C8 1px solid;">
	<input type="checkbox" name="EasyMail_SendBackup" <% if userweb.EnableBackupAllSendMail = true then response.write "checked"%>>备份
	</td></tr>
    <tr><td noWrap align="right" class="tr_st">邮件等级：</td>
	<td class="tr_st" align="left">
	<select name="EasyMail_Priority" class=drpdwn>
	<option value="Normal">普通邮件</option>
	<option value="Low">慢件</option>
	<option value="High">紧急邮件</option>
	</select>
	</td></tr>
	<tr><td noWrap align="right" class="tr_st">签名：</td>
	<td class="tr_st" align="left">
	<select name="EasyMail_SignNo" class=drpdwn>
<%
ds = userweb.defaultSign
if ds = -1 then
%>
	<option value="-1" selected>不使用</option>
<%
else
%>
	<option value="-1">不使用</option>
<%
end if


dim sm
set sm = server.createobject("easymail.SignManager")
sm.Load Session("wem")

allnum = sm.count
i = 0

do while i < allnum
	sm.get i, s_title, s_text, shtmltext

	if i <> ds then
		response.write "<option value='" & i & "'>" & server.htmlencode(s_title) & "</option>"
	else
		response.write "<option value='" & i & "' selected>" & server.htmlencode(s_title) & "</option>"
	end if

	s_title = NULL
	s_text = NULL
	shtmltext = NULL

	i = i + 1
loop

set sm = nothing
%>
	</select>
	</td></tr>
	<tr><td noWrap align="right" class="tr_st">回复地址：</td>
	<td class="tr_st" align="left">
	<input name="EasyMail_BackAddress" type="text" size="30" value="<%=userweb.ReMail %>" class='n_textbox'>
	</td></tr>
	<tr><td align="right" noWrap class="tr_st">数字证书：</td>
	<td noWrap class="tr_st" align="left">
<select name="EasyMail_CertServer" class=drpdwn LANGUAGE=javascript onchange="return sec_onchange()">
	<option value="0" selected>不使用</option>
	<option value="1">数字签名并加密</option>
	<option value="2">数字签名</option>
</select>
<br>
<div id="cert_check_1" style="display:none;">
<a href='javascript:checkcanenc()'>验证是否可以向接收者发送加密邮件</a><%
if Session("scpw") = "" and wemcert.IsNeedSecCertPassword = true then
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;请输入您数字证书的密码：<input type='password' id='CertPW' class='n_textbox' size='11'>"
end if
%></div>
<div id="cert_check_2" style="display:none;">
<%
	if Session("scpw") = "" and wemcert.IsNeedSecCertPassword = true then
		Response.Write "请输入您数字证书的密码：<input type='password' id='CertPW' class='n_textbox' size='11'>"
	end if
%></div>
	</td></tr>
	<tr><td align="right" noWrap class="tr_st">由网络存储加入的附件：</td>
	<td class="tr_st" align="left">
	<select name="NetAtts" class=drpdwn size="1" multiple>
	</select> 
	<input type="button" value=" >> " onclick="javascript:delNetAtt()" class=sbttn>
	<input type="button" value=" << " onclick="javascript:addNetAtt()" class=sbttn>
	<select name="NetSaveAtts" class=drpdwn>
<%
dim nas
set nas = server.createobject("easymail.InfoList")
nas.LoadMailBox Session("wem"), "att"
allnum = nas.getMailsCount
i = 0

do while i < allnum
	nas.getMailInfoEx allnum - i - 1, idname, isread, priority, sendMail, sendName, subject, size, etime, mstate

	Response.Write "<option value='" & idname & "'>" & server.htmlencode(subject) & "</option>" & Chr(13)

	idname = NULL
	isread = NULL
	priority = NULL
	sendMail = NULL
	sendName = NULL
	subject = NULL
	size = NULL
	etime = NULL
	mstate = NULL

	i = i + 1
loop

set nas = nothing
%>
	</select> 
	</td></tr>
	<tr><td noWrap align="right" class="tr_st">读取确认：</td>
	<td class="tr_st" align="left">
	<input type="checkbox" name="EasyMail_ReadBack">读取确认(系统内用户)
	</td></tr>
	<tr><td noWrap align="right" class="tr_st">定时发送：</td>
	<td class="tr_st" align="left">
<select name="t_year" class="drpdwn">
<%
	now_temp = Year(Now())

	i = now_temp
	do while i < now_temp + 10
		response.write "<option value='" & i & "'>" & i & "年</option>" & Chr(13)
		i = i + 1
	loop
%>
</select>&nbsp;
<select name="t_month" class="drpdwn">
<%
	now_temp = Month(Now())
	i = 1
	do while i < 13
		if i <> now_temp then
			response.write "<option value='" & i & "'>" & i & "月</option>" & Chr(13)
		else
			response.write "<option value='" & i & "' selected>" & i & "月</option>" & Chr(13)
		end if
		i = i + 1
	loop
%>
</select>&nbsp;
<select name="t_day" class="drpdwn">
<%
	now_temp = Day(Now())
	i = 1
	do while i < 32
		if i <> now_temp then
			response.write "<option value='" & i & "'>" & i & "日</option>" & Chr(13)
		else
			response.write "<option value='" & i & "' selected>" & i & "日</option>" & Chr(13)
		end if
		i = i + 1
	loop
%>
</select>&nbsp;
<select name="t_hour" class="drpdwn">
<%
	i = 0
	do while i < 24
		response.write "<option value='" & i & "'>" & i & "时</option>" & Chr(13)
		i = i + 1
	loop
%>
</select>
	</td>
	</tr>
<%
if isadmin() = true then
%>
	<tr><td noWrap align="right" class="tr_st">系统邮件：</td>
	<td class="tr_st" align="left">
	<input type="checkbox" name="EasyMail_SystemMail">是否是系统邮件
	</td></tr>
    <tr><td noWrap align="right" class="tr_st">垃圾邮件投诉：</td>
	<td class="tr_st" align="left">
	<input type="checkbox" name="needAddInDebarList">此邮件不可被投诉为垃圾邮件
	</td></tr>
<%
end if
%>
	</table>
</div>
	</td>
	</tr>
</table>
<div id="sending" class="wwm_msg" style="position:absolute; top:62%; left:50%; margin:-100px 0 0 -100px; z-index:100; display:none;">邮件正在发送, 请稍候...</div>
<div id="esave" class="wwm_msg" style="position:absolute; top:62%; left:50%; margin:-100px 0 0 -100px; z-index:100; display:none;">邮件正在保存中, 请稍候...</div>
<div style="position:absolute; top:10; left:10; z-index:15; display:none;">
<textarea name="RichEdit_Text" cols="0" rows="0"></textarea>
<textarea name="RichEdit_Html" cols="0" rows="0"></textarea>
<select name="zAttName" id="zAttName"></select>
<table>
<tr>
<td id="add0"></td>
<td id="add1"></td>
<td id="add2"></td>
<td id="add3"></td>
<td id="add4"></td>
<td id="add5"></td>
<td id="add6"></td>
<td id="add7"></td>
<td id="add8"></td>
<td id="add9"></td>
</tr>
<tr>
<td id="add10"></td>
<td id="add11"></td>
<td id="add12"></td>
<td id="add13"></td>
<td id="add14"></td>
<td id="add15"></td>
<td id="add16"></td>
<td id="add17"></td>
<td id="add18"></td>
<td id="add19"></td>
</tr>
<tr>
<td id="add20"></td>
<td id="add21"></td>
<td id="add22"></td>
<td id="add23"></td>
<td id="add24"></td>
<td id="add25"></td>
<td id="add26"></td>
<td id="add27"></td>
<td id="add28"></td>
<td id="add29"></td>
</tr>
</table>
</div>
<br>
</FORM>

<div id="pop_ads_div" class="mydiv" style="display:none;">
	<div class="pop_base"><div class="pop_bd bd"><div class="title">
		<div class="title_left">选择邮件地址</div>
		<div class="title_right" title="关闭" onclick="javascript:close_ads(0);"><span>&nbsp;</span></div>
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
<div id="ads_name_id" class="s_ads_name">地址簿</div>
<div id="left_ads_div"></div>
<div id="adg_name_id" class="s_ads_name" style="border-top:1px solid #d3e1f0; display:none;">通迅组</div>
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
	<a id="pop_ok" class="wwm_btnDownload btn_gray" href="#" onclick="javascript:close_ads(1);">确定</a>&nbsp;
	<a class="wwm_btnDownload btn_gray" href="#" onclick="javascript:close_ads(0);">取消</a>
	</div></div></div></div>
</div>

<%
if Application("em_EnableEntAddress") = true then
%>
<div id="pop_entads_div" class="mydiv" style="display:none;">
	<div class="pop_base"><div class="pop_bd bd"><div class="title">
		<div class="title_left">选择企业地址簿中的邮件地址</div>
		<div class="title_right" title="关闭" onclick="javascript:close_ent_ads(0);"><span>&nbsp;</span></div>
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
<div id="ent_ads_name_id" class="s_ads_name">企业地址簿</div>
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
	<a id="btex_id" class="wwm_btnDownload btn_gray" href="#" onclick="javascript:exall();"><span id="btex">展开</span></a>&nbsp;
	<a class="wwm_btnDownload btn_gray" href="#" onclick="javascript:show_ent_find();"><span id="entf_bt">查找</span></a>&nbsp;
	<a id="pop_ok" class="wwm_btnDownload btn_gray" href="#" onclick="javascript:close_ent_ads(1);">确定</a>&nbsp;
	<a class="wwm_btnDownload btn_gray" href="#" onclick="javascript:close_ent_ads(0);">取消</a>
	</div></div></div></div>
</div>
<%
end if
%>

<div id="pop_msg_div" class="mydiv" style="display:none;">
	<div class="pop_base"><div class="pop_bd bd"><div class="title">
		<div class="title_left">信息</div>
		<div class="title_right" title="关闭" onclick="javascript:close_pop_msg();"><span>&nbsp;</span></div>
	</div>
	<div id="pop_msg_id" class="pop_content" style="height:140px; text-align:center; overflow-x:hidden; overflow-y:auto;">
	</div>
	<div class="title_bottom">
	<div class="title_ok_cancel_div">
	<a class="wwm_btnDownload btn_gray" href="#" onclick="javascript:close_pop_msg();">确定</a>
	</div></div></div></div>
</div>

<div id="bg" class="bg" style="display:none;"></div>
<iframe id='popIframe' class='popIframe' frameborder='0'></iframe>

</BODY>
</HTML>

<%
set userweb = nothing
set ads = nothing
set wemcert = nothing
%>
