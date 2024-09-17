<!--#include file="passinc.asp" --> 

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

dim dm
set dm = server.createobject("easymail.Domain")
dm.Load

if dm.GetUserManagerDomainCount(Session("wem")) < 1 then
	if isadmin() = false then
		set dm = nothing
		response.redirect "noadmin.asp"
	end if
end if



gourl = trim(request("gourl"))

dim userweb
set userweb = server.createobject("easymail.UserWeb")
'-----------------------------------------

userweb.Load Session("wem")

if userweb.useRichEditer = false then
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
ads.Load Session("wem")
%>

<!DOCTYPE html>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
<link rel="stylesheet" href="images/autocompleter.css" type="text/css" media="screen">
<script language="JavaScript" type="text/javascript" src="rte/wrte1.js"></script>
<script language="JavaScript" type="text/javascript" src="rte/wrte2.js"></script>

<style type="text/css">
html {overflow: scroll; overflow-x: hidden; overflow-y: auto !important;}
body {margin:5px 10px 5px 10px;}
.textarea_wwm {padding:2px 8px 0pt 3px; border:1px solid #999; font-size:14px;}
.wwm_msg {padding:8px; margin:-6px 0 14px 0; color:#7E4F05; line-height:18px; background:#FFF3C3;border-radius:4px; -webkit-border-radius:4px;padding-left:20px;padding-right:20px;text-align:left;border: #7E4F05 1px solid;}
.wwm_ar_msg {padding:3px; margin:-6px 0 14px 0; color:#104A7B; line-height:18px; background:#e0ecf9;border-radius:4px; -webkit-border-radius:4px;padding-left:8px;padding-right:8px;text-align:left;border: #8db8e7 1px solid;}
.tr_st {height:26px; border-bottom: #A5B6C8 1px solid;}
.tn_textbox {font-size:14px; padding:5px 5px 3px 5px; outline:none; border:1px solid #999999; background-color:#FFFFEE;}
.sbttn {font-family:宋体,MS SONG,SimSun,tahoma,sans-serif;font-size:9pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5;color:#000066;height:19px;text-decoration:none;cursor:pointer}
</style>
</HEAD>

<script LANGUAGE=javascript>
<!--
<%
if useRichEditer = true then
%>
initRTE("./rte/images/", "./rte/", "", false);
<%
end if
%>

var TIMEOUT_ALERT_STR = "您已经超时, 请立即保存信息";

var allminutes = <%=Session.TimeOut %>;
setTimeout("Time()", 60000);

function Time() {
	allminutes--;

	if (document.layers)
	{
		document.layers.minpt.document.write(allminutes.toString());
		document.layers.minpt.document.close();
	}
	else if (document.getElementById)
	{
		var theObj = document.getElementById("minpt");
		theObj.innerHTML = allminutes.toString();
	}

	if (allminutes > 0)
		setTimeout("Time()", 60000);
	else
		alert(TIMEOUT_ALERT_STR);
}


function showSending() {
if (document.f1.recdomains.length < 1)
	alert("请先添加接收邮件域名 ");
else
	sub(1);
}

function cutz(inval)
{
	var rval = "";

	for (var i = 0; i < inval.length; i++)
	{
		if (inval.charAt(i) != '0')
			break;
	}

	rval = inval.substring(i);

	return rval;
}

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

document.f1.EasyMail_DomainList.value = "\t";
i = 0;
for (i; i < document.f1.recdomains.length; i++)
{
	document.f1.EasyMail_DomainList.value = document.f1.EasyMail_DomainList.value + document.f1.recdomains[i].value + "\t";
}

	document.f1.submit();
}

function delNetAtt(){
	var i = 0;
	for (i; i < document.f1.NetAtts.length; i++)
	{
		if (document.f1.NetAtts.selectedIndex != -1)
		{
			document.f1.NetAtts.remove(document.f1.NetAtts.selectedIndex);
			i--;
		}
	}

	if (document.f1.NetAtts.length > 0)
		document.f1.NetAtts.size = document.f1.NetAtts.length;
	else
		document.f1.NetAtts.size = 1;

	document.f1.NetAtts.selectedIndex = 0;
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

function dec_EasyMail_Text()
{
	var count = 0;
	var theObj;
	var FormLimit = 50000;

	var TempVar = new String;
	TempVar = document.f1.EasyMail_Text.value;

	if (TempVar.length > FormLimit)
	{
		while (TempVar.length > 0 && count < 10)
		{
			theObj = document.getElementById("add" + count);
			theObj.innerHTML = "<Textarea rows=1 cols=1 name='Mdec_EasyMail_Text" + count + "'></Textarea>";

			theObj = eval("document.f1.Mdec_EasyMail_Text" + count);
			theObj.value = TempVar.substr(0, FormLimit);

			TempVar = TempVar.substr(FormLimit);

			count++;
		}
	}
	else
	{
		theObj = document.getElementById("add1");
		theObj.innerHTML = "<Textarea rows=1 cols=1 name='Mdec_EasyMail_Text1'></Textarea>";

		theObj = eval("document.f1.Mdec_EasyMail_Text1");
		theObj.value = TempVar;
	}
}

function dec_RichEdit_Text()
{
	var count = 10;
	var theObj;
	var FormLimit = 50000;

	var TempVar = new String;
	TempVar = document.f1.RichEdit_Text.value;

	if (TempVar.length > FormLimit)
	{
		while (TempVar.length > 0 && count < 20)
		{
			theObj = document.getElementById("add" + count);
			theObj.innerHTML = "<Textarea rows=1 cols=1 name='Mdec_RichEdit_Text" + count + "'></Textarea>";

			theObj = eval("document.f1.Mdec_RichEdit_Text" + count);
			theObj.value = TempVar.substr(0, FormLimit);

			TempVar = TempVar.substr(FormLimit);

			count++;
		}
	}
	else
	{
		theObj = document.getElementById("add10");
		theObj.innerHTML = "<Textarea rows=1 cols=1 name='Mdec_RichEdit_Text10'></Textarea>";

		theObj = eval("document.f1.Mdec_RichEdit_Text10");
		theObj.value = TempVar;
	}
}

function dec_RichEdit_Html()
{
	var count = 20;
	var theObj;
	var FormLimit = 50000;

	var TempVar = new String;
	TempVar = document.f1.RichEdit_Html.value;

	if (TempVar.length > FormLimit)
	{
		while (TempVar.length > 0 && count < 30)
		{
			theObj = document.getElementById("add" + count);
			theObj.innerHTML = "<Textarea rows=1 cols=1 name='Mdec_RichEdit_Html" + count + "'></Textarea>";

			theObj = eval("document.f1.Mdec_RichEdit_Html" + count);
			theObj.value = TempVar.substr(0, FormLimit);

			TempVar = TempVar.substr(FormLimit);

			count++;
		}
	}
	else
	{
		theObj = document.getElementById("add20");
		theObj.innerHTML = "<Textarea rows=1 cols=1 name='Mdec_RichEdit_Html20'></Textarea>";

		theObj = eval("document.f1.Mdec_RichEdit_Html20");
		theObj.value = TempVar;
	}
}

function isinlist(name)
{
	var i = 0;
	for (i; i < document.f1.recdomains.length; i++)
	{
		if (document.f1.recdomains[i].value == name)
		{
			return true;
		}
	}
	
	return false;
}

function addin()
{
	var i = 0;
	for (i; i < document.f1.alldomains.length; i++)
	{
		if (document.f1.alldomains[i].selected == true)
		{
			if (isinlist(document.f1.alldomains[i].value) == false)
			{
				var oOption = document.createElement("OPTION");
				oOption.text = document.f1.alldomains[i].value;
				oOption.value = document.f1.alldomains[i].value;
<%
if isMSIE = true then
%>
				document.f1.recdomains.add(oOption);
<%
else
%>
				document.f1.recdomains.appendChild(oOption);
<%
end if
%>
			}
		}
	}
}

function delout()
{
	var i = 0;
	for (i; i < document.f1.recdomains.length; i++)
	{
		if (document.f1.recdomains[i].selected == true)
		{
			document.f1.recdomains.remove(i);
			i--;
		}
	}
}

function window_onload() {
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
	obj_textarea.style.height = document.getElementById("tdtt").height + "px";
	obj_textarea.style.width = document.getElementById("tdtt").offsetWidth + "px";

	if (window.screen.width < 900)
		obj_textarea.style.width = "570px";
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

	hide_ex();
}

exfunc_is_show = false;
function hide_ex()
{
	Stag = document.getElementById("ex_showstr");
	if (exfunc_is_show == true)
	{
		Stag.innerHTML = "隐藏扩展功能";
		document.getElementById("ex_function_div").style.display = "inline";
	}
	else
	{
		Stag.innerHTML = "展示扩展功能";
		document.getElementById("ex_function_div").style.display = "none";
	}

	exfunc_is_show = !exfunc_is_show;
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
	<input name="EasyMail_DomainList" type="hidden">
    <INPUT NAME="AddFromAttFileString" TYPE="hidden">
    <INPUT NAME="MailName" TYPE="hidden" Value="<%=MailName %>">
    <INPUT NAME="SendMode" TYPE="hidden" Value="domainslistmail">
    <INPUT NAME="EasyMail_From" TYPE="hidden" value="<%= Session("wem")%>" maxlength="64">
    <INPUT NAME="EasyMail_TimerSend" TYPE="hidden" maxlength="10">
  <table width="94%" border="0" bgColor="white" align="center" cellspacing="0">
	<tr> 
	<td colspan="2" height="30" align="right">
		<table width="100%" border="0" cellspacing="0">
		<tr><td align="left" width="40%">
&nbsp;&nbsp;&nbsp;<span class="wwm_ar_msg">提醒:无动作超时时间还剩<font color="#FF3333"><b><span id="minpt"><%=Session.TimeOut %></span></b></font>分钟</span>
		</td><td align="right" style="padding-right:20px;">
		<a class="wwm_btnDownload btn_blue" style="width: 60px" href="#" onclick="javascript:showSending()">发送</a>&nbsp;
<%
if isadmin() = false then
%>
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:parent.location.href='domainright.asp?<%=getGRSN() %>'">取消</a>
<%
else
%>
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:parent.location.href='right.asp?<%=getGRSN() %>'">取消</a>
<%
end if
%>
		</td></tr>
		</table>
	</td>
	</tr>
	<tr>
	<td colspan="2" align="center">
	<hr size="1" Color="#A5B6C8">
  <table align="center" border="0" width="90%" cellspacing="0">
    <tr>
      <td rowspan="2" width="45%"> 
        <div align="center"><font class="s">所辖域</font></div>
      </td>
      <td rowspan="2"> 
        <div align="center"> </div>
        <div align="center"> </div>
      </td>
      <td rowspan="2" width="45%"> 
        <div align="center"><font class="s">接收邮件域</font></div>
      </td>
    </tr>
    <tr></tr>
    <tr> 
      <td rowspan="2" width="45%"> 
        <div align="center">
          <select name="alldomains" size="5" class="drpdwn" style="width: 200px; background-color:#FFFFEE;" multiple LANGUAGE=javascript ondblclick="return addin()">
<%
i = 0

if isadmin() = false then
	allnum = dm.GetUserManagerDomainCount(Session("wem"))

	do while i < allnum
		domain = dm.GetUserManagerDomain(Session("wem"), i)

		response.write "<option value='" & server.htmlencode(domain) & "'>" & server.htmlencode(domain) & "</option>" & Chr(13)

		domain = NULL

		i = i + 1
	loop
else
	allnum = dm.GetCount()

	do while i < allnum
		domain = dm.GetDomain(i)

		response.write "<option value='" & server.htmlencode(domain) & "'>" & server.htmlencode(domain) & "</option>" & Chr(13)

		domain = NULL

		i = i + 1
	loop
end if
%>
          </select>
        </div>
		</td>
      <td width="10%"> 
        <div align="center"> 
          <input type="button" value=" ==&gt; " class="sbttn" LANGUAGE=javascript onclick="addin()">
        </div>
      </td>
      <td rowspan="2" width="45%"> 
        <div align="center"> 
		<select name="recdomains" size="5" class="drpdwn" style="width: 200px; background-color:#FFFFEE;" multiple LANGUAGE=javascript ondblclick="return delout()">
		</select>
        </div>
      </td>
    </tr>
    <tr> 
      <td width="10%"> 
        <div align="center"> 
          <input type="button" value=" &lt;== " class="sbttn" LANGUAGE=javascript onclick="delout()">
        </div>
      </td>
    </tr>
  </table>
<hr size="1" Color="#A5B6C8">
	</td>
	</tr>
	<tr> 
	<td colspan="2" align="left"> 
<table cellspacing=0 cellpadding=2>
  <tbody>
  <tr> 
    <td noWrap align="right" width="78" style="padding-top:6px;">回复地址：</td>
    <td align="left">
<input name="EasyMail_BackAddress" type="text" size="60" value="<%=userweb.ReMail %>" class='tn_textbox'>
	</td></tr>
	<tr>
    <td noWrap align="right" width="78" style="padding-top:6px;">主题：</td>
    <td align="left">
<input name="EasyMail_Subject" type="text" size="60" maxlength="512" class='tn_textbox'>
	</td></tr>
  </tbody>
</table>
      </td>
    </tr>
    <tr>
	<td colspan="2" align="left">
<%
if useRichEditer = false then
%>
<table width="98%"><tr><td height="315" width="100%" name="tdtt" id="tdtt">
	<textarea name="EasyMail_Text" id="EasyMail_Text" cols="76" rows="12" class='textarea_wwm'></textarea>
</td></tr></table>
<%
else
%>
<table width="100%"><tr><td>
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
	</td></tr>
	<tr>
	<td height="32" width="17%" align="left" noWrap>
&nbsp;[<a href="javascript:hide_ex()"><span id="ex_showstr"></span></a>]
	</td>
	<td align="right" style="padding-right:20px;">
		<a class="wwm_btnDownload btn_blue" style="width: 60px" href="#" onclick="javascript:showSending()">发送</a>&nbsp;
<%
if isadmin() = false then
%>
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:parent.location.href='domainright.asp?<%=getGRSN() %>'">取消</a>
<%
else
%>
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:parent.location.href='right.asp?<%=getGRSN() %>'">取消</a>
<%
end if
%>
      </td>
    </tr>
    <tr>
    <td noWrap colspan="2">
<div id="ex_function_div">
	<table width="100%" border="0" cellspacing="0">
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

	Response.Write "<option value='" & idname & "'>" & server.htmlencode(subject) & "</option>"

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
<%
if isadmin() = true then
%>
	<tr><td noWrap align="right" class="tr_st">系统邮件：</td>
	<td class="tr_st" align="left">
	<input type="checkbox" name="EasyMail_SystemMail">是否是系统邮件
	</td></tr>
<%
end if
%>
    <tr><td noWrap align="right" class="tr_st">垃圾邮件投诉排除：</td>
	<td class="tr_st" align="left">
	<input type="checkbox" name="needAddInDebarList" checked>此邮件不可被投诉为垃圾邮件
	</td></tr>
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
</BODY>
</HTML>

<%
set dm = nothing
set userweb = nothing
set ads = nothing
%>
