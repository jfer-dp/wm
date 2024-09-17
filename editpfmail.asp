<!--#include file="passinc.asp" --> 

<%
if isadmin() = false then
	dim pfvl
	set pfvl = server.createobject("easymail.PubFolderViewLimit")
	pfvl.Load trim(request("iniid"))

	if pfvl.IsShow(Session("mail")) = false then
		set pfvl = nothing
		Response.Redirect "noadmin.asp"
	end if

	set pfvl = nothing
end if


isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

if Application("em_EnableBBS") = false then
	response.redirect "noadmin.asp?errstr=此公共文件夹已关闭&" & getGRSN()
end if

if IsNumeric(trim(request("face"))) = true then
	face = CInt(trim(request("face")))
else
	face = 0
end if

gourl = trim(request("gourl"))
searchkey = trim(request("searchkey"))

dim ei
set ei = server.createobject("easymail.emmail")
ei.IsInPublicFolder = true
'-----------------------------------------

filename = trim(request("filename"))

ei.LoadAll Session("wem"), filename


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


function showSave() {
	document.f1.SendMode.value = "save"

	sub(0);
}

function showSending() {
	document.f1.SendMode.value = "editpost";
	sub(1);
}

function timerSending() {
	if (checkTime() == true)
	{
		document.f1.SendMode.value = "timersend"
		sub(0);
	}
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

<%
if useRichEditer = false then
%>
		if (smode == 1)
			document.getElementById("sending").style.display = "inline";
		else
			document.getElementById("esave").style.display = "inline";

		dec_EasyMail_Text();
<%
end if
%>

		addOAtt();

<%
if useRichEditer = true then
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

	document.f1.submit();
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
}

function delOAtt(){
	if (document.f1.OrAtt.selectedIndex != -1)
	{
		document.f1.OrAtt.remove(document.f1.OrAtt.selectedIndex);
		document.f1.OrAtt.selectedIndex = 0;
	}
}

function addOAtt(){
	var i = 0;
	var al = "";

	for (i; i < document.f1.OrAtt.length; i++)
	{
		al = al + document.f1.OrAtt[i].value + '\t';
	}

	document.f1.EasyMail_OrAtt.value = al;
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
//-->
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
   <FORM ACTION="sendmail.asp?<%=getGRSN() %>" METHOD=POST NAME="f1" target="_parent">
    <INPUT NAME="EasyMail_CharSet" TYPE="hidden" Value="<%=userweb.CharSet %>">
    <INPUT NAME="pid" TYPE="hidden" Value="<%=trim(request("pid")) %>">
    <INPUT NAME="iniid" TYPE="hidden" Value="<%=trim(request("iniid")) %>">
	<input name="EasyMail_To" type="hidden" value="post">
    <INPUT NAME="gourl" TYPE="hidden" Value="<%=gourl %>">
    <INPUT NAME="oname" TYPE="hidden" Value="<%=filename %>">
    <INPUT NAME="SendMode" TYPE="hidden" Value="editpost">
    <INPUT NAME="EasyMail_OrMailName" TYPE="hidden" value="<%= filename%>">
    <INPUT NAME="EasyMail_R_F_MailName" TYPE="hidden" value="<%= filename%>">
    <INPUT NAME="EasyMail_OrAtt" TYPE="hidden">
    <INPUT NAME="useRichEditer" TYPE="hidden" Value="<%
if useRichEditer = true then
	Response.Write "true"
else
	Response.Write "false"
end if
%>">
    <INPUT NAME="AddFromAttFileString" TYPE="hidden">
    <INPUT NAME="MailName" TYPE="hidden" Value="<%=MailName %>">
    <INPUT NAME="EasyMail_From" TYPE="hidden" value="<%= Session("wem")%>" maxlength="64">
    <INPUT NAME="EasyMail_TimerSend" TYPE="hidden" maxlength="10">
  <table width="94%" border="0" bgColor="white" align="center" cellspacing="0" style="margin-top:4px;">
    <tr>
		<td colspan="2" height="40" align="right" style="padding-right:20px;">
		<a class="wwm_btnDownload btn_blue" style="width: 60px" href="#" onclick="javascript:showSending()">发送</a>&nbsp;
<% if gourl = "" then %>
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:history.back();">取消</a>
<% else %>
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:parent.location.href='<%=gourl %>?<%=getGRSN() %>'">取消</a>
<% end if %>
		</td>
	</tr>
  <tr>
	<td colspan="2" height="30" noWrap align="left" style='border-top:1px #A5B6C8 solid; padding-top:6px;'>&nbsp;&nbsp;主&nbsp;&nbsp;题：
<input name="EasyMail_Subject" type="text" size="60" maxlength="512" class='tn_textbox' value="<%=server.htmlencode(ei.subject) %>">
	</td>
	</tr>
  <tr>
    <td colspan="2" height="24" noWrap align="left">&nbsp;&nbsp;关键字：
<input name="searchkey" type="text" size="60" maxlength="64" class='tn_textbox' value="<%=searchkey %>">
      </td>
	</tr>
	<tr>
	<td colspan="2" align="left">
<%
if useRichEditer = false then
%>
<table width="98%"><tr><td height="315" width="100%" name="tdtt" id="tdtt">
	<textarea name="EasyMail_Text" id="EasyMail_Text" cols="76" rows="12" class='textarea_wwm'><%
	Response.Write RemoveEndRN(ei.Text)
%>
</textarea>
</td></tr></table>
<%
else
%>
<table width="100%"><tr><td>
<script language="JavaScript" type="text/javascript">
<!--
<%
	if isMSIE = true then
		rtf_height = "270"
	else
		rtf_height = "253"
	end if

	if ei.HTML_Text <> "" then
		Response.Write "writeRichText('richedit', RemoveScript('" & RTESafe(ei.HTML_Text) & "'), ""100%"", " & rtf_height & ", true, false);"
	else
		if LCase(ei.ContentType) = "text/html" then
			Response.Write "writeRichText('richedit', RemoveScript('" & RTESafe(ei.Text) & "'), ""100%"", " & rtf_height & ", true, false);"
		else
			html_text = replace(RemoveEndRN(ei.Text), "'", "&#39;")
			html_text = replace(html_text, "<", "&lt;")
			html_text = replace(html_text, ">", "&gt;")
			html_text = replace(html_text, Chr(13) & Chr(10), "<br>")
			html_text = replace(html_text, Chr(10) & Chr(13), "<br>")
			html_text = replace(html_text, Chr(13), "<br>")
			html_text = replace(html_text, Chr(10), "<br>")
			html_text = replace(html_text, "\", "\\")
			Response.Write "writeRichText('richedit', '" & html_text & "', ""100%"", " & rtf_height & ", true, false);"
		end if
	end if
%>
//-->
</script>
</td></tr></table>
<%
end if
%>
      </td>
    </tr>
	<tr>
		<td colspan="2" align="left" style="padding-top:2px; padding-bottom:10px;">
		<table width="100%" border="0" cellspacing="0">
		<tr><td align="left" width="40%">
&nbsp;&nbsp;<span align="left" class="wwm_ar_msg">提醒:无动作超时时间还剩<font color="#FF3333"><b><span id="minpt"><%=Session.TimeOut %></span></b></font>分钟</span>
		</td><td align="right" style="padding-right:20px;">
		<a class="wwm_btnDownload btn_blue" style="width: 60px" href="#" onclick="javascript:showSending()">发送</a>&nbsp;
<% if gourl = "" then %>
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:history.back();">取消</a>
<% else %>
		<a class="wwm_btnDownload btn_blue" style="width: 60px;" href="#" onclick="javascript:parent.location.href='<%=gourl %>?<%=getGRSN() %>'">取消</a>
<% end if %>
		</td></tr>
		</table>
		</td>
	</tr>
	<tr>
	<td colspan="2" align="left" class="tr_st" style='border-top:1px #A5B6C8 solid; height:75px;'>&nbsp;<%
i = 1

do while i < 31
	if i <> face then
		Response.Write "<img src='images\face\" & i & ".gif' align='absmiddle' border='0'><input type='radio' value='" & i & "' name='face'>"
	else
		Response.Write "<img src='images\face\" & i & ".gif' align='absmiddle' border='0'><input type='radio' checked value='" & i & "' name='face'>"
	end if

	if i = 10 or i = 20 then
		Response.Write "<br>&nbsp;"
	elseif i < 30 then
		Response.Write "&nbsp;&nbsp;" & Chr(13)
	end if

	i = i + 1
loop
%>
	</td></tr>
    <tr>
	<td width="20%" noWrap align="right" class="tr_st">由网络存储加入的附件：
	</td>
	<td align="left" class="tr_st">
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
	<tr><td noWrap align="right" class="tr_st">签名：</td>
	<td align="left" class="tr_st">
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
	<tr><td noWrap align="right" class="tr_st">邮件等级：</td>
	<td align="left" class="tr_st">
	<select name="EasyMail_Priority" class=drpdwn>
	<option value="Normal">普通邮件</option>
	<option value="Low">慢件</option>
	<option value="High">紧急邮件</option>
	</select>
	</td></tr>
	<tr><td noWrap align="right" class="tr_st">原件中的附件：</td>
	<td align="left" class="tr_st">
	<select name="OrAtt" class=drpdwn>
<%
if ei.IsHtmlMail = true then
	i = 1
else
	i = 0
end if

allnum = ei.AttachmentCount

do while i < allnum
	Response.Write "<option value='" & i & "'>" & ei.GetAttachmentName(i) & "</option>"
	i = i + 1
loop
%>
	</select> 
	<input type="button" value="删 除" onclick="javascript:delOAtt()" class=sbttn>
	</td></tr>
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
set userweb = nothing
set ei = nothing


function RTESafe(strText)
	dim tmpString
	tmpString = replace(strText, "'", "&#39;")
	tmpString = replace(tmpString, Chr(10), "")
	tmpString = replace(tmpString, Chr(13), " ")
	tmpString = replace(tmpString, "&lt;", "&#11;")
	tmpString = replace(tmpString, "<", "&lt;")
	RTESafe = replace(tmpString, "\", "\\")
end function

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
%>
