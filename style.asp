<%
Response.CacheControl = "no-cache" 
%>

<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" -->

<%
dim ei
set ei = server.createobject("easymail.UserWeb")
ei.Load Session("wem")

if trim(request("mbsetname")) <> "" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	ei.MailName = trim(request("mbsetname"))
	ei.save
	set ei = nothing
	Response.End
end if

dim mam
set mam = server.createobject("easymail.AdminManager")
mam.Load

dim ischangeicon
dim lang
ischangeicon = FALSE

if trim(request("maxlist")) <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	MailName = trim(request("MailName"))
	MailName = replace(MailName, """", "'")
	ei.MailName = MailName

	ReMail = trim(request("ReMail"))
	ReMail = replace(ReMail, """", "'")
	ei.ReMail = ReMail

	BccMail = trim(request("BccMail"))
	BccMail = replace(BccMail, """", "'")
	ei.BccMail = BccMail

	if trim(request("EnableBackupAllSendMail")) = "" then
		ei.EnableBackupAllSendMail = false
	else
		ei.EnableBackupAllSendMail = true
	end if

	if trim(request("EnableClearWhenFull")) = "" then
		ei.EnableClearWhenFull = false
	else
		ei.EnableClearWhenFull = true
	end if

	if trim(request("EnableClearSendBox")) = "" then
		ei.EnableClearSendBox = false
	else
		ei.EnableClearSendBox = true
	end if

	if trim(request("enableAutoAdaptCharSet")) = "" then
		ei.enableAutoAdaptCharSet = false
	else
		ei.enableAutoAdaptCharSet = true
	end if

	ei.CharSet = trim(request("CharSet"))

	if trim(request("enableRichEditer")) = "" then
		ei.useRichEditer = false
	else
		ei.useRichEditer = true
	end if

	if trim(request("EnableShowHtmlMail")) = "" then
		ei.EnableShowHtmlMail = false
	else
		ei.EnableShowHtmlMail = true
	end if

	if trim(request("Enable_AutoReply_OneDay")) = "" then
		ei.Enable_AutoReply_OneDay = false
	else
		ei.Enable_AutoReply_OneDay = true
	end if

	if trim(request("EnableShowDateECMailList")) = "" then
		ei.EnableShowDateECMailList = false
	else
		ei.EnableShowDateECMailList = true
	end if

	if trim(request("EnableSession")) = "" then
		ei.EnableSession = false
		Session("EnableSession") = false
	else
		ei.EnableSession = true
		Session("EnableSession") = true
	end if

	if trim(request("defaultSign")) <> "" and IsNumeric(trim(request("defaultSign"))) = true then
		ei.defaultSign = CInt(trim(request("defaultSign")))
	end if

	if trim(request("ShowLanguage")) <> "" and IsNumeric(trim(request("ShowLanguage"))) = true then
		if ei.ShowLanguage <> CInt(trim(request("ShowLanguage"))) then
			lang = b_lang_194
		end if

		ei.ShowLanguage = CInt(trim(request("ShowLanguage")))
	end if

	if trim(request("enableAutoClear")) = "" then
		ei.useAutoClearTrashBox = false
	else
		ei.useAutoClearTrashBox = true
	end if

	if trim(request("autoClearDays")) <> "" and IsNumeric(trim(request("autoClearDays"))) = true then
		ei.autoClearTrashBoxDays = CInt(trim(request("autoClearDays")))
	else
		ei.autoClearTrashBoxDays = 15
	end if

	if trim(request("maxlist")) <> "" then
		ei.pageLines = CInt(trim(request("maxlist")))
	else
		ei.pageLines = 10
	end if

	Session("pl") = ei.pageLines


	if trim(request("addo")) = 1 then
		ei.orMailForReply = true
	else
		ei.orMailForReply = false
	end if

	Session("addomail") = ei.orMailForReply


	if trim(request("replyf")) <> "" then
		ei.addInSubjectForReply = CInt(trim(request("replyf")))

		if ei.addInSubjectForReply = 0 then
			Session("addsubject") = "> "
		elseif ei.addInSubjectForReply = 1 then
			Session("addsubject") = "Re: "
		elseif ei.addInSubjectForReply = 2 then
			Session("addsubject") = b_lang_195
		end if
	end if

	if trim(request("delProc")) <> "" then
		ei.delProc = CInt(trim(request("delProc")))
	else
		ei.delProc = 0
	end if

	Session("delProc") = ei.delProc

	ei.save

	set ei = nothing
	set mam = nothing


	if ischangeicon = false and lang = "" then
		if err.number = 0 then
			response.redirect "ok.asp?" & getGRSN() & "&gourl=style.asp"
		else
			response.redirect "err.asp?" & getGRSN() & "&gourl=style.asp"
		end if
	else
		Response.Write "<SCRIPT LANGUAGE=javascript>"

		if lang <> "" then
			response.write "parent.parent.window.location.href = """ & lang & "welcome.asp?" & getGRSN() & """;"
		else
			if ischangeicon = TRUE then
				Response.Write "parent.parent.window.location.href = ""welcome.asp?" & getGRSN() & """;"
			end if
		end if

		Response.Write "</script>"
		Response.End
	end if

end if
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<STYLE type=text/css>
<!--
html {overflow: scroll; overflow-x: hidden; overflow-y: auto !important;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.td_line_l {text-align:right; white-space:nowrap; background-color:#EFF7FF; border-bottom:1px #A5B6C8 solid; height:30px; color:#303030;}
.td_line_r {text-align:left; background-color:white; border-bottom:1px #A5B6C8 solid; height:30px; padding-left:6px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function wximgerr() {
	document.getElementById("wx").style.display = "none";
}

function window_onload() {
	var temp_charset = "<%=ei.CharSet %>";
	document.fm1.charSet.value = temp_charset.toLowerCase();
	EnableClearWhenFull_onclick();
}

function EnableClearWhenFull_onclick() {
	if (document.fm1.EnableClearWhenFull.checked == true)
		document.fm1.EnableClearSendBox.disabled = false;
	else
		document.fm1.EnableClearSendBox.disabled = true;
}

function gosub() {
	if (document.fm1.ReMail.value != "" && remailisok() == false)
	{
		alert("<%=b_lang_196 %>");
		document.fm1.ReMail.focus();
		return ;
	}

	document.fm1.submit();
}

function remailisok() {
	var mailisok = true;
	var sp = document.fm1.ReMail.value.indexOf("@");
	if (sp == -1)
		mailisok = false;
	else
	{
		sp = document.fm1.ReMail.value.indexOf("@", sp + 1);
		if (sp != -1)
			mailisok = false;
		else
		{
			if (document.fm1.ReMail.value.charAt(0) == '@' || document.fm1.ReMail.value.charAt(document.fm1.ReMail.value.length - 1) == '@')
			{
				mailisok = false;
			}
		}
	}

	return mailisok;
}
//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload()">
<form method="post" action="style.asp" name="fm1">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_197 %>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"><%
dim wx
set wx = server.createobject("easymail.WXSet")
wx.Load
wx_is_enabled = wx.is_enabled
set wx = nothing

if wx_is_enabled = true then
%>
<div id="wx" style="text-align:left; padding-left:16px;">
<table width="60%" border="0" align="left" cellspacing="0">
<tr>
<td width="5%" style="padding:6px 2px 6px 2px;"><img src="wx.jpg" align='absmiddle' border="0" onerror="wximgerr();"></td>
<td nowrap align="left" style="padding-left:20px;"><%=b_lang_392 %></td>
</tr>
</table>
</div>
<%
end if
%></td></tr>
</table>

<table width="88%" border="0" align="center" cellspacing="0">
	<tr>
	<td width="46%" valign=center class="td_line_l" style="border-top:1px #A5B6C8 solid;"><%=b_lang_198 %><%=s_lang_mh %></td>
	<td class="td_line_r" style="border-top:1px #A5B6C8 solid;"><input type="text" name="MailName" class='n_textbox' value="<%=ei.MailName %>" maxlength="256"></td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_199 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="checkbox" name="EnableShowHtmlMail" value="checkbox" <% if ei.EnableShowHtmlMail = true then response.write "checked"%>></td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_200 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="checkbox" name="enableRichEditer" value="checkbox" <% if ei.useRichEditer = true then response.write "checked"%>></td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_201 %><%=s_lang_mh %></td>
	<td class="td_line_r">
<select name="ShowLanguage" class="drpdwn">
<%
i = ei.ShowLanguage

if i = 0 then
	Response.Write "<option value='0' selected>Simplified Chinese</option>"
else
	Response.Write "<option value='0'>Simplified Chinese</option>"
end if

if i = 1 then
	Response.Write "<option value='1' selected>English</option>"
else
	Response.Write "<option value='1'>English</option>"
end if
%>
</select>
	</td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_202 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="text" name="ReMail" class='n_textbox' value="<%=ei.ReMail %>" maxlength="256"></td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=s_lang_0119 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="checkbox" name="EnableBackupAllSendMail" id="EnableBackupAllSendMail" value="checkbox" <% if ei.EnableBackupAllSendMail = true then response.write "checked"%>></td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=s_lang_0120 %><%=s_lang_mh %></td>
	<td class="td_line_r">
		<input type="checkbox" name="EnableClearWhenFull" value="checkbox" <% if ei.EnableClearWhenFull = true then response.write "checked"%> LANGUAGE=javascript onclick="return EnableClearWhenFull_onclick()">
		<%=s_lang_0121 %><br>
		<input type="checkbox" name="EnableClearSendBox" value="checkbox" <% if ei.EnableClearSendBox = true then response.write "checked"%>>
		<%=s_lang_0122 %>
	</td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_203 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="checkbox" name="EnableShowDateECMailList" value="checkbox" <% if ei.EnableShowDateECMailList = true then response.write "checked"%>></td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_204 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="checkbox" name="EnableSession" value="checkbox" <% if ei.EnableSession = true then response.write "checked"%>></td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_367 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="checkbox" name="Enable_AutoReply_OneDay" value="checkbox" <% if ei.Enable_AutoReply_OneDay = true then response.write "checked"%>></td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_205 %><%=s_lang_mh %></td>
	<td class="td_line_r">
<select name="defaultSign" class="drpdwn">
<%
ds = ei.defaultSign
if ds = -1 then
%>
	<option value="-1" selected><%=b_lang_206 %></option>
<%
else
%>
	<option value="-1"><%=b_lang_206 %></option>
<%
end if

dim sm
set sm = server.createobject("easymail.SignManager")
sm.Load Session("wem")

allnum = sm.count
i = 0

do while i < allnum
	sm.get i, s_title, s_text, s_htmltext

if i <> ds then
	Response.Write "<option value='" & i & "'>" & server.htmlencode(s_title) & "</option>"
else
	Response.Write "<option value='" & i & "' selected>" & server.htmlencode(s_title) & "</option>"
end if

	s_title = NULL
	s_text = NULL
	s_htmltext = NULL

	i = i + 1
loop

set sm = nothing
%>
</select>
	</td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_207 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="checkbox" name="enableAutoClear" value="checkbox" <% if ei.useAutoClearTrashBox = true then response.write "checked"%>></td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_208 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="text" name="autoClearDays" class='n_textbox' value="<%=ei.autoClearTrashBoxDays %>" size="4" maxlength="4"> <%=b_lang_230 %></td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_209 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="text" name="BccMail" class='n_textbox' value="<%=ei.BccMail %>" maxlength="64"></td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_210 %><%=s_lang_mh %></td>
	<td class="td_line_r">
		<input type=radio <% if ei.pageLines = 10 then response.write "checked"%> value="10" name=maxlist> 10<br>
		<input type=radio <% if ei.pageLines = 20 then response.write "checked"%> value="20" name=maxlist> 20<br>
		<input type=radio <% if ei.pageLines = 50 then response.write "checked"%> value="50" name=maxlist> 50<br>
		<input type=radio <% if ei.pageLines = 100 then response.write "checked"%> value="100" name=maxlist> 100<br>
		<input type=radio <% if ei.pageLines = 200 then response.write "checked"%> value="200" name=maxlist> 200<br>
		<input type=radio <% if ei.pageLines = 500 then response.write "checked"%> value="500" name=maxlist> 500
	</td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_211 %><%=s_lang_mh %></td>
	<td class="td_line_r">
		<input type=radio <% if ei.orMailForReply = true then response.write "checked"%> value="1" name=addo><%=b_lang_212 %><br>
		<input type=radio <% if ei.orMailForReply = false then response.write "checked"%> value="0" name=addo><%=b_lang_213 %>
	</td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_214 %><%=s_lang_mh %></td>
	<td class="td_line_r">
		<input type=radio <% if ei.addInSubjectForReply = 0 then response.write "checked"%> value=0 name=replyf>&gt;<br>
		<input type=radio <% if ei.addInSubjectForReply = 1 then response.write "checked"%> value=1 name=replyf>Re:<br>
		<input type=radio <% if ei.addInSubjectForReply = 2 then response.write "checked"%> value=2 name=replyf><%=b_lang_195 %>
	</td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_215 %><%=s_lang_mh %></td>
	<td class="td_line_r">
		<input type=radio <% if ei.delProc = 0 then response.write "checked"%> value="0" name="delproc"><%=b_lang_216 %><br>
		<input type=radio <% if ei.delProc = 1 then response.write "checked"%> value="1" name="delproc"><%=b_lang_217 %>
	</td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_218 %><%=s_lang_mh %></td>
	<td class="td_line_r">
		<select name="charSet" class="drpdwn">
		<option value="gb2312"><%=b_lang_219 %></option>
		<option value="big5"><%=b_lang_220 %></option>
		<option value="iso-8859-1"><%=b_lang_221 %></option>
		<option value="euc-jp"><%=b_lang_222 %></option>
		<option value="shift-jis"><%=b_lang_223 %></option>
		<option value="iso-2022-jp"><%=b_lang_224 %></option>
		<option value="euc-kr"><%=b_lang_225 %></option>
		<option value="iso-2022-kr"><%=b_lang_226 %></option>
		<option value="chn-utf-8">Unicode(UTF-8)</option>
		<option value="gb2312_iso-2022-cn"><%=b_lang_227 %></option>
		</select>
	</td>
	</tr>

	<tr>
	<td valign=center class="td_line_l"><%=b_lang_228 %><%=s_lang_mh %></td>
	<td class="td_line_r"><input type="checkbox" name="enableAutoAdaptCharSet" value="checkbox" <% if ei.enableAutoAdaptCharSet = true then response.write "checked"%>> <%=b_lang_229 %></td>
	</tr>
	<tr><td colspan="2" align="left" height="28"><br>
	<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
	<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
	</td></tr>
</table>
<br>
</Form>
</BODY>
</HTML>

<%
set ei = nothing
set mam = nothing
%>
