<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.CodePage=65001%>  
<% Response.Charset="UTF-8" %> 

<%
isback = trim(request("isback"))
wid = replace(trim(request("wid")), " ", "+")
is_send_ok = false

dim wx
set wx = server.createobject("easymail.WXSet")
wx.GetInfo wid, wx_user, wx_filename
is_expires = wx.IsExpires(wx_filename, 3)
set wx = nothing

if is_expires = false and isback <> "1" and wid <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim mailsend
	set mailsend = server.createobject("easymail.MailSend")

	dim ei
	set ei = server.createobject("easymail.emmail")
	ei.LoadAll wx_user, wx_filename

	dim userweb
	set userweb = server.createobject("easymail.UserWeb")
	userweb.Load wx_user

	mailsend.createnew wx_user, Left(wx_filename, 18)
	mailsend.EM_To = ei.FromMail
	mailsend.CharSet = userweb.CharSet
	mailsend.MailName = userweb.MailName

	if userweb.addInSubjectForReply = 0 then
		mailsend.EM_Subject = "> " & ei.subject
	elseif userweb.addInSubjectForReply = 1 then
		mailsend.EM_Subject = "Re: " & ei.subject
	elseif userweb.addInSubjectForReply = 2 then
		mailsend.EM_Subject = "Reply: " & ei.subject
	end if

	mailsend.EM_Text = Request.Form("mailtext")
	mailsend.EM_HTML_Text = Request.Form("mailhtml")
	mailsend.useRichEditer = true

	is_send_ok = mailsend.Send

	if is_send_ok = true then
		mailsend.SetReply wx_filename
	end if

	set userweb = nothing
	set ei = nothing
	set mailsend = nothing
end if
%>

<!doctype html>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0,user-scalable=no">
<style>
html, body {
    height: 98%;
    width: 97%;
	margin: 8px 6px 0px 5px;;
}

body {
	overflow-x:hidden; 
	overflow-y:hidden; 
}

.test_box {
    width:100%; 
    height:97%; 
    margin:0px;
    padding:0px; 
    font-size: 14px; 
	padding: 8px 5px 0px 8px;
    word-wrap: break-word;
    overflow-x: hidden;
    overflow-y: auto;
    -webkit-user-modify: read-write-plaintext-only;
	outline: none;
	color: #333;
	background-color: #fff;
	background-repeat: no-repeat;
	background-position: right center;
	border: 1px solid #999;
	border-radius: 3px;
	box-shadow: inset 0 1px 2px rgba(0,0,0,0.075);
	-moz-box-sizing: border-box;
	box-sizing: border-box;
	transition: all 0.15s ease-in;
	-webkit-transition: all 0.15s ease-in 0;
	border-color: #51a7e8;
	box-shadow: inset 0 1px 2px rgba(0,0,0,.075), 0 0 5px rgba(81,167,232,.5);
}

.wrap {
	margin-bottom: 18px;
	padding: 8px;
	background: #fff;
	border-radius: 2px 2px 2px 2px;
	-webkit-box-shadow: 0 10px 6px -6px #777;
	-moz-box-shadow: 0 10px 6px -6px #777;
	box-shadow: 0 10px 6px -6px #777;
}
</style>

<script language="JavaScript">
function getText(strHtml)
{
	var cv_sh = strHtml.replace(/[\r\n]*/g, "");
	cv_sh = cv_sh.replace(/<[\s\/]*br(\s[^>]*>|[^>]*>)/gi, "\r\n");

	cv_sh = cv_sh.replace(/^<[\s]*p[\s]*>/i, "");
	cv_sh = cv_sh.replace(/<[\s]*\/+[\s]*p(\s[^>]*>|[^>]*>)/gi, "\r\n");

	cv_sh = cv_sh.replace(/<[\s]*li(\s[^>]*>|[^>]*>)/gi, "\r\n");
	cv_sh = cv_sh.replace(/<[\s\/]*blockquote(\s[^>]*>|[^>]*>)/gi, "\r\n");
	cv_sh = cv_sh.replace(/<[\s]*hr(\s[^>]*>|[^>]*>)/gi, "\r\n");
	cv_sh = cv_sh.replace(/<[\s]*\/+[\s]*tr(\s[^>]*>|[^>]*>)/gi, "\r\n");
	cv_sh = cv_sh.replace(/<[\s]*\/+[\s]*div(\s[^>]*>|[^>]*>)/gi, "\r\n");

	cv_sh = cv_sh.replace(/<[\s]*\/+[\s]*td(\s[^>]*>|[^>]*>)/gi, " ");

	cv_sh = cv_sh.replace(/<!--[\s\S]*?-->/g, "");

	cv_sh = cv_sh.replace(/<head>[\s\S]*?<\/head>/gi, "");
	cv_sh = cv_sh.replace(/<scrip[^>]*>[\s\S]*?<\/script>/gi, "");
	cv_sh = cv_sh.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "");

	cv_sh = cv_sh.replace(/<[^>]*>/g, "");

	cv_sh = cv_sh.replace(/&amp;/g, "&");
	cv_sh = cv_sh.replace(/&#38;/g, "&");

	cv_sh = cv_sh.replace(/&lt;/g, "<");
	cv_sh = cv_sh.replace(/&#60;/g, "<");

	cv_sh = cv_sh.replace(/&gt;/g, ">");
	cv_sh = cv_sh.replace(/&#62;/g, ">");

	cv_sh = cv_sh.replace(/&quot;/g, "\"");
	cv_sh = cv_sh.replace(/&nbsp;/g, " ");

	cv_sh = cv_sh.replace(/^[\r\n]{0,2}|[\r\n]{0,2}$/g, "");

	return cv_sh;
}

function window_onload() {
<%
if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	Response.Write "document.body.style.backgroundColor = ""#e6e6e6"";"

	if is_send_ok = true then
		Response.Write "parent.document.getElementById(""pageframe"").rows=""0,*,0"";"
	else
		Response.Write "parent.f3.show_back();"
	end if
end if
%>
	parent.f3.sendend();
}

function send() {
	if (document.getElementById("divtext").innerHTML != "<br>")
	{
		document.getElementById("mailhtml").value = document.getElementById("divtext").innerHTML;
		document.getElementById("mailtext").value = getText(document.getElementById("divtext").innerHTML);
		document.f1.submit();
	}
}

function back2write() {
	document.f1.submit();
}

javascript:window.history.forward(1); 
</script>
</head>

<%
if isback = "1" or Request.ServerVariables("REQUEST_METHOD") <> "POST" then
%>
<body>
<div id="divtext" class="test_box" contenteditable="true"><%
if isback <> "1" then
	Response.Write "<br>"
else
	Response.Write Request.Form("mailhtml")
end if
%></div>
<form action="#" method=post name="f1" style="display:none;">
<input id="wid" name="wid" type="hidden" value="<%=wid %>">
<textarea id="mailtext" name="mailtext" cols="0" rows="0" type="hidden"></textarea>
<textarea id="mailhtml" name="mailhtml" cols="0" rows="0" type="hidden"></textarea>
</form>
<%
else
%>
<body language=javascript onload="return window_onload()">
<form action="#" method=post name="f1" style="display:none;">
<input name="isback" type="hidden" value="1">
<input id="wid" name="wid" type="hidden" value="<%=wid %>">
<textarea id="mailhtml" name="mailhtml" cols="0" rows="0" type="hidden"><%=Request.Form("mailhtml") %></textarea>
</form>
<div class="wrap">
<div style="padding-top:40px; min-height:100px; text-align:center;">
<font style="color:#5fa207; font-weight:bold; font-size:14px;"><%
if is_send_ok = false then
	Response.Write "邮件发送失败"
else
	Response.Write "邮件发送成功"
end if
%></font><%
if is_send_ok = false then
	Response.Write "<br><font style=""color:#999999; font-size:12px;"">可能原因：链接过期或原邮件不存在</font>"
end if
%></div>
</div>
<%
end if
%>
</body>
</html>
