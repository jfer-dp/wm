<%
bf_isok = false
zck_isok = false

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if trim(request.Cookies("zatt_checkcode")) = trim(request("zck")) then
		zck_isok = true
	end if

	bf_url = Request.ServerVariables("HTTP_REFERER")
	bf_index = InStr(bf_url, "?")

	if bf_index > 15 then
		if LCase(Mid(bf_url, bf_index - 12, 12)) = "/downatt.asp" then 
			bf_isok = true
		end if
	end if
else
	bf_isok = true
	zck_isok = true
end if

k = trim(request("k"))

dim mt
set mt = server.createobject("easymail.WMethod")
mt.GetZattFileInfoWithKey k, ret_Val, ret_filename, ret_file_length, ret_lasttime, ret_username, ret_id

if bf_isok = true and zck_isok = true and Request.ServerVariables("REQUEST_METHOD") = "POST" and ret_Val = 0 and Len(ret_id) > 0 and Len(ret_username) > 0 then
	dim a
	set a = server.createobject("easymail.emmail")

	a.LoadAll ret_username, ret_id
	Response.ContentType = a.GetContentType(0)
	a.ShowAttachment 0, true

	mt.AddLog 1, Request.ServerVariables("REMOTE_ADDR"), ret_username, getShowSize(ret_file_length)

	set a = nothing
	set mt = nothing
	Response.End
end if
%>

<!DOCTYPE html>
<html>
<head>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<title>WinWebMail Server 邮件系统 - 附件下载</title>
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<style>
html, body, p, ul, li, h1, input, form{margin:0; padding:0;}
body {font:12px Verdana,Arial,Helvetica,sans-serif;}
body {background:white;}
.input_wwm {padding:2px 8px 0pt 3px;height:18px;border:1px solid #999;background-color:#FFFFEE;}
.wwm_headerContainer {height:28px;padding:21px 0 0;text-align:center;background-color:#d6d6d6;border-bottom: #a0a0a0 1px solid;}
.wwm_title {display:inline;float:center;font-size:16px;color:#606060;}
.wwm_container {margin-left:auto;margin-right:auto;width:500px;padding-bottom:100px;}
.wwm_headerTips {padding:10px 0;margin:30px 0 20px;border-bottom:1px solid #ccc;visibility:hidden;}
.wwm_main {position:relative;min-height:116px;margin-bottom:28px;clear:both;_height:116px;}
.wwm_fileIcon {float:left;margin-right:20px;}
.wwm_fileInfo {padding-top:10px;}
.wwm_filename {font-size:16px;font-weight:bold;line-height:20px;max-height:40px;word-wrap:break-word;overflow:hidden;}
.wwm_fileDetail {color:#a0a0a0;margin-top:6px;padding-bottom:26px;}
.wwm_footer {position:absolute;left:0;bottom:0;right:0;_width:100%;height:69px;text-align:center;border-top: #a0a0a0 1px solid;background-color:#dddddd;}
.wwm_copyright {margin-top:27px;color:#a0a0a0;}
.wwm_msg {padding:8px; margin:-6px 0 14px 0; color:#7E4F05; line-height:18px; background:#FFF3C3;border-radius:4px; -webkit-border-radius:4px;padding-left:20px;padding-right:20px;text-align:left;border: #7E4F05 1px solid;}
.wwm_span_vc {display:-moz-inline-box; display:inline-block; width:210px;}
</style>
</head>

<script LANGUAGE=javascript>
<!--
if (top.location !== self.location)
{
	top.location = self.location;
}

var two_click = false;
function sub()
{
	if (two_click == false)
	{
		if (document.getElementById("zck").value.length < 1)
		{
			document.getElementById("zck").focus();
			alert("验证码填写错误!");
			return ;
		}
		else
		{
			document.f1.submit();
			document.getElementById("bt_text").innerHTML = "返回首页";
			document.getElementById("spvc").style.visibility = "hidden";
		}
	}
	else
		window.location.href = bt_hp();

	two_click = true;
}

function bt_hp()
{
	var sindex = window.location.href.indexOf('?');
	if (sindex > -1)
	{
		var tmpstr = window.location.href.substr(0, sindex);
		sindex = tmpstr.lastIndexOf('/');

		if (sindex > -1)
		{
			tmpstr = tmpstr.substr(0, sindex + 1);
			return tmpstr;
		}
	}

	return "";
}
<%
if zck_isok = false and bf_isok = true and ret_Val = 0 then
%>
window.onload = function()
{
	document.getElementById("zck").focus();
	alert("验证码填写错误!");
}
<%
end if
%>
//-->
</script>

<body>
<form name="f1" method="post" action="downatt.asp?<%=getGRSN() %>">
<div class="container"><div class="wwm_headerContainer">
<h1 class="wwm_title"><a href="http://www.winwebmail.com" style="text-decoration: none; color:#606060;" target="_blank">WinWebMail Server 邮件系统</a> - 附件下载</h1>
</div></div>
<div class="wwm_container"><div class="wwm_headerTips" style="visibility: hidden;"></div>
<div class="wwm_main"><img class="wwm_fileIcon" src="images/zfile.gif">
<div class="wwm_fileInfo"><div class="wwm_filename"><%
if bf_isok = false then
	errstr = "系统忙碌，请稍后再试。"
else
	if ret_Val = 0 then
		Response.Write server.htmlencode(ret_filename)
	elseif ret_Val = 5 then
		errstr = "已超过有效期。"
	elseif ret_Val = 1 then
		errstr = "此功能已关闭。"
	else
		errstr = "提取文件出错。"
	end if
end if
%></div>
<%
if Len(errstr) < 1 then
%>
<div class="wwm_fileDetail"><span><%=server.htmlencode(getShowSize(ret_file_length)) %></span>&nbsp;&nbsp;&nbsp;&nbsp;<span><%=server.htmlencode(conv_show_date(ret_lasttime)) %> 到期</span></div>
<%
else
%>
<div class="wwm_fileDetail"><span class="wwm_msg"><%=server.htmlencode(errstr) %></span></div>
<%
end if
%>
<div>
<%
if Len(errstr) < 1 then
%>
<span id="spvc" class="wwm_span_vc">输入验证码：<img src="tu.asp" align="absmiddle" border="0"><input type="text" name="zck" id="zck" class="input_wwm" value="" size="2" maxlength="2"></span><a id="downbt" class="wwm_btnDownload btn_blue" href="#" onclick="javascript:sub();"><span id="bt_text">&nbsp;下&nbsp;载&nbsp;</span></a>
<%
else
%>
<span id="spvc" class="wwm_span_vc"></span><a class="wwm_btnDownload btn_blue" href="#" onclick="javascript:bt_hp();">返回首页</a>
<%
end if
%>
</div></div></div></div>
<div class="wwm_footer"><div class="wwm_copyright"><a href="http://www.winwebmail.com" style="text-decoration: none; color:#a0a0a0;" target="_blank">&copy; 2000 - Present by Ma Jian. All rights reserved.</a></div></div>
<input id="k" name="k" type="hidden" value="<%=k %>">
</form>
</body></html>

<%
set mt = nothing

ret_Val = NULL
ret_filename = NULL
ret_file_length = NULL
ret_lasttime = NULL
ret_username = NULL
ret_id = NULL

function getShowSize(bytesize)
	if bytesize < 1000 then
		getShowSize = "1K"
	else
		if bytesize < 1000000 then
			getShowSize = CLng(bytesize/1000) & "K"
		else
			tmpSize = CStr(CDbl(bytesize/1000000))
			tmpindex = InStr(1, tmpSize, ".")
			if tmpindex = 0 then
				getShowSize = tmpSize & "M"
			else
				getShowSize = Left(tmpSize, tmpindex + 2) & "M"
			end if
		end if
	end if
end function

function conv_show_date(datastr)
	sl_y = Left(datastr, 4)
	sl_m = Mid(datastr, 5, 2)
	sl_d = Mid(datastr, 7, 2)
	sl_h = Mid(datastr, 9, 2)

	if Left(sl_m, 1) = "0" then
		sl_m = Right(sl_m, 1)
	end if

	if Left(sl_d, 1) = "0" then
		sl_d = Right(sl_d, 1)
	end if

	if Left(sl_h, 1) = "0" then
		sl_h = Right(sl_h, 1)
	end if

	conv_show_date = sl_y & "年" & sl_m & "月" & sl_d & "日 " & sl_h & "时"
end function

function getGRSN()
	dim theGRSN
	Randomize
	theGRSN = Int((9999999 * Rnd) + 1)

	getGRSN = "GRSN=" & CStr(theGRSN)
end function
%>
