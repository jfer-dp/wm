<%
wid = replace(trim(request("wid")), " ", "+")

dim wx
set wx = server.createobject("easymail.WXSet")
wx.GetInfo wid, wx_user, wx_filename

is_load_ok = false
is_expires = wx.IsExpires(wx_filename, 3)

dim ei
set ei = server.createobject("easymail.emmail")

if is_expires = false then
	filename = wx_filename
	ei.LoadAll wx_user, filename

	if Len(ei.HeadMessage) > 1 then
		is_load_ok = true
	end if
end if

dim allnum
allnum = 0
%>

<!doctype html>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<meta name="viewport" content="width=device-width,initial-scale=1.0,user-scalable=no">
<style>
html, body {
	margin: 8px 6px 3px 5px;;
}

body {
	color: #4A4A4A ;
	background: #e6e6e6;
}

A.cha:link {COLOR:#a0a0a0;}
A.cha:hover {COLOR:#a0a0a0;}
A.cha:visited {COLOR:#a0a0a0;}
A.cha:active {COLOR: #a0a0a0;}

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
function window_onload() {
<%
if is_load_ok = false then
	Response.Write "parent.document.getElementById(""pageframe"").rows=""*,0,0"";"
end if
%>
}

function iFrameHeight() {
	var ifm= document.getElementById("iframepage");
	var subWeb = document.frames ? document.frames["iframepage"].document : ifm.contentDocument;
	if(ifm != null && subWeb != null) {
		ifm.height = subWeb.body.scrollHeight;
	}
}
</script>
</head>

<%
if is_load_ok = true then
%>
<body>
<div class="wrap">
<div style="border-bottom:1px solid #909090; padding-bottom:3px;">
<%
if ei.FromName = ei.FromMail then
	receiver = "<font color='#5fa207' style='font-weight:bold;'>" & server.htmlencode(ei.FromMail) & "</font>"
else
	receiver = "<font color='#5fa207' style='font-weight:bold;'>" & server.htmlencode(ei.FromName) & "</font>" & server.htmlencode(" <" & ei.FromMail & ">")
end if
Response.Write "发件人：" & receiver & "<br>"
Response.Write "主题：" & server.htmlencode(ei.subject) & "<br>"
%>
</div>
<%
if ei.IsHtmlMail = true then
%>
<div style="padding-top:3px; min-height:100px;">
<iframe src="<%="wxshowatt.asp?ishtml=1&wid=" & wid & "&count=0&" & getGRSN() %>" id="iframepage" name="iframepage" frameBorder=0 scrolling=no width="100%" onLoad="iFrameHeight()"></iframe>
</div>
<%
else
%>
<div id="mailtext" style="padding-top:3px; padding-bottom:3px; min-height:100px;">
<%
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
end if
%>
</div>
<%
end if

is_show_att_str = false
allnum = ei.AttachmentCount

if allnum = 1 and ei.IsHtmlMail = true then
	allnum = 0
end if

i = 0
if allnum > 0 then

if ei.IsHtmlMail = true then
	i = 1
	allnum = ei.AllAttachmentCount
	show_i = 1

	do while i < allnum
		if is_show_att_str = false then
			Response.Write "<div style=""border-top:1px solid #909090; padding-top:3px; padding-bottom:3px; font-size:12px; background-color:#f9f9f9"">附件：<br>"
			is_show_att_str = true
		end if

		if ei.AttachmentCanShow(i) = true then
			if ei.GetAttachmentName(show_i) = "" then
			    Response.Write "<a href=""wxshowatt.asp?wid=" & wid & "&count=" & i & "&" & getGRSN() & """ target='_blank'>" & "html" & "</a>"
			else
				if ei.AttachmentIsMessage(show_i) = false then
			    	Response.Write "<a href=""wxshowatt.asp?wid=" & wid & "&count=" & i & "&" & getGRSN() & """ target='_blank'>" & server.htmlencode(ei.GetAttachmentName(show_i)) & "</a>"
				else
			    	Response.Write "<a href=""default.asp?" & getGRSN() & """ target='_blank'>" & server.htmlencode(ei.GetAttachmentName(show_i)) & "</a>"
				end if
			end if
			Response.Write "<br>" & Chr(13)

			show_i = show_i + 1
		end if

	    i = i + 1
	loop
else
	do while i < allnum
		if is_show_att_str = false then
			Response.Write "<div style=""border-top:1px solid #909090; padding-top:3px; padding-bottom:3px; font-size:12px; background-color:#f9f9f9"">附件：<br>"
			is_show_att_str = true
		end if

		if ei.GetAttachmentName(i) = "" then
		    Response.Write "<a href=""wxshowatt.asp?wid=" & wid & "&count=" & i & "&" & getGRSN() & """ target='_blank'>" & "html" & "</a>"
		else
			if ei.AttachmentIsMessage(i) = false then
		    	Response.Write "<a href=""wxshowatt.asp?wid=" & wid & "&count=" & i & "&" & getGRSN() & """ target='_blank'>" & server.htmlencode(ei.GetAttachmentName(i)) & "</a>"
			else
		    	Response.Write "<a href=""default.asp?" & getGRSN() & """ target='_blank'>" & server.htmlencode(ei.GetAttachmentName(i)) & "</a>"
			end if
		end if
		Response.Write "<br>" & Chr(13)

	    i = i + 1
	loop
end if

if is_show_att_str = true then
	Response.Write "</div>"
end if
end if


wx.load

if is_show_att_str = true then
	Response.Write "<div style=""padding-top:3px; padding-bottom:3px; font-size:12px; text-align:center;""><a href=""http://" & wx.url_mail_default & """ class=""cha"">如果邮件有显示不完整的现象时，请使用完整版界面</a>"
else
	Response.Write "<div style=""border-top:1px solid #909090; padding-top:3px; padding-bottom:3px; font-size:12px; text-align:center;""><a href=""http://" & wx.url_mail_default & """ class=""cha"">如果邮件有显示不完整的现象时，请使用完整版界面</a>"
end if
%>
</div>
</div>
<%
else
%>
<body language=javascript onload="return window_onload()">
<div class="wrap">
<div style="padding-top:40px; min-height:100px; text-align:center;">
<font style="color:#5fa207; font-weight:bold; font-size:14px;">访问失效</font><br>
<font style="color:#999999; font-size:12px;">可能原因：链接过期或原邮件不存在</font>
</div>
</div>
<%
end if
%>
</body>
</html>

<%
set ei = nothing
set wx = nothing

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

function getGRSN()
	dim theGRSN
	Randomize
	theGRSN = Int((9999999 * Rnd) + 1)

	getGRSN = "GRSN=" & CStr(theGRSN)
end function
%>
