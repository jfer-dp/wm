<%
dim mam
set mam = server.createobject("easymail.AdminManager")
mam.LoadExt

if IsEmpty(Application("em_MaxMPOP3")) and IsEmpty(Application("em_MaxSigns")) then
	if mam.IsLoadOK = false then
		tmp_num = 0
		do while tmp_num < 30
			mam.LoadExt

			if mam.IsLoadOK = true then
				Exit Do
			end if

			mam.Sleep 500
			tmp_num = tmp_num + 1
		loop
	end if

	if mam.IsLoadOK = true then
		Application("em_MaxMPOP3") = mam.MaxMPOP3
		Application("em_MaxSigns") = mam.MaxSigns
		Application("em_SystemAdmin") = LCase(mam.SystemAdmin)
		Application("em_EnableBBS") = mam.EnableBBS
		Application("em_Enable_SignHold") = mam.Enable_SignHold
		Application("em_Enable_FreeSign") = mam.Enable_FreeSign
		Application("em_Enable_SignWithDomainUser") = mam.Enable_SignWithDomainUser
		Application("em_Enable_SignNumberLimit") = mam.Enable_SignNumberLimit
		Application("em_SignNumberLimitDays") = mam.SignNumberLimitDays
		Application("em_Enable_ShareFolder") = mam.Enable_ShareFolder
		Application("em_Enable_SignEnglishName") = mam.Enable_SignEnglishName
		Application("em_LogPageKSize") = mam.LogPageKSize
		Application("em_TestAccounts") = LCase(mam.TestAccounts)
		Application("em_SignMode") = mam.SignMode
		Application("em_SignWaitDays") = mam.SignWaitDays
		Application("em_am_Name") = mam.am_Name
		Application("em_am_Accounts") = LCase(mam.am_Accounts)
		Application("em_AccountsAdmin") = LCase(mam.AccountsAdmin)
		Application("em_EnableEntAddress") = mam.Enable_Show_EntAddress
		Application("em_SpamAdmin") = LCase(mam.SpamAdmin)

		Application("em_EnableTrap") = mam.EnableTrap
		if mam.EnableTrap = true then
			Application("em_TrapMail") = mam.TrapMail
		end if
	else
		set mam = nothing
		Response.Redirect "default.asp?errstr=" & Server.URLEncode("超时, 请重试") & "&" & getGRSN()
	end if
end if

Enable_SignWithInputMoreInfo = mam.Enable_SignWithInputMoreInfo
AssumpsitString = mam.AssumpsitString
Sign_AccountMinLen = mam.Sign_AccountMinLen
Sign_PassWordMinLen = mam.Sign_PassWordMinLen
Sign_AccessMode = mam.Sign_AccessMode
Enable_Puny_DBCS_SignName = mam.Enable_Puny_DBCS_SignName

set mam = nothing


if Enable_SignWithInputMoreInfo = true then
	set mri = server.createobject("easymail.MoreRegInfo")
	mri.LoadSetting
	mri_count = mri.Count_Setting
	set mri = nothing

	if mri_count > 0 then
		if (Session("Reg") <> "step 1 over" and Session("Reg") <> "step 2 over") or Request.ServerVariables("REQUEST_METHOD") <> "POST" then
			Response.Redirect "reginfo.asp?" & getGRSN()
		end if
	else
		Enable_SignWithInputMoreInfo = false
	end if
end if


if Application("em_Enable_FreeSign") <> true then
	Response.Redirect "default.asp?errstr=" & Server.URLEncode("邮箱申请功能被禁止") & "&" & getGRSN()
	Response.End
end if

if Application("em_Enable_SignNumberLimit") = true and Request.Cookies("SignOk") = "1" then
	Response.Redirect "default.asp?errstr=" & Server.URLEncode("您已申请过邮箱") & "&" & getGRSN()
end if

if Application("em_Enable_SignNumberLimit") = false then
	Response.Cookies("SignOk") = ""
end if


dim webkill
set webkill = server.createobject("easymail.WebKill")
webkill.Load

rip = Request.ServerVariables("REMOTE_ADDR")

if webkill.IsKill(rip) = true then
	set webkill = nothing
	Response.Redirect "outerr.asp?errstr=" & Server.URLEncode("拒绝IP地址 " & rip & " 访问") & "&" & getGRSN()
end if

set webkill = nothing



dim errstr

if trim(request("errstr")) <> "" then
	errstr = trim(request("errstr"))
else
	errstr = "请输入您想申请的用户名以及密码信息, 并选择域名"
end if


username = LCase(trim(request("username")))
if Enable_Puny_DBCS_SignName = true and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim wmethod
	set wmethod = server.createobject("easymail.WMethod")

	if wmethod.isHaveDBCS(username) = true then
		username = wmethod.Str_To_Puny(username)
	end if

	set wmethod = nothing
end if

domain = LCase(trim(request("domain")))
pw = trim(request("pw"))
pw1 = trim(request("pw1"))
regemail = LCase(trim(request("regemail")))

if pw <> pw1 then
	errstr = "输入的密码不相同"
end if

if username <> "" and domain <> "" then
	if pw = "" or pw1 = "" then
		errstr = "密码不可为空"
	end if
end if


'-----------------------------------------
dim ei
set ei = server.createobject("easymail.domain")
ei.Load

dim comeinadd
comeinadd = false

if Session("Reg") = "step 2 over" and username <> "" and domain <> "" and pw <> "" and pw1 <> "" and pw = pw1 and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim isok
	isok = true

	if AssumpsitString <> trim(request("myAssumpsitString")) and Application("em_SignMode") = 3 then
		isok = false
		errstr = "校验码输入错误"
	end if

	if regemail = "" and Application("em_SignMode") = 2 then
		isok = false
		errstr = "接收确认信的有效Email地址不可为空"
	end if

	if isok = true then
		if trim(request.Cookies("zatt_checkcode")) <> trim(request("zck")) then
			isok = false
			errstr = "验证码填写错误"
		end if
	end if

	if isok = true then
		dim isdomain
		isdomain = false

		ei.GetControlMsg domain, isshow, maxuser, manager
		mdn = ei.GetUserNumberInDomain(domain)
		isdomain = ei.IsDomain(domain)

		if mdn >= maxuser then
			errstr="当前域中的用户数已满"
			isok = false
		end if

		if isdomain = false then
			errstr="无效域名"
			isok = false
		end if
	end if

if isok = true then
	if InStr(username, "!") or InStr(username, """") or InStr(username, "#") or InStr(username, "$") or InStr(username, "%") then
		errstr="用户名中包含非法字符"
		isok = false
	end if

	if InStr(username, "&") or InStr(username, "`") or InStr(username, "(") or InStr(username, ")") or InStr(username, "*") then
		errstr="用户名中包含非法字符"
		isok = false
	end if

	if InStr(username, "+") or InStr(username, ",") or InStr(username, "/") or InStr(username, ":") then
		errstr="用户名中包含非法字符"
		isok = false
	end if

	if InStr(username, ";") or InStr(username, "<") or InStr(username, "=") or InStr(username, ">") or InStr(username, "?") then
		errstr="用户名中包含非法字符"
		isok = false
	end if

	if InStr(username, "@") or InStr(username, "[") or InStr(username, "\") or InStr(username, "]") or InStr(username, "^") then
		errstr="用户名中包含非法字符"
		isok = false
	end if

	if InStr(username, "'") or InStr(username, "{") or InStr(username, "|") or InStr(username, "}") or InStr(username, "~") then
		errstr="用户名中包含非法字符"
		isok = false
	end if

	if InStr(username, " ") or InStr(username, Chr(9)) then
		errstr="用户名中包含非法字符"
		isok = false
	end if
end if



	'-----
	if isok = true then
		Set easymail = Application("em")
		if Application("em_Enable_SignWithDomainUser") = true then
			if easymail.isUser(username & "@" & domain) = true then
				errstr="系统中已有此用户"
				isok = false
			end if
		else
			if easymail.isUser(username) = true then
				errstr="系统中已有此用户"
				isok = false
			end if    
		end if
		Set easymail = nothing
	end if

	'-----
	if isok = true then
		if Application("em_Enable_SignHold") = true then
			dim sh
			Set sh = server.createobject("easymail.SignHold")
			sh.Load()

			if sh.IsHold(username) = true then
				errstr="系统中已有此用户"
				isok = false
			end if

			Set sh = nothing
		end if
	end if
	'-----


	if isok = true then
		comeinadd = true
	end if

end if



'------------------------------------------------
if comeinadd = true and Request.ServerVariables("REQUEST_METHOD") = "POST" and isshow = true then
	reg_Text = trim(request("reg_Text"))
	set mri = server.createobject("easymail.MoreRegInfo")

	if Enable_SignWithInputMoreInfo = true and reg_Text <> "" then
		if Application("em_Enable_SignWithDomainUser") = true then
			mri.NewTempRegInfo LCase(username) & "@" & LCase(domain) & Chr(9) & pw & Chr(9) & LCase(domain)
		else
			mri.NewTempRegInfo LCase(username) & Chr(9) & pw & Chr(9) & LCase(domain)
		end if

		if Len(reg_Text) > 0 then
			dim item
			dim ss
			dim se
			ss = 1
			se = 1

			Do While 1
				se = InStr(ss, reg_Text, Chr(13))

				If se <> 0 Then
					item = Mid(reg_Text, ss, se - ss)
					mri.AddLine_RegInfo item
				Else
					Exit Do
				End If

				ss = se + 2
			Loop
		end if
	end if


	Session("Reg") = ""

	if Application("em_SignMode") > 0 and Application("em_SignMode") <> 3 then
		errstr = ""

		dim pr
		set pr = server.createobject("easymail.PendRegister")
		pr.Load Application("em_SignWaitDays")

		if Application("em_Enable_SignWithDomainUser") = true then
			if Application("em_SignMode") = 1 then
				accode = pr.SignPend(username & "@" & domain, pw, domain, Request.ServerVariables("REMOTE_ADDR"), username & " (" & CStr(createGRSN()) & ")")
			else
				accode = pr.SignPend(username & "@" & domain, pw, domain, Request.ServerVariables("REMOTE_ADDR"), regemail)
			end if
		else
			if Application("em_SignMode") = 1 then
				accode = pr.SignPend(username, pw, domain, Request.ServerVariables("REMOTE_ADDR"), username & " (" & CStr(createGRSN()) & ")")
			else
				accode = pr.SignPend(username, pw, domain, Request.ServerVariables("REMOTE_ADDR"), regemail)
			end if
		end if

		if accode <> "" then
			if Application("em_SignMode") = 2 then
				if Application("em_Enable_SignWithDomainUser") = true then
					amText = pr.GetActionMailText(username & "@" & domain, username & "@" & domain, Request.ServerVariables("REMOTE_ADDR"), accode)
				else
					amText = pr.GetActionMailText(username, username & "@" & domain, Request.ServerVariables("REMOTE_ADDR"), accode)
				end if

				Dim mailsend
				Dim sender
				sender = Application("em_am_Accounts")

				Set mailsend = Server.CreateObject("easymail.MailSend")
				mailsend.CreateNew sender, "temp"

				mailsend.MailName = Application("em_am_Name")

				mailsend.EM_To = regemail

				mailsend.EM_Subject = pr.acSubject
				mailsend.EM_Text = amText

				if mailsend.Send() = false then
					errstr = "确认邮件发送失败"
				end if

				Set mailsend = Nothing
			end if

			pr.Save
		else
			if Application("em_SignMode") = 1 then
				errstr = "邮箱注册失败: 此用户名已被申请"
			else
				errstr = "邮箱注册失败: 可能是此用户名已申请或接收的外部信箱已被使用"
			end if
		end if

		set pr = nothing
		set ei = nothing

		if accode = "" then
			set mri = nothing

			if Enable_SignWithInputMoreInfo = false or reg_Text = "" then
				Response.Redirect "outerr.asp?gourl=default.asp&errstr=" & Server.URLEncode(errstr) & "&" & getGRSN()
			else
				Session("Reg") = "step 1 over"
%>
<html>
<body>
<form action="create.asp?<%=getGRSN() %>" name="f1" METHOD="POST">
<div style="position:absolute; top:10; left:10; z-index:15; visibility:hidden">
<textarea name="reg_Text" cols="0" rows="0"><%=reg_Text %></textarea>
<input type="hidden" name="errstr" value="<%=errstr %>">
</div>
</form>
</body>

<script type="text/javascript">
<!--
f1.submit();
//-->
</script>
</html>
<%
			end if
		else
			if errstr = "" then
				if Application("em_Enable_SignNumberLimit") = true then
					Response.Cookies("SignOk") = "1"
					Response.Cookies("SignOk").Expires = DateAdd("d", Application("em_SignNumberLimitDays"), Now())
				end if

				if Application("em_SignMode") = 1 then
					mri.SaveRegInfo
					set mri = nothing

					Response.Redirect "outok.asp?gourl=default.asp&errstr=" & Server.URLEncode("申请成功，请等待管理员审批") & "&" & getGRSN()
				end if

				if Application("em_SignMode") = 2 then
					mri.SaveRegInfo
					set mri = nothing

					Response.Redirect "outok.asp?gourl=default.asp&errstr=" & Server.URLEncode("申请成功，请在 " & regemail & " 处接收确认邮件") & "&" & getGRSN()
				end if
			else
				set mri = nothing

				if Enable_SignWithInputMoreInfo = false or reg_Text = "" then
					Response.Redirect "outerr.asp?gourl=default.asp&errstr=" & Server.URLEncode(errstr) & "&" & getGRSN()
				else
					Session("Reg") = "step 1 over"
%>
<html>
<body>
<form action="create.asp?<%=getGRSN() %>" name="f1" METHOD="POST">
<div style="position:absolute; top:10; left:10; z-index:15; visibility:hidden">
<textarea name="reg_Text" cols="0" rows="0"><%=reg_Text %></textarea>
<input type="hidden" name="errstr" value="<%=errstr %>">
</div>
</form>
</body>

<script type="text/javascript">
<!--
f1.submit();
//-->
</script>
</html>
<%
				end if
			end if

			set mri = nothing

			response.redirect "outok.asp?gourl=default.asp&" & getGRSN()
		end if

	else
		Set easymail = Application("em")
		if Application("em_Enable_SignWithDomainUser") = true then
			easymail.AddUserPublic username & "@" & domain, pw, domain, "From: " & Request.ServerVariables("REMOTE_ADDR"), Sign_AccessMode
		else
			easymail.AddUserPublic username, pw, domain, "From: " & Request.ServerVariables("REMOTE_ADDR"), Sign_AccessMode
		end if
		Set easymail = nothing

		mri.SaveRegInfo

		if Application("em_Enable_SignNumberLimit") = true then
			Response.Cookies("SignOk") = "1"
			Response.Cookies("SignOk").Expires = DateAdd("d", Application("em_SignNumberLimitDays"), Now())
		end if
	end if


	set mri = nothing
%>
<!DOCTYPE html>
<html>
<head>
<title>申请邮箱</title>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<STYLE type=text/css>
<!--
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
-->
</STYLE>
</head>

<body>
<br>
<table width="82%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
创建成功
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="left" style="padding-left:12px; padding-right:12px;">
	邮箱 <%=username & "@" & domain %> 创建成功。<br><br>
<%
if Application("em_Enable_SignWithDomainUser") = true then
%>
	您的登录用户名是：<font color="#901111"><%=username & "@" & domain %></font>
<%
else
%>
	您的登录用户名是：<font color="#901111"><%=username %></font>
<%
end if
%>
	</td></tr>

	<tr><td class="block_top_td" style="height:10px;"></td></tr>

	<tr><td align="left" style="background-color:white; padding-top:18px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="create.asp?<%=getGRSN() %>"><< 上一步</a>
<a class='wwm_btnDownload btn_blue' href="default.asp?<%=getGRSN() %>">返回首页</a>
	</td></tr>
</table>
</body>
</html>
<%
else
%>
<!DOCTYPE html>
<html>
<head>
<title>申请邮箱</title>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<STYLE type=text/css>
<!--
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.wwm_ar_msg {padding:3px; color:#222222; line-height:18px; background:#FFF8D3; border-radius:4px; -webkit-border-radius:4px; padding:6px 8px 4px 8px; text-align:left; border:#ff0000 1px solid;}
.td_l {white-space:nowrap; height:30px; padding-top:4px; padding-right:4px;}
-->
</STYLE>
</head>

<script type="text/javascript">
<!--
function isCharsInBag (s, bag)
{
	var i,c;
	for (i = 0; i < s.length; i++)
	{
		c = s.charAt(i);

		if (bag.indexOf(c) == -1)
			return false;
	}

	return true;
}

function ischinese(s)
{
	if (s.charAt(s.length - 1) == '.')
		return true;

	var badChar = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_.";

	return !isCharsInBag(s, badChar);
}

function GetStringRealLength(tstr) {
	var reallen = 0;

	for (var i = 0; i < tstr.length; i++)
	{
		if (escape(tstr.charAt(i)).length < 4)
			reallen++;
		else
			reallen = reallen + 2;
	}

	return reallen;
}

function checkpw(){
	if (fc.username.value == "")
	{
		alert("用户名不可为空");
		fc.username.focus();
		return ;
	}

	if (GetStringRealLength(fc.username.value) < <%=Sign_AccountMinLen %>)
	{
		alert("用户名长度不可以小于 <%=Sign_AccountMinLen %>个字符");
		fc.username.focus();
		return ;
	}
<%
if Application("em_Enable_SignEnglishName") = true then
%>
	if (ischinese(fc.username.value) == true)
	{
		alert("用户名输入字符非法");
		fc.username.focus();
		return ;
	}
<%
end if
%>
	if (fc.pw.value == "")
	{
		alert("密码不可为空");
		fc.pw.focus();
		return ;
	}

	if (fc.pw1.value == "")
	{
		alert("密码不可为空");
		fc.pw1.focus();
		return ;
	}
<%
if Application("em_SignMode") = 2 then
%>
	if (fc.regemail.value == "")
	{
		alert("接收确认信的有效Email地址不可为空");
		fc.regemail.focus();
		return ;
	}

	var mailisok = true;
	var sp = fc.regemail.value.indexOf("@");
	if (sp == -1)
		mailisok = false;
	else
	{
		sp = fc.regemail.value.indexOf("@", sp + 1);
		if (sp != -1)
			mailisok = false;
		else
		{
			if (fc.regemail.value.charAt(0) == '@' || fc.regemail.value.charAt(fc.regemail.value.length - 1) == '@')
			{
				mailisok = false;
			}
		}
	}

	if (mailisok == false)
	{
		alert("接收确认信的Email地址无效");
		fc.regemail.focus();
		return ;
	}
<%
end if
%>
	if (fc.pw.value != fc.pw1.value)
	{
		alert("输入的密码不相同");
		fc.pw1.focus();
		return ;
	}

	if (fc.pw.value.length < <%=Sign_PassWordMinLen %>)
	{
		alert("密码长度不可以小于 <%=Sign_PassWordMinLen %>个字符");
		fc.pw.focus();
		return ;
	}

	if (now_rt < 2)
	{
		alert("您输入的密码强度不足! ");
		return ;
	}

	if (document.getElementById("zck").value.length < 1)
	{
		document.getElementById("zck").focus();
		alert("验证码填写错误!");
		return ;
	}

	fc.submit();
}

	ratingMsgs = new Array(4);
	ratingMsgColors = new Array(4);
	barColors = new Array(4);
	ratingMsgs[0] = "太短";
	ratingMsgs[1] = "弱";
	ratingMsgs[2] = "一般";
	ratingMsgs[3] = "极佳";
	ratingMsgColors[0] = "#676767";
	ratingMsgColors[1] = "#aa0033";
	ratingMsgColors[2] = "#f5ac00";
	ratingMsgColors[3] = "#008000";
	barColors[0] = "#dddddd";
	barColors[1] = "#aa0033";
	barColors[2] = "#ffcc33";
	barColors[3] = "#008000";
	var now_rt = 0;

function CreateRatePasswdReq(pwd){
	if (!isBrowserCompatible) return;

	if(!pwd) return; 
	passwd=pwd.value;
	var min_passwd_len = 6; 
	if (passwd.length < min_passwd_len){
		if (passwd.length > 0)
			DrawBar(0);
		else
			ResetBar();
	} else {
		rating = checkPasswdRate(passwd);
		DrawBar(rating);
	}
}

function getElement(name){
	if (document.all)
		return document.all(name);

	return document.getElementById(name);
}

function DrawBar(rating){
	now_rt = rating;
	var posbar = getElement('posBar');
	var negbar = getElement('negBar');
	var passwdRating = getElement('passwdRating');
	var barLength = getElement('passwdBar').width;
	if (rating >= 0 && rating < 4) {
		posbar.style.width = barLength / 3 * rating + "px";
		negbar.style.width = barLength / 3 * (3 - rating) + "px";
	}

	posbar.style.background = barColors[rating];
 	passwdRating.innerHTML = "<font color='" + ratingMsgColors[rating] + "'>" + ratingMsgs[rating] + "</font>";
}

function ResetBar(){
	var posbar = getElement('posBar');
	var negbar = getElement('negBar');
	var passwdRating = getElement('passwdRating');
	var barLength = getElement('passwdBar').width;
	posbar.style.width = "0px";
	negbar.style.width = barLength + "px";
	passwdRating.innerHTML = "";
}

	var agt = navigator.userAgent.toLowerCase();
	var is_op = (agt.indexOf("opera") != -1);
	var is_ie = (agt.indexOf("msie") != -1) && document.all && !is_op;
	var is_mac = (agt.indexOf("mac") != -1);
	var is_gk = (agt.indexOf("gecko") != -1);
	var is_sf = (agt.indexOf("safari") != -1);

function gff(str, pfx){
	var i = str.indexOf(pfx);
	if (i != -1)
	{
		var v = parseFloat(str.substring(i + pfx.length));
		if (!isNaN(v))
			return v;
	}
	return null;
}

function Compatible(){
	if (is_ie && !is_op && !is_mac)
	{
		var v = gff(agt, "msie ");
		if (v != null)
			return (v >= 6.0);
	}

	if (is_gk && !is_sf)
	{
		var v = gff(agt, "rv:");
		if (v != null)
			return (v >= 1.4);
		else
		{
			v = gff(agt, "galeon/");
			if (v != null)
				return (v >= 1.3);
		}
	}

	if (is_sf)
	{
		var v = gff(agt, "applewebkit/");
		if (v != null)
			return (v >= 124);
	}
	return false;
}

var isBrowserCompatible = Compatible();

function CharMode(iN){
	if (iN>=48 && iN <=57)
		return 1;
	if ((iN>=65 && iN <=90) || (iN>=97 && iN <=122))
		return 2;
	else
		return 4;
}

function bitTotal(num){
	modes=0;
	for (i=0;i<4;i++)
	{
		if (num & 1) modes++;
			num>>>=1;
	}
	return modes; 
}

function checkPasswdRate(sPW){
	if (sPW.length < 6)
		return 0;

	Modes=0;
	for (i=0;i<sPW.length;i++){
		Modes|=CharMode(sPW.charCodeAt(i));
	}

	return bitTotal(Modes);
}
//-->
</script>

<body>
<br>
<form name="fc" METHOD="POST">
<table width="82%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
申请邮箱
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="center" style="padding-left:16px; padding-right:16px; word-break:break-all; word-wrap:break-word;">

<div class="wwm_ar_msg">&nbsp;<%=server.htmlencode(errstr) %>.
<%
if Application("em_SignMode") = 2 then
%>
<br>&nbsp;注意: 一个外部邮件地址, 仅可用于申请一个邮箱.
<%
elseif Application("em_SignMode") = 3 then
%>
<br>&nbsp;注意: 服务器要求输入正确的验证码后, 才可以申请邮箱.
<%
end if
%>
</div>
	</td></tr>

	<tr><td align="left" style="padding: 8px 16px 0 16px;">
		<table width="100%" border="0" align="left" cellspacing="0" bgcolor="white">
		<tr><td align="right" width="12%" class="td_l">用户名：</td>
		<td align="left">
		<input type="text" name="username" value="<%=username %>" maxlength="32" size="30" class="n_textbox">
		</td>
		</tr>

		<tr><td align="right" class="td_l">
		域名：</td>
		<td align="left">
<select name="domain" class="drpdwn" size="1">
<%
i = 0
allnum = ei.getcount

set wmethod = server.createobject("easymail.WMethod")

do while i < allnum
	domainname = ei.GetDomain(i)

	ei.GetControlMsg domainname, isshow, maxuser, manager

	if isshow = true then
		if domainname <> domain then
			response.write "<option value='" & domainname & "'>" & wmethod.Puny_To_Domain(domainname) & "</option>"
		else
			response.write "<option value='" & domainname & "' selected>" & wmethod.Puny_To_Domain(domainname) & "</option>"
		end if
	end if

	domainname = NULL
	isshow = NULL
	maxuser = NULL
	manager = NULL

	i = i + 1
loop

set wmethod = nothing
%>
</select>
		</td>
		</tr>

		<tr><td align="right" class="td_l">密码：</td>
		<td align="left">
<table border="0" align="left" cellspacing="0" cellPadding="0" bgcolor="white">
	<tr>
	<td>
	<input type="password" name="pw" maxlength="32" class="n_textbox" onkeyup="CreateRatePasswdReq(this)">
	</td>

	<td style="padding-left:12px;">
	<table width="10" cellSpacing=0 cellPadding=0 border=0>
	<tr>
	<td vAlign=top noWrap width="0"><font color="#444444">密码强度：</font></td>
	<td vAlign=top noWrap><font color="#808080">
	<strong><div id=passwdRating></div></strong></font>
	</td></tr>
	<tr><td colspan="2">
		<table cellSpacing=0 cellPadding=0 border=0><tr>
		<td>
			<table id=passwdBar cellSpacing=0 cellPadding=0 width="120" bgColor=#ffffff border=0>
			<tr>
			<td id=posBar width=0% bgColor=#e0e0e0 height=4></td>
			<td id=negBar width="100%" bgColor=#e0e0e0 height=4></td>
			</tr>
			</table>
		</td></tr>
		</table>
	</td></tr>
	</table>
</td></tr>
</table>
		</td>
		</tr>

		<tr><td align="right" class="td_l">确认密码：</td>
		<td align="left">
		<input type="password" name="pw1" maxlength="32" class="n_textbox">
		</td>
		</tr>
<%
if Application("em_SignMode") = 2 then
%>
		<tr><td align="right" class="td_l">接收确认信的外部邮件地址：</td>
		<td align="left">
		<input type="text" name="regemail" maxlength="64" size="30" class="n_textbox">
		</td>
		</tr>
<%
elseif Application("em_SignMode") = 3 then
%>
		<tr><td align="right" class="td_l">校验码：</td>
		<td align="left">
		<input type="text" name="myAssumpsitString" maxlength="128" class="n_textbox">
		</td>
		</tr>
<%
end if
%>
		<tr><td align="right" class="td_l">输入验证码：</td>
		<td align="left">
		<img src="tu.asp" align="absmiddle" border="0"><input type="text" name="zck" id="zck" class="n_textbox" size="2" maxlength="2">
		</td>
		</tr>
	</table>

</td></tr>

<tr><td class="block_top_td" style="height:8px;"></td></tr>

<tr><td align="left" style="background-color:white; padding-top:18px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="default.asp?<%=getGRSN() %>">取消</a>
<a class='wwm_btnDownload btn_blue' href="javascript:checkpw();">提交</a>
</td></tr>
</table>

<div style="position:absolute; top:10; left:10; z-index:15; display:none;">
<textarea name="reg_Text" cols="0" rows="0"><%=trim(request("reg_Text")) %></textarea>
</div>
</form>
<%
if Application("em_EnableTrap") = true then
%>
<div style="position:absolute; top:0; left:0; z-index:0; display:none;">
<a href="mailto:<%=Application("em_TrapMail") %>"><%=Application("em_TrapMail") %></a>
</div>
<%
end if
%>
</body>
</html>

<%
	Session("Reg") = "step 2 over"
end if
%>

<%
set ei = nothing
%>

<%
function getGRSN()
	dim theGRSN
	Randomize
	theGRSN = Int((9999999 * Rnd) + 1)

	getGRSN = "GRSN=" & CStr(theGRSN)
end function


function createGRSN()
	Randomize
	createGRSN = Int((9999999 * Rnd) + 1)
end function
%>
