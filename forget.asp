<%
Response.CacheControl = "no-cache" 

mode = trim(request("mode"))

if mode <> "sub" then
	server_v1=CStr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=CStr(Request.ServerVariables("SERVER_NAME"))

	isok = false
	if Mid(server_v1, 8, Len(server_v2)) = server_v2 then
		isok = true
	end if

	if isok = false and LCase(Left(server_v1, 6)) = "https:" and Mid(server_v1, 9, Len(server_v2)) = server_v2 then
		isok = true
	end if

	if isok = false then
		Response.Redirect "forgetbf.asp?errstr=" & Server.URLEncode("URL错误") & "&" & getGRSN()
	end if
end if


dim webkill
set webkill = server.createobject("easymail.WebKill")
webkill.Load

rip = Request.ServerVariables("REMOTE_ADDR")

if webkill.IsKill(rip) = true then
	set webkill = nothing
	response.redirect "outerr.asp?gourl=default.asp&errstr=" & Server.URLEncode("拒绝IP地址 " & rip & " 访问") & "&" & getGRSN()
end if

set webkill = nothing
%>


<%
un = trim(request("username"))
dim errmsg

if Request.ServerVariables("REQUEST_METHOD") <> "POST" then
	response.redirect "forgetbf.asp?errstr=" & Server.URLEncode("操作错误") & "&" & getGRSN()
end if

if un = "" then
	response.redirect "forgetbf.asp?errstr=" & Server.URLEncode("请输入用户名") & "&" & getGRSN()
end if

if mode <> "sub" then
	set euser = Application("em")

	dim tmp_un
	tmp_un = euser.GetRealUser(un)
	if IsNull(tmp_un) = false and Len(tmp_un) > 0 then
		un = LCase(tmp_un)
	end if

	set euser = nothing
end if

dim ei
set ei = server.createobject("easymail.UserWeb")
'-----------------------------------------

ei.Load un

if ei.QuestionInfo = "" then
	set ei = nothing
	response.redirect "forgetbf.asp?errstr=" & Server.URLEncode("此用户未设置帐号保护功能或无此用户名") & "&" & getGRSN()
end if


if mode = "sub" then
	isok = ei.CheckAnswer(trim(request("AnswerInfo")))
	set ei = nothing

	if isok = true then
		set euser = Application("em")

		euser.GetUserByName2 un, outname, outdomain, outcomment, outforbid, outlasttime, outaccessmode

		if outforbid = false then
			outname = NULL
			outdomain = NULL
			outcomment = NULL
			outforbid = NULL
			outlasttime = NULL
			outaccessmode = NULL

			Session("tid") = euser.Login(un)
			Session("wem") = un
			Session("mail") = euser.GetUserMail(un)
			set euser = nothing

			dim ul
			set ul = server.createobject("easymail.UserLog")
			ul.Load Session("wem")
			ul.Add 1, Request.ServerVariables("REMOTE_ADDR")
			ul.Save
			set ul = nothing

			Session("changepw") = un
			Response.Redirect "welcome.asp?noticemsg=" & Server.URLEncode("请及时修改您的密码") & "&" & getGRSN()
		else
			outname = NULL
			outdomain = NULL
			outcomment = NULL
			outforbid = NULL
			outlasttime = NULL
			outaccessmode = NULL

			set euser = nothing
			response.redirect "default.asp?errstr=" & Server.URLEncode("此帐号已被禁用") & "&" & getGRSN()
		end if
	end if

	response.redirect "forgetbf.asp?errstr=" & Server.URLEncode("验证失败") & "&" & getGRSN()
end if 
%>

<!DOCTYPE html>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<STYLE type=text/css>
<!--
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.td_l {text-align:right; white-space:nowrap; height:26px; padding-top:4px; padding-left:16px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function window_onload() {
	document.getElementById("showAnswerInfo").focus();
}

function gosub() {
	if (document.getElementById("showAnswerInfo").value == "")
	{
		alert("请填写 [帐号保护答案]");
		document.getElementById("showAnswerInfo").focus();
		return ;
	}

	document.fm1.AnswerInfo.value = document.getElementById("showAnswerInfo").value;
	document.fm1.submit();
}
//-->
</script>

<body LANGUAGE=javascript onload="return window_onload()">
<form method="post" action="forget.asp" name="fm1">
<input type="hidden" name="username" value="<%=un %>">
<input type="hidden" name="mode" value="sub">
<input type="hidden" name="AnswerInfo">
</form>
<br>
<table width="82%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td colspan="2" class="block_top_td" style="height:4px;"></td></tr>
<tr><td colspan="2" style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
忘记密码
</td></tr>
<tr><td colspan="2" class="block_top_td" style="height:10px; _height:12px;"></td></tr>

<tr>
<td width="14%" class="td_l">帐号保护问题：</td>
<td align="left"><font color="#901111" style="font-size:14px;"><%=ei.QuestionInfo %></font></td>
</tr>

<tr>
<td class="td_l">提示信息：</td>
<td align="left"><font color="#901111" style="font-size:14px;"><%=ei.HintInfo %></font></td>
</tr>

<tr>
<td class="td_l">帐号保护答案：</td>
<td><input type="text" name="showAnswerInfo" id="showAnswerInfo" class='n_textbox' maxlength="256"></td>
</tr>

<tr><td colspan="2" class="block_top_td" style="height:10px;"></td></tr>

<tr><td colspan="2" align="left" style="background-color:white; padding-top:18px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="default.asp?<%=getGRSN() %>">取消</a>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();">提交</a>
</td></tr>
</table>
</BODY>
</HTML>

<%
set ei = nothing


function getGRSN()
	dim theGRSN
	Randomize
	theGRSN = Int((9999999 * Rnd) + 1)

	getGRSN = "GRSN=" & CStr(theGRSN)
end function
%>
