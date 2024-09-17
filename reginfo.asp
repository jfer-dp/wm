<%
Session("Reg") = ""

if Application("em_Enable_FreeSign") = false then
	Response.Redirect "default.asp?errstr=" & Server.URLEncode("邮箱申请功能被禁止") & "&" & getGRSN()
end if

if Application("em_Enable_SignNumberLimit") = true and Request.Cookies("SignOk") = "1" then
	Response.Redirect "default.asp?errstr=" & Server.URLEncode("您已申请过邮箱") & "&" & getGRSN()
end if

if Application("em_Enable_SignNumberLimit") = false then
	Response.Cookies("SignOk") = ""
end if


dim ei
set ei = server.createobject("easymail.MoreRegInfo")
ei.LoadSetting

allnum = ei.Count_Setting
errline = -1

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	i = 0
	reg_Text = ""

	do while i < allnum
		rgline = trim(request("rgline" & i))

		if errline = -1 and Len(rgline) < 3 then
			errline = i
			Exit Do
		end if

		if errline = -1 then
			if ei.AddLine_RegInfo(rgline) = false then
				errline = i
				Exit Do
			else
				reg_Text = reg_Text & rgline & Chr(13)
			end if
		end if 

	    i = i + 1
	loop

	if errline = -1 then
		set ei = nothing
		Session("Reg") = "step 1 over"
%>
<html>
<body>
<form action="create.asp?<%=getGRSN() %>" name="f1" METHOD="POST">
<div style="position:absolute; top:10; left:10; z-index:15; visibility:hidden">
<textarea name="reg_Text" cols="0" rows="0"><%=reg_Text %></textarea>
</div>
</form>
</body>

<script type="text/javascript">
<!--
document.f1.submit();
//-->
</script>
</html>
<%
		Response.End
	end if
end if
%>

<!DOCTYPE html>
<html>
<head>
<TITLE>WinWebMail</TITLE>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<STYLE type=text/css>
<!--
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.cont_td {white-space:nowrap; height:24px; border-bottom:1px solid #A5B6C8; padding-left:2px; padding-right:2px;}
-->
</STYLE>
</head>

<script type="text/javascript">
<!--
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

function checksub() {
<%
i = 0

do while i < allnum
	ei.Get_Setting i, s_name, s_sel, s_len

	Response.Write "	if (checkone(" & i & ", """ & s_name & """, " & s_sel & ", " & s_len & ") == false) return ;" & Chr(13)

	s_name = NULL
	s_sel = NULL
	s_len = NULL

	i = i + 1
loop
%>
	document.f1.submit();
}

function checkone(cnum, cname, csel, clen) {
	var isok = true;
	var inputObj = document.getElementById("input" + cnum);

	if (csel == 0 && GetStringRealLength(inputObj.value) >= clen)
		isok = false;
	else if (csel == 1 && GetStringRealLength(inputObj.value) != clen)
		isok = false;
	else if (csel == 2 && GetStringRealLength(inputObj.value) <= clen)
		isok = false;

	var tempstr = "";
	if (isok == false)
	{
		alert("输入错误.");
		inputObj.focus();
	}
	else
	{
		tempstr = cname + '\t' + csel + '\t' + clen + '\t' + inputObj.value;

		inputObj = document.getElementById("rgline" + cnum);
		inputObj.value = tempstr;
	}

	return isok;
}
//-->
</script>

<body>
<br>
<form action="reginfo.asp?<%=getGRSN() %>" name="f1" METHOD="POST">
<table width="82%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
填写个人资料
</td></tr>
<tr><td class="block_top_td" style="height:10px; _height:12px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
<%
i = 0

do while i < allnum
	ei.Get_Setting i, s_name, s_sel, s_len

	if errline <> i then
		if s_sel = 0 then
			Response.Write "<tr><td align='right' width='12%' class='cont_td'>" & server.htmlencode(s_name) & "：</td><td align='left' class='cont_td'>"
			Response.Write "<input id=""input" & i & """ name=""input" & i & """ type=""text"" class=""n_textbox"" size=""30"" maxlength=""" & s_len - 1 & """ value=""" & trim(request("input" & i)) & """><input id=""rgline" & i & """ name=""rgline" & i & """ type=""hidden""></td></tr>" & Chr(13)
		elseif s_sel = 1 then
			Response.Write "<tr><td align='right' width='12%' class='cont_td'>" & server.htmlencode(s_name) & "：</td><td align='left' class='cont_td'>"
			Response.Write "<input id=""input" & i & """ name=""input" & i & """ type=""text"" class=""n_textbox"" size=""30"" maxlength=""" & s_len & """ value=""" & trim(request("input" & i)) & """>&nbsp;*<input id=""rgline" & i & """ name=""rgline" & i & """ type=""hidden""></td></tr>" & Chr(13)
		elseif s_sel = 2 then
			Response.Write "<tr><td align='right' width='12%' class='cont_td'>" & server.htmlencode(s_name) & "：</td><td align='left' class='cont_td'>"
			Response.Write "<input id=""input" & i & """ name=""input" & i & """ type=""text"" class=""n_textbox"" size=""30"" value=""" & trim(request("input" & i)) & """ maxlength=""128"">&nbsp;*<input id=""rgline" & i & """ name=""rgline" & i & """ type=""hidden""></td></tr>" & Chr(13)
		end if
	else
		if s_sel = 0 then
			Response.Write "<tr><td align='right' width='12%' class='cont_td'><font color='#901111'>" & server.htmlencode(s_name) & "：</td><td align='left' class='cont_td'>"
			Response.Write "<input id=""input" & i & """ name=""input" & i & """ type=""text"" class=""n_textbox"" size=""30"" maxlength=""" & s_len - 1 & """ value=""" & trim(request("input" & i)) & """><input id=""rgline" & i & """ name=""rgline" & i & """ type=""hidden""></td></tr>" & Chr(13)
		elseif s_sel = 1 then
			Response.Write "<tr><td align='right' width='12%' class='cont_td'><font color='#901111'>" & server.htmlencode(s_name) & "：</td><td align='left' class='cont_td'>"
			Response.Write "<input id=""input" & i & """ name=""input" & i & """ type=""text"" class=""n_textbox"" size=""30"" maxlength=""" & s_len & """ value=""" & trim(request("input" & i)) & """>&nbsp;*<input id=""rgline" & i & """ name=""rgline" & i & """ type=""hidden""></td></tr>" & Chr(13)
		elseif s_sel = 2 then
			Response.Write "<tr><td align='right' width='12%' class='cont_td'><font color='#901111'>" & server.htmlencode(s_name) & "：</td><td align='left' class='cont_td'>"
			Response.Write "<input id=""input" & i & """ name=""input" & i & """ type=""text"" class=""n_textbox"" size=""30"" value=""" & trim(request("input" & i)) & """ maxlength=""128"">&nbsp;*<input id=""rgline" & i & """ name=""rgline" & i & """ type=""hidden""></td></tr>" & Chr(13)
		end if
	end if

	s_name = NULL
	s_sel = NULL
	s_len = NULL

	i = i + 1
loop
%>
	<tr><td colspan="2" height="40" bgcolor="#ffffff" align="left"><br>
<a class='wwm_btnDownload btn_blue' href="default.asp?<%=getGRSN() %>">取消</a>
<a class='wwm_btnDownload btn_blue' href="javascript:checksub();">下一步 >></a>
	</td></tr>
</table>

</td></tr>
</table>
</form>
</body>
</html>

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
