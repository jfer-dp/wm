<%
Response.CacheControl = "no-cache" 
%>

<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
pw = trim(request("pwhidden"))

if pw <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	pw = strDecode(pw, trim(request("picnum")))

	dim wemcert
	set wemcert = server.createobject("easymail.WebEasyMailCert")
	wemcert.Load Session("wem"), Session("mail")

	isok = wemcert.ChangeSecCertPassword(trim(request("oldpw")), pw, CLng(trim(request("save_day"))))

	set wemcert = nothing

	if isok = true then
		Response.Redirect "ok.asp?gourl=cert_mysec.asp&" & getGRSN()
	else
		Response.Redirect "err.asp?gourl=cert_changepw.asp&errstr=" & Server.URLEncode(a_lang_017) & "&" & getGRSN()
	end if
end if
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<STYLE type=text/css>
<!--
.block_top_td {white-space:nowrap; background:white; font-size:0pt; height:1px;}
.td_l {white-space:nowrap; background-color:white; height:30px; text-align:right; padding-top:6px;}
.td_r {white-space:nowrap; background-color:white; height:30px; text-align:left; padding-left:6px;}
-->
</STYLE>
</head>

<script type="text/javascript">
<!--
function checkpw(){
	if (document.form1.pw1.value.length < 8 || document.form1.pw2.value.length < 8)
	{
		document.getElementById("pw1").focus();
		alert("<%=a_lang_018 %>");
		return ;
	}

	if (document.form1.pw1.value != "" && document.form1.pw2.value != "")
	{
		if (document.form1.pw1.value != document.form1.pw2.value)
		{
			alert("<%=a_lang_019 %>");
			document.getElementById("pw2").focus();
		}
		else
		{
			document.form1.pwhidden.value = encode(document.form1.pw1.value, parseInt(document.form1.picnum.value));
			document.form1.submit();
		}
	}
}

function encode(datastr, bassnum) {
	var tempstr;
	var tchar;
	var newdata = "";

	for (var i = 0; i < datastr.length; i++)
	{
		tchar = 65535 + bassnum - datastr.charCodeAt(i);
		tchar = tchar.toString();

		while(tchar.length < 5)
		{
			tchar = "0" + tchar;
		}

		newdata = newdata + tchar;
	}

	return newdata;
}
//-->
</script>

<body>
<form name="form1" method="post" action="cert_changepw.asp">
<input type="hidden" name="forget" value="<%=trim(request("forget")) %>">
<input type="hidden" name="pwhidden">
<input type="hidden" name="picnum" value="<%=createRnd() %>">

<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white" style="margin-top:6px;">
<tr><td colspan="2" style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=a_lang_020 %>
</td></tr>
<tr><td colspan=2 class="block_top_td" style="height:10px; _height:12px;"></td></tr>

<tr><td width="10%" class="td_l">
<%=a_lang_021 %><%=s_lang_mh %>
</td>
<td class="td_r">
<input type="password" id="oldpw" name="oldpw" maxlength="64" size="36" class="n_textbox">
</td></tr>

<tr><td class="td_l">
<%=a_lang_022 %><%=s_lang_mh %>
</td>
<td class="td_r">
<input type="password" id="pw1" name="pw1" maxlength="64" size="36" class="n_textbox">
</td></tr>

<tr><td class="td_l">
<%=a_lang_023 %><%=s_lang_mh %>
</td>
<td class="td_r">
<input type="password" id="pw2" name="pw2" maxlength="64" size="36" class="n_textbox">
</td></tr>

<tr>
	<td colspan="2" height="30" class="td_r">
<%=a_lang_024 %><%=s_lang_mh %>
<select name="save_day" class="drpdwn">
<option value="-1" selected><%=a_lang_025 %></option>
<option value="0"><%=a_lang_026 %></option>
<%
	now_temp = 10

	do while now_temp < 999
		response.write "<option value='" & now_temp & "'>" & now_temp & a_lang_027 & "</option>" & Chr(13)

		now_temp = now_temp + 10
	loop
%>
</select>
	</td>
</tr>

<tr><td colspan=2 class="block_top_td" style="height:8px;"></td></tr>

<tr><td colspan="2" align="left" style="background-color:white; padding-right:16px; padding-top:14px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="cert_mysec.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:checkpw();"><%=s_lang_ok %></a>
</td></tr>
</table>

</form>
</body>
</html>


<%
function createRnd()
	dim retval
	retval = getGRSN()

	if Len(retval) > 4 then
		retval = Right(retval, 4)
	end if

	if Left(retval, 1) = "0" then
		retval = "5" & Right(retval, 3)
	end if

	createRnd = retval
end function

function strDecode(sd_Data, sd_bassnum)
	dim sd_vChar
	dim sd_NewData
	dim sd_TempChar
	sd_vChar = 1

	do
		if sd_vChar > Len(sd_Data) then
			exit do
		end if

	    sd_TempChar = CLng(Mid(sd_Data, sd_vChar, 5))
		sd_TempChar = ChrW(65535 + sd_bassnum - sd_TempChar)

        sd_NewData = sd_NewData & sd_TempChar
		sd_vChar = sd_vChar + 5
	loop

	strDecode = sd_NewData
end function
%>
