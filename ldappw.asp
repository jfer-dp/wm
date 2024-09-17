<%
Response.CacheControl = "no-cache" 
%>

<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
pw = trim(request("pwhidden"))

if pw <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	pw = strDecode(pw, trim(request("picnum")))

	if Session("wem") <> Application("em_TestAccounts") then
		dim ldap
		set ldap = server.createobject("easymail.LDAP")

		ldap.Submit_Password Session("wem"), pw

		set ldap = nothing

		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("myreginfo.asp")
	else
		Response.Redirect "err.asp?errstr=" & Server.URLEncode(s_lang_0048) & "&" & getGRSN() & "&gourl=" & Server.URLEncode("myreginfo.asp")
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
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
function checkpw(){
	if (f1.pw1.value != "" && f1.pw2.value != "")
	{
		if (f1.pw1.value != f1.pw2.value)
			alert("<%=s_lang_0047 %>");
		else
		{
			document.form1.pwhidden.value = encode(f1.pw1.value, parseInt(document.form1.picnum.value));
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
<form name="form1" method="post" action="ldappw.asp">
<input type="hidden" name="pwhidden">
<input type="hidden" name="picnum" value="<%=createRnd() %>">
</form>
<form name="f1">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td colspan="2" class="block_top_td" style="height:4px;"></td></tr>
<tr><td colspan="2" style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=s_lang_0033 %>
</td></tr>

<tr><td colspan="2" class="block_top_td" style="height:10px; _height:12px;"></td></tr>

<tr><td width="8%" nowrap align="right" height="30" style="padding-left:12px;">
<%=s_lang_0049 %><%=s_lang_mh %>
</td>
<td align="left">
<input type="password" name="pw1" maxlength="64" size="45" class="n_textbox">
</td></tr>

<tr><td nowrap align="right" height="30" style="padding-left:12px;">
<%=s_lang_0050 %><%=s_lang_mh %>
</td>
<td align="left">
<input type="password" name="pw2" maxlength="64" size="45" class="n_textbox">
</td></tr>

<tr><td colspan="2" class="block_top_td" style="height:8px;"></td></tr>

<tr><td colspan="2" align="left" style="background-color:white; padding-top:18px; padding-bottom:10px; border-top:1px #a7c5e2 solid;">
<a class='wwm_btnDownload btn_blue' href="myreginfo.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:checkpw();"><%=s_lang_save %></a>
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
