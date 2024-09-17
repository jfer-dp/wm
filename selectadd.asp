<!--#include file="passinc.asp" --> 
<!--#include file="language-2.asp" --> 

<%
	mode = trim(request("mode"))
	ofm = trim(request("ofm"))

	if Len(ofm) > 0 then
		modestr = "opener." + ofm + ".value"
	else
		modestr = "opener.value"
	end if

	if mode = "To" then
		dispmode = b_lang_068
	elseif mode = "Cc" then
		dispmode = b_lang_069
	elseif mode = "Bcc" then
		dispmode = b_lang_070
	elseif mode = "deliver" then
		modestr = "opener.document.f1.to.value"
		dispmode = b_lang_071
	end if

dim ads
set ads = server.createobject("easymail.Addresses")
ads.Load Session("wem")

pop_toccbcc = false
if InStr(modestr, "EasyMail_") > 0 then
	pop_toccbcc = true
end if
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail <%=b_lang_072 %></TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
</HEAD>

<script type="text/javascript" src="images/mglobal.js"></script>

<script language="JavaScript">
<!-- 
function window_onload() {
	var sstr = <%=modestr %>;
	var sbegin = 0;
	var send = 0;
	var stemp = "";
	var in_begin = 0;
	var in_end = 0;
	var in_name;
	var in_email;
	var t_len = 0;
	var sf_end = 0;

	while(1)
	{
		send = sstr.indexOf(',', sbegin);
		sf_end = sstr.indexOf(';', sbegin);

		if ((sf_end != -1 && sf_end < send) || send == -1)
			send = sf_end;

		if (send != -1)
			stemp = sstr.substring(sbegin, send);
		else
			stemp = sstr.substring(sbegin);

		if (stemp != "")
		{
			in_begin = stemp.indexOf('<', 0);

			if (in_begin != -1)
			{
				in_end = stemp.indexOf('>', in_begin);

				if (in_end != -1)
				{
					in_email = stemp.substring(in_begin + 1, in_end);

					if (in_begin > 0)
						in_name = stemp.substring(0, in_begin);
					else
						in_name = in_email;

					t_len = in_name.length - 1;

					for (; t_len >= 0; t_len--)
					{
						if (in_name.charAt(t_len) == ' ')
							in_name = in_name.substring(0, t_len);
					}

					addnewaddress(in_name, in_email);
				}
				else
					addnewaddress(stemp, stemp);
			}
			else
				addnewaddress(stemp, stemp);
		}

		if (send == -1)
			break;

		sbegin = send + 1;
	}
}

function addnewaddress(in_name, in_email)
{
	var oOption = document.createElement("OPTION");
	oOption.text = in_name;
	oOption.value = in_email;

	if (ie == false)
		document.f1.selectusers.appendChild(oOption);
	else
		document.f1.selectusers.add(oOption);
}

function addin()
{
	var i = 0;
	for (i; i < document.f1.selectallusers.length; i++)
	{
		if (document.f1.selectallusers[i].selected == true)
		{
			if (isinlist(document.f1.selectallusers[i].value) == false)
			{
				var oOption = document.createElement("OPTION");
<%
if pop_toccbcc = true then
%>
				oOption.text = document.f1.selectallusers[i].text;
<%
else
%>
				oOption.text = document.f1.selectallusers[i].value;
<%
end if
%>
				oOption.value = document.f1.selectallusers[i].value;

				if (ie == false)
					document.f1.selectusers.appendChild(oOption);
				else
					document.f1.selectusers.add(oOption);
			}
		}
	}
}

function isinlist(name)
{
	var i = 0;
	for (i; i < document.f1.selectusers.length; i++)
	{
		if (document.f1.selectusers[i].value == name)
		{
			return true;
		}
	}
	
	return false;
}

function delout()
{
	var i = 0;
	for (i; i < document.f1.selectusers.length; i++)
	{
		if (document.f1.selectusers[i].selected == true)
		{
			document.f1.selectusers.remove(i);
			i--;
		}
	}
}

function gook()
{
	var str = "";
	var i = 0;
	for (i; i < document.f1.selectusers.length; i++)
	{
<%
if pop_toccbcc = true then
%>
		if (document.f1.selectusers[i].text == document.f1.selectusers[i].value)
		{
			if (str.length > 0)
				str = str + "," + document.f1.selectusers[i].value;
			else
				str = document.f1.selectusers[i].value;
		}
		else
		{
			if (str.length > 0)
				str = str + "," + document.f1.selectusers[i].text + " <" + document.f1.selectusers[i].value + ">";
			else
				str = document.f1.selectusers[i].text + " <" + document.f1.selectusers[i].value + ">";
		}
<%
else
%>
		if (str.length > 0)
			str = str + "," + document.f1.selectusers[i].value;
		else
			str = document.f1.selectusers[i].value;
<%
end if
%>
	}

	<%=modestr %> = str;

	self.close();
}
// -->
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
<form name="f1">
<table border="0" width="92%" cellpadding=2 cellspacing=0 align="center">
	<tr bgcolor="#104A7B">
	<td height="20" rowspan="2" align="center" width="45%" style="color:#ffffff;"><%=b_lang_073 %></td>
	<td height="20" rowspan="2"></td>
	<td height="20" rowspan="2" align="center" width="45%" style="color:#ffffff;"><%=dispmode %></td>
	</tr>
	<tr></tr>
	<tr>
	<td height="94" rowspan="2" width="45%" align="center">
	<br>
	<select name="selectallusers" size="14" class="drpdwn" style="width:160px;" multiple LANGUAGE=javascript ondblclick="return addin()">
<%
allnum = ads.EmailCount
i = 0
do while i < allnum
	ads.MoveTo i
	Response.Write "<option value='" & server.htmlencode(replace(ads.email, """", "")) & "'>" & server.htmlencode(replace(ads.nickname, """", "")) & "</option>" & Chr(13)
	i = i + 1
loop

allnum = ads.GroupCount
i = 0
do while i < allnum
	ads.GetGroupInfo i, nickname, emails
	Response.Write "<option value='" & server.htmlencode(replace(nickname, """", "")) & "'>" & server.htmlencode(replace(nickname, """", "")) & "</option>" & Chr(13)

	nickname = NULL
	emails = NULL

	i = i + 1
loop
%>
	</select>
	</td>
	<td height="47" width="10%" valign="bottom" align="center" style="padding-bottom:12px;">
	<a class="wwm_btnDownload btn_gray" href="#" onclick="addin()">>></a>
	</td>
	<td height="94" rowspan="2" width="45%" align="center">
	<br>
	<select id="selectusers" name="selectusers" size="14" class="drpdwn" style="width:160px;" multiple LANGUAGE=javascript ondblclick="return delout()">
	</select>
	</td></tr>
	<tr><td height="47" width="10%" valign="top" align="center" style="padding-top:12px;">
	<a class="wwm_btnDownload btn_gray" href="#" onclick="delout()"><<</a>
	</td></tr>
</table>

<table width="92%" cellpadding=0 cellspacing=0 align="center">
	<tr><td style="border-bottom:1px solid #8CA5B5; height:16px;">&nbsp;</td></tr>
	<tr><td style="height:14px;">&nbsp;</td></tr>
	<tr><td align="right" style="padding-right:30px;">
	<a class="wwm_btnDownload btn_blue" href="#" onclick="gook()"><%=b_lang_074 %></a>&nbsp;
	<a class="wwm_btnDownload btn_blue" href="#" onclick="javascript:self.close();"><%=s_lang_cancel %></a>
	</td></tr>
</table>
</form>
</BODY>
</HTML>

<%
set ads = nothing
%>
