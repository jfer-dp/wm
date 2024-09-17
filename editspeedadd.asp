<!--#include file="passinc.asp" --> 
<!--#include file="language-1.asp" --> 

<%
dim ads
set ads = server.createobject("easymail.Addresses")
ads.Load Session("wem")

if trim(request("save")) = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim msg
	msg = trim(request("speedstr"))
	ads.DisableSpeedForAll

	if Len(msg) > 0 then
		dim ss
		dim se
		ss = 1
		se = 1

	    Do While 1
	        se = InStr(ss, msg, Chr(9))

	        If se <> 0 Then
    	        item = Mid(msg, ss, se - ss)
    	        ads.SetIsSpeedByNickName item, true
			Else
	            Exit Do
    	    End If

	        ss = se + 1
	    Loop
	end if

	ads.Save

	set ads = nothing
%>
<HTML>
<script language="JavaScript">
<!-- 
self.close();
// -->
</script>
</HTML>
<%
	Response.End
end if
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
</HEAD>

<script type="text/javascript" src="images/mglobal.js"></script>

<script language="JavaScript">
<!-- 
function addnewaddress(addstr)
{
	var oOption = document.createElement("OPTION");
	oOption.text = addstr;
	oOption.value = addstr;

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
				oOption.text = document.f1.selectallusers[i].value;
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

function savespeed(){
	var str = "";
	var i = 0;
	for (i; i < document.f1.selectusers.length; i++)
	{
		str = str + document.f1.selectusers[i].value + '\t';
	}

	document.f1.speedstr.value = str;
	document.f1.submit();
}
// -->
</script>

<BODY>
<FORM ACTION="editspeedadd.asp" METHOD=POST NAME="f1">
<INPUT NAME="speedstr" TYPE="hidden">
<INPUT NAME="save" TYPE="hidden" value="1">
<table border="0" width="92%" cellpadding=0 cellspacing=0 align="center">
	<tr bgcolor="#104A7B">
	<td colspan=3 nowrap style="color:white; height:22px; padding-left:6px;"><%=a_lang_198 %></td>
	</tr>
	<tr><td style="padding-top:12px; padding-left:16px; color:#444;"><%=a_lang_199 %><%=s_lang_mh %></td>
	<td>&nbsp;</td>
	<td style="padding-top:12px; padding-left:16px; color:#444;"><%=a_lang_200 %><%=s_lang_mh %></td>
	</tr>
	<tr>
	<td valign="top" align="center">
	<select name="selectallusers" size="13" class="drpdwn" style="width:170px;" multiple ondblclick="return addin()">
<%
dim allstr
allnum = ads.EmailCount
i = 0
do while i < allnum
	ads.MoveTo i
	Response.Write "<option value='" & server.htmlencode(replace(ads.nickname, """", "")) & "'>" & server.htmlencode(replace(ads.nickname, """", "")) & "</option>" & Chr(13)
	i = i + 1
loop
%>
	</select>
	</td>
	<td valign="middle" align="center">
	<a class="wwm_btnDownload btn_gray" href="javascript:addin();">>></a><p>
	<a class="wwm_btnDownload btn_gray" href="javascript:delout();"><<</a>
	</td>
	<td valign="top" align="center">
	<select name="selectusers" size="13" class="drpdwn" style="width: 170px;" multiple LANGUAGE=javascript ondblclick="return delout()">
<%
allnum = ads.SpeedCount
i = 0

do while i < allnum
	ads.GetSpeedInfoByIndex i, nickname, email
 	Response.Write "<option value=""" & server.htmlencode(replace(nickname, """", "")) & """>" & server.htmlencode(replace(nickname, """", "")) & "</option>" & Chr(13)

	nickname = NULL
	email = NULL

	i = i + 1
loop
%>
	</select>
	</td></tr>
</table>

<table width="92%" cellpadding=0 cellspacing=0 align="center">
	<tr><td style="border-bottom:1px solid #8CA5B5; height:16px;">&nbsp;</td></tr>
	<tr><td style="height:14px;">&nbsp;</td></tr>
	<tr><td align="right" style="padding-right:30px;">
	<a class="wwm_btnDownload btn_blue" href="#" onclick="savespeed()"><%=s_lang_ok %></a>&nbsp;
	<a class="wwm_btnDownload btn_blue" href="#" onclick="javascript:window.close();"><%=s_lang_cancel %></a>
	</td></tr>
</table>
</FORM>
</HTML>

<%
set ads = nothing
%>
