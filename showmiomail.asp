<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
if isadmin() = false then
	Response.Redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.emmail")

filename = trim(request("filename"))
user = trim(request("user"))
inout = trim(request("inout"))
gourl = trim(request("gourl"))
pt = trim(request("pt"))
bd = trim(request("bd"))

subismessage = false

if pt = "" then
	pt = "0"
else
	subismessage = true
end if

if inout = "in" then
	ei.LoadAll_MonInMail user, filename, CDbl(pt), bd
else
	ei.LoadAll_MonOutMail user, filename, CDbl(pt), bd
	inout = "out"
end if

charset = UCase(ei.Text_CharSet)
if charset = "" or charset = "DEFAULT_CHARSET" then
	charset = s_lang_0553
end if

dim userweb
set userweb = server.createobject("easymail.UserWeb")
userweb.Load Session("wem")

enableAutoAdaptCharSet = userweb.enableAutoAdaptCharSet
EnableShowHtmlMail = userweb.EnableShowHtmlMail

set userweb = nothing

dim allnum
allnum = 0

dim app_url
app_url = "&user=" & Server.URLEncode(user) & "&inout=" & inout
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html<%
if enableAutoAdaptCharSet = true then
	Response.Write "; charset=" & charset
else
	Response.Write "; charset=" & s_lang_0553
end if
%>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
</HEAD>

<script type="text/javascript" src="images/mglobal.js"></script>

<script type="text/javascript">
<!-- 
var isshow = false;

function window_onload() {
	if (ie != false)
		document.body.focus();

	hide_rads();
}

function back() {
<% if gourl = "" then %>
	history.back();
<% else %>
	location.href = "<%=gourl %>";
<% end if %>
}

function delthis() {
	location.href="mulmonmail.asp?filename=<%=Server.URLEncode(filename) %>&user=<%=Server.URLEncode(user) %>&inout=<%=inout %>&<%=getGRSN() %>&gourl=<%=Server.URLEncode(gourl) %>";
}

function headmessage() {
	var theObj;
	theObj = document.getElementById("headMessage");

	if (isshow == false)
	{
		var instr = "<%
		ht = server.htmlencode(ei.HeadMessage)
		ht = replace(ht, Chr(13), "<br>\")
		ht = replace(ht, Chr(32), "&nbsp;")
		ht = replace(ht, Chr(9), "&nbsp;&nbsp;&nbsp;&nbsp;")

		Response.Write ht %>";

		instr = "<table width='95%' align='center' border='0' bgcolor='#DBEAF5' cellspacing='0' style='border:1px solid #8CA5B5; word-break:break-all; word-wrap:break-word;'><tr><td>" + instr + "</td></tr></table><br>"
		theObj.innerHTML = instr;
		isshow = true;
	}
	else
	{
		theObj.innerHTML = "";
		isshow = false;
	}
}

function doZoom(size){
<%
if ei.IsHtmlMail = false or EnableShowHtmlMail = false then
%>
	document.getElementById('zoom').style.fontSize=size+'px'
<%
end if
%>
}

var rads_is_show = <%
if inout = "out" then
	Response.Write "true"
else
	Response.Write "false"
end if
%>;
function hide_rads()
{
	var Stag = document.getElementById("rads_showstr");
	if (rads_is_show == true)
	{
		Stag.innerHTML = "<%=s_lang_0554 %>";
		rads_function_div.style.display = "inline";
	}
	else
	{
		Stag.innerHTML = "<%=s_lang_0555 %>";
		rads_function_div.style.display = "none";
	}

	rads_is_show = !rads_is_show;
}

function iFrameHeight() {
	var ifm= document.getElementById("iframepage");
	var subWeb = document.frames ? document.frames["iframepage"].document : ifm.contentDocument;
	if(ifm != null && subWeb != null) {
		ifm.height = subWeb.body.scrollHeight;
	}
}
// -->
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
<form name="f1">
<table width="95%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:24px;">
  <tr>
	<td height="20" align="center" colspan="2">
	<a href="javascript:headmessage()"><%=s_lang_0503 %></a>
<%
if subismessage = false then
%>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="javascript:delthis()"><%=s_lang_del %></a>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="javascript:back()"><%=s_lang_return %></a>
<%
end if
%>
	</td>
  </tr>
</table>
<br>
<span id="headMessage"></span>
<table width="95%" border="0" bgColor="#93BEE2" align="center" cellspacing="0" style='border:1px solid #336699;'>
  <tr> 
    <td width="17%" height="20">
      <div align="right"><font class="s" color="#104A7B"><b><%=s_lang_0128 %></b><%=s_lang_mh %></font></div>
    </td>
    <td width="83%" bgcolor="#93BEE2" style='border-bottom:1px solid #336699;'><table width="100%" border="0" cellspacing="0"><tr><td><%=ei.Time %></td><td><div align="right">[<a href="javascript:hide_rads()"><span id="rads_showstr"></span></a>]&nbsp;&nbsp;<span id="go_att"></span><font class="s" color="#104A7B"><b><%=s_lang_0483 %>:</b></font>[<a href="javascript:doZoom(16)"><%=s_lang_0484 %></a> <a href="javascript:doZoom(14)"><%=s_lang_0485 %></a> <a href="javascript:doZoom(12)"><%=s_lang_0486 %></a>]</div></td></tr></table>
  </tr>
  <tr>
    <td width="17%" height="20"> 
      <div align="right"><font class="s" color="#104A7B"><b><%=s_lang_0501 %></b><%=s_lang_mh %></font></div>
    </td>
    <td width="83%" bgcolor="#93BEE2" style='border-bottom:1px solid #336699;'><%
xmsp = ei.XMSMailPriority

if xmsp = "High" then
	Response.Write "<font color='#901111'>" & s_lang_0130 & "</font>"
elseif xmsp = "Low" then
	Response.Write s_lang_0131
else
	Response.Write s_lang_0146
end if
%>&nbsp;</td>
  </tr>
<%
if inout = "out" then
%>
  <tr>
    <td width="17%" height="20">
      <div align="right"><font class="s" color="#104A7B"><b><%=s_lang_0556 %></b><%=s_lang_mh %></font></div>
    </td>
    <td width="83%" bgcolor="#93BEE2" style='border-bottom:1px solid #336699;'><b><%=ei.SendNumber %></b>&nbsp;</td>
  </tr>
<%
end if
%>
  <tr>
    <td width="17%" height="20">
      <div align="right"><font class="s" color="#104A7B"><b><%=s_lang_0147 %></b><%=s_lang_mh %></font></div>
    </td>
    <td width="83%" bgcolor="#93BEE2" style='border-bottom:1px solid #336699;'><%=server.htmlencode(ei.FromName) %>&nbsp;</td>
  </tr>
  <tr> 
    <td width="17%" height="20" nowrap>
      <div align="right"><font class="s" color="#104A7B"><b><%=s_lang_0148 %></b><%=s_lang_mh %></font></div>
    </td>
    <td width="83%" bgcolor="#93BEE2" style='border-bottom:1px solid #336699;'><%
receiver = server.htmlencode(ei.FromMail)
receiver = replace(receiver, "'", "")
receiver = replace(receiver, """", "")
Response.Write receiver
%>&nbsp;</td>
  </tr>
  <tr><td noWrap colspan="2" align="right" width="100%">
<div id="rads_function_div">
<table width="100%" border="0" cellspacing="0"><tr>
    <td width="17%" height="20">
      <div align="right"><font class="s" color="#104A7B"><b><span id="cal_mode"><%=s_lang_0557 %></span></b><%=s_lang_mh %></font></div>
    </td>
    <td width="83%" align="left" bgcolor="#93BEE2" style='border-bottom:1px solid #336699;'><span id="cal_msg"><%
allnum = ei.to_count
i = 0
first_show = true

do while i < allnum
	ei.GetToAds i, ret_name, ret_email

	if InStr(ret_email, "@") then
		if first_show = false then
			Response.Write "<br>"
		else
			first_show = false
		end if

		Response.Write server.htmlencode(ret_name & " <" & ret_email & ">")
	end if

	ret_name = NULL
	ret_email = NULL

	i = i + 1
loop
%></span>&nbsp;</td>
  </tr>
<%
allnum = ei.cc_count
if allnum > 0 then
%>
  <tr> 
    <td width="17%" height="20">
      <div align="right"><font class="s" color="#104A7B"><b><%=s_lang_0502 %></b><%=s_lang_mh %></font></div>
    </td>
    <td width="83%" align="left" bgcolor="#93BEE2" style='border-bottom:1px solid #336699;'><%
i = 0
first_show = true

do while i < allnum
	ei.GetCcAds i, ret_name, ret_email

	if InStr(ret_email, "@") then
		if first_show = false then
			Response.Write "<br>"
		else
			first_show = false
		end if

		Response.Write server.htmlencode(ret_name & " <" & ret_email & ">")
	end if

	ret_name = NULL
	ret_email = NULL

	i = i + 1
loop
%>&nbsp;</td></tr>
<%
end if
%>
</table>
</div>
  </td></tr>
  <tr> 
<%
if ei.IsHtmlMail = false then
%>
    <td width="17%" height="20" style='border-bottom:1px solid #336699;'>
<%else%>
    <td width="17%" height="20">
<%end if%>
      <div align="right"><font class="s" color="#104A7B"><b><%=s_lang_0127 %></b><%=s_lang_mh %></font></div>
    </td>
    <td width="83%" bgcolor="#93BEE2" style='border-bottom:1px solid #336699; word-break:break-all; word-wrap:break-word;'> <%=server.htmlencode(ei.subject) %>&nbsp;</td>
  </tr>
<%
i = 0

if ei.IsHtmlMail = true then
	i = 1
%>
  <tr> 
    <td width="17%" height="20" style='border-bottom:1px solid #336699;'>
      <div align="right"><font class="s" color="#104A7B"><b><%=s_lang_0495 %></b><%=s_lang_mh %></font></div>
    </td>
    <td width="83%" bgcolor="#93BEE2" style='border-bottom:1px solid #336699;'><font class="s" color="#104A7B"><b>
<%
    Response.Write "<a href=""showmioatt.asp?ishtml=1&filename=" & filename & "&count=0&pt=" & pt & "&" & getGRSN() & app_url & """ target='_blank'>" & s_lang_0496 & "</a>"
%>
</b></font></td>
  </tr>
<%
	if EnableShowHtmlMail = true then
%>
	<tr bgcolor="white"><td colspan="2">
<iframe src="<%="showmioatt.asp?ishtml=1&filename=" & filename & "&count=0&pt=" & pt & app_url & "&" & getGRSN() %>" id="iframepage" name="iframepage" frameBorder=0 scrolling=no width="100%" onLoad="iFrameHeight()"></iframe>
	</td></tr>
<%
	end if
end if

if ei.IsHtmlMail = false or EnableShowHtmlMail = false then
%>
  <tr bgcolor="#FFFFFF">
    <td colspan="2" id="zoom" style="word-break:break-all; word-wrap:break-word;">
<%
end if

isok = true

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
%>&nbsp;</td>
  </tr>
<%
end if

allnum = ei.AttachmentCount

if allnum = 1 and ei.IsHtmlMail = true then
	allnum = 0
end if

i = 0

if allnum > 0 then
%>
  <tr> 
    <td colspan="2" height="20" bgcolor="#DBEAF5" style='border-top:1px solid #336699; border-bottom:1px solid #336699;'>
      <div align="center"><font class="s" color="#104A7B"><b><%=s_lang_0505 %></b></font></div>
    </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td colspan="2"> 
<%
if ei.IsHtmlMail = true then
	i = 1

	do while i < allnum
		if ei.GetAttachmentName(i) = "" then
		    response.write i & ".<a href=""showmioatt.asp?filename=" & filename & "&count=" & i & app_url & "&" & getGRSN() & """ target='_blank'>" & "html" & "</a>"
			response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""showmioatt.asp?isdown=1&filename=" & filename & "&count=" & i & app_url & "&" & getGRSN() & """ target='_blank'><img src='images/downatt.gif' border='0' align='absmiddle' title='" & s_lang_0373 & "'></a><br>" & chr(13)
		else
			if ei.AttachmentIsMessage(i) = false then
		    	response.write i & ".<a href=""showmioatt.asp?filename=" & filename & "&count=" & i & "&pt=" & pt & app_url & "&" & getGRSN() & """ target='_blank'>" & ei.GetAttachmentName(i) & "</a>"
				response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""showmioatt.asp?isdown=1&filename=" & filename & "&count=" & i & "&pt=" & pt & app_url & "&" & getGRSN() & """ target='_blank'><img src='images/downatt.gif' border='0' align='absmiddle' title='" & s_lang_0373 & "'></a><br>" & chr(13)
			else
		    	response.write i & ".<a href=""showmiomail.asp?filename=" & filename & "&count=" & i & "&pt=" & ei.GetAttachmentPT(i) & "&bd=" & Server.URLEncode(ei.GetEmlAttachmentBD(i)) & app_url & "&" & getGRSN() & """ target='_blank'>" & ei.GetAttachmentName(i) & "</a>"
				response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""showmioatt.asp?isdown=1&filename=" & filename & "&count=" & i & "&pt=" & pt & app_url & "&" & getGRSN() & """ target='_blank'><img src='images/downatt.gif' border='0' align='absmiddle' title='" & s_lang_0373 & "'></a><br>" & chr(13)
			end if
		end if

	    i = i + 1
	loop
else
	do while i < allnum
		if ei.GetAttachmentName(i) = "" then
		    response.write i+1 & ".<a href=""showmioatt.asp?filename=" & filename & "&count=" & i & app_url & "&" & getGRSN() & """ target='_blank'>" & "html" & "</a>"
			response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""showmioatt.asp?isdown=1&filename=" & filename & "&count=" & i & app_url & "&" & getGRSN() & """ target='_blank'><img src='images/downatt.gif' border='0' align='absmiddle' title='" & s_lang_0373 & "'></a><br>" & chr(13)
		else
			if ei.AttachmentIsMessage(i) = false then
		    	response.write i+1 & ".<a href=""showmioatt.asp?filename=" & filename & "&count=" & i & "&pt=" & pt & app_url & "&" & getGRSN() & """ target='_blank'>" & ei.GetAttachmentName(i) & "</a>"
				response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""showmioatt.asp?isdown=1&filename=" & filename & "&count=" & i & "&pt=" & pt & app_url & "&" & getGRSN() & """ target='_blank'><img src='images/downatt.gif' border='0' align='absmiddle' title='" & s_lang_0373 & "'></a><br>" & chr(13)
			else
		    	response.write i+1 & ".<a href=""showmiomail.asp?filename=" & filename & "&count=" & i & "&pt=" & ei.GetAttachmentPT(i) & "&bd=" & Server.URLEncode(ei.GetEmlAttachmentBD(i)) & app_url & "&" & getGRSN() & """ target='_blank'>" & ei.GetAttachmentName(i) & "</a>"
				response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""showmioatt.asp?isdown=1&filename=" & filename & "&count=" & i & "&pt=" & pt & app_url & "&" & getGRSN() & """ target='_blank'><img src='images/downatt.gif' border='0' align='absmiddle' title='" & s_lang_0373 & "'></a><br>" & chr(13)
			end if
		end if

	    i = i + 1
	loop
end if
%> </td>
  </tr>
</table>
<%
end if
%>
<table width="95%" border="0"><tr><td>
<br><div align="right"><a href="#"><img src='images/gotop.gif' border='0' title="<%=s_lang_0152 %>"></a></div>
</td></tr></table>
</td></tr></table>
<%
if allnum > 0 then
%>
<a name="goatt"></a>
<script language="JavaScript">
<!--
document.getElementById("go_att").innerHTML = "<a href='#goatt'><img src='images/attach.gif' border='0' align='absmiddle'></a>&nbsp;&nbsp;";
//-->
</script>
<%
end if
%>
</form>
</BODY>
</HTML>

<%
set ei = nothing


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
%>
