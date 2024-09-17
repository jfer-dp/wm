<!--#include file="passinc.asp" --> 
<!--#include file="language-2.asp" --> 

<%
if isadmin() = false then
	dim pfvl
	set pfvl = server.createobject("easymail.PubFolderViewLimit")
	pfvl.Load trim(request("iniid"))

	if pfvl.IsShow(Session("mail")) = false then
		set pfvl = nothing
		Response.Redirect "noadmin.asp"
	end if

	set pfvl = nothing
end if

dim ei
set ei = server.createobject("easymail.emmail")
ei.IsInPublicFolder = true
ei.PublicFolderName = trim(request("iniid"))
ei.needAddReadCount = true

'-----------------------------------------
filename = trim(request("filename"))
mode = trim(request("mode"))

dim pf
set pf = server.createobject("easymail.PubFolderManager")
pf.load trim(request("iniid"))

pf.GetItemInfoByName filename, gn_filename, gn_ownID, gn_nextstep, gn_postuser, gn_subject, gn_time, gn_length, gn_state, gn_searchkey, gn_readcount, gn_face, gn_istop

mystep = gn_nextstep
myistop = gn_istop
searchkey = gn_searchkey
face = gn_face

gn_filename = NULL
gn_ownID = NULL
gn_nextstep = NULL
gn_postuser = NULL
gn_subject = NULL
gn_time = NULL
gn_length = NULL
gn_state = NULL
gn_searchkey = NULL
gn_readcount = NULL
gn_face = NULL
gn_istop = NULL

permission = pf.Permission
isok = false

if isadmin() = false and LCase(pf.admin) <> LCase(Session("wem")) then
	if (permission = 0 or permission = 1) and Session("mail") = pf.GetPostName(filename) then
		isok = true
	end if
else
	isok = true
end if

if mode = "del" then
	isok = false

	if isadmin() = false and LCase(pf.admin) <> LCase(Session("wem")) then
		if permission = 1 and Session("mail") = pf.GetPostName(filename) then
			isok = true
		end if
	else
		isok = true
	end if

	if isok = true then
		pf.RemoveItem filename
	end if

	set pf = nothing
	set ei = nothing

	gourl = trim(request("gourl"))

	if gourl <> "" then
		response.redirect gourl & "&" & getGRSN()
	else
		response.redirect "showallpf.asp?" & getGRSN()
	end if
elseif mode = "settop" then
	if isadmin() = true or LCase(pf.admin) = LCase(Session("wem")) then
		if isok = true then
			if pf.TopItem(filename, true) = true then
				myistop = true
			else
				myistop = false
			end if
		end if
	end if
elseif mode = "nosettop" then
	if isadmin() = true or LCase(pf.admin) = LCase(Session("wem")) then
		if isok = true then
			if pf.TopItem(filename, false) = false then
				myistop = true
			else
				myistop = false
			end if
		end if
	end if
elseif mode = "delall" then
	if isadmin() = true or LCase(pf.admin) = LCase(Session("wem")) then
		if isok = true then
			isok = pf.RemoveParentItemAndAllChildItems(filename)

			if isok = true then
				set pf = nothing
				set ei = nothing

				gourl = trim(request("gourl"))

				if gourl <> "" then
					response.redirect gourl & "&" & getGRSN()
				else
					response.redirect "showallpf.asp?" & getGRSN()
				end if
			end if
		end if
	end if
end if


pt = trim(request("pt"))
subismessage = false

if pt <> "" then
	bd = trim(request("bd"))

	if bd <> "" then
		ei.LoadAll2 Session("wem"), filename, CDbl(pt), bd
	else
		ei.LoadAll1 Session("wem"), filename, CDbl(pt)
	end if

	subismessage = true
else
	ei.LoadAll Session("wem"), filename
end if

charset = UCase(ei.Text_CharSet)
if charset = "" or charset = "DEFAULT_CHARSET" then
	charset = b_lang_103
end if

Response.Cookies("name") = Session("wem")

dim userweb
set userweb = server.createobject("easymail.UserWeb")
userweb.Load Session("wem")

enableAutoAdaptCharSet = userweb.enableAutoAdaptCharSet
EnableShowHtmlMail = userweb.EnableShowHtmlMail

set userweb = nothing

dim allnum
allnum = 0
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html<%
if enableAutoAdaptCharSet = true then
	Response.Write "; charset=" & charset
else
	Response.Write "; charset=" & b_lang_103
end if
%>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
<link rel="stylesheet" type="text/css" href="images/showpfmail.css">

<STYLE type=text/css>
<!--
body {margin-top:23px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript" src="images/mglobal.js"></script>
<script type="text/javascript" src="images/jquery.min.js"></script>

<script type="text/javascript">
<!-- 
if (ie == 6)
	document.execCommand("BackgroundImageCache", false, true); 

var pdeladd;
var isshow = true;

function back() {
<% if trim(request("gourl")) = "" then %>
	history.back();
<% else %>
	location.href = "<%=trim(request("gourl")) %>&<%=getGRSN() %>";
<% end if %>
}

function saveatt(anum) {
	location.href = "mail2att.asp?<%=getGRSN() %>&mode=post&filename=<%=filename %><%
if pt <> "" then
	Response.Write "&pt=" & pt
end if

if bd <> "" then
	Response.Write "&bd=" & Server.URLEncode(bd)
end if
%>&attnum=" + anum;
}

function add2ads(vname, vemail) {
	var post_date = "setmode=3&<%=getGRSN() %>&cname=" + escape(vname) + "&email=" + escape(vemail);

$.ajax({
	type:"POST",
	url:"showmail.asp",
	data:post_date,
	success:function(data){
		alert_msg("<%=b_lang_104 %>");
	},
	error:function(){
		alert_msg("<%=b_lang_105 %>");
	}
});
}

function add2kill(vemail) {
	var post_date = "setmode=4&<%=getGRSN() %>&kill=" + escape(vemail);

$.ajax({
	type:"POST",
	url:"showmail.asp",
	data:post_date,
	success:function(data){
		alert_msg("<%=b_lang_104 %>");
	},
	error:function(){
		alert_msg("<%=b_lang_105 %>");
	}
});
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

<%
if isok = true then
%>
function delit() {
	if (confirm("<%=b_lang_036 %>") == false)
		return ;

    location.href = "showpfmail.asp?mode=del&iniid=<%=trim(request("iniid")) %>&filename=<%= filename & "&" & getGRSN()%>&gourl=<%=Server.URLEncode(trim(request("gourl"))) %>";
}
<%
	if isadmin() = true or LCase(pf.admin) = LCase(Session("wem")) then
		if mystep = 0 then
%>
function delall() {
	if (confirm("<%=b_lang_106 %>") == false)
		return ;

    location.href = "showpfmail.asp?mode=delall&iniid=<%=trim(request("iniid")) %>&filename=<%= filename & "&" & getGRSN()%>&gourl=<%=Server.URLEncode(trim(request("gourl"))) %>";
}
<%
		end if
	end if
end if
%>

function m_over(tag_obj) {
	tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj) {
	tag_obj.style.backgroundColor = "white";
}

function iFrameHeight() {
	var ifm= document.getElementById("iframepage");
	var subWeb = document.frames ? document.frames["iframepage"].document : ifm.contentDocument;
	if(ifm != null && subWeb != null) {
		ifm.height = subWeb.body.scrollHeight;
	}
}

function alert_msg(amsg) {
	$("#top_show_msg").text(amsg);
	document.getElementById("top_show_msg").style.display = "inline";
	setTimeout("close_alert()", 3000);
}

function close_alert() {
	document.getElementById("top_show_msg").style.display = "none";
}
// -->
</script>

<BODY>
<a name="gotop" style="font-size:0pt; height:0px;"></a>
<table id="table_main" class="table_main" align="center" cellspacing="0" cellpadding="0">
	<tr><td class="block_top_td"><div class="table_min_width"></div></td></tr>
<%
if subismessage = false then
%>
	<tr><td class="tool_top_td">

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:back()"><< <%=s_lang_return %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

<%
if isok = true then
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=post&pid=<%=trim(request("pid")) %>&iniid=<%=trim(request("iniid")) & "&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl"))) %>"><%=b_lang_107 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=editpost&face=<%=face %>&searchkey=<%=Server.URLEncode(searchkey) %>&iniid=<%=trim(request("iniid")) %>&filename=<%= filename & "&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl"))) %>"><%=b_lang_047 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

<%
	if mystep = 0 then
		if myistop = true then
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:delit()"><%=s_lang_del %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
			if isadmin() = true or LCase(pf.admin) = LCase(Session("wem")) then
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:delall()"><%=b_lang_108 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="showpfmail.asp?mode=nosettop&iniid=<%=trim(request("iniid")) %>&filename=<%= filename & "&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl"))) %>"><%=b_lang_109 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
			end if
		else
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:delit()"><%=s_lang_del %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
			if isadmin() = true or LCase(pf.admin) = LCase(Session("wem")) then
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:delall()"><%=b_lang_108 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="showpfmail.asp?mode=settop&iniid=<%=trim(request("iniid")) %>&filename=<%= filename & "&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl"))) %>"><%=b_lang_110 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
			end if
		end if
	else
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:delit()"><%=s_lang_del %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
	end if
else
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=post&pid=<%=trim(request("pid")) %>&iniid=<%=trim(request("iniid")) & "&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl"))) %>"><%=b_lang_107 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
end if
%>
	</td></tr>
<%
end if
%>
	<tr class="head_tr"><td id="subject_td" class="head_subject_td" style="border-top:1px solid #aac1de;">
	<span class="head_subject_span"><%=server.htmlencode(ei.subject) %></span>
	</td></tr>

	<tr class="head_tr"><td class="head_td">
	<span style="float:left;"><%=b_lang_111 %><%=s_lang_mh %></span>
	<span style="float:left;"><%
if ei.FromName = ei.FromMail then
	receiver = "<font color='#5fa207' style='font-weight:bold;'>" & server.htmlencode(ei.FromMail) & "</font>"
else
	receiver = "<font color='#5fa207' style='font-weight:bold;'>" & server.htmlencode(ei.FromName) & "</font>" & server.htmlencode(" <" & ei.FromMail & ">")
end if

dim item
item = server.htmlencode(ei.FromName)
item = replace(item, "'", "")
item = replace(item, """", "")

rec_email = ei.FromMail
rec_email = replace(rec_email, "'", "")
rec_email = replace(rec_email, """", "")

is_own = false
if LCase(ei.FromMail) = LCase(Session("mail")) then
	Response.Write "<img src='images/null.gif' class='head_own'>&nbsp;"

	is_own = true
end if

Response.Write receiver

if InStr(ei.FromMail, "@") then
	if is_own = false then
		Response.Write "&nbsp;&nbsp;&nbsp;[<a href='javascript:add2ads(""" & item & """,""" & rec_email & """)'>" & b_lang_112 & "</a>&nbsp;&nbsp;<a href='javascript:add2kill(""" & rec_email & """)'>" & b_lang_113 & "</a>]"
	end if
end if
%></span>
	</td></tr>

	<tr class="head_tr"><td class="head_td">
	<span style="float:left;"><%=b_lang_114 %><%=s_lang_mh %><%=ei.Time %></span>
<%
if ei.IsHtmlMail = false or EnableShowHtmlMail = false then
%>
	<span style="float:right;"><%=b_lang_045 %><%=s_lang_mh %>[<a href="javascript:doZoom(16)"><%=b_lang_115 %></a> <a href="javascript:doZoom(14)"><%=b_lang_116 %></a> <a href="javascript:doZoom(12)"><%=b_lang_117 %></a>]</span>
<%
end if
%>
	</td></tr>

	<tr class="head_tr"><td class="head_td">
	<span style="float:left;"><%=b_lang_118 %><%=s_lang_mh %><font color="black"><%=getShowSize(ei.Size) %></font></span>
	<span id="go_att"></span>
	</td></tr>

	<tr class="head_tr"><td class="head_td">
	<span style="float:left;"><%=b_lang_119 %><%=s_lang_mh %><%=server.htmlencode(searchkey) %></span>
	</td></tr>
<%
if ei.IsHtmlMail = true then
%>
	<tr class="head_tr"><td class="head_td">
	<span style="float:left;"><%=b_lang_120 %><%=s_lang_mh %><%
	Response.Write "<a href=""showatt.asp?ishtml=1&mode=post&filename=" & filename & "&count=0&pt=" & pt & "&" & getGRSN() & """ target='_blank'>" & b_lang_121 & "</a>"
%></span>
	</td></tr>

<%
	if EnableShowHtmlMail = true then
%>
	<tr bgcolor="white"><td class="iframe_td">
<iframe src="<%="showatt.asp?ishtml=1&mode=post&filename=" & filename & "&count=0&pt=" & pt & "&" & getGRSN() %>" id="iframepage" name="iframepage" frameBorder=0 scrolling=no width="100%" onLoad="iFrameHeight()"></iframe>
	</td></tr>
<%
	end if
end if

if ei.IsHtmlMail = false or EnableShowHtmlMail = false then
%>
	<tr bgcolor="white">
    <td id="zoom" class="zoom_td">
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
	t = server.htmlencode(ei.Text)

	if Len(t) < 100000 then
		t = ei.ConvText2Html(t)
	end if

	t = replace(RemoveEndRN(t), Chr(10), "<br>")
	t = replace(t, Chr(32) & Chr(32), "&nbsp;&nbsp;")
	t = replace(t, Chr(9), "&nbsp;&nbsp;&nbsp;&nbsp;")
end if

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
	<tr><td>
<table width="100%" border="0" cellspacing="0" align="center" class='table_att_in'>
<tr><td class='att_in_td'><img src='images/attach.gif' border='0' align='absmiddle' style='padding-right:8px;'><%=b_lang_122 %></td></tr>
<%
if ei.IsHtmlMail = true then
	i = 1

	do while i < allnum
		Response.Write "<tr style='cursor:default;' onmouseover='m_over(this);' onmouseout='m_out(this);'><td class='att_td'><span class='att_span'>"
		if ei.GetAttachmentName(i) = "" then
		    Response.Write "<a class='att' href=""showatt.asp?mode=post&filename=" & filename & "&count=" & i & "&" & getGRSN() & """ target='_blank'>" & "html" & "</a>"
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<a href=""showatt.asp?isdown=1&mode=post&filename=" & filename & "&count=" & i & "&" & getGRSN() & """ target='_blank'>" & b_lang_123 & "</a>]"
			Response.Write "&nbsp;&nbsp;[<a href=""JavaScript:saveatt(" & i & ")"")>" & b_lang_124 & "</a>]"
		else
			if ei.AttachmentIsMessage(i) = false then
				Response.Write "<a class='att' href=""showatt.asp?mode=post&filename=" & filename & "&count=" & i & "&pt=" & pt & "&" & getGRSN() & """ target='_blank'>" & server.htmlencode(ei.GetAttachmentName(i)) & "</a>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<a href=""showatt.asp?isdown=1&mode=post&filename=" & filename & "&count=" & i & "&pt=" & pt & "&" & getGRSN() & """ target='_blank'>" & b_lang_123 & "</a>]"
				Response.Write "&nbsp;&nbsp;[<a href=""JavaScript:saveatt(" & i & ")"")>" & b_lang_124 & "</a>]"
			else
				Response.Write "<a class='att' href=""showpfmail.asp?filename=" & filename & "&count=" & i & "&pt=" & ei.GetAttachmentPT(i) & "&bd=" & Server.URLEncode(ei.GetEmlAttachmentBD(i)) & "&" & getGRSN() & """ target='_blank'>" & server.htmlencode(ei.GetAttachmentName(i)) & "</a>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<a href=""showatt.asp?isdown=1&mode=post&filename=" & filename & "&count=" & i & "&pt=" & pt & "&" & getGRSN() & """ target='_blank'>" & b_lang_123 & "</a>]"
				Response.Write "&nbsp;&nbsp;[<a href=""JavaScript:saveatt(" & i & ")"")>" & b_lang_124 & "</a>]"
			end if
		end if
		Response.Write "<span></td></tr>" & Chr(13)

	    i = i + 1
	loop
else
	do while i < allnum
		Response.Write "<tr style='cursor:default;' onmouseover='m_over(this);' onmouseout='m_out(this);'><td class='att_td'><span class='att_span'>"
		if ei.GetAttachmentName(i) = "" then
			Response.Write "<a class='att' href=""showatt.asp?mode=post&filename=" & filename & "&count=" & i & "&" & getGRSN() & """ target='_blank'>" & "html" & "</a>"
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<a href=""showatt.asp?isdown=1&mode=post&filename=" & filename & "&count=" & i & "&" & getGRSN() & """ target='_blank'>" & b_lang_123 & "</a>]"
			Response.Write "&nbsp;&nbsp;[<a href=""JavaScript:saveatt(" & i & ")"")>" & b_lang_124 & "</a>]"
		else
			if ei.AttachmentIsMessage(i) = false then
		    	Response.Write "<a class='att' href=""showatt.asp?mode=post&filename=" & filename & "&count=" & i & "&pt=" & pt & "&" & getGRSN() & """ target='_blank'>" & server.htmlencode(ei.GetAttachmentName(i)) & "</a>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<a href=""showatt.asp?isdown=1&mode=post&filename=" & filename & "&count=" & i & "&pt=" & pt & "&" & getGRSN() & """ target='_blank'>" & b_lang_123 & "</a>]"
				Response.Write "&nbsp;&nbsp;[<a href=""JavaScript:saveatt(" & i & ")"")>" & b_lang_124 & "</a>]"
			else
				Response.Write i+1 & ".<a href=""showpfmail.asp?filename=" & filename & "&count=" & i & "&pt=" & ei.GetAttachmentPT(i) & "&bd=" & Server.URLEncode(ei.GetEmlAttachmentBD(i)) & "&" & getGRSN() & """ target='_blank'>" & server.htmlencode(ei.GetAttachmentName(i)) & "</a>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<a href=""showatt.asp?isdown=1&mode=post&filename=" & filename & "&count=" & i & "&pt=" & pt & "&" & getGRSN() & """ target='_blank'>" & b_lang_123 & "</a>]"
				Response.Write "&nbsp;&nbsp;[<a href=""JavaScript:saveatt(" & i & ")"")>" & b_lang_124 & "</a>]"
			end if
		end if
		Response.Write "<span></td></tr>" & Chr(13)

	    i = i + 1
	loop
end if
%>
</table>
	</td></tr>
<%
end if

if subismessage = false then
%>
	<tr><td class="block_top_td" style="height:12px;"></td></tr>
	<tr><td class="tool_top_td">

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:back()"><< <%=s_lang_return %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

<%
if isok = true then
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=post&pid=<%=trim(request("pid")) %>&iniid=<%=trim(request("iniid")) & "&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl"))) %>"><%=b_lang_107 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=editpost&face=<%=face %>&searchkey=<%=Server.URLEncode(searchkey) %>&iniid=<%=trim(request("iniid")) %>&filename=<%= filename & "&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl"))) %>"><%=b_lang_047 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

<%
	if mystep = 0 then
		if myistop = true then
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:delit()"><%=s_lang_del %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
			if isadmin() = true or LCase(pf.admin) = LCase(Session("wem")) then
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:delall()"><%=b_lang_108 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="showpfmail.asp?mode=nosettop&iniid=<%=trim(request("iniid")) %>&filename=<%= filename & "&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl"))) %>"><%=b_lang_109 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
			end if
		else
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:delit()"><%=s_lang_del %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
			if isadmin() = true or LCase(pf.admin) = LCase(Session("wem")) then
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:delall()"><%=b_lang_108 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>

	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="showpfmail.asp?mode=settop&iniid=<%=trim(request("iniid")) %>&filename=<%= filename & "&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl"))) %>"><%=b_lang_110 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
			end if
		end if
	else
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="javascript:delit()"><%=s_lang_del %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
	end if
else
%>
	<span class="st_span">
	<a class='wwm_btnDownload btn_gray' href="wframe.asp?mode=post&pid=<%=trim(request("pid")) %>&iniid=<%=trim(request("iniid")) & "&" & getGRSN() & "&gourl=" & Server.URLEncode(trim(request("gourl"))) %>"><%=b_lang_107 %></a>
	</span>
	<span style='float:left; width:3px;'>&nbsp;</span>
<%
end if
%>
	</td></tr>
<%
end if
%>

	<tr><td class="block_top_td" style="height:16px;"></td></tr>
	<tr><td align="right">
	<span style="margin-right:16px;"><a href="#gotop"><img src='images/gotop.gif' border='0' title="<%=b_lang_125 %>"></a></span>
	</td></tr>
</table>
<%
if allnum > 0 then
%>
<a name="goatt">&nbsp;</a>
<script type="text/javascript">
<!--
document.getElementById("go_att").innerHTML = "&nbsp;<a href='#goatt'><img src='images/atta.gif' border='0' title='<%=b_lang_126 %>' align='absmiddle'></a>";
//-->
</script>
<%
end if
%>
<div id="top_show_msg" class="top_show_msg"></div>
</BODY>
</HTML>

<%
set pf = nothing
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

function getShowSize(bytesize)
	if bytesize < 1000 then
		getShowSize = "1K"
	else
		if bytesize < 1000000 then
			tmpSize = CDbl(bytesize/1000)
			tmpindex = InStr(1, tmpSize, ".")
			if tmpindex = 0 then
				getShowSize = tmpSize & "K"
			else
				getShowSize = CDbl(Left(tmpSize, tmpindex + 1)) & "K"
			end if
		else
			tmpSize = CStr(CDbl(bytesize/1000000))
			tmpindex = InStr(1, tmpSize, ".")
			if tmpindex = 0 then
				getShowSize = tmpSize & "M"
			else
				getShowSize = CDbl(Left(tmpSize, tmpindex + 2)) & "M"
			end if
		end if
	end if
end function
%>
