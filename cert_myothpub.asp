<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
Session("cert_imp_type") = ""
Session("cert_imp_pw") = ""
Session("cert_imp_save_day") = ""

dim wemcert
set wemcert = server.createobject("easymail.WebEasyMailCert")
wemcert.Load Session("wem"), Session("mail")

allnum = wemcert.PubCertCount

if trim(request("page")) = "" then
	page = 0
else
	page = CInt(request("page"))
end if

if page < 0 then
	page = 0
end if

allpage = CInt((allnum - (allnum mod pageline))/ pageline)

if allnum mod pageline <> 0 then
	allpage = allpage + 1
end if

if page >= allpage then
	page = allpage - 1
end if

if page < 0 then
	page = 0
end if

if allpage = 0 then
	allpage = 1
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
.title_tr {white-space:nowrap; background:#f2f4f6; height:24px;}
.st_l, .st_r {height:24px; text-align:center; white-space:nowrap; border-left:1px solid #A5B6C8; border-top:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8;}
.st_r {border-right:1px solid #A5B6C8;}
.cont_tr {background:white; height:24px;}
.cont_td {height:24px; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px;}
.cont_td_word {height:24px; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px; word-break:break-all; word-wrap:break-word;}
-->
</STYLE>
</head>

<script type="text/javascript">
<!--
function mdel()
{
	if (ischeck() == true)
	{
		if (confirm("<%=a_lang_065 %>") == false)
			return ;

		document.form1.action = "cert_muldel.asp";
		document.form1.submit();
	}
}

function allcheck_onclick() {
	if (document.form1.allcheck.checked == true)
		checkall(true);
	else
		checkall(false);
}

function checkall(check) {
	var theall = <%
if allnum > pageline then
	if page > 0 then
		Response.Write allnum - page * pageline
	else
		Response.Write allnum
	end if
else
	Response.Write allnum
end if
%>;
	var i = theall - <%=pageline %> - 1;
	if (i < 0)
		i = 0;

	var theObj;

	for(; i<theall; i++)
	{
		theObj = eval("document.form1.check" + i);

		if (theObj != null)
			theObj.checked = check;
	}
}

function ischeck() {
	var theall = <%
if allnum > pageline then
	if page > 0 then
		Response.Write allnum - page * pageline
	else
		Response.Write allnum
	end if
else
	Response.Write allnum
end if
%>;
	var i = theall - <%=pageline %> - 1;
	if (i < 0)
		i = 0;

	var theObj;

	for(; i<theall; i++)
	{
		theObj = eval("document.form1.check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}

function selectpage_onchange()
{
	location.href = "cert_myothpub.asp?<%=getGRSN() %>&page=" + document.form1.page.value;
}

function delall(){
	if (confirm("<%=a_lang_066 %>") == false)
		return ;

	location.href = "cert_del.asp?<%=getGRSN() %>&delmode=allpub&retstr=cert_myothpub.asp";
}

function m_over(tag_obj) {
	tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj) {
	tag_obj.style.backgroundColor = "white";
}
//-->
</SCRIPT>

<BODY>
<form action="cert_myothpub.asp" method=post id=form1 name=form1>
<input type="hidden" name="mdel">
<input type="hidden" name="thispage" value="<%=page %>">

<table width="90%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr>
	<td align="left" height="28" width="55%" nowrap style="padding-left:4px;">
<a class='wwm_btnDownload btn_blue' href="cert_index.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="cert_imp.asp?<%=getGRSN() %>&im=pub"><%=a_lang_067 %></a>
<a class='wwm_btnDownload btn_blue' href="cert_pubstradd.asp?<%=getGRSN() %>&page=<%=page %>"><%=a_lang_068 %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:mdel();"><%=s_lang_del %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:delall();"><%=a_lang_069 %></a>
	</td>
	<td align="left" width="20%" nowrap>
<%
if page > 0 then
	Response.Write "<a href=""cert_myothpub.asp?" & getGRSN() & "&page=" & page - 1 & """><img src='images\prep.gif' border='0' align='absmiddle'></a>&nbsp;"
else
	Response.Write "<img src='images\gprep.gif' border='0' align='absmiddle'>&nbsp;"
end if
%>
<select name="page" class="drpdwn" size="1" LANGUAGE=javascript onchange="selectpage_onchange()">
<%
i = 0

do while i < allpage
	if i <> page then
		Response.Write "<option value=""" & i & """>" & i + 1 & "</option>"
	else
		Response.Write "<option value=""" & i & """ selected>" & i + 1 & "</option>"
	end if
	i = i + 1
loop
%></select>
<%
if page < allpage - 1 then
	Response.Write "&nbsp;<a href=""cert_myothpub.asp?" & getGRSN() & "&page=" & page + 1 & """><img src='images\nextp.gif' border='0' align='absmiddle'></a>"
else
	Response.Write "&nbsp;<img src='images\gnextp.gif' border='0' align='absmiddle'>"
end if
%>
	</td>
	<td align="right" width="25%" nowrap style="padding-right:8px; color:#444444;"><%=a_lang_070 %></td>
	</tr>
</table>
<br>

<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr class="title_tr">
    <td width="5%" height="24" class="st_l"><input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()"></td>
    <td width="6%" class="st_l"><%=a_lang_071 %></td>
    <td width="45%" class="st_l"><%=a_lang_072 %></td>
    <td width="30%" class="st_l"><%=a_lang_073 %></td>
    <td width="7%" class="st_l"><%=a_lang_074 %></td>
    <td width="7%" class="st_r"><%=s_lang_del %></td>
	</tr>
<%
i = page * pageline
li = 0

do while i < allnum and li < pageline
	si = allnum - i - 1
	wemcert.GetPubCertInfo si, uid, email

	Response.Write "<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);'>"
	Response.Write "	<td align='center' height='25' class='cont_td'><input type='checkbox' name='check" & si & "' value='" & server.htmlencode(email) & "'></td>"
	Response.Write "	<td align='center' class='cont_td'>" & i + 1 & "</td>"

	Response.Write "	<td align='left' class='cont_td_word'><a href='cert_showcert.asp?email=" & Server.URLEncode(email) & "&page=" & page & "&" & getGRSN() & "'>" & server.htmlencode(uid) & "</a></td>"
	Response.Write "	<td align='left' class='cont_td_word'>" & server.htmlencode(email) & "</td>"
	Response.Write "	<td align='center' class='cont_td'><a href='cert_exp.asp?" & getGRSN() & "&mode=pub&pub_email=" & Server.URLEncode(email) & "' target='_blank'>" & a_lang_074 & "</a></td>"
	Response.Write "	<td align='center' class='cont_td'><a href='cert_del.asp?delmode=pub&pub_email=" & Server.URLEncode(email) & "&retstr=" & Server.URLEncode("cert_myothpub.asp?page=" & page) & "&" & getGRSN() & "'><img src='images\del.gif' border='0' alt='" & s_lang_del & "'></a></td>"
	Response.Write "</tr>" & Chr(13)

	uid = NULL
	email = NULL

    i = i + 1
    li = li + 1
loop

%>
</table>
</FORM>
</BODY>
</HTML>

<%
set wemcert = nothing
%>
