<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" --> 

<%
if Application("em_EnableBBS") = false then
	Response.Redirect "noadmin.asp?errstr=" & Server.URLEncode(b_lang_087) & "&" & getGRSN()
end if

dim pf
set pf = server.createobject("easymail.PubFolderManager")

dim pfvl
set pfvl = server.createobject("easymail.PubFolderViewLimit")

allnum = pf.PubFoldersCount

if isadmin() = true then
	mode = trim(request("mode"))
	fileid = trim(request("fileid"))

	if mode <> "" and fileid <> "" then
		if mode = "up" then
			pf.UpPubFolder fileid
		elseif mode = "down" then
			pf.DownPubFolder fileid
		end if
	end if
end if

dim ei
set ei = Application("em")
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
.font_top_title {font-size:11pt; color:#104A7B; font-weight:bold;}

.title_tr {white-space:nowrap; background:#f2f4f6; height:24px;}
.st_left {text-align:center; white-space:nowrap; border-left:1px solid #A5B6C8; border-top:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8;}
.st_right {text-align:center; white-space:nowrap; border:1px solid #A5B6C8;}

.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.table_min_width {width:660px; font-size:0pt; height:0px; width:0px; border:0px;}
.cont_tr {background:white; height:24px;}
.cont_td {height:24px; border-bottom:1px solid #A5B6C8; padding-left:4px; padding-right:4px; white-space:nowrap;}
-->
</STYLE>
</HEAD>

<script type="text/javascript" src="images/sc_left.js"></script>

<script LANGUAGE=javascript>
<!--
function del(fid) {
	if (confirm("<%=b_lang_036 %>") == false)
		return ;

	location.href = "deletepf.asp?fileid=" + fid + "&<%=getGRSN() %>";
}

function m_over(tag_obj) {
	tag_obj.style.backgroundColor = "#ecf9ff";
}

function m_out(tag_obj) {
	tag_obj.style.backgroundColor = "white";
}
//-->
</script>

<body>
<table width="98%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_088 %>
</td></tr>
<tr><td class="block_top_td" style="height:12px; _height:14px;"></td></tr>
<tr><td align="center">

<table width="95%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr><td colspan="9" class="block_top_td"><div class="table_min_width"></div></td></tr>

	<tr class="title_tr">
    <td width="5%" height="25" class="st_left"><%=b_lang_040 %></td>
    <td width="37%" class="st_left"><%=b_lang_089 %></td>
    <td width="8%" class="st_left"><%=b_lang_090 %></td>
    <td width="16%" class="st_left"><%=b_lang_091 %></td>
    <td width="10%" class="st_left"><%=b_lang_092 %></td>
<%
if isadmin() = true then
%>
	<td width="5%" class="st_left"><%=b_lang_045 %></td>
	<td width="5%" class="st_left"><%=b_lang_047 %></td>
    <td width="8%" class="st_left"><%=b_lang_093 %></td>
    <td width="8%" class="st_right"><%=s_lang_del %></td>
<%
else
%>
	<td width="8%" class="st_left"><%=b_lang_045 %></td>
	<td width="8%" class="st_right"><%=b_lang_047 %></td>
<%
end if
%>
  </tr>
<%
i = 0

dim pfilename
dim admin
dim permission
dim name
dim createTime
dim count
dim maxid
dim maxitem
dim maxsize
vi = 0

do while i < allnum
	pf.GetFolderInfoByIndex i, pfilename, admin, permission, name, createTime, count, maxid, maxitem, maxsize

	canView = false

	if admin = Session("wem") or isadmin() = true then
		canView = true
		vi = vi + 1
	else
		pfvl.Load pfilename
		if pfvl.IsShow(Session("mail")) = true then
			canView = true
			vi = vi + 1
		end if
	end if

if canView = true then
	Response.Write "<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);'><td class='cont_td' align='center'>" & vi & "</td>"
	Response.Write "<td class='cont_td' align='left'><a href='showpf.asp?fileid=" & Server.URLEncode(pfilename) & "&" & getGRSN() & "'>" & server.htmlencode(name) & "</a>&nbsp;</td>"
	Response.Write "<td class='cont_td' align='center'>" & count & "</td>"
	Response.Write "<td class='cont_td' align='left'>" & ei.GetUserMail(admin) & "&nbsp;</td>"
	Response.Write "<td class='cont_td' align='center'>" & CLng(maxsize/1000) & "K</td>"

	if admin = Session("wem") or isadmin() = true then
		Response.Write "    <td align='center' class='cont_td'><a href='editpfpm.asp?fileid=" & pfilename & "&" & getGRSN() & "'>" & b_lang_094 & "</a></td>"
	else
		Response.Write "    <td align='center' class='cont_td'>&nbsp;</td>"
	end if

	if admin = Session("wem") or isadmin() = true then
		Response.Write "    <td align='center' class='cont_td'><a href='editpf.asp?fileid=" & pfilename & "&" & getGRSN() & "'><img src='images/edit.gif' border='0'></a></td>"
	else
		Response.Write "    <td align='center' class='cont_td'>&nbsp;</td>"
	end if

	if isadmin() = true then
		if allnum = 1 then
			Response.Write "    <td align='center' class='cont_td'>&nbsp;</td>"
		elseif i = 0 then
			Response.Write "    <td align='center' class='cont_td'>&nbsp;&nbsp;&nbsp;&nbsp;<a href='showallpf.asp?mode=down&fileid=" & Server.URLEncode(pfilename) & "'><img src='images/arrow_down.gif' border='0' align='absmiddle'></a></td>"
		elseif i = allnum - 1 then
			Response.Write "    <td align='center' class='cont_td'><a href='showallpf.asp?mode=up&fileid=" & Server.URLEncode(pfilename) & "'><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>&nbsp;&nbsp;&nbsp;&nbsp;</td>"
		else
			Response.Write "    <td align='center' class='cont_td'><a href='showallpf.asp?mode=up&fileid=" & Server.URLEncode(pfilename) & "'><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>&nbsp;&nbsp;<a href='showallpf.asp?mode=down&fileid=" & Server.URLEncode(pfilename) & "'><img src='images/arrow_down.gif' border='0' align='absmiddle'></a></td>"
		end if

		Response.Write "    <td align='center' class='cont_td'><a href=""javascript:del('" & pfilename & "')""><img src='images/del.gif' border='0'></a></td>"
	end if

	Response.Write "</tr>"
end if

pfilename = NULL
admin = NULL
permission = NULL
name = NULL
createTime = NULL
count = NULL
maxid = NULL
maxitem = NULL
maxsize = NULL

    i = i + 1
loop

if isadmin() = true then
%>
	<tr><td colspan="9" align="left" bgcolor="white" style="padding-top:16px;">
	<a class='wwm_btnDownload btn_blue' href="javascript:location.href='createpf.asp?<%=getGRSN() %>';"><%=b_lang_095 %></a>
	</td></tr>
<%
end if
%>
</table>

</BODY>
</HTML>

<%
set ei = nothing
set pf = nothing
set pfvl = nothing
%>
