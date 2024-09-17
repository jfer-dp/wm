<!--#include file="passinc.asp" --> 
<!--#include file="language-1.asp" --> 

<%
dim eads
set eads = server.createobject("easymail.EntAddress")
eads.Load

if trim(request("delall")) = "1" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	eads.RemoveAll

	if eads.Save() = true then
		set eads = nothing
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("ea_brow.asp?" & getGRSN())
	else
		set eads = nothing
		Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("ea_brow.asp?" & getGRSN())
	end if
end if

mp = -1
if IsNumeric(trim(request("mt_pid"))) = true then
	mp = CLng(trim(request("mt_pid")))
end if

if mp >= 0 and Request.ServerVariables("REQUEST_METHOD") = "POST" and isadmin() = true then
	eads.IndexAll

	dim msg
	dim item
	dim ss
	dim se
	dim stm

	msg = trim(request("mt_ads_ids"))
	if Len(msg) > 0 then
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(9))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)

				stm = InStr(item, "|")
				If stm <> 0 Then
					eads.SetAdsPid CLng(Left(item, stm - 1)), Right(item, Len(item) - stm), mp
				end if
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	eads.Save()
end if

allnum = eads.Count
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
.clsNode, .clsNode:visited{
	color: #101010;
	font-size : 9pt;
	text-decoration : none;
}
.clsNode:hover{
	color: #101010;
	font-size : 9pt;
	text-decoration : underline;
}
.clsNodeAds, .clsNodeAds:visited{
	color: #1e5494;
	font-size : 9pt;
	text-decoration : none;
}
.clsNodeAds:hover{
	color: #1e5494;
	font-size : 9pt;
	text-decoration : underline;
}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.wwm_line_msg {padding:18px; width:320px; margin:0 17px 0 17px; color:#202020; font-size:10pt; line-height:10px; background:#e0ecf9; border-radius:4px; -webkit-border-radius:4px; padding-left:20px; padding-right:20px; text-align:center;}
.sbttn {font-family:<%=s_lang_font %>; font-size:9pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5;color:#000066;height:19px;text-decoration:none;cursor:pointer}
-->
</STYLE>
</HEAD>

<script type="text/javascript" src="images/sc_left.js"></script>

<script type="text/javascript">
<!--
parent.f1.document.leftval.to.value = "";

var Stag;
var folderArray = new Array(<%=eads.FolderCount %>);
var folderChooseArray = new Array(<%=eads.FolderCount %>);

for (var i = 0; i < <%=eads.FolderCount %>; i++)
{
	folderArray[i] = false;
	folderChooseArray[i] = false;
}

function hide_folder(fdindex)
{
	Stag = document.getElementById("folder_" + fdindex);
	var divtag = document.getElementById("folder_" + fdindex + "_div");
	if (folderArray[fdindex] == true)
	{
		Stag.innerHTML = "<img align='absmiddle' src='images/e.gif' border=0><img align='absmiddle' src='images/fe.gif' border=0>";
		divtag.style.display = "inline";
	}
	else
	{
		Stag.innerHTML = "<img align='absmiddle' src='images/c.gif' border=0><img align='absmiddle' src='images/fc.gif' border=0>";
		divtag.style.display = "none";
	}

	folderArray[fdindex] = !folderArray[fdindex];
}

function window_onload()
{
	for (var i = 0; i < <%=eads.FolderCount %>; i++)
	{
		hide_folder(i);
	}
}

var isexall = false;
function exall()
{
	if (isexall == false)
		document.getElementById("ep").innerHTML = "<%=s_lang_co %>";
	else
		document.getElementById("ep").innerHTML = "<%=s_lang_ex %>";

	for (var i = 0; i < <%=eads.FolderCount %>; i++)
	{
		if (isexall == false)
			folderArray[i] = true;
		else
			folderArray[i] = false;

		hide_folder(i);
	}

	isexall = !isexall;
}

function checkit(num, islink)
{
	Stag = document.getElementById("cbx_" + num);

	if (islink == true)
		Stag.checked = !Stag.checked;
<%
if isadmin() = true then
%>
	if (Stag.checked == true || ischeck() == true)
	{
		document.getElementById("sfid").style.display = "inline-block";
		document.getElementById("btmove").style.display = "inline-block";
	}
	else
	{
		document.getElementById("sfid").style.display = "none";
		document.getElementById("btmove").style.display = "none";
	}
<%
end if
%>
}

function gosend()
{
	parent.f1.document.leftval.to.value = "";
	var tmpstr = "";

	for (var i = 0; i < <%=eads.Count %>; i++)
	{
		Stag = document.getElementById("cbx_" + i);

		if (Stag != null)
		{
			if (Stag.checked == true)
			{
				if (tmpstr.length > 0)
					tmpstr = tmpstr + "," + Stag.value;
				else
					tmpstr = Stag.value;
			}
		}
	}

	if (tmpstr.length > 0)
		parent.f1.document.leftval.to.value = tmpstr;

	location.href = "wframe.asp?<%=getGRSN() %>&gourl=<%=Server.URLEncode("ea_brow.asp") %>";
}

function choose_folder(fdindex, start_num)
{
	if (folderArray[fdindex] == true)
		hide_folder(fdindex);

	var i = start_num + 1;
	for (; i < <%=allnum %>; i++)
	{
		Stag = document.getElementById("cbx_" + i);
		if (Stag == null)
			break;

		if (folderChooseArray[fdindex] == false)
			Stag.checked = true;
		else
			Stag.checked = false;
	}

	folderChooseArray[fdindex] = !folderChooseArray[fdindex];
<%
if isadmin() = true then
%>
	if (ischeck() == true)
	{
		document.getElementById("sfid").style.display = "inline-block";
		document.getElementById("btmove").style.display = "inline-block";
	}
	else
	{
		document.getElementById("sfid").style.display = "none";
		document.getElementById("btmove").style.display = "none";
	}
<%
end if
%>
}

function expall() {
	var theObj = document.getElementById("exp_id");
	if (theObj != null)
	{
		if (theObj.style.display == "inline-block")
			theObj.style.display = "none";
		else
			theObj.style.display = "inline-block";
	}
	else
		location.href = "ea_brow.asp?showexp=1&<%=getGRSN() %>";
}

function delall() {
	if (confirm("<%=a_lang_182 %>") == false)
		return ;

	location.href = "ea_brow.asp?delall=1&<%=getGRSN() %>";
}

function ischeck() {
	var theObj;
	var i = 1;
	for (; i < <%=allnum %>; i++)
	{
		theObj = document.getElementById("cbx_" + i);
		if (theObj == null)
			continue;

		if (theObj.checked == true)
			return true;
	}

	return false;
}

<%
if isadmin() = true then
%>
function moveads() {
	if (document.getElementById("sfid").value.length < 1)
		return ;

	var tmp_val = "";
	var theObj;
	var i = 1;
	for (; i < <%=allnum %>; i++)
	{
		theObj = document.getElementById("cbx_" + i);
		if (theObj == null)
			continue;

		if (theObj.checked == true)
			tmp_val += i.toString() + "|" + theObj.value + "\t";
	}

	document.f1.mt_ads_ids.value = tmp_val;
	document.f1.mt_pid.value = document.getElementById("sfid").value;
	document.f1.submit();
}
<%
end if
%>
//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload()">
<table width="96%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr><td width="50%" nowrap align="left" height="28" style="padding-left:4px;">
	<a class='wwm_btnDownload btn_blue' href="javascript:gosend();"><%=a_lang_183 %></a>
	<a class='wwm_btnDownload btn_blue' href="javascript:exall();"><span id="ep"><%=s_lang_ex %></span></a>
<%
if isadmin() = true then
	Response.Write "<a class='wwm_btnDownload btn_blue' href='ea_ads.asp?" & getGRSN() & "'>" & a_lang_175 & "</a>" & Chr(13)
	Response.Write "<a class='wwm_btnDownload btn_blue' href='ea_folder.asp?" & getGRSN() & "'>" & a_lang_184 & "</a>" & Chr(13)
	Response.Write "<a class='wwm_btnDownload btn_blue' href='javascript:expall();'>" & a_lang_185 & "</a>" & Chr(13)
	Response.Write "<a class='wwm_btnDownload btn_blue' href='javascript:delall();'>" & a_lang_186 & "</a>"
%>
	</td>
	<td width="30%" nowrap align="left" height="28" style="padding-left:12px;">
<select id='sfid' name='sfid' class='drpdwn' size='1' style='display:none;'>
<option value="">---<%=a_lang_347 %>---</option>
<%
	tmp_allnum = eads.FolderCount
	i = 0
	li = 0

	do while i < tmp_allnum
		eads.GetFolderInfo i, fid, layer, name, comment

		tempstr = ""
		li = 0
		do while li < layer
			tempstr = tempstr & "|&nbsp;"
			li = li + 1
		loop

		Response.Write "<option value='" & fid & "'>" & tempstr & server.htmlencode(name) & "</option>" & Chr(13)

		fid = NULL
		layer = NULL
		name = NULL
		comment = NULL

		i = i + 1
	loop
%>
</select>
	<input type="button" id="btmove" value="<%=a_lang_348 %>" onclick="javascript:moveads();" class="sbttn" style="display:none;">
<%
end if
%>
	</td>
	<td nowrap align="right" style="padding-right:8px; color:#444444;"><%=a_lang_163 %></td>
	</tr>
</table>
<%
if trim(request("showexp")) = "1" then
%>
<div id="exp_id" style="padding-left:17px; padding-right:17px; padding-top:6px; display:inline-block;">
<textarea rows="9" cols="100" wrap="virtual" class="n_textarea"><%
i = 0
tmp_allnum = eads.FolderCount

do while i < tmp_allnum
	eads.GetFolderInfoEx i, pid, fid, layer, name, comment
	Response.Write server.htmlencode(name) & "|" & eads.GetFolderName(pid) & "|" & server.htmlencode(comment) & Chr(13) & Chr(10)

	pid = NULL
	fid = NULL
	layer = NULL
	name = NULL
	comment = NULL

	i = i + 1
loop

i = 0
tmp_allnum = eads.AdsCount
Response.Write Chr(13) & Chr(10)

do while i < tmp_allnum
	eads.GetAdsInfo i, fid, name, email, comment
	Response.Write eads.GetFolderName(fid) & "|" & server.htmlencode(email) & "|" & server.htmlencode(name) & "|" & server.htmlencode(comment) & Chr(13) & Chr(10)

	fid = NULL
	name = NULL
	email = NULL
	comment = NULL

	i = i + 1
loop
%></textarea>
</div>
<%
end if
%>
<div style="padding-left:16px; padding-right:16px; padding-top:6px;">
<%
i = 0
fi = 0
tmplayer = -1

do while i < allnum
	eads.Get i, isFolder, fid, layer, name, email, comment

	if isFolder = true then
		if layer <= tmplayer then
			li = layer
			do while li <= tmplayer
				Response.Write "</div>" & Chr(13)
				li = li + 1
			loop
		end if

		Response.Write "<table><tr><td height='2'></td></tr></table>" & Chr(13)
		Response.Write "<img src=""images/b.gif"" width=""" & 12 * layer & """ height=1 border=0>" & Chr(13)

		Response.Write "<a class=""clsNode"" href=""javascript:hide_folder(" & fi & ")""><span id=""folder_" & fi & """></span></a>&nbsp;<a class=""clsNode"" href=""javascript:choose_folder(" & fi & "," & i & ")"">" & server.htmlencode(name) & "</a>"
		if Len(comment) > 0 then
			Response.Write "&nbsp;&nbsp;<font style='color:#444444'>[" & server.htmlencode(comment) & "]</font>"
		end if
		Response.Write "<div id=""folder_" & fi & "_div""><br>" & Chr(13)

		tmplayer = layer
		fi = fi + 1
	else
		if Len(name) < 1 then
			name = email
		end if

		Response.Write "<img src=""images/b.gif"" width=""" & 30 + tmplayer * 12 & """ height=1 border=0><input id=""cbx_" & i & """ name=""cbx_" & i & """ value=""" & server.htmlencode(email) & """ type=""checkbox"" onclick=""javascript:checkit(" & i & ", false)"">&nbsp;<a class=""clsNodeAds"" title=""" & server.htmlencode(email) & """ href=""javascript:checkit(" & i & ", true)"">" & server.htmlencode(name) & "</a>"
		if Len(comment) > 0 then
			Response.Write "&nbsp;&nbsp;<font style='color:#444444'>[" & server.htmlencode(comment) & "]</font>"
		end if
		Response.Write "<br>" & Chr(13)
	end if

	isFolder = NULL
	fid = NULL
	layer = NULL
	name = NULL
	email = NULL
	comment = NULL

	i = i + 1
loop

li = 0
do while li <= tmplayer
	Response.Write "</div>" & Chr(13)
	li = li + 1
loop
%>
</div>

<%
if allnum < 1 then
%>
<table width="96%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:12px;"></td></tr>
<tr><td align="center">
<div class="wwm_line_msg"><%=s_lang_nodate %></div>
</td></tr></table>
<%
end if
%>
<form action="ea_brow.asp" method="post" name="f1">
<input type="hidden" name="mt_ads_ids">
<input type="hidden" name="mt_pid">
</form>
</BODY>
</HTML>

<%
set eads = nothing
%>
