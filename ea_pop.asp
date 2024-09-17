<!--#include file="passinc.asp" --> 
<!--#include file="language-1.asp" --> 

<%
mode = trim(request("mode"))
ofm = trim(request("ofm"))

if Len(ofm) > 0 then
	modestr = "opener." + ofm + ".value"
else
	modestr = "opener.value"
end if

if mode = "deliver" then
	modestr = "opener.document.f1.to.value"
end if

dim eads
set eads = server.createobject("easymail.EntAddress")
eads.Load

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
-->
</STYLE>
</HEAD>

<SCRIPT LANGUAGE=javascript>
<!--
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
	{
		document.getElementById("btex").innerHTML = "<%=s_lang_co %>";
		document.getElementById("btex1").innerHTML = "<%=s_lang_co %>";
	}
	else
	{
		document.getElementById("btex").innerHTML = "<%=s_lang_ex %>";
		document.getElementById("btex1").innerHTML = "<%=s_lang_ex %>";
	}

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
}

function gook()
{
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
	{
		if (<%=modestr %>.length > 0)
		{
			if (end_have_dh(<%=modestr %>) == false)
				<%=modestr %> = <%=modestr %> + "," + tmpstr;
			else
				<%=modestr %> = <%=modestr %> + tmpstr;
		}
		else
			<%=modestr %> = tmpstr;
	}

	self.close();
}

function end_have_dh(ads_str)
{
	if (ads_str.length > 0)
	{
		var echr = ads_str.substr(ads_str.length - 1);
		if (echr == "," || echr == ";")
			return true;
	}

	return false;
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
}
//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload()">
<div align="right" style="padding-right:20px;">
<a class="wwm_btnDownload btn_blue" href="#" onclick="gook()"><%=s_lang_ok %></a>&nbsp;
<a class="wwm_btnDownload btn_blue" href="#" onclick="exall()"><span id="btex"><%=s_lang_ex %></span></a>&nbsp;
<a class="wwm_btnDownload btn_blue" href="#" onclick="javascript:self.close();"><%=s_lang_cancel %></a>
</div>
<hr size="1" color="#A5B6C8">
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

		Response.Write "<table><tr><td height=""2""></td></tr></table>" & Chr(13)
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

<%
if allnum < 1 then
%>
<table width="96%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:12px;"></td></tr>
<tr><td align="center">
<div class="wwm_line_msg"><%=s_lang_nodate %></a>
</td></tr></table>
<%
end if
%>

<br>
<hr size="1" color="#A5B6C8">
<div align="right" style="padding-right:20px;">
<a class="wwm_btnDownload btn_blue" href="#" onclick="gook()"><%=s_lang_ok %></a>&nbsp;
<a class="wwm_btnDownload btn_blue" href="#" onclick="exall()"><span id="btex1"><%=s_lang_ex %></span></a>&nbsp;
<a class="wwm_btnDownload btn_blue" href="#" onclick="javascript:self.close();"><%=s_lang_cancel %></a>
</div>
</BODY>
</HTML>

<%
set eads = nothing
%>
