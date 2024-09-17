<!--#include file="passinc.asp" --> 
<!--#include file="language-1.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim eads
set eads = server.createobject("easymail.EntAddress")
eads.Load

if Request.ServerVariables("REQUEST_METHOD") = "POST" and trim(request("fid")) <> "" then
	isok = false
	optmode = trim(request("optmode"))

	savevalue = trim(request("savevalue"))
	save2value = trim(request("save2value"))

	if optmode = "0" then
		isok = eads.RemoveFolder(trim(request("fid")))
	elseif optmode = "1" then
		isok = eads.UpFolder(trim(request("fid")), false)
	elseif optmode = "2" then
		isok = eads.UpFolder(trim(request("fid")), true)
	elseif optmode = "3" then
		isok = eads.DownFolder(trim(request("fid")), false)
	elseif optmode = "4" then
		isok = eads.DownFolder(trim(request("fid")), true)
	elseif optmode = "5" then
		isok = eads.SetFolderName(trim(request("fid")), savevalue)
	elseif optmode = "6" then
		isok = eads.SetFolderComment(trim(request("fid")), savevalue)
	elseif optmode = "7" then
		isok = eads.SetFolderPid(trim(request("fid")), savevalue)
	elseif optmode = "8" then
		isok = eads.AddFolder(savevalue, save2value, trim(request("fid")))
	elseif optmode = "9" then
		isok = eads.AddFolder(savevalue, save2value, -1)
	end if

	if isok = true then
		isok = eads.Save()
	end if

	if isok = false then
		set eads = nothing
		Response.Redirect "err.asp?" & getGRSN() & "&gourl=ea_folder.asp"
	elseif optmode = "8" or optmode = "9" then
		set eads = nothing

		if isok = true then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=ea_folder.asp"
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=ea_folder.asp"
		end if
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
.Bsbttn {font-family:<%=s_lang_font %>; font-size:10pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5; color:#000066;text-decoration:none;cursor:pointer}
.sbttn {font-family:<%=s_lang_font %>; font-size:9pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5;color:#000066;height:19px;text-decoration:none;cursor:pointer}
.textbox { BORDER:1px solid #555555;}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.wwm_line_msg {padding:18px; width:320px; margin:0 17px 0 17px; color:#202020; font-size:10pt; line-height:10px; background:#e0ecf9; border-radius:4px; -webkit-border-radius:4px; padding-left:20px; padding-right:20px; text-align:center;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
var clickfid = -1;
var Stag;

function showset(fid)
{
	if (clickfid != fid)
	{
		if (clickfid > -1)
		{
			Stag = document.getElementById("folder_option_" + clickfid);
			Stag.innerHTML = "";
		}

		Stag = document.getElementById("folder_option_" + fid);
		Stag.innerHTML = "<select id='msel' name='msel' class='drpdwn' size='1' LANGUAGE=javascript onchange='selopt()'><option value=''>---<%=a_lang_187 %>---</option><option value='0'><%=s_lang_del %></option><option value=''>------</option><option value='1'><%=a_lang_188 %></option><option value='2'><%=a_lang_189 %></option><option value='3'><%=a_lang_190 %></option><option value='4'><%=a_lang_191 %></option><option value=''>------</option><option value='5'><%=a_lang_192 %></option><option value='6'><%=a_lang_193 %></option><option value='7'><%=a_lang_194 %>...</option><option value=''>------</option><option value='8'><%=s_lang_add %></option><option value='9'><%=a_lang_195 %></option></select>&nbsp;<span id='forinput'></span>";

		clickfid = fid;
	}
}

function selopt()
{
	var optnum = document.getElementById("msel").value;
	document.f1.fid.value = clickfid;

	if (optnum == "0")
	{
		document.getElementById("msel").selectedIndex = 0;
		show_folder_sel(false);
		show_input(false);
		show_add(false);

		if (confirm("<%=a_lang_196 %>") == false)
			return ;

		document.f1.optmode.value = "0";
		document.f1.savevalue.value = "";
		document.f1.save2value.value = "";
		document.f1.submit();
	}
	else if (optnum == "1" || optnum == "2" || optnum == "3" || optnum == "4")
	{
		show_folder_sel(false);
		show_input(false);
		show_add(false);

		document.f1.optmode.value = optnum;
		document.f1.savevalue.value = "";
		document.f1.save2value.value = "";
		document.f1.submit();
	}
	else if (optnum == "5" || optnum == "6")
	{
		show_folder_sel(false);
		show_add(false);
		show_input(true);

		document.f1.optmode.value = optnum;
	}
	else if (optnum == "7")
	{
		show_input(false);
		show_add(false);
		show_folder_sel(true);

		document.f1.optmode.value = optnum;
	}
	else if (optnum == "8" || optnum == "9")
	{
		show_folder_sel(false);
		show_input(false);
		show_add(true);

		document.f1.optmode.value = optnum;
	}
	else
	{
		document.getElementById("msel").selectedIndex = 0;
		show_folder_sel(false);
		show_input(false);
		show_add(false);
	}
}

function save_name_comment()
{
	document.f1.savevalue.value = document.getElementById("name_comment").value;
	document.f1.save2value.value = "";
	document.f1.submit();
}

function save_add()
{
	document.f1.savevalue.value = document.getElementById("eaname").value;
	document.f1.save2value.value = document.getElementById("eacomment").value;
	document.f1.submit();
}

function save_move()
{
	document.f1.savevalue.value = document.getElementById("sfid").value;
	document.f1.save2value.value = "";
	document.f1.submit();
}

function show_input(isshow)
{
	if (isshow == true)
		document.getElementById("forinput").innerHTML = "<input id='name_comment' name='name_comment' type='text' class='textbox' size=10>&nbsp;<input type='button' value='<%=s_lang_ok %>' class='sbttn' LANGUAGE=javascript onclick='save_name_comment()'>";
	else
		document.getElementById("forinput").innerHTML = "";
}

function show_add(isshow)
{
	if (isshow == true)
		document.getElementById("forinput").innerHTML = "<%=a_lang_178 %>:<input id='eaname' name='eaname' type='text' class='textbox' size=10>&nbsp;<%=a_lang_180 %>:<input id='eacomment' name='eacomment' type='text' class='textbox' size=10>&nbsp;<input type='button' value='<%=s_lang_ok %>' class='sbttn' LANGUAGE=javascript onclick='save_add()'>";
	else
		document.getElementById("forinput").innerHTML = "";
}

function show_folder_sel(isshow)
{
	if (isshow == true)
		document.getElementById("forinput").innerHTML = foldersel + "&nbsp;<input type='button' value='<%=s_lang_ok %>' class='sbttn' LANGUAGE=javascript onclick='save_move()'>";
	else
		document.getElementById("forinput").innerHTML = "";
}

var isshowheadadd = false;
function headadd()
{
	Stag = document.getElementById("head_add");

	if (isshowheadadd == false)
		Stag.innerHTML = "<%=a_lang_178 %>:<input id='heaname' name='heaname' type='text' class='n_textbox' size=10>&nbsp;<%=a_lang_180 %>:<input id='heacomment' name='heacomment' type='text' class='n_textbox' size=10>&nbsp;<input type='button' value='<%=s_lang_ok %>' class='Bsbttn' LANGUAGE=javascript onclick='save_head_add()'>";
	else
		Stag.innerHTML = "";

	isshowheadadd = !isshowheadadd;
}

function save_head_add()
{
	if (document.getElementById("heaname").value.length > 0)
	{
		document.f1.optmode.value = "9";
		document.f1.fid.value = "-1";
		document.f1.savevalue.value = document.getElementById("heaname").value;
		document.f1.save2value.value = document.getElementById("heacomment").value;
		document.f1.submit();
	}
}
//-->
</SCRIPT>

<BODY>
<table width="96%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr>
	<td nowrap width="10%" align="left" height="28" style="padding-left:4px;">
	<a class='wwm_btnDownload btn_blue' href="ea_brow.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
	<a class='wwm_btnDownload btn_blue' href="javascript:headadd();"><%=s_lang_add %></a>
	</td>
	<td nowrap align="left">
	<span id="head_add" style="padding-left:8px; color:#444444;"></span>
	</td>
	<td width="10%" nowrap align="right" style="padding-right:8px; color:#444444;"><%=a_lang_184 %></td>
	</tr>
</table>
<div style="padding-left:16px; padding-right:16px; padding-top:6px;">
<%
allnum = eads.FolderCount
i = 0
li = 0

do while i < allnum
	eads.GetFolderInfo i, fid, layer, name, comment

	Response.Write "<table><tr><td height='2'></td></tr></table>" & Chr(13)
	Response.Write "<img src=""images/b.gif"" width=""" & 12 * layer & """ height=1 border=0>" & Chr(13)
	Response.Write "<img align='absmiddle' src='images/e.gif' border=0><a class=""clsNode"" href=""javascript:showset(" & fid & ")""><img align='absmiddle' src='images/fe.gif' border=0></a>&nbsp;<a class=""clsNode"" href=""javascript:showset(" & fid & ")"">" & server.htmlencode(name) & "</a>"

	if Len(comment) > 0 then
		Response.Write "&nbsp;&nbsp;<font style='color:#444444'>[" & server.htmlencode(comment) & "]</font>"
	end if
	Response.Write "&nbsp;<span id=""folder_option_" & fid & """></span>" & Chr(13)

	tempstr = ""
	li = 0
	do while li < layer
		tempstr = tempstr & "|&nbsp;"
		li = li + 1
	loop

	fstr = fstr & "<option value='" & fid & "'>" & tempstr & server.htmlencode(name) & "</option>\" & Chr(13)

	fid = NULL
	layer = NULL
	name = NULL
	comment = NULL

	i = i + 1
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

<FORM ACTION="ea_folder.asp" METHOD="POST" NAME="f1">
<input id="optmode" name="optmode" type="hidden">
<input id="fid" name="fid" type="hidden">
<input id="savevalue" name="savevalue" type="hidden">
<input id="save2value" name="save2value" type="hidden">
</FORM>

</BODY>

<SCRIPT LANGUAGE=javascript>
<!--
var foldersel = "<select id='sfid' name='sfid' class='drpdwn' size='1'><option value='-1'><%=a_lang_197 %></option><%=fstr %></select>";
//-->
</SCRIPT>
</HTML>

<%
set eads = nothing
%>
