<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" --> 

<%
ofm = trim(request("ofm"))

if Len(ofm) > 0 then
	modestr = "opener.document.f1." + ofm + ".value"
else
	modestr = "opener.value"
end if

dim eads
set eads = server.createobject("easymail.EntAddress")
eads.Load

allnum = eads.Count
%>

<HTML>
<HEAD>
<%=s_lang_meta %>
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>

<STYLE type=text/css>
<!--
.clsNode, .clsNode:visited{
	color: #6C6C6C;
	font-weight : bold;
	font-size : 9pt;
	text-decoration : none;
}
.clsNode:hover{
	color: #6C6C6C;
	font-weight : bold;
	font-size : 9pt;
	text-decoration : underline;
}
.clsNodeAds, .clsNodeAds:visited{
	color: #002f72;
	font-weight : bold;
	font-size : 9pt;
}
.clsNodeAds:hover{
	color: #BC131A;
	font-weight : bold;;
	font-size : 9pt;
}
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
		Stag.innerHTML = "<img align='absmiddle' src='images\\e.gif' border=0><img align='absmiddle' src='images\\fe.gif' border=0>";
		divtag.style.display = "inline";
	}
	else
	{
		Stag.innerHTML = "<img align='absmiddle' src='images\\c.gif' border=0><img align='absmiddle' src='images\\fc.gif' border=0>";
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
		document.getElementById("btex").value = "<%=s_lang_0057 %>";
		document.getElementById("btex1").value = "<%=s_lang_0057 %>";
	}
	else
	{
		document.getElementById("btex").value = "<%=s_lang_0058 %>";
		document.getElementById("btex1").value = "<%=s_lang_0058 %>";
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

var ret_value = "";
function gook()
{
	if (ret_value.length > 0)
		<%=modestr %> = ret_value;

	self.close();
}

function check_one(cknum)
{
	var ck_obj;
	for (var i = 0; i < <%=eads.FolderCount %>; i++)
	{
		ck_obj = document.getElementById("cbx_" + i);
		ck_obj.checked = false;
	}

	ck_obj = document.getElementById("cbx_" + cknum);
	ck_obj.checked = true;
	ret_value = ck_obj.value
}
//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload()">
<br>
<div align="right">
<input type="button" value="<%=s_lang_0059 %>" style="WIDTH: 55px" LANGUAGE=javascript onclick="gook()" class="Bsbttn">&nbsp;
<input type="button" id="btex" name="btex" value="<%=s_lang_0058 %>" style="WIDTH: 55px" LANGUAGE=javascript onclick="exall()" class="Bsbttn">&nbsp;
<input type="button" value="<%=s_lang_close %>" style="WIDTH: 55px" LANGUAGE=javascript onclick="javascript:self.close();" class="Bsbttn">&nbsp;
</div>
<hr size="1" color="<%=MY_COLOR_1 %>">
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
		Response.Write "<img src=""images\b.gif"" width=""" & 12 * layer & """ height=1 border=0>" & Chr(13)

		Response.Write "<a class=""clsNode"" href=""javascript:hide_folder(" & fi & ")""><span id=""folder_" & fi & """></span></a>&nbsp;<input type=""checkbox"" value=""" & server.htmlencode(name) & """  name=""cbx_" & fi & """ id=""cbx_" & fi & """ LANGUAGE=javascript onclick=""check_one(" & fi & ")""><font class=""clsNode"">" & server.htmlencode(name) & "</font>"
		if Len(comment) > 0 then
			Response.Write "&nbsp;&nbsp;<font style='color:#8C8C8C'>[" & server.htmlencode(comment) & "]</font>"
		end if
		Response.Write "<div id=""folder_" & fi & "_div"">" & Chr(13)

		tmplayer = layer
		fi = fi + 1
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
<br>
<hr size="1" color="<%=MY_COLOR_1 %>">
<div align="right">
<input type="button" value="<%=s_lang_0059 %>" style="WIDTH: 55px" LANGUAGE=javascript onclick="gook()" class="Bsbttn">&nbsp;
<input type="button" id="btex1" name="btex1" value="<%=s_lang_0058 %>" style="WIDTH: 55px" LANGUAGE=javascript onclick="exall()" class="Bsbttn">&nbsp;
<input type="button" value="<%=s_lang_close %>" style="WIDTH: 55px" LANGUAGE=javascript onclick="javascript:self.close();" class="Bsbttn">&nbsp;
</div>
<br>
</BODY>
</HTML>

<%
set eads = nothing
%>
