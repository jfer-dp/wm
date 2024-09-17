<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
dim mlb
set mlb = server.createobject("easymail.Labels")
mlb.Load Session("wem")

mode = trim(request("mode"))
if mode <> "" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	id = trim(request("id"))

	if mode = 1 then
		nid = trim(request("nid"))
		if nid <> "last" then
			newindex = mlb.GetIndex(nid)
			if trim(request("bf")) = "1" then
				newindex = newindex - 1
			else
				if newindex < mlb.GetIndex(id) then
					newindex = newindex + 1
				end if
			end if
		else
			newindex = mlb.Count - 1
		end if

		if mlb.GetIndex(id) = newindex then
			Response.Write "1"
		else
			if mlb.MoveTo(id, newindex) = true then
				Response.Write "1"
				mlb.Save
			else
				Response.Write "0"
			end if
		end if
	elseif mode = 2 then
		newname = trim(request("newname"))
		newname = replace(newname, "'", """")

		if mlb.ChangeTitle(id, newname) = true then
			Response.Write "1"
			mlb.Save
		else
			Response.Write "0"
		end if
	elseif mode = 3 then
		newcolor = trim(request("newcolor"))
		if mlb.ChangeColor(id, newcolor) = true then
			Response.Write "1"
			mlb.Save
		else
			Response.Write "0"
		end if
	elseif mode = 4 then
		if mlb.Del(id) = true then
			Response.Write "1"
			mlb.Save
		else
			Response.Write "0"
		end if
	elseif mode = 5 then
		newname = trim(request("newname"))
		newname = replace(newname, "'", """")
		newcolor = trim(request("newcolor"))

		if newcolor = "" then
			newcolor = mlb.Get_ZS_Color()
		end if

		old_count = mlb.Count

		if mlb.Create(newname, newcolor) = true then
			mlb.Save
			if old_count <> mlb.Count then
				mlb.GetByIndex mlb.Count - 1, ret_id, ret_title, ret_color
				Response.Write ret_id & ret_color
				ret_id = NULL
				ret_title = NULL
				ret_color = NULL
			end if
		else
			Response.Write "0"
		end if
	end if

	set mlb = nothing
	Response.End
end if

needrt = trim(request("needrt"))
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<title><%=s_lang_0300 %></title>

<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
<link rel="stylesheet" type="text/css" href="images/labels.css">

<script type="text/javascript" src="images/sc_left.js"></script>
<script type="text/javascript" src="images/mglobal.js"></script>

<style type="text/css">
.mydiv {
background:transparent;
text-align: left;
line-height: 14px;
font-size: 12px;
z-index:999;
width: 300px;
height: 180px;
left:50%;
top:34%;
margin-left:-150px!important;
margin-top:-90px!important;
position:fixed!important;
position:absolute;
_top:expression(documentElement.scrollTop + (document.documentElement.clientHeight*34)/100);
}
</style>

<script type="text/javascript">
if (ie == 6)
	document.execCommand("BackgroundImageCache", false, true); 

var offsetYInsertDiv = -3;
if(!document.all)offsetYInsertDiv = offsetYInsertDiv - 7;

var arrParent = false;
var arrMoveCont = false;
var arrMoveCounter = -1;
var arrTarget = false;
var arrNextSibling = false;
var leftPosArrangableNodes = false;
var widthArrangableNodes = false;
var nodePositionsY = new Array();
var nodeHeights = new Array();
var arrInsertDiv = false;
var insertAsFirstNode = false;
var arrNodesDestination = false;

var is_in_menu = false;
var is_menu_show = false;
var my_menu_time;
var doc = document.documentElement;
var body = document.body;

function cancelEvent()
{
	return false;
}

function getTopPos(inputObj)
{
	var returnValue = inputObj.offsetTop;
	while((inputObj = inputObj.offsetParent) != null){
	returnValue += inputObj.offsetTop;
	}
	return returnValue;
}

function getLeftPos(inputObj)
{
	var returnValue = inputObj.offsetLeft;
	while((inputObj = inputObj.offsetParent) != null)returnValue += inputObj.offsetLeft;
	return returnValue;
}

function clearMovableDiv()
{
	if(arrMoveCont.getElementsByTagName('div').length>0){
		if(arrNextSibling)arrParent.insertBefore(arrTarget,arrNextSibling); else arrParent.appendChild(arrTarget);
	}
}

function initMoveNode(e)
{
	clearMovableDiv();
	if(document.all)e = event;
	arrMoveCounter = 0;
	arrTarget = this.parentNode;
	if(this.nextSibling)arrNextSibling = this.parentNode.nextSibling; else arrNextSibling = false;
	timerMoveNode();
	closePopColorMenu();
	arrMoveCont.parentNode.style.left = e.clientX + 'px';
	arrMoveCont.parentNode.style.top = getY(e) + 'px';
	arrMoveCont.parentNode.style.width = "100%";
	arrMoveCont.parentNode.style.lineHeight = "32px";
	document.body.style.cursor = "move";
	arrTarget.style.backgroundColor = '#eee';
	return false;
}

function timerMoveNode()
{
	if(arrMoveCounter>=0 && arrMoveCounter<10){
		arrMoveCounter = arrMoveCounter +1;
		setTimeout('timerMoveNode()',20);
	}

	if(arrMoveCounter>=10){
		arrMoveCont.appendChild(arrTarget);
		arrTarget.style.borderBottomWidth = "1px";
		arrTarget.style.borderBottomStyle = "solid";
		arrTarget.style.borderBottomColor = "#ececec";
	}
}

function arrangeNodeMove(e)
{
	if(document.all)e = event;
	if(arrMoveCounter<10)return;
	if(document.all && arrMoveCounter>=10 && e.button!=1 && navigator.userAgent.indexOf('Opera')==-1){
		arrangeNodeStopMove();
	}

	arrMoveCont.parentNode.style.left = e.clientX + 'px';
	arrMoveCont.parentNode.style.top = getY(e) + 'px';

	var tmpY = getY(e);
	arrInsertDiv.style.display='none';
	arrNodesDestination = false;

	if(e.clientX<leftPosArrangableNodes || e.clientX>leftPosArrangableNodes + widthArrangableNodes)return;

	var subs = arrParent.getElementsByTagName('div');
	for(var no=0;no<subs.length;no++)
	{
		var topPos =getTopPos(subs[no]);
		var tmpHeight = subs[no].offsetHeight;

		if(no==0)
		{
			if(tmpY<=topPos && tmpY>=topPos-5)
			{
				arrInsertDiv.style.top = (topPos + offsetYInsertDiv) + 'px';
				arrInsertDiv.style.display = 'block';
				arrNodesDestination = subs[no];
				insertAsFirstNode = true;
				return;
			}
		}

		if(tmpY>=topPos && tmpY<=(topPos+tmpHeight))
		{
			arrInsertDiv.style.top = (topPos+tmpHeight + offsetYInsertDiv) + 'px';
			arrInsertDiv.style.display = 'block';
			arrNodesDestination = subs[no];
			insertAsFirstNode = false;
			return;
		}
	}
}

function arrangeNodeStopMove()
{
	if(arrTarget)
	{
		arrTarget.style.borderBottomWidth = "0px";
		arrTarget.style.borderBottomStyle = "none";
		arrTarget.style.borderBottomColor = "white";
		arrTarget.style.backgroundColor = 'white';
	}

	arrMoveCounter = -1;
	arrInsertDiv.style.display='none';

	if(arrNodesDestination)
	{
		mode = 1;
		var subs = arrParent.getElementsByTagName('div');
		if(arrNodesDestination==subs[0] && insertAsFirstNode){
			arrParent.insertBefore(arrTarget,arrNodesDestination);
			add_url = "mode=1&id=" + arrTarget.id.substring(3, 11) + "&nid=" + arrNodesDestination.id.substring(3, 11) + "&bf=1";
		}else{
			if(arrNodesDestination.nextSibling){
				arrParent.insertBefore(arrTarget,arrNodesDestination.nextSibling);
				add_url = "mode=1&id=" + arrTarget.id.substring(3, 11) + "&nid=" + arrNodesDestination.id.substring(3, 11);
			}else{
				arrParent.appendChild(arrTarget);
				add_url = "mode=1&id=" + arrTarget.id.substring(3, 11) + "&nid=last";
			}
		}
		SendInfo();
	}

	arrNodesDestination = false;
	clearMovableDiv();
	document.body.style.cursor = "default";
}

function initArrangableNodes()
{
	arrParent = document.getElementById('arrangableNodes');
	arrMoveCont = document.getElementById('movableNode').getElementsByTagName('UL')[0];
	arrInsertDiv = document.getElementById('arrDestInditcator');

	leftPosArrangableNodes = getLeftPos(arrParent);
	arrInsertDiv.style.left = leftPosArrangableNodes - 5 + 'px';
	widthArrangableNodes = arrParent.offsetWidth;

	var subs = arrParent.getElementsByTagName('div');
	for(var no=0;no<subs.length;no++)
	{
		if (subs[no].id.substring(0, 4) == "node")
		{
			subs[no].onmousedown = initMoveNode;
			subs[no].onselectstart = cancelEvent;
		}
	}

	document.documentElement.onmouseup = arrangeNodeStopMove;
	document.documentElement.onmousemove = arrangeNodeMove;
	arrParent.onselectstart = cancelEvent;
	init_pop_win();
}

window.onload = initArrangableNodes;

function mouse_over_div(div_obj)
{
	if (arrMoveCounter < 0 && is_in_menu == false && is_menu_show == false)
		div_obj.style.backgroundColor='#eee';
	else
		div_obj.style.backgroundColor='white';
}

function mouse_out_div(div_obj)
{
	if (arrMoveCounter < 0 && is_in_menu == false && is_menu_show == false)
		div_obj.style.backgroundColor='white';
}

function bindListener(dal_id){
	$("#" + dal_id).unbind().powerFloat({
		width: 140,
		eventType: "click",
		target: "#color_box",
		showCall: function() {
			if (is_menu_show == true)
			{
				cg_cl_tag_id = "";
				$.powerFloat.hide();
			}
			else
			{
				cg_cl_tag_id = this.attr("id");
				is_menu_show = true;
				clearTimeout(my_menu_time);
			}

			$(".color_qmpanel_shadow").mouseover(function() {
				is_in_menu = true;
				clearTimeout(my_menu_time);
			});

			$(".color_qmpanel_shadow").mouseout(function() {
				is_in_menu = false;
				my_menu_time = setTimeout("setTimeClose()", 1000);
			});
		},
		hideCall: function() {
			setTimeout("set_menu_close()", 300);
		}
	});
}

parent.f1.document.leftval.purl.value = "";
function showlabel(lb_id) {
	parent.f1.document.leftval.purl.value = "labels.asp?<%=getGRSN() %>";
	location.href = "listlabel.asp?<%=getGRSN() %>&lbid=" + lb_id;
}

function lb_get_str(lb_id, lb_title, lb_color, lb_mailnum, lb_newmailnum)
{
	return "<div class='dpline' id='dp_" + lb_id + "' onmouseover='mouse_over_div(this);' onmouseout='mouse_out_div(this);'><div class='movebar' id='node_" + lb_id + "'><span class='ico_move'></span></div>\
<span style='float:left; padding-top:4px;'><span class=\"color_pop\" id='cp_" + lb_id + "'>\
<div class='attbg'><img align='absmiddle' style='margin: 3px 0pt 0pt;' src='images/popshow.gif'></div>\
<div class='color_pop_text'><span id='cl_" + lb_id + "' class='wwm_color_in_line' style='background:#" + lb_color + ";'>&nbsp;</span></div>\
</span></span>\
<span class='col1'><a href='#' style='color:white;text-decoration:none;' onclick=\"showlabel('" + lb_id + "')\"><span id='tt_" + lb_id + "' class='wwm_color_text' style='background:#" + lb_color + ";'>" + lb_title + "</span></a></span>\
<span class='col2'>" + lb_newmailnum + "</span>\
<span class='col2'>" + lb_mailnum + "</span>\
<span class='col_r'><a href='javascript:void(0)' onClick='change_name(\"" + lb_id + "\");'><%=s_lang_0301 %></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href='javascript:void(0)' onClick='return del_tag(\"" + lb_id + "\")'><%=s_lang_del %></a></span>\
</div>";
}

function lb_add(lb_id, lb_title, lb_color, lb_mailnum, lb_newmailnum)
{
	$("#arrangableNodes").append(lb_get_str(lb_id, lb_title, lb_color, lb_mailnum, lb_newmailnum));

	bindListener("cp_" + lb_id);
	initArrangableNodes();
}

var oWin;
var oLay;
var oClose;

function init_pop_win()
{
	oWin = document.getElementById("pop_win");
	oLay = document.getElementById("pop_overlay");
	oClose = document.getElementById("pop_close");

	oClose.onclick = function ()
	{
		pop_close();
	}
};

function pop_show()
{
	oLay.style.height = document.documentElement.scrollHeight + "px";
	oLay.style.width = document.documentElement.scrollWidth + "px";

	var ie_h = doc && doc.clientHeight  || body && body.clientHeight  || 0;
	var ie_w = doc && doc.clientWidth  || body && body.clientWidth  || 0;

	if (ie_h > document.documentElement.scrollHeight)
		oLay.style.height = ie_h + "px";

	if (ie_w > document.documentElement.scrollWidth)
		oLay.style.width = ie_w + "px";

	oLay.style.display = "block";
	oWin.style.display = "block"	
}

function pop_close()
{
	oLay.style.display = "none";
	oWin.style.display = "none"	
}

var tag_id = "";
var mode = 0;
var add_url = "";
var cg_cl_tag_id = "";
var cg_cl_select = "";
var tag_title = "";

function have_this_text(f_text)
{
	var h_find_it = false;
	$(".wwm_color_text").each(function() {
		if (f_text.toLowerCase() == this.innerHTML.toLowerCase())
		{
			h_find_it = true;
			return false;
		}
	});
	return h_find_it;
}

function have_this_color(t_color)
{
	var h_find_it = false;
	$(".wwm_color_in_line").each(function() {
		if (t_color.toLowerCase() == this.style.backgroundColor.toLowerCase())
		{
			h_find_it = true;
			return false;
		}
	});
	return h_find_it;
}

function get_new_color()
{
	var ret_new_color = "";
	$(".wwm_color_in_box").each(function() {
		if (have_this_color(this.style.backgroundColor) == false)
		{
			ret_new_color = this.style.backgroundColor.colorHex();
			return false;
		}
	});

	if (ret_new_color.length > 6)
		ret_new_color = ret_new_color.substring(1, 7);

	return ret_new_color;
}

function pop_create()
{
	$(".title_left")[0].innerHTML = "<%=s_lang_0302 %>";
	$("#pop_ctmsg")[0].innerHTML = "<%=s_lang_0303 %>";

	document.getElementById('NewName').value = "";
	document.getElementById('pop_ok').onclick = create_label;
	pop_show();
	document.getElementById('NewName').focus();
}

function create_label()
{
	var g_newname = document.getElementById('NewName').value;

	if (g_newname.length < 1)
	{
		alert("<%=s_lang_0304 %>.");
		document.getElementById('NewName').focus();
		return ;
	}

	if (have_this_text(g_newname) == true)
	{
		alert("<%=s_lang_0305 %>.");
		document.getElementById('NewName').focus();
		return ;
	}

	g_newname = g_newname.replace(/\'/g,"\"");

	mode = 5;
	tag_title = htmlEscape(g_newname);
	add_url = "mode=5&newcolor=" + get_new_color() + "&newname=" + escape(g_newname);
	SendInfo();
}

function change_name(obid)
{
	$(".title_left")[0].innerHTML = "<%=s_lang_0306 %>";
	$("#pop_ctmsg")[0].innerHTML = "<%=s_lang_0307 %>";

	tag_id = obid;
	add_url = "id=" + obid;
	document.getElementById('NewName').value = "";
	document.getElementById('pop_ok').onclick = change_name_over;
	pop_show();
	document.getElementById('NewName').focus();
}

function change_name_over(obid)
{
	var g_newname = document.getElementById('NewName').value;

	if (g_newname.length < 1)
	{
		alert("<%=s_lang_0304 %>.");
		document.getElementById('NewName').focus();
		return ;
	}

	if (have_this_text(g_newname) == true)
	{
		alert("<%=s_lang_0305 %>.");
		document.getElementById('NewName').focus();
		return ;
	}

	g_newname = g_newname.replace(/\'/g,"\"");

	mode = 2;
	tag_title = htmlEscape(g_newname);
	add_url += "&mode=2&newname=" + escape(g_newname);
	SendInfo();
}

function change_color(s_col)
{
	if (s_col.length == 6 && cg_cl_tag_id.length == 11)
	{
		obid = cg_cl_tag_id.substring(3, 11)
		tag_id = obid;
		mode = 3;
		cg_cl_select = s_col;

		add_url = "id=" + obid + "&mode=3&newcolor=" + s_col;
		SendInfo();
	}
}

function del_tag(obid)
{
	tag_id = obid;
	mode = 4;
	add_url = "id=" + obid + "&mode=4";
	SendInfo();
}


var request = false;
try {
	request = new XMLHttpRequest();
} catch (trymicrosoft) {
try {
	request = new ActiveXObject("Msxml2.XMLHTTP");
} catch (othermicrosoft) {
try {
	request = new ActiveXObject("Microsoft.XMLHTTP");
} catch (failed) {
	request = false;
}}}

if (!request)
	alert("Error initializing XMLHttpRequest!");


function SendInfo()
{
	var url = "labels.asp?" + add_url + "&<%=getGRSN() %>";
	request.open("GET", url, true);
	request.onreadystatechange = updatePage;
	request.send(null);
}

function updatePage()
{
	if (request.readyState == 4)
	{
		if (request.status == 200)
		{
			if (request.responseText == "0")
				document.location.reload(true);
			else
			{
				if (mode == 2)
				{
					pop_close();

					if (tag_id.length == 8)
						document.getElementById("tt_" + tag_id).innerHTML = tag_title;
				}
				else if (mode == 3)
				{
					if (cg_cl_select.length == 6 && tag_id.length == 8)
					{
						closePopColorMenu();
						document.getElementById("cl_" + tag_id).style.background = "#" + cg_cl_select;
						document.getElementById("tt_" + tag_id).style.background = "#" + cg_cl_select;
					}
				}
				else if (mode == 4)
				{
					if (tag_id.length == 8)
					{
						$("#dp_" + tag_id).unbind();
						$("#dp_" + tag_id).remove();
					}
				}
				else if (mode == 5)
				{
					if (tag_title.length > 0 && request.responseText.length > 13)
					{
						lb_add(request.responseText.substring(0, 8), tag_title, request.responseText.substring(8, 14), "0", "0");
						pop_close();
					}
				}

				parent.f1.window.location.href = "left.asp?<%=getGRSN() %>&asp=" + escape(document.location.href);
			}

			tag_id = "";
			mode = 0;
			add_url = "";
			cg_cl_tag_id = "";
			cg_cl_select = "";
			tag_title = "";
		}
	}
}
</script>

</head>

<body>
<div class="create_title">
	<table width="98%" border="0" cellspacing="0" align="center"><tr><td width="50%" align="left">
<%
gourl = trim(request("gourl"))

if needrt = "1"  then
	if Len(gourl) < 4 then
		Response.Write "<a class='wwm_btnDownload btn_gray' href='javascript:history.back();'><< " & s_lang_return & "</a>"
	else
		Response.Write "<a class='wwm_btnDownload btn_gray' href='" & gourl & "'><< " & s_lang_return & "</a>"
	end if
end if
%>
	&nbsp;
	</td><td width="50%">
	<a class="wwm_btnDownload btn_gray" href="#" onclick="javascript:pop_create();"><%=s_lang_0308 %></a>
	</td></tr></table>
</div>
<br>
<div class="ie6-out">
<div class="ie6-in">
<div class="dpline" style="border-top:2px solid #a7c5e2;background-color:#e0ecf9;">
<span class="col1_title"><%=s_lang_0309 %></span>
<span class="col2"><%=s_lang_0310 %></span>
<span class="col2"><%=s_lang_0311 %></span>
<span class="col_r"><%=s_lang_0312 %></span>
</div>
<ul id="arrangableNodes" class="folder_manage" style="">
<script type="text/javascript">
<%
i = 0
allnum = mlb.Count
do while i < allnum
	mlb.GetByIndex i, ret_id, ret_title, ret_color
	mlb.GetMailsCountInLabel ret_id, ret_mailnum, ret_newmailnum

	Response.Write "document.write(lb_get_str(""" & server.htmlencode(ret_id) & """, """ & server.htmlencode(ret_title) & """, """ & server.htmlencode(ret_color) & """, """ & ret_mailnum & """, """ & ret_newmailnum & """));" & Chr(13)

	ret_mailnum = NULL
	ret_newmailnum = NULL
	ret_id = NULL
	ret_title = NULL
	ret_color = NULL

	i = i + 1
loop
%>
</script>
</ul>
</div>
</div>
<div id="color_box" class="color_qmpanel_shadow" style="display:none; position:absolute;">
	<div class="menu_base">
    	<div class="menu_bd bd">
	<div class="target_list">
		<a href="javascript:change_color('609022');"><span class="wwm_color_in_box" style="background:#609022;">&nbsp;</span></a>
		<a href="javascript:change_color('aa41cd');"><span class="wwm_color_in_box" style="background:#aa41cd;">&nbsp;</span></a>
		<a href="javascript:change_color('35909e');"><span class="wwm_color_in_box" style="background:#35909e;">&nbsp;</span></a>
		<a href="javascript:change_color('3d6aaa');"><span class="wwm_color_in_box" style="background:#3d6aaa;">&nbsp;</span></a>
		<a href="javascript:change_color('4d53a5');"><span class="wwm_color_in_box" style="background:#4d53a5;">&nbsp;</span></a><br>
		<a href="javascript:change_color('b48e43');"><span class="wwm_color_in_box" style="background:#b48e43;">&nbsp;</span></a>
		<a href="javascript:change_color('c26502');"><span class="wwm_color_in_box" style="background:#c26502;">&nbsp;</span></a>
		<a href="javascript:change_color('b3341a');"><span class="wwm_color_in_box" style="background:#b3341a;">&nbsp;</span></a>
		<a href="javascript:change_color('c24d96');"><span class="wwm_color_in_box" style="background:#c24d96;">&nbsp;</span></a>
		<a href="javascript:change_color('b21414');"><span class="wwm_color_in_box" style="background:#b21414;">&nbsp;</span></a><br>
		<a href="javascript:change_color('e59c00');"><span class="wwm_color_in_box" style="background:#e59c00;">&nbsp;</span></a>
		<a href="javascript:change_color('ec6928');"><span class="wwm_color_in_box" style="background:#ec6928;">&nbsp;</span></a>
		<a href="javascript:change_color('9d569d');"><span class="wwm_color_in_box" style="background:#9d569d;">&nbsp;</span></a>
		<a href="javascript:change_color('955959');"><span class="wwm_color_in_box" style="background:#955959;">&nbsp;</span></a>
		<a href="javascript:change_color('ae7841');"><span class="wwm_color_in_box" style="background:#ae7841;">&nbsp;</span></a><br>
		<a href="javascript:change_color('abab4e');"><span class="wwm_color_in_box" style="background:#abab4e;">&nbsp;</span></a>
		<a href="javascript:change_color('ec5105');"><span class="wwm_color_in_box" style="background:#ec5105;">&nbsp;</span></a>
		<a href="javascript:change_color('ab4646');"><span class="wwm_color_in_box" style="background:#ab4646;">&nbsp;</span></a>
		<a href="javascript:change_color('950695');"><span class="wwm_color_in_box" style="background:#950695;">&nbsp;</span></a>
		<a href="javascript:change_color('703b70');"><span class="wwm_color_in_box" style="background:#703b70;">&nbsp;</span></a><br>
		<a href="javascript:change_color('3b994f');"><span class="wwm_color_in_box" style="background:#3b994f;">&nbsp;</span></a>
		<a href="javascript:change_color('21b1b1');"><span class="wwm_color_in_box" style="background:#21b1b1;">&nbsp;</span></a>
		<a href="javascript:change_color('1e87ef');"><span class="wwm_color_in_box" style="background:#1e87ef;">&nbsp;</span></a>
		<a href="javascript:change_color('4b9d8f');"><span class="wwm_color_in_box" style="background:#4b9d8f;">&nbsp;</span></a>
		<a href="javascript:change_color('7c657c');"><span class="wwm_color_in_box" style="background:#7c657c;">&nbsp;</span></a><br>
		<a href="javascript:change_color('5487ed');"><span class="wwm_color_in_box" style="background:#5487ed;">&nbsp;</span></a>
		<a href="javascript:change_color('354b66');"><span class="wwm_color_in_box" style="background:#354b66;">&nbsp;</span></a>
		<a href="javascript:change_color('2768ea');"><span class="wwm_color_in_box" style="background:#2768ea;">&nbsp;</span></a>
		<a href="javascript:change_color('7044b2');"><span class="wwm_color_in_box" style="background:#7044b2;">&nbsp;</span></a>
		<a href="javascript:change_color('1f28df');"><span class="wwm_color_in_box" style="background:#1f28df;">&nbsp;</span></a><br>
		<a href="javascript:change_color('a59f79');"><span class="wwm_color_in_box" style="background:#a59f79;">&nbsp;</span></a>
		<a href="javascript:change_color('8899ab');"><span class="wwm_color_in_box" style="background:#8899ab;">&nbsp;</span></a>
		<a href="javascript:change_color('585858');"><span class="wwm_color_in_box" style="background:#585858;">&nbsp;</span></a>
		<a href="javascript:change_color('343434');"><span class="wwm_color_in_box" style="background:#343434;">&nbsp;</span></a>
		<a href="javascript:change_color('000000');"><span class="wwm_color_in_box" style="background:#000000;">&nbsp;</span></a><br>
    </div>
</div>
</div>
</div>

<div id="movableNode"><ul></ul></div>	
<div id="arrDestInditcator"><img src="images/lbinsert.gif"></div>

<script type="text/javascript" src="images/jquery.min.js"></script>
<script type="text/javascript" src="images/jquery-powerFloat-min.js"></script>
<script type="text/javascript">
$(function() {
	$(".color_pop").powerFloat({
		width: 140,
		eventType: "click",
		target: "#color_box",
		showCall: function() {
			if (is_menu_show == true)
			{
				cg_cl_tag_id = "";
				$.powerFloat.hide();
			}
			else
			{
				cg_cl_tag_id = this.attr("id");
				is_menu_show = true;
				clearTimeout(my_menu_time);
			}

			$(".color_qmpanel_shadow").mouseover(function() {
				is_in_menu = true;
				clearTimeout(my_menu_time);
			});

			$(".color_qmpanel_shadow").mouseout(function() {
				is_in_menu = false;
				my_menu_time = setTimeout("setTimeClose()", 1000);
			});
		},
		hideCall: function() {
			setTimeout("set_menu_close()", 300);
		}
	});
});

function set_menu_close()
{
	cg_cl_tag_id = "";
	is_menu_show = false;
}

function setTimeClose()
{
	if (is_in_menu == false)
	{
		cg_cl_tag_id = "";
		$.powerFloat.hide();
	}
}

function closePopColorMenu()
{
	cg_cl_tag_id = "";
	$.powerFloat.hide();
}
</script>


<div id="pop_overlay">
</div>

<div id="pop_win" style="display:none; position:absolute;" class="mydiv">
	<div class="pop_base"><div class="pop_bd bd"><div class="title">
		<div class="title_left"><%=s_lang_0306 %></div>
		<div class="title_right" title="<%=s_lang_close %>" id="pop_close"><span>&nbsp;</span></div>
	</div>
	<div class="pop_content"><span id="pop_ctmsg"><%=s_lang_0307 %></span><br>
	<input type="text" id="NewName" size="30" maxlength="50" class='b_input'>
	</div>
	<div class="title_bottom">
	<div class="title_ok_cancel_div">
	<a id="pop_ok" class="wwm_btnDownload btn_gray" href="#" onclick="javascript:change_name_over();"><%=s_lang_0313 %></a>&nbsp;
	<a class="wwm_btnDownload btn_gray" href="#" onclick="javascript:pop_close()"><%=s_lang_cancel %></a>
	</div></div></div></div>
</div>

</body>
</html>

<%
set mlb = nothing
%>
