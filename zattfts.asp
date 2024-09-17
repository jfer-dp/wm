<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" --> 

<%
dim ei
set ei = server.createobject("easymail.InfoList")
ei.isLoadZatt = true
ei.timeMode = 5

ei.LoadMailBox Session("wem"), "att"
allnum = ei.getMailsCount
zallnum = allnum
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail <%=s_lang_0284  %></TITLE>
<link rel="stylesheet" type="text/css" href="images/hwem.css">
<link rel="stylesheet" type="text/css" href="images/slstyle.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
<link rel="stylesheet" type="text/css" href="images/tab_content.css">
</HEAD>

<BODY>
<form name="f1">
<body>

<ul class="tabs" persist="false">
	<li><a href="#" rel="view1"><%=s_lang_0292 %></a></li>
	<li><a href="#" rel="view2"><%=s_lang_0296 %></a></li>
</ul>
<div class="tabcontents" style="padding: 8px 8px 8px;">
	<div id="view1" class="tabcontent">

<div style="padding-left:3px; padding-bottom:2px; font-family:<%=s_lang_font %>; font-size:9pt">
<select id="columns" style="display:none;"><option value="1"></option></select>
<%=s_lang_0295 %><%=s_lang_mh %><input type="text" class='textbox_wwm' id="query" onkeyup="sorter.search('query')">
</div>

	<table align="center" id="table" class="tinytable" border="0" cellpadding="0" cellspacing="0">
		<thead>
			<tr>
				<th width="4%" class="nosort"><h3>&nbsp;</h3></th>
				<th width="63%" noWrap><h3><%=s_lang_0067 %></h3></th>
				<th width="13%" noWrap><h3><%=s_lang_0179 %></h3></th>
				<th width="20%" noWrap><h3><%=s_lang_0128 %></h3></th>
			</tr>
		</thead>
		<tbody>
<%
i = 0
do while i < allnum
	ei.getMailInfoEx allnum - i - 1, idname, isread, priority, sendMail, sendName, subject, size, etime, mstate

	if subject = "" then
		subject = s_lang_0129
	end if

	Response.Write "<tr><td><input type=""checkbox"" id=""check_" & i & """ value=""" & server.htmlencode(getZattAppSize(size)) & "|" & server.htmlencode(idname) & """></td><td>"

	Response.Write server.htmlencode(subject) & "</td>"
	Response.Write "<td align=""right"" noWrap>" & server.htmlencode(getShowSize(size)) & "</td>"
	Response.Write "<td noWrap><span style=""display:none;"">" & etime & "</span>" & server.htmlencode(conv_show_date(etime)) & "</td></tr>" & Chr(13)

	idname = NULL
	isread = NULL
	priority = NULL
	sendMail = NULL
	sendName = NULL
	subject = NULL
	size = NULL
	etime = NULL
	mstate = NULL

	i = i + 1
loop
%></tbody>
  </table>
	</div>

	<div id="view2" class="tabcontent">

<div style="padding-left:3px; padding-bottom:2px; font-family:<%=s_lang_font %>; font-size:9pt">
<select id="columns2" style="display:none;"><option value="1"></option></select>
<%=s_lang_0295 %><%=s_lang_mh %><input type="text" class="textbox_wwm" id="query2" onkeyup="sorter2.search('query2')">
</div>

	<table align="center" id="table2" class="tinytable" border="0" cellpadding="0" cellspacing="0">
		<thead>
			<tr>
				<th width="4%" class="nosort"><h3>&nbsp;</h3></th>
				<th width="63%" noWrap><h3><%=s_lang_0067 %></h3></th>
				<th width="13%" noWrap><h3><%=s_lang_0179 %></h3></th>
				<th width="20%" noWrap><h3><%=s_lang_0128 %></h3></th>
			</tr>
		</thead>
		<tbody>
<%
ei.isLoadZatt = false
ei.LoadMailBox Session("wem"), "att"
allnum = ei.getMailsCount
i = 0

do while i < allnum
	ei.getMailInfoEx allnum - i - 1, idname, isread, priority, sendMail, sendName, subject, size, etime, mstate

	if subject = "" then
		subject = s_lang_0129
	end if

	Response.Write "<tr><td><input type=""checkbox"" id=""check_n" & i & """ value=""" & server.htmlencode(getZattAppSize(size)) & "|" & server.htmlencode(idname) & """></td><td>"

	Response.Write server.htmlencode(subject) & "</td>"
	Response.Write "<td align=""right"" noWrap>" & server.htmlencode(getShowSize(size)) & "</td>"
	Response.Write "<td noWrap><span style=""display:none;"">" & etime & "</span>" & server.htmlencode(conv_show_date(etime)) & "</td></tr>" & Chr(13)

	idname = NULL
	isread = NULL
	priority = NULL
	sendMail = NULL
	sendName = NULL
	subject = NULL
	size = NULL
	etime = NULL
	mstate = NULL

	i = i + 1
loop
%></tbody>
  </table>
	</div>
</div>

</form>
<table width="100%" cellpadding=0 cellspacing=0 align="center">
<tr>
<td align="right" style="padding-top: 10px;">
<a class="wwm_btnDownload btn_blue" style="WIDTH: 40px" href="javascript:gook()"><%=s_lang_0059 %></a>&nbsp;
<a class="wwm_btnDownload btn_blue" style="WIDTH: 40px" href="javascript:self.close();"><%=s_lang_cancel %></a>
</td>
</tr>
</table>
</BODY>

<script src="images/tab_content.js" type="text/javascript"></script>
<script type="text/javascript" src="images/slscript.js"></script>
<script type="text/javascript">
var ie = function () {
	var v = 4,
		div = document.createElement('div'),
		i = div.getElementsByTagName('i');
	do {
		div.innerHTML = '<!--[if gt IE ' + (++v) + ']><i></i><![endif]-->';
	} while (i[0]);
	return v > 5 ? v : false;
}();

if (ie == 6)
	document.execCommand("BackgroundImageCache", false, true); 

	var sorter = new TINY.table.sorter('sorter','table',{
		headclass:'head',
		ascclass:'asc',
		descclass:'desc',
		evenclass:'evenrow',
		oddclass:'oddrow',
		evenselclass:'evenselected',
		oddselclass:'oddselected',
		paginate:true,
		size:9999,
		colddid:'columns',
		hoverid:'selectedrow',
		sortcolumn:1,
		sortdir:1,
		init:true
	});

sorter.sort(3, false);

	var sorter2 = new TINY.table.sorter('sorter2','table2',{
		headclass:'head',
		ascclass:'asc',
		descclass:'desc',
		evenclass:'evenrow',
		oddclass:'oddrow',
		evenselclass:'evenselected',
		oddselclass:'oddselected',
		paginate:true,
		size:9999,
		colddid:'columns2',
		hoverid:'selectedrow',
		sortcolumn:1,
		sortdir:1,
		init:true
	});

sorter2.sort(3, false);

function _save_mode(s_col, s_mode){}
</script>

<script language="JavaScript">
<!--
function have_it(ck_array, one_val)
{
	for (var i = 0; i < ck_array.length; i++)
	{
		if (ck_array[i] == one_val)
			return true;
	}

	return false;
}

function gook()
{
	var ck_array = [];
	var i = 0;
	var theObj;
	var is_app = false;
	var f_index;
	var tempstr;

	for(; i<<%=zallnum %>; i++)
	{
		theObj = eval("document.f1.check_" + i);

		if (theObj != null)
		{
			if (theObj.checked == true)
			{
				f_index = theObj.value.indexOf("|");
				if (f_index > 0)
				{
					tempstr = theObj.value.substr(f_index + 1, theObj.value.length - f_index);
					if (have_it(ck_array, tempstr) == false)
					{
						ck_array.push(tempstr);
						opener.add_zatt_select_from_pop(tempstr, theObj.value.substr(0, f_index));
						is_app = true;
					}
				}
			}
		}
	}

	for(i = 0; i<<%=allnum %>; i++)
	{
		theObj = eval("document.f1.check_n" + i);

		if (theObj != null)
		{
			if (theObj.checked == true)
			{
				f_index = theObj.value.indexOf("|");
				if (f_index > 0)
				{
					tempstr = theObj.value.substr(f_index + 1, theObj.value.length - f_index);
					if (have_it(ck_array, tempstr) == false)
					{
						ck_array.push(tempstr);
						opener.add_zatt_select_from_pop(tempstr, theObj.value.substr(0, f_index));
						is_app = true;
					}
				}
			}
		}
	}

	if (is_app == true)
		opener.flash_zatt_div();

	self.close();
}
// -->
</script>
</HTML>

<%
set ei = nothing

function getShowSize(bytesize)
	if bytesize < 1000 then
		getShowSize = "1K"
	else
		getShowSize = CLng(bytesize/1000) & "K"
	end if
end function

function conv_show_date(datastr)
	sl_y = Left(datastr, 4)
	sl_m = Mid(datastr, 5, 2)
	sl_d = Mid(datastr, 7, 2)

	if Left(sl_m, 1) = "0" then
		sl_m = Right(sl_m, 1)
	end if

	if Left(sl_d, 1) = "0" then
		sl_d = Right(sl_d, 1)
	end if

	conv_show_date = sl_y & s_lang_0139 & sl_m & s_lang_0140 & sl_d & s_lang_0141
end function

function getZattAppSize(bytesize)
	if bytesize < 1000 then
		getZattAppSize = "1K"
	else
		if bytesize < 1000000 then
			getZattAppSize = CLng(bytesize/1000) & "K"
		else
			tmpSize = CStr(CDbl(bytesize/1000000))
			tmpindex = InStr(1, tmpSize, ".")
			if tmpindex = 0 then
				getZattAppSize = tmpSize & "M"
			else
				getZattAppSize = Left(tmpSize, tmpindex + 2) & "M"
			end if
		end if
	end if
end function
%>
