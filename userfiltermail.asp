<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" -->

<%
issave = trim(Request("issave"))
mode = trim(Request("mode"))

dim fm
set fm = server.createobject("easymail.MailFilterManager")
fm.Load Session("wem")

dim mlb
set mlb = server.createobject("easymail.Labels")
mlb.Load Session("wem")

if mode = "del" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	id = trim(Request("id"))

	if id <> "" then
		if IsNumeric(id) = true then
			if fm.Remove(CInt(id)) = true then
				fm.Save
				set fm = nothing
				set mlb = nothing
				response.redirect "ok.asp?" & getGRSN() & "&gourl=userfiltermail.asp?" & getGRSN()
			end if
		end if
	end if
end if

if mode = "up" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	id = trim(Request("id"))

	if id <> "" then
		if IsNumeric(id) = true then
			if fm.MoveUp(CInt(id)) = true then
				fm.Save
			end if
		end if
	end if
end if

if mode = "down" and Request.ServerVariables("REQUEST_METHOD") = "GET" then
	id = trim(Request("id"))

	if id <> "" then
		if IsNumeric(id) = true then
			if fm.MoveDown(CInt(id)) = true then
				fm.Save
			end if
		end if
	end if
end if


if mode = "mdel" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	themax = fm.count
	isedit = false

	do while themax >= 0
		if trim(request("check" & themax)) <> "" then
			md = trim(request("check" & themax))

			if IsNumeric(md) = true then
				if fm.Remove(CInt(md)) = true then
					isedit = true
				end if
			end if
		end if 

	    themax = themax - 1
	loop

	if isedit = true then
		fm.Save
		set fm = nothing
		set mlb = nothing
		response.redirect "ok.asp?" & getGRSN() & "&gourl=userfiltermail.asp?" & getGRSN()
	end if
end if

if issave = "1" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	n_content = trim(Request("s_content"))
	n_mode = trim(Request("s_mode"))
	n_deal = trim(Request("s_deal"))
	n_msg = trim(Request("s_msg"))
	n_continue = trim(Request("s_continue"))
	isok = false

	if n_content <> "" and n_mode <> "" and n_deal <> "" then
		if IsNumeric(n_content) = true and IsNumeric(n_mode) = true then
			if IsNumeric(n_deal) = true then
				isok = fm.Add(CInt(n_content), CInt(n_mode), CInt(n_deal), CInt(n_continue), n_msg)
			else
				if Mid(n_deal, 2, 1) = "~" then
					tn_deal = CInt(Mid(n_deal, 1, 1))
					if tn_deal = 3 then
						isok = fm.Add2(CInt(n_content), CInt(n_mode), tn_deal, CInt(n_continue), n_msg, "", Mid(n_deal, 3))
					elseif tn_deal = 4 then
						isok = fm.Add2(CInt(n_content), CInt(n_mode), 4, CInt(n_continue), n_msg, Mid(n_deal, 3), "")
					end if
				else
					isok = fm.Add2(CInt(n_content), CInt(n_mode), 2, CInt(n_continue), n_msg, Mid(n_deal, 5), "")
				end if
			end if
		end if
	end if

	if isok = true then
		fm.Save
		set fm = nothing
		set mlb = nothing
		response.redirect "ok.asp?" & getGRSN() & "&gourl=userfiltermail.asp?" & getGRSN()
	end if
end if

allnum = fm.count

dim arex
set arex = server.createobject("easymail.AutoReplyEx")
arex.Load Session("wem")
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
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.cont_td {white-space:nowrap; height:28px; padding-left:5px; padding-right:5px;}
.cont_td_word {height:28px; padding-left:5px; padding-right:5px; word-break:break-all; word-wrap:break-word;}
.ctd {height:24px; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px;}
.ctd_word {height:24px; border-bottom:1px solid #A5B6C8; padding-left:5px; padding-right:5px; word-break:break-all; word-wrap:break-word;}
.wwm_color_in_line{padding:1px; border-radius:3px; -webkit-border-radius:3px; display:-moz-inline-box; display:inline-block; padding-left:2px; padding-right:2px; height:8px; width:6px; font-size:0pt;}
-->
</STYLE>
</HEAD>

<script type="text/javascript" src="images/mglobal.js"></script>

<script type="text/javascript">
<!--
function mdel() {
	if (ischeck() == true)
	{
		document.f1.mode.value = "mdel";
		document.f1.submit();
	}
}

function ischeck() {
	var i = 0;
	var theObj;

	for(; i<<%=allnum %>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}

function addnew() {
	var isok = false;

	if (document.fnew.s_msg.value == "" && document.fnew.s_content.value != 3 && (document.fnew.s_mode.value == 1 || document.fnew.s_mode.value == 2))
		isok = true;

	if (document.fnew.s_msg.value != "")
		isok = true;

	if (isok == true)
	{
		if (document.fnew.s_content.value != 3)
		{
			document.fnew.submit();
		}
		else
		{
			if (document.fnew.s_msg.value >= 0)
				document.fnew.submit();
		}
	}
}

function add_s_mode(a_value, a_text)
{
	var oOption = document.createElement("OPTION");
	oOption.text = a_text;
	oOption.value = a_value;

	if (ie == false)
		document.getElementById("s_mode").appendChild(oOption);
	else
		document.getElementById("s_mode").add(oOption);
}

function select_content_onchange()
{
	var i = 0;
	for (i; i < document.fnew.s_mode.length; i++)
	{
		document.fnew.s_mode.remove(i);
		i--;
	}

	if (document.fnew.s_content.value != "3")
	{
		add_s_mode("1", "<%=b_lang_240 %>");
		add_s_mode("2", "<%=b_lang_241 %>");
		add_s_mode("3", "<%=b_lang_242 %>");
		add_s_mode("4", "<%=b_lang_243 %>");
		add_s_mode("5", "<%=b_lang_244 %>");
	}
	else
	{
		add_s_mode("1", "<%=b_lang_245 %>");
		add_s_mode("6", "<%=b_lang_246 %>");
		add_s_mode("7", "<%=b_lang_247 %>");
	}
}

function allcheck_onclick() {
	if (document.f1.allcheck.checked == true)
		checkall(true);
	else
		checkall(false);
}

function checkall(check) {
	var i = 0;
	var theObj;

	for(; i<<%=allnum %>; i++)
	{
		theObj = eval("document.f1.check" + i);

		if (theObj != null)
			theObj.checked = check;
	}
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
<table width="98%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_248 %>
</td></tr>
<tr><td class="block_top_td" style="height:8px; _height:10px;"></td></tr>
<tr><td align="center">

<FORM ACTION="userfiltermail.asp" METHOD="POST" NAME="fnew">
<input type="hidden" name="issave" value="1">
	<table width="98%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr><td align="left" class="cont_td">
	<%=b_lang_249 %><%=s_lang_mh %>
	<select name="s_content" class=drpdwn LANGUAGE=javascript onchange="select_content_onchange()">
<option value="1"><%=b_lang_250 %></option>
<option value="2"><%=b_lang_251 %></option>
<option value="3"><%=b_lang_252 %></option>
<option value="4"><%=b_lang_253 %></option>
</select>
<select id="s_mode" name="s_mode" class=drpdwn>
<option value="1"><%=b_lang_240 %></option>
<option value="2"><%=b_lang_241 %></option>
<option value="3"><%=b_lang_242 %></option>
<option value="4"><%=b_lang_243 %></option>
<option value="5"><%=b_lang_244 %></option>
</select>
<input type="input" name="s_msg" class="n_textbox">
	</td></tr>
	<tr><td align="left" class="cont_td">
	<%=b_lang_254 %><%=s_lang_mh %>
	<select name="s_deal" class=drpdwn>
<option value="1"><%=s_lang_del %></option>
<option value="2"><%=b_lang_176 %></option>
<%
i = 0
allnum = mlb.Count
do while i < allnum
	mlb.GetByIndex i, ret_id, ret_title, ret_color
	Response.Write "<option value='4~" & server.htmlencode(ret_id) & "'>" & b_lang_255 & s_lang_mh & server.htmlencode(ret_title) & "</option>"

	ret_id = NULL
	ret_title = NULL
	ret_color = NULL

	i = i + 1
loop
%>
<option value="3~del"><%=b_lang_256 %></option>
<option value="3~in"><%=b_lang_257 %></option>
<option value="3~out"><%=b_lang_258 %></option>
<option value="3~sed"><%=b_lang_259 %></option>
<%
dim pf
set pf = server.createobject("easymail.PerFolders")
pf.Load Session("wem")

pfNumber = pf.FolderCount
i = 0

do while i < pfNumber
	spfname = pf.GetFolderName(i)
	response.write "<option value='3~" & spfname & "'>" & b_lang_260 & " " & server.htmlencode(spfname) & "</option>" & Chr(13)

	spfname = NULL
	i = i + 1
loop

set pf = nothing


i = 0
arallnum = arex.count

do while i < arallnum
	arex.Get i, are_name, are_subject, are_text

	response.write "<option value='arx:" & server.htmlencode(are_name) & "'>" & b_lang_261 & s_lang_mh & server.htmlencode(are_name) & "</option>" & Chr(13)

	are_name = NULL
	are_subject = NULL
	are_text = NULL

	i = i + 1
loop
%>
</select>
<select name="s_continue" class=drpdwn>
<option value="0"><%=b_lang_262 %></option>
<option value="1"><%=b_lang_263 %></option>
</select>
	</td></tr>
	<tr><td align="left" class="cont_td">
<a class='wwm_btnDownload btn_gray' href="javascript:addnew();"><%=b_lang_264 %></a>
	</td></tr>
</table>
</form>
</td></tr>
</table>
<br>

<table width="98%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_265 %>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="center">

<FORM ACTION="userfiltermail.asp" METHOD="POST" NAME="f1">
<input type="hidden" name="mode">
	<table width="98%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr class="title_tr">
	<td width="5%" class="st_l"><input type="checkbox" id="allcheck" LANGUAGE=javascript onclick="return allcheck_onclick()"></td>
	<td width="5%" class="st_l"><%=b_lang_040 %></td>
	<td width="61%" class="st_l"><%=b_lang_266 %></td>
	<td width="14%" class="st_l"><%=b_lang_267 %></td>
	<td width="10%" class="st_l"><%=b_lang_268 %></td>
	<td width="5%" class="st_r"><%=s_lang_del %></td>
	</tr>
<%
i = 0
allnum = fm.count

do while i < allnum
	fm.GetInfo2 i, filter_content, filter_mode, filter_deal, filter_continue, filter_msg, filter_autoreply_name, filter_perfolder_name

	Response.Write "<tr class='cont_tr' onmouseover='m_over(this);' onmouseout='m_out(this);'>"
	Response.Write "	<td align='center' class='ctd'><input type='checkbox' name='check" & i & "' value='" & i & "'></td>"
	Response.Write "	<td align='center' class='ctd'>" & i + 1 & "</td>"

	Response.Write "	<td align='left' class='ctd_word'>" & getFilterStr(filter_content, filter_mode, filter_deal, filter_msg, filter_autoreply_name, filter_perfolder_name) & "</td>"

	if filter_continue = 0 then
		Response.Write "	<td align='center' class='ctd'><font class='s'>" & b_lang_262 & "</font></td>"
	else
		Response.Write "	<td align='center' class='ctd'><font class='s'>" & b_lang_263 & "</font></td>"
	end if

	if allnum = 1 then
		Response.Write "	<td align='center' class='ctd'>&nbsp;</td>"
	elseif i = 0 then
		Response.Write "	<td align='center' class='ctd'>&nbsp;&nbsp;&nbsp;&nbsp;<a href='userfiltermail.asp?mode=down&id=" & i & "'><img src='images/arrow_down.gif' border='0' align='absmiddle'></a></td>"
	elseif i = allnum - 1 then
		Response.Write "	<td align='center' class='ctd'><a href='userfiltermail.asp?mode=up&id=" & i & "'><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>&nbsp;&nbsp;&nbsp;&nbsp;</td>"
	else
		Response.Write "	<td align='center' class='ctd'><a href='userfiltermail.asp?mode=up&id=" & i & "'><img src='images/arrow_up.gif' border='0' align='absmiddle'></a>&nbsp;&nbsp;<a href='userfiltermail.asp?mode=down&id=" & i & "'><img src='images/arrow_down.gif' border='0' align='absmiddle'></a></td>"
	end if

	Response.Write "	<td align='center' class='ctd'><a href='userfiltermail.asp?mode=del&id=" & i & "&" & getGRSN() & "'><img src='images/del.gif' border='0' title='" & s_lang_del & "'></a></td>"
	Response.Write "</tr>" & Chr(13)

	filter_content = NULL
	filter_mode = NULL
	filter_deal = NULL
	filter_msg = NULL
	filter_autoreply_name = NULL
	filter_perfolder_name = NULL

    i = i + 1
loop
%>
	</td></tr>

<tr><td colspan="6" align="left" style="background-color:white; padding-right:16px; padding-top:16px; padding-bottom:10px;">
<a class='wwm_btnDownload btn_blue' href="user_right.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
<a class='wwm_btnDownload btn_blue' href="javascript:mdel();"><%=s_lang_del %></a>
<a class='wwm_btnDownload btn_blue' href="showautoreplyex.asp?returl=userfiltermail.asp&<%=getGRSN() %>"><%=b_lang_269 %></a>
	</td></tr>
	</table>
</form>
</td></tr>
</table>

<table width="96%" border="0" align="center" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px #8CA5B5 solid; margin-top:30px; padding-bottom:10px;'>
	<tr>
	<td valign="top" style="padding:4px; width:82px;"><font color="#901111">*<%=b_lang_270 %></font></td>
	<td style="padding:4px; color:#444444;"><%=b_lang_271 %><br>
	</td></tr>
</table>

<div style="position:absolute; left:8px; top:6px;">
<a href="help.asp#userfiltermail" target="_blank"><img src="images/help.gif" border="0" title="<%=s_lang_help %>"></a></div>
</BODY>
</HTML>

<%
set fm = nothing
set arex = nothing


function getFilterStr(f_content, f_mode, f_deal, f_msg, filter_autoreply_name, filter_perfolder_name)
	getFilterStr = ""

	if f_msg = "" then
		Exit Function
	end if

	if f_content = 1 then
		getFilterStr = "邮件地址"
	elseif f_content = 2 then
		getFilterStr = "发件人"
	elseif f_content = 3 then
		getFilterStr = "邮件大小"
	elseif f_content = 4 then
		getFilterStr = "主题"
	else
		Exit Function
	end if

	if f_mode = 1 then
		getFilterStr = getFilterStr & "等于"
	elseif f_mode = 2 then
		getFilterStr = getFilterStr & "不等于"
	elseif f_mode = 3 then
		getFilterStr = getFilterStr & "包含"
	elseif f_mode = 4 then
		getFilterStr = getFilterStr & "不包含"
	elseif f_mode = 5 then
		getFilterStr = getFilterStr & "通配符等于"
	elseif f_mode = 6 then
		getFilterStr = getFilterStr & "大于"
	elseif f_mode = 7 then
		getFilterStr = getFilterStr & "小于"
	else
		getFilterStr = ""
		Exit Function
	end if

	if msg = "" then
		msg = "[Empty]"
	end if

	getFilterStr = getFilterStr & f_msg

	if f_deal = 1 then
		getFilterStr = getFilterStr & "时, 则删除"
		getFilterStr = server.htmlencode(getFilterStr)
	elseif f_deal = 2 then
		if filter_autoreply_name = "" then
			getFilterStr = getFilterStr & "时, 则自动回复"
		else
			getFilterStr = getFilterStr & "时, 则使用增强型自动回复: " & filter_autoreply_name
		end if
		getFilterStr = server.htmlencode(getFilterStr)
	elseif f_deal = 3 then
		if filter_perfolder_name = "" then
			getFilterStr = getFilterStr & "时, 则移到垃圾箱"
		else
			if filter_perfolder_name = "del" then
				getFilterStr = getFilterStr & "时, 则移到垃圾箱"
			elseif filter_perfolder_name = "in" then
				getFilterStr = getFilterStr & "时, 则移到收件箱"
			elseif filter_perfolder_name = "out" then
				getFilterStr = getFilterStr & "时, 则移到草稿箱"
			elseif filter_perfolder_name = "sed" then
				getFilterStr = getFilterStr & "时, 则移到发件箱"
			else
				getFilterStr = getFilterStr & "时, 则移到" & filter_perfolder_name
			end if
		end if
		getFilterStr = server.htmlencode(getFilterStr)
	elseif f_deal = 4 then
		ret_id = ""
		ret_title = ""
		ret_color = ""

		mlb.GetByID filter_autoreply_name, ret_id, ret_title, ret_color
		getFilterStr = getFilterStr & "时, 则添加标签:"
		getFilterStr = server.htmlencode(getFilterStr)

		if Len(ret_id) > 0 then
			if Len(ret_title) > 0 then
				getFilterStr = getFilterStr & "&nbsp;<span class='wwm_color_in_line' style='background:#" & ret_color & ";'></span>&nbsp;" & server.htmlencode(ret_title)
			else
				getFilterStr = getFilterStr & "&nbsp;<span class='wwm_color_in_line' style='background:#" & ret_color & ";'></span>"
			end if
		else
			getFilterStr = getFilterStr & "&nbsp;" & server.htmlencode("[标签已失效]")
		end if

		ret_id = NULL
		ret_title = NULL
		ret_color = NULL
	else
		getFilterStr = ""
		Exit Function
	end if
end function

set mlb = nothing
%>
