<!--#include file="passinc.asp" --> 
<!--#include file="language-1.asp" --> 

<%
if isadmin() = false then
	Response.Redirect "noadmin.asp"
end if
%>

<%
dim eads
set eads = server.createobject("easymail.EntAddress")
eads.Load

mode = trim(request("mode"))
rq_pid = trim(request("pid"))

if IsNumeric(trim(request("pid"))) = true then
	rq_pid = CLng(trim(request("pid")))
else
	rq_pid = -1
end if

if rq_pid > -1 and mode <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	allnum = eads.GetEmailsCountByFolderID(rq_pid)
	eads.RemoveEmailsByFolderID rq_pid

	i = 0
	if mode = "save" or mode = "add" then
		do while i < allnum + 1
			eemail = trim(request("email" & i))

			if eemail <> "" then
				ename = trim(request("name" & i))
				ecomment = trim(request("comment" & i))
				eemail = replace(eemail, "'", "")
				eemail = replace(eemail, "|", "")
				eemail = replace(eemail, """", "")
				ename = replace(ename, """", "'")
				ecomment = replace(ecomment, "|", "")
				ecomment = replace(ecomment, """", "'")

				eads.AddAds eemail, ename, ecomment, rq_pid
			end if

		    i = i + 1
		loop
	elseif mode = "del" then
		do while i < allnum + 1
			if trim(request("check" & i)) = "" then
				eemail = trim(request("email" & i))

				if eemail <> "" then
					ename = trim(request("name" & i))
					ecomment = trim(request("comment" & i))
					eemail = replace(eemail, "'", "")
					eemail = replace(eemail, "|", "")
					eemail = replace(eemail, """", "")
					ename = replace(ename, """", "'")
					ecomment = replace(ecomment, "|", "")
					ecomment = replace(ecomment, """", "'")

					eads.AddAds eemail, ename, ecomment, rq_pid
				end if
			end if

		    i = i + 1
		loop
	end if

	isok = eads.Save()

	if trim(request("mode")) <> "add" then
		set eads = nothing

		if isok = true then
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("ea_ads.asp?pid=" & rq_pid & "&" & getGRSN())
		else
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=" & Server.URLEncode("ea_ads.asp?pid=" & rq_pid & "&" & getGRSN())
		end if
	end if
end if


set_mode = trim(Request.Form("setmode"))
if set_mode <> "" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	tmp_max = trim(Request.Form("max"))
	i = 0

	if IsNull(tmp_max) = false and IsNumeric(tmp_max) = true then
		themax = CLng(tmp_max)
	end if

	if set_mode = "1" then
		do while i < themax
			tmp_n = UnEscape(trim(Request.Form("n_" & i)))
			tmp_p = UnEscape(trim(Request.Form("p_" & i)))
			tmp_c = UnEscape(trim(Request.Form("c_" & i)))

			tmp_pid = eads.GetFolderID(tmp_p)
			if tmp_pid >= -1 then
				eads.AddFolder tmp_n, tmp_c, tmp_pid
			end if

		    i = i + 1
		loop
	elseif set_mode = "2" then
		do while i < themax
			tmp_f = UnEscape(trim(Request.Form("f_" & i)))
			tmp_e = UnEscape(trim(Request.Form("e_" & i)))
			tmp_n = UnEscape(trim(Request.Form("n_" & i)))
			tmp_c = UnEscape(trim(Request.Form("c_" & i)))

			tmp_fid = eads.GetFolderID(tmp_f)
			if tmp_fid >= -1 then
				eads.AddAds tmp_e, tmp_n, tmp_c, tmp_fid
			end if

		    i = i + 1
		loop
	end if

	if eads.Save() = true then
		Response.Write "1"
	else
		Response.Write "0"
	end if

	set eads = nothing
	Response.End
end if

allnum = eads.GetEmailsCountByFolderID(rq_pid)
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
<link rel="stylesheet" type="text/css" href="images/selads.css">

<STYLE type=text/css>
<!--
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.title_tr {white-space:nowrap; background:#f2f4f6; height:26px;}
.title_td {text-align:center; white-space:nowrap; border-left:1px solid #A5B6C8; border-top:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8;}
.con_td {text-align:center; white-space:nowrap; height:24px; border-bottom:1px solid #A5B6C8;}
.Bsbttn {font-family:<%=s_lang_font %>; font-size:10pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5; color:#000066;text-decoration:none;cursor:pointer}
.wwm_msg {padding:8px; margin:-6px 0 14px 0; color:#7E4F05; line-height:18px; background:#FFF3C3;border-radius:4px; -webkit-border-radius:4px;padding-left:20px;padding-right:20px;text-align:left;border: #7E4F05 1px solid;}
.wwm_line_msg {padding:18px; width:320px; margin:0 17px 0 17px; color:#202020; font-size:10pt; line-height:10px; background:#e0ecf9; border-radius:4px; -webkit-border-radius:4px; padding-left:20px; padding-right:20px; text-align:center;}
-->
</STYLE>
</HEAD>

<script type="text/javascript" src="images/mglobal.js"></script>
<script type="text/javascript" src="images/jquery.min.js"></script>

<script type="text/javascript">
<!--
var foldersel = "<select id='sfid' name='sfid' class='drpdwn' size='1'><%=fstr %></select>";

var clickfid = -1;
var Stag;

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

function allcheck_onclick() {
	if (document.f1.allcheck.checked == true)
		checkall(true);
	else
		checkall(false);
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

function selopt()
{
	var optval = document.getElementById("sfid").value;
	location.href = "ea_ads.asp?<%=getGRSN() %>&pid=" + optval;

	if (optval.length < 1)
	{
		document.getElementById("abt_span").style.display = "none";
		document.getElementById("abt_imp").style.display = "inline";
		document.getElementById("imp_id").style.display = "none";
	}
	else
	{
		document.getElementById("abt_span").style.display = "inline";
		document.getElementById("abt_imp").style.display = "none";
		document.getElementById("imp_id").style.display = "none";
	}
}

function del() {
	if (ischeck() == true)
	{
		document.f1.pid.value = document.getElementById("sfid").value;
		document.f1.mode.value = "del";
		document.f1.submit();
	}
}

function add() {
	document.f1.pid.value = document.getElementById("sfid").value;
	document.f1.mode.value = "add";
	document.f1.submit();
}

function save() {
	document.f1.pid.value = document.getElementById("sfid").value;
	document.f1.mode.value = "save";
	document.f1.submit();
}

function select_convfrom() {
	if (document.f1.convfrom.selectedIndex == 5)
		document.getElementById("ldap_com").innerHTML = "<select name=\"ldapsel\" id=\"ldapsel\" class=\"drpdwn\" size=\"1\"><option selected><%=s_lang_0069 %><\/option><option value=\"homePhone\"><%=s_lang_0037 %><\/option><option value=\"pager\"><%=s_lang_0038 %><\/option><option value=\"mobile\"><%=s_lang_0039 %><\/option><option value=\"telephoneNumber\"><%=s_lang_0040 %><\/option><option value=\"facsimileTelephoneNumber\"><%=s_lang_0041 %><\/option><option value=\"title\"><%=s_lang_0042 %><\/option><option value=\"physicalDeliveryOfficeName\"><%=s_lang_0043 %><\/option><option value=\"sn\"><%=s_lang_0045 %><\/option><option value=\"givenName\"><%=s_lang_0046 %><\/option><\/select>";
	else
		document.getElementById("ldap_com").innerHTML = "<input type=\"hidden\" name=\"ldapsel\" id=\"ldapsel\">";

	if (document.f1.convfrom.selectedIndex > 0)
		document.f1.bt_conv.disabled = false;
	else
		document.f1.bt_conv.disabled = true;
}

function convads() {
	if (document.f1.convfrom.selectedIndex == 5)
		location.href = "ea_conv.asp?<%=getGRSN() %>&pid=<%=rq_pid %>&cf=" + document.f1.convfrom.value + "&ldapsel=" + document.f1.ldapsel.value + "&gourl=<%=Server.URLEncode("ea_ads.asp?pid=" & rq_pid) %>";
	else if (document.f1.convfrom.selectedIndex == 4)
	{
		document.getElementById("tb_systemuser").style.display = "block";
		document.getElementById("pop_msg_id").style.height = "126px";
		show_pop_msg();
	}
	else
	{
		document.getElementById("tb_systemuser").style.display = "none";
		document.getElementById("pop_msg_id").style.height = "30px";
		show_pop_msg();
	}
}

function window_onload() {
<%
if rq_pid < 0 then
%>
	document.getElementById("abt_span").style.display = "none";
	document.getElementById("abt_imp").style.display = "inline";
	document.getElementById("imp_id").style.display = "none";
<%
else
%>
	document.getElementById("abt_span").style.display = "inline";
	document.getElementById("abt_imp").style.display = "none";
	document.getElementById("imp_id").style.display = "none";
<%
end if
%>

	var theObj = eval("document.f1.name<%=allnum %>");
	if (theObj != null)
		theObj.focus();
}

function impall() {
	var theObj = document.getElementById("imp_id");
	if (theObj.style.display == "inline")
		theObj.style.display = "none";
	else
		theObj.style.display = "inline";
}

var folder_arr;
var ads_arr;

function split_all(instr) {
	var result = instr.split("\n");
	folder_arr = new Array();
	ads_arr = new Array();
	var find_folder_end = false;

	for(var i = 0; i < result.length; i++)
	{
		if (find_folder_end == false && result[i].length < 1)
			find_folder_end = true;

		if (find_folder_end == false)
			folder_arr.push(result[i]);
		else
		{
			if (result[i].length > 1)
				ads_arr.push(result[i]);
		}
	}
}

function split_one(instr) {
	var result = instr.split("|");
	var onearr = new Array();
	for(var i = 0; i < result.length; i++)
	{
		onearr.push(result[i]);
	}
	return onearr;
}

function get_folder_str()
{
	var ret_val = "max=" + folder_arr.length;

	for(var i = 0; i < folder_arr.length; i++)
	{
		var arr = split_one(folder_arr[i]);

		ret_val += "&n_" + i + "=" + escape(arr[0]);
		ret_val += "&p_" + i + "=" + escape(arr[1]);
		ret_val += "&c_" + i + "=" + escape(arr[2]);
	}

	return ret_val;
}

var ads_read_index = 0;
function get_ads_str()
{
	var show_i = 0;
	var max_read = 100;
	var ret_val;

	if ((ads_read_index + max_read) < ads_arr.length)
		ret_val = "max=" + max_read;
	else
		ret_val = "max=" + (ads_arr.length - ads_read_index);

	for(var i = ads_read_index; i < ads_arr.length; i++)
	{
		var arr = split_one(ads_arr[i]);

		ret_val += "&f_" + show_i + "=" + escape(arr[0]);
		ret_val += "&e_" + show_i + "=" + escape(arr[1]);
		ret_val += "&n_" + show_i + "=" + escape(arr[2]);
		ret_val += "&c_" + show_i + "=" + escape(arr[3]);

		ads_read_index = i + 1;
		show_i++;

		if (show_i >= max_read)
			break;
	}

	return ret_val;
}

function do_imp() {
	var str = document.getElementById("imp_text").value;
	if (str.length < 1)
		return ;

	document.getElementById("show_proc").style.display = "inline";

	str = str.replace(/\n\r/g,"\n");
	str = str.replace(/\r\n/g,"\n");
	str = str.replace(/\r/g,"\n");

	split_all(str);
	var post_data;

	if (folder_arr.length > 0)
	{
		post_data = "setmode=1&<%=getGRSN() %>&" + get_folder_str();
		send_ajax(post_data);
	}

	setTimeout("do_imp_2()", 2000);
}

function do_imp_2() {
	ads_read_index = 0;
	while(ads_read_index < ads_arr.length)
	{
		post_data = "setmode=2&<%=getGRSN() %>&" + get_ads_str();
		send_ajax(post_data);
	}

	setTimeout("showproc()", 500);
}

function send_ajax(post_data)
{
$.ajax({
	type:"POST",
	url:"ea_ads.asp",
	data:post_data,
	success:function(data){
	},
	error:function(){
	}
});
}

function showproc() {
	location.href = "ea_brow.asp?<%=getGRSN() %>";
}

function show_pop_msg() {
	document.getElementById('pop_msg_div').style.display='block';
	document.getElementById('popIframe').style.display='block';
	document.getElementById('bg').style.display='block';

	var v_height = document.body.clientHeight

	if (document.documentElement.clientHeight > v_height)
		v_height = document.documentElement.clientHeight;

	if (ie == 6 && v_height > document.documentElement.clientHeight)
		v_height += 21;

	document.getElementById('popIframe').style.height = v_height + "px";
	document.getElementById('popIframe').style.width = document.documentElement.scrollWidth + "px";

	document.getElementById('bg').style.height = v_height + "px";
	document.getElementById('bg').style.width = document.documentElement.scrollWidth + "px";
}

function close_pop_msg(need_sub){
	document.getElementById('pop_msg_div').style.display='none';
	document.getElementById('bg').style.display='none';
	document.getElementById('popIframe').style.display='none';

	if (need_sub == 1)
	{
		var url_add = "&ns=";
		if (document.getElementById("cv_nosame").checked == true)
			url_add += "1";
		else
			url_add += "0";

		url_add += "&cm=";
		if (document.getElementById("cv_comm").checked == true)
			url_add += "1";
		else
			url_add += "0";

		url_add += "&nd=";
		if (document.getElementById("cv_nodisabled").checked == true)
			url_add += "1";
		else
			url_add += "0";

		url_add += "&dm=" + escape(document.getElementById("cv_domain").value);
		location.href = "ea_conv.asp?<%=getGRSN() %>&pid=<%=rq_pid %>&cf=" + document.f1.convfrom.value + "&ldapsel=" + document.f1.ldapsel.value + url_add + "&gourl=<%=Server.URLEncode("ea_ads.asp?pid=" & rq_pid) %>";
	}
}
//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload()">
<FORM ACTION="ea_ads.asp" METHOD="POST" NAME="f1">
<input type="hidden" name="mode">
<input type="hidden" name="pid">
<table width="96%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr>
	<td nowrap width="8%" align="left" height="28" style="padding-left:4px;">
	<a class='wwm_btnDownload btn_blue' href="ea_brow.asp?<%=getGRSN() %>"><< <%=s_lang_return %></a>
	<span id='abt_imp'>
	<a class='wwm_btnDownload btn_blue' href="javascript:impall();"><%=a_lang_167 %></a>
	</span>
	<span id='abt_span'>
	<a class='wwm_btnDownload btn_blue' href="javascript:add();"><%=s_lang_add %></a>
	<a class='wwm_btnDownload btn_blue' href="javascript:del();"><%=s_lang_del %></a>
	<a class='wwm_btnDownload btn_blue' href="javascript:save();"><%=s_lang_save %></a>
	</span>
	</td>
	<td nowrap width="10%" align="left" style="padding-left:8px;">
<select id='sfid' name='sfid' class='drpdwn' size='1' LANGUAGE=javascript onchange="selopt()">
<option value="">-----<%=a_lang_168 %>-----</option>
<%
	allnum = eads.FolderCount
	i = 0
	li = 0

	do while i < allnum
		eads.GetFolderInfo i, fid, layer, name, comment

		tempstr = ""
		li = 0
		do while li < layer
			tempstr = tempstr & "|&nbsp;"
			li = li + 1
		loop

		if fid <> rq_pid then
			Response.Write "<option value='" & fid & "'>" & tempstr & server.htmlencode(name) & "</option>" & Chr(13)
		else
			Response.Write "<option value='" & fid & "' selected>" & tempstr & server.htmlencode(name) & "</option>" & Chr(13)
		end if

		fid = NULL
		layer = NULL
		name = NULL
		comment = NULL

		i = i + 1
	loop
%>
</select>
	</td>
	<td nowrap width="40%" style="padding-left:8px;">
	<select name="convfrom" id="convfrom" class="drpdwn" size="1" LANGUAGE=javascript onchange="select_convfrom()"<%if rq_pid < 0 then Response.Write " disabled" %>>
	<option value="0" selected>--<%=a_lang_169 %>--</option>
	<option value="1"><%=a_lang_170 %></option>
	<option value="2"><%=a_lang_171 %></option>
	<option value="3"><%=a_lang_172 %></option>
	<option value="4"><%=a_lang_173 %></option>
	<option value="5"><%=s_lang_0070 %></option>
	</select><span name="ldap_com" id="ldap_com"></span>
	<input type="button" name="bt_conv" id="bt_conv" value="<%=a_lang_174 %>" onclick="javascript:convads()" class="Bsbttn" disabled>
	</td>
	<td nowrap align="right" style="padding-right:8px; color:#444444;"><%=a_lang_175 %></td>
	</tr>
</table>

<div id="imp_id">
<table width="96%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr><td class="block_top_td" style="height:6px;"></td></tr>
	<tr><td><textarea id="imp_text" rows="9" cols="100" wrap="virtual" class="n_textarea"></textarea>
	</td></tr>
	<tr><td class="block_top_td" style="height:2px;"></td></tr>
	<tr><td>
	<a class='wwm_btnDownload btn_gray' href="javascript:do_imp();"><%=a_lang_176 %></a>
	<span style="padding-left:20px; color:#666666">[<%=a_lang_177 %>]</span>
	</td></tr>
</table>
</div>

<table width="96%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr><td class="block_top_td" style="height:8px;"></td></tr>
	<tr class="title_tr">
	<td width="4%" class="title_td"><input type="checkbox" id="allcheck" name="allcheck" onclick="return allcheck_onclick()"></td>
	<td width="6%" class="title_td"><%=a_lang_071 %></td>
	<td width="30%" class="title_td"><%=a_lang_178 %></td>
	<td width="30%" class="title_td"><%=a_lang_179 %></td>
	<td width="30%" class="title_td" style="border-right:1px solid #A5B6C8;"><%=a_lang_180 %></td>
	</tr>
<%
allnum = eads.GetEmailsCountByFolderID(rq_pid)

if mode = "add" then
	Response.Write "<tr>"
	Response.Write "<td class='con_td'>&nbsp;</td>"
	Response.Write "<td class='con_td'>&nbsp;</td>"
	Response.Write "<td class='con_td'><input type='text' id='name" & allnum & "' name='name" & allnum & "' size='26' maxlength='128' class='n_textbox' value=""" & name & """></td>"
	Response.Write "<td class='con_td'><input type='text' name='email" & allnum & "' size='26' maxlength='128' class='n_textbox' value=""" & email & """></td>"
	Response.Write "<td class='con_td'><input type='text' name='comment" & allnum & "' size='26' maxlength='256' class='n_textbox' value=""" & comment & """></td>"
	Response.Write "</tr>" & Chr(13)
end if

i = 0

do while i < allnum
	eads.GetAdsInfoByFolderID rq_pid, i, name, email, comment

	Response.Write "<tr>"
	Response.Write "<td class='con_td'><input type='checkbox' name='check" & i & "'></td>"
	Response.Write "<td class='con_td'>" & i+1 & "</td>"
	Response.Write "<td class='con_td'><input type='text' name='name" & i & "' size='26' maxlength='128' class='n_textbox' value=""" & name & """></td>"
	Response.Write "<td class='con_td'><input type='text' name='email" & i & "' size='26' maxlength='128' class='n_textbox' value=""" & email & """></td>"
	Response.Write "<td class='con_td'><input type='text' name='comment" & i & "' size='26' maxlength='250' class='n_textbox' value=""" & comment & """></td>"
	Response.Write "</tr>" & Chr(13)

	name = NULL
	email = NULL
	comment = NULL

	i = i + 1
loop
%>
</table>

<%
if allnum < 1 then
%>
<table width="96%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:18px;"></td></tr>
<tr><td align="center">
<div class="wwm_line_msg"><%=s_lang_nodate %></div>
</td></tr></table>
<%
end if
%>

</FORM>
<div id="show_proc" class="wwm_msg" style="position:absolute; top:17%; left:50%; margin:0 0 0 -150px; z-index:100; display:none;"><%=a_lang_181 %>...</div>

<div id="pop_msg_div" class="mydiv" style="display:none;">
	<div class="pop_base"><div class="pop_bd bd"><div class="title">
		<div class="title_left"><%=a_lang_341 %></div>
		<div class="title_right" title="<%=s_lang_close %>" onclick="javascript:close_pop_msg(0);"><span>&nbsp;</span></div>
	</div>
	<div id="pop_msg_id" class="pop_content" style="height:126px; text-align:center; overflow-x:hidden; overflow-y:auto;">
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="white">
		<tr><td align="left">
		<input type="checkbox" id="cv_nosame" value="checkbox" checked><%=a_lang_342 %>
		</td></tr>
	</table>
	<table id="tb_systemuser" width="90%" border="0" align="center" cellspacing="0" bgcolor="white">
		<tr><td align="left">
		<input type="checkbox" id="cv_comm" value="checkbox" checked><%=a_lang_343 %>
		</td></tr>
		<tr><td align="left" style="padding-bottom:10px;">
		<input type="checkbox" id="cv_nodisabled" value="checkbox" checked><%=a_lang_344 %>
		</td></tr>
		<tr><td align="left" style="padding-top:10px; padding-left:6px; border-top: 1px solid #999999;">
		<%=a_lang_345 %><br>
<select id="cv_domain" class="drpdwn" size="1">
<option value="" selected><%=a_lang_346 %></option>
<%
dim wmethod
set wmethod = server.createobject("easymail.WMethod")

dim ei
set ei = server.createobject("easymail.domain")
ei.Load

i = 0
allnum = ei.GetCount()

do while i < allnum
	cdomainstr = ei.GetDomain(i)

	Response.Write "<option value='" & server.htmlencode(cdomainstr) & "'>" & server.htmlencode(wmethod.Puny_To_Domain(cdomainstr)) & "</option>" & Chr(13)

	cdomainstr = NULL
	i = i + 1
loop

set ei = nothing
set wmethod = nothing
%>
</select>
		</td></tr>
	</table>
	</div>
	<div class="title_bottom">
	<div class="title_ok_cancel_div">
	<a class="wwm_btnDownload btn_gray" href="#" onclick="javascript:close_pop_msg(1);"><%=s_lang_ok %></a>&nbsp;
	<a class="wwm_btnDownload btn_gray" href="#" onclick="javascript:close_pop_msg(0);"><%=s_lang_cancel %></a>
	</div></div></div></div>
</div>

<div id="bg" class="bg" style="display:none;"></div>
<iframe id='popIframe' class='popIframe' frameborder='0'></iframe>
</BODY>
</HTML>

<%
set eads = nothing

Function UnEscape(str)
    dim i,s,c
    s=""
    For i=1 to Len(str)
        c=Mid(str,i,1)
        If Mid(str,i,2)="%u" and i<=Len(str)-5 Then
            If IsNumeric("&H" & Mid(str,i+2,4)) Then
                s = s & CHRW(CInt("&H" & Mid(str,i+2,4)))
                i = i+5
            Else
                s = s & c
            End If
        ElseIf c="%" and i<=Len(str)-2 Then
            If IsNumeric("&H" & Mid(str,i+1,2)) Then
                s = s & CHRW(CInt("&H" & Mid(str,i+1,2)))
                i = i+2
            Else
                s = s & c
            End If
        Else
            s = s & c
        End If
    Next
    UnEscape = s
End Function
%>
