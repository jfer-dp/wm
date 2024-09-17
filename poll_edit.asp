<!--#include file="passinc.asp" -->
<!--#include file="language-2.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

fid = trim(request("fid"))
gourl = trim(request("gourl"))

dim poll
set poll = server.createobject("easymail.Poll")
poll.LoadOne fid

allnum = poll.PI_ChooseCount

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if trim(request("isgoend")) = "true" then
		poll.PI_End_Poll

		set poll = nothing
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
	end if

	if IsNumeric(trim(request("maxsel"))) = true then
		Limit_Choose_Number = CInt(trim(request("maxsel")))
	else
		Limit_Choose_Number = 1
	end if

	if Limit_Choose_Number > 1 then
		poll.PI_Is_Choose_One = false
	else
		poll.PI_Is_Choose_One = true
	end if

	poll.PI_Limit_Choose_Number = Limit_Choose_Number
	poll.PI_Title = trim(request("p_title"))
	poll.PI_EndTime = trim(request("p_date"))
	poll.PI_Remove_All_Domains
	poll.PI_Remove_All_Items


	msg = trim(request("p_domains"))
	if Len(msg) > 0 then
		dim item
		dim ss
		dim se
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(9))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				poll.PI_AddDomain item
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if


	i = 0
	do while i < 15
		s_name = trim(request("name" & i))
		s_name = replace(s_name, """", "'")
		s_name = replace(s_name, Chr(9), "")

		s_len = 0
		if IsNumeric(trim(request("len" & i))) = true then
			s_len = CLng(trim(request("len" & i)))
		end if

		if s_name <> "" then
			poll.PI_AddName s_name, s_len
		end if 

	    i = i + 1
	loop


	poll.PI_Poll_BBS = trim(request("mto"))
	poll.Save
	set poll = nothing

	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
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
.sbttn {font-family:<%=s_lang_font %>;font-size:9pt;background: #D6E7EF;border-bottom: 1px solid #104A7B;border-right: 1px solid #104A7B;border-left: 1px solid #AFC4D5;border-top:1px solid #AFC4D5;color:#000066;height:19px;text-decoration:none;cursor:pointer}
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.st_l {height:26px; text-align:right; white-space:nowrap; background-color:#f2f4f6; border-right:1px solid #A5B6C8; border-bottom:1px solid #A5B6C8;}
.st_r {height:26px; white-space:nowrap; background-color:white; border-bottom:1px solid #A5B6C8; padding-left:6px;}
-->
</STYLE>
</HEAD>

<script type="text/javascript">
<!--
var maxnumber = 15;
var curnumber = <%=allnum %>;

function addnew() {
	if (curnumber >= maxnumber)
		return ;

document.getElementById("myadd").innerHTML = document.getElementById("myadd").innerHTML + "<div id=\"aim" + curnumber + "\"><%=b_lang_011 %>:&nbsp;<input name=\"name" + curnumber + "\" type=\"text\" class=\"n_textbox\" maxlength=\"64\">\
&nbsp;&nbsp;<%=b_lang_012 %>:&nbsp;<input name=\"len" + curnumber + "\" type=\"text\" class=\"n_textbox\" maxlength=\"5\" size=\"5\" value=\"0\">\
&nbsp;&nbsp;<a href='javascript:upit(" + curnumber + ")'><img src='images\\arrow_up.gif' border='0' align='absmiddle' title='<%=b_lang_013 %>'></a>&nbsp;&nbsp;<a href='javascript:downit(" + curnumber + ")'><img src='images\\arrow_down.gif' border='0' align='absmiddle' title='<%=b_lang_014 %>'></a>&nbsp;&nbsp;<a href='javascript:delit(" + curnumber + ")'><img src='images\\del.gif' border='0' align='absmiddle' title='<%=s_lang_del %>'></a><br></div>";

	curnumber++;
}

function delit(mid) {
	var curid = mid;
	var nextid;

	var i = 0;
	for (; i < curnumber; i++)
	{
		nextid = curid + 1;

		if (curid == nextid)
			break ;

		if (nextid > maxnumber || nextid >= curnumber)
			break ;

		downit(curid, nextid);

		curid = nextid;
	}

	curnumber--;
	var curObj = document.getElementById("aim" + curnumber);
	var dellen = curObj.innerHTML.length;

	var tstr = "<div id=\"" + curnumber + "\"></div>"

	var alllen = document.getElementById("myadd").innerHTML.length;
	document.getElementById("myadd").innerHTML = document.getElementById("myadd").innerHTML.substr(0, alllen - dellen - tstr.length);
}

function moveit(curid, nextid) {
	if (curid == nextid)
		return ;

	if (nextid < 0 || nextid > maxnumber || nextid >= curnumber)
		return ;

	var bfObj = eval("document.f1.name" + nextid);
	var curObj = eval("document.f1.name" + curid);

	var tempstr = "";
	tempstr = bfObj.value;
	bfObj.value = curObj.value;
	curObj.value = tempstr;

	tempstr = "";
	bfObj = eval("document.f1.len" + nextid);
	curObj = eval("document.f1.len" + curid);
	tempstr = bfObj.value;
	bfObj.value = curObj.value;
	curObj.value = tempstr;
}

function upit(mid) {
	moveit(mid, mid - 1);
}

function downit(mid) {
	moveit(mid, mid + 1);
}

function isinlist(name)
{
	var i = 0;
	for (i; i < document.f1.recdomains.length; i++)
	{
		if (document.f1.recdomains[i].value == name)
		{
			return true;
		}
	}
	
	return false;
}

function addin()
{
	var i = 0;
	for (i; i < document.f1.alldomains.length; i++)
	{
		if (document.f1.alldomains[i].selected == true)
		{
			if (isinlist(document.f1.alldomains[i].value) == false)
			{
				var oOption = document.createElement("OPTION");
				oOption.text = document.f1.alldomains[i].value;
				oOption.value = document.f1.alldomains[i].value;
<%
if isMSIE = true then
%>
				document.f1.recdomains.add(oOption);
<%
else
%>
				document.f1.recdomains.appendChild(oOption);
<%
end if
%>
			}
		}
	}
}

function delout()
{
	var i = 0;
	for (i; i < document.f1.recdomains.length; i++)
	{
		if (document.f1.recdomains[i].selected == true)
		{
			document.f1.recdomains.remove(i);
			i--;
		}
	}
}

function radioall_onclick() {
	document.f1.recdomains.disabled = true;
	document.f1.bdel.disabled = true;
	document.f1.badd.disabled = true;
	document.f1.alldomains.disabled = true;
}

function radiosel_onclick() {
	document.f1.recdomains.disabled = false;
	document.f1.bdel.disabled = false;
	document.f1.badd.disabled = false;
	document.f1.alldomains.disabled = false;
}

function gosub() {
	if (document.f1.p_title.value.length == 0)
	{
		alert("<%=s_lang_inputerr %>")
		document.f1.p_title.focus();
		return ;
	}

	var theObj = eval("document.f1.name0");
	if (theObj == null)
	{
		alert("<%=s_lang_inputerr %>")
		document.f1.btadd.focus();
		return ;
	}
	else if (theObj.value.length == 0)
	{
		alert("<%=s_lang_inputerr %>")
		theObj.focus();
		return ;
	}

	var curObj = eval("document.f1.name1");
	if (curObj == null)
	{
		alert("<%=s_lang_inputerr %>")
		document.f1.btadd.focus();
		return ;
	}
	else if (curObj.value.length == 0)
	{
		alert("<%=s_lang_inputerr %>")
		curObj.focus();
		return ;
	}


	document.f1.p_domains.value = "";
	if (document.f1.rd2.checked == true)
	{
		for (var i = 0; i < document.f1.recdomains.length; i++)
		{
			if (document.f1.recdomains[i].value.length > 0)
				document.f1.p_domains.value = document.f1.p_domains.value + document.f1.recdomains[i].value + "\t";
		}
	}


	if (parseInt(document.f1.maxsel.value) > curnumber)
	{
		alert("<%=s_lang_inputerr %>")
		document.f1.maxsel.focus();
		return ;
	}


	if (document.f1.t_year.value != "" && document.f1.t_month.value != "" && document.f1.t_day.value != "" && document.f1.t_hour.value != "")
	{
		var nowdate = new Date(<%=Year(now()) & "," & Month(now()) - 1 & "," & Day(now()) & "," & Hour(now()) & "," & Minute(now()) %>);
		var mydate = new Date(document.f1.t_year.value, document.f1.t_month.value - 1, document.f1.t_day.value, document.f1.t_hour.value, 1);
		if (mydate < nowdate)
		{
			alert("<%=s_lang_inputerr %>")
			document.f1.t_year.focus();
			return ;
		}


		if (document.f1.t_month.value < 10)
			nmonth = "0" + cutz(document.f1.t_month.value);
		else
			nmonth = document.f1.t_month.value

		if (document.f1.t_day.value < 10)
			nday = "0" + cutz(document.f1.t_day.value);
		else
			nday = document.f1.t_day.value

		if (document.f1.t_hour.value < 10)
			nhour = "0" + cutz(document.f1.t_hour.value);
		else
			nhour = document.f1.t_hour.value

		if (nhour == "0")
			nhour = "00"

		document.f1.p_date.value = document.f1.t_year.value + nmonth + nday + nhour;
	}

	document.f1.submit();
}

function cutz(inval)
{
	var rval = "";

	for (var i = 0; i < inval.length; i++)
	{
		if (inval.charAt(i) != '0')
			break;
	}

	rval = inval.substring(i);
	return rval;
}

function window_onload() {
	if (document.f1.rd1.checked == true)
		radioall_onclick();
	else if (document.f1.rd2.checked == true)
		radiosel_onclick();
}

function goend() {
	if (confirm("<%=b_lang_033 %>") == false)
		return ;

	document.f1.isgoend.value = "true";
	document.f1.submit();
}
//-->
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
<FORM ACTION="poll_edit.asp" METHOD="POST" NAME="f1">
<input name="p_domains" type="hidden">
<input name="p_date" type="hidden">
<input name="isgoend" type="hidden">
<input name="fid" type="hidden" value="<%=fid %>">
<input name="gourl" type="hidden" value="<%=gourl %>">
<table width="92%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
<%=b_lang_034 %>
</td></tr>
<tr><td class="block_top_td" style="height:14px; _height:16px;"></td></tr>
<tr><td align="center">

<table width="97%" border="0" align="center" cellspacing="0" bgcolor="white">
	<tr><td width="30%" class="st_l" style="border-top:1px #A5B6C8 solid;"><%=b_lang_016 %><%=s_lang_mh %>
	</td>
	<td align="left" class="st_r" style="border-top:1px #A5B6C8 solid;">
	<input name="p_title" type="text" class="n_textbox" size="40" maxlength="512" value="<%=poll.PI_Title %>">
	</td>
	</tr>

	<tr>
	<td rowspan="2" class="st_l"><%=b_lang_017 %><%=s_lang_mh %>
	</td>
	<td align="center" class="st_r">
	<input name="btadd" type="button" value="<%=b_lang_018 %>" onclick="javascript:addnew()" class="sbttn">
	</td>
	</tr>

	<tr>
	<td id="myadd" align="left" class="st_r">
<%
i = 0
do while i < allnum
	poll.PI_GetNameAndNumber i, v_name, v_num

	Response.Write "<div id='aim" & i & "'>" & b_lang_011 & ":&nbsp;<input name='name" & i & "' type='text' class='n_textbox' maxlength='64' value='" & v_name & "'>"
	Response.Write "&nbsp;&nbsp;" & b_lang_012 & ":&nbsp;<input name='len" & i & "' type='text' class='n_textbox' maxlength='5' size='5' value='" & v_num & "'>"
	Response.Write "&nbsp;&nbsp;<a href='javascript:upit(" & i & ")'><img src='images\arrow_up.gif' border='0' align='absmiddle' title='" & b_lang_013 & "'></a>&nbsp;&nbsp;<a href='javascript:downit(" & i & ")'><img src='images\\arrow_down.gif' border='0' align='absmiddle' title='" & b_lang_014 & "'></a>&nbsp;&nbsp;<a href='javascript:delit(" & i & ")'><img src='images\del.gif' border='0' align='absmiddle' title='" & s_lang_del & "'></a><br></div>"

	i = i + 1

	v_name = NULL
	v_num = NULL
loop
%>
	</td>
	</tr>

	<tr>
	<td class="st_l"><%=b_lang_019 %><%=s_lang_mh %>
	</td>
	<td align="left" class="st_r">
	<input type="text" class="n_textbox" size="5" maxlength="2" name="maxsel" value="<%=poll.PI_Limit_Choose_Number %>">
	</td>
	</tr>

	<tr>
	<td class="st_l"><%=b_lang_020 %><%=s_lang_mh %>
	</td>
	<td align="left" class="st_r">
<%
if poll.PI_Domains <> "" then
%>
	<input type=radio id="rd1" value="True" name="domainmode" LANGUAGE=javascript onclick="return radioall_onclick()"><%=b_lang_021 %><br>
	<input type=radio id="rd2" checked value="False" name="domainmode" LANGUAGE=javascript onclick="return radiosel_onclick()"><%=b_lang_022 %>
<%
else
%>
	<input type=radio id="rd1" checked value="True" name="domainmode" LANGUAGE=javascript onclick="return radioall_onclick()"><%=b_lang_021 %><br>
	<input type=radio id="rd2" value="False" name="domainmode" LANGUAGE=javascript onclick="return radiosel_onclick()"><%=b_lang_022 %>
<%
end if
%>
	</td>
	</tr>

	<tr>
	<td class="st_l"><%=b_lang_023 %><%=s_lang_mh %>
	</td>
	<td align="left" class="st_r">
		<table align="center" border="0" width="90%" cellspacing="0">
		<tr>
		<td height="20" rowspan="2" width="45%" align="center"><%=b_lang_024 %></td>
		<td height="20" rowspan="2">&nbsp;</td>
		<td height="20" rowspan="2" width="45%" align="center"><%=b_lang_021 %></td>
		</tr>
	    <tr></tr>
	    <tr> 
      	<td rowspan="2" width="45%" align="center">
		<select name="recdomains" disabled size="9" class="drpdwn" style="width:200px;" multiple LANGUAGE=javascript ondblclick="return delout()">
<%
if poll.PI_Domains <> "" then
	msg = poll.PI_Domains & Chr(9)
	if Len(msg) > 0 then
		ss = 2
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(9))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				Response.Write "<option value='" & server.htmlencode(item) & "'>" & server.htmlencode(item) & "</option>" & Chr(13)
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if
end if
%>
		</select>
		</td>
		<td width="10%" align="center">
		<input name="bdel" disabled type="button" value=" ==&gt; " class="sbttn" LANGUAGE=javascript onclick="delout()">
		</td>
		<td rowspan="2" width="45%" align="center">
		<select name="alldomains" disabled size="9" class="drpdwn" style="width:200px;" multiple LANGUAGE=javascript ondblclick="return addin()">
<%
dim dm
set dm = server.createobject("easymail.Domain")
dm.Load

i = 0
allnum = dm.GetCount()

do while i < allnum
	domain = dm.GetDomain(i)
	Response.Write "<option value='" & server.htmlencode(domain) & "'>" & server.htmlencode(domain) & "</option>" & Chr(13)
	domain = NULL

	i = i + 1
loop

set dm = nothing
%>
		</select>
		</td>
		</tr>

		<tr> 
		<td width="10%" align="center">
		<input name="badd" disabled type="button" value=" &lt;== " class="sbttn" LANGUAGE=javascript onclick="addin()">
		</td>
		</tr>
	</table>
	</td>
	</tr>

	<tr>
	<td class="st_l"><%=b_lang_025 %><%=s_lang_mh %>
	</td>
	<td align="left" class="st_r">
		<select name="mto" class="drpdwn" size="1">
		<option value=""><%=b_lang_007 %></option>
<%
dim pf
set pf = server.createobject("easymail.PubFolderManager")

allnum = pf.PubFoldersCount
i = 0

do while i < allnum
	pf.GetFolderInfoByIndex i, pf_pfilename, pf_admin, pf_permission, pf_name, pf_createTime, pf_count, pf_maxid, pf_maxitem, pf_maxsize

	if LCase(poll.PI_Poll_BBS) <> LCase(pf_pfilename) then
		Response.Write "<option value=""" & pf_pfilename & """>" & server.htmlencode(pf_name) & "</option>" & Chr(13)
	else
		Response.Write "<option value=""" & pf_pfilename & """ selected>" & server.htmlencode(pf_name) & "</option>" & Chr(13)
	end if

	pf_pfilename = NULL
	pf_admin = NULL
	pf_permission = NULL
	pf_name = NULL
	pf_createTime = NULL
	pf_count = NULL
	pf_maxid = NULL
	pf_maxitem = NULL
	pf_maxsize = NULL

    i = i + 1
loop

set pf = nothing
%>
		</select>&nbsp;&nbsp;&nbsp;<input type="button" value="<%=b_lang_026 %>" class="sbttn" LANGUAGE=javascript onclick="javascript:location.href='createpf.asp?<%=getGRSN() %>'">
	</td>
	</tr>

	<tr>
	<td class="st_l"><%=b_lang_027 %><%=s_lang_mh %>
	</td>
	<td align="left" class="st_r">
<select name="t_year" class="drpdwn">
<option value="">------</option>
<%
	PI_EndTime = poll.PI_EndTime

	if PI_EndTime <> "0" then
		PI_EndTime_Year = CInt(Mid(PI_EndTime, 1, 4))
	end if

	now_temp = Year(Now())

	i = now_temp
	do while i < now_temp + 5
		if PI_EndTime_Year <> i then
			Response.Write "<option value='" & i & "'>" & i & b_lang_028 & "</option>"
		else
			Response.Write "<option value='" & i & "' selected>" & i & b_lang_028 & "</option>"
		end if

		i = i + 1
	loop
%>
</select>&nbsp;
<select name="t_month" class="drpdwn">
<option value="">----</option>
<%
	if PI_EndTime <> "0" then
		PI_EndTime_Month = CInt(Mid(PI_EndTime, 5, 2))
	end if

	i = 1
	do while i < 13
		if PI_EndTime_Month <> i then
			Response.Write "<option value='" & i & "'>" & i & b_lang_029 & "</option>"
		else
			Response.Write "<option value='" & i & "' selected>" & i & b_lang_029 & "</option>"
		end if

		i = i + 1
	loop
%>
</select>&nbsp;
<select name="t_day" class="drpdwn">
<option value="">----</option>
<%
	if PI_EndTime <> "0" then
		PI_EndTime_Day = CInt(Mid(PI_EndTime, 7, 2))
	end if

	i = 1
	do while i < 32
		if PI_EndTime_Day <> i then
			Response.Write "<option value='" & i & "'>" & i & b_lang_030 & "</option>"
		else
			Response.Write "<option value='" & i & "' selected>" & i & b_lang_030 & "</option>"
		end if

		i = i + 1
	loop
%>
</select>&nbsp;
<select name="t_hour" class="drpdwn">
<option value="">----</option>
<%
	PI_EndTime_Hour = -1
	if PI_EndTime <> "0" then
		PI_EndTime_Hour = CInt(Mid(PI_EndTime, 9, 2))
	end if

	i = 0
	do while i < 24
		if PI_EndTime_Hour <> i then
			Response.Write "<option value='" & i & "'>" & i & b_lang_031 & "</option>"
		else
			Response.Write "<option value='" & i & "' selected>" & i & b_lang_031 & "</option>"
		end if

		i = i + 1
	loop
%>
</select>
	</td></tr>
	<tr>
	<td colspan="2" align="left" height="26"><br>
<a class='wwm_btnDownload btn_blue' href="javascript:history.back();"><< <%=s_lang_return %></a>
<%
if poll.PI_IsEnd = false then
%>
<a class='wwm_btnDownload btn_blue' href="javascript:goend();"><%=b_lang_035 %></a>
<%
end if
%>
<a class='wwm_btnDownload btn_blue' href="javascript:gosub();"><%=s_lang_save %></a>
	</td>
	</tr>
</table>

</td></tr>
</table>

</form>
</BODY>
</HTML>

<%
	set poll = nothing
%>
