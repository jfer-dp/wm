<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" --> 

<%
if isadmin() = false then
	Response.Redirect "noadmin.asp"
end if
%>

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

dim rou
set rou = server.createobject("easymail.ReadOnlyUsers")
rou.Load
allnum = rou.Count

dim is_add
is_add = false

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	if trim(request("EnableSH")) <> "" then
		rou.IsEnabled = true
	else
		rou.IsEnabled = false
	end if

	if trim(request("mode")) = "add" then
		is_add = true
		rou.Add trim(request("rouadd")), trim(request("pwadd"))
		rou.Save
	elseif trim(request("mode")) = "save" then
		rouname = trim(request("rouadd"))
		if rouname <> "" then
			rou.Add rouname, trim(request("pwadd"))
		end if
		rou.Save
		set rou = nothing
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=setrou.asp"
	elseif trim(request("mode")) = "editpw" then
		isok = false
		isok = rou.ModifyPassword(trim(request("edit_rou")), trim(request("edit_pw")))
		if isok = true then
			rou.Save
			set rou = nothing
			Response.Redirect "ok.asp?" & getGRSN() & "&gourl=setrou.asp"
		else
			set rou = nothing
			Response.Redirect "err.asp?" & getGRSN() & "&gourl=setrou.asp"
		end if
	elseif trim(request("mode")) = "del" then
		i = 0
		do while i < allnum
			rouname = trim(request("rou" & i))
			if trim(request("check" & i)) <> "" and rouname <> "" then
				rou.DelByName rouname
			end if 

		    i = i + 1
		loop

		rou.Save
		set rou = nothing
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=setrou.asp"
	end if
end if

allnum = rou.Count
%>


<HTML>
<HEAD>
<%=s_lang_meta %>
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
<script type="text/javascript" src="images/jquery.min.js"></script>

<STYLE type=text/css>
<!--
.block_top_td {white-space:nowrap; background:white; font-size: 0pt; height:1px;}
.routdend {padding-top:6px; padding-bottom:2px; border-bottom:1px <%=MY_COLOR_1 %> solid;}
.jtextbox {border:1px solid #555;}
-->
</STYLE>
</HEAD>

<script LANGUAGE=javascript>
<!--
function window_onload() {
<%
if is_add = true then
%>
document.getElementById("rouadd").focus();
<%
end if
%>
}

var before_changed = -1;

function editpw(index) {
	if (before_changed > -1)
	{
		var oldObj = document.getElementById("pwtd_" + before_changed);
		if (oldObj != null)
			oldObj.innerHTML = "<a href='javascript:editpw(" + before_changed + ")'><img src='images/pedit.gif' border=0 title='<%=s_lang_0194 %>'></a>";
	}

	var theObj = document.getElementById("pwtd_" + index);
	if (theObj != null)
	{
		theObj.innerHTML = "<input type='password' id='pw' name='pw' class='jtextbox'>&nbsp;<input type='button' value='<%=s_lang_0194 %>' class='sbttn' LANGUAGE=javascript onclick='save_pw()'>";
		before_changed = index;
	}

	document.getElementById("pw").focus();
}

function add() {
	document.f1.mode.value = "add";
	document.f1.submit();
}

function save() {
	document.f1.mode.value = "save";
	document.f1.submit();
}

function del() {
	if (ischeck() == true)
	{
		document.f1.mode.value = "del";
		document.f1.submit();
	}
}

function ischeck() {
	var i = 0;
	var theObj;

	for(; i<<%=allnum %>; i++)
	{
		theObj = document.getElementById("check" + i);

		if (theObj != null)
			if (theObj.checked == true)
				return true;
	}

	return false;
}

function save_pw() {
	if (before_changed > -1)
	{
		document.f1.edit_rou.value = document.getElementById("rou" + before_changed).value;
		document.f1.edit_pw.value = document.getElementById("pw").value;
		document.f1.mode.value = "editpw";
		document.f1.submit();
	}
}
//-->
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
<br>
<FORM ACTION="#" METHOD="POST" NAME="f1">
<input type="hidden" name="mode">
<input type="hidden" name="edit_rou">
<input type="hidden" name="edit_pw">
<div align="center"><br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>" style="border:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="15%" height="28">&nbsp;</td>
      <td width="60%"><a href="right.asp?<%=getGRSN() %>"><%=s_lang_return %></a></td>
      <td><font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0615 %></b></font></td>
	</tr>
	<tr bgcolor="<%=MY_COLOR_3 %>">
	<td height="26" colspan="4" style="border-top:1px <%=MY_COLOR_1 %> solid;">
	<input type="checkbox" name="EnableSH" id="EnableSH" <% if rou.isEnabled = true then response.write "checked"%>>
	<%=s_lang_0617 %>
	</td>
	</tr>
  </table>
</div>
  <div align="center">
  <table align="center" border="0" width="90%" cellspacing="0">
  <tr><td height="30" style="padding-top:12px; padding-left:8px; padding-bottom:12px;"></td></tr>
	<tr><td class="block_top_td" style="height:4px;"></td></tr>
	<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
	<%=s_lang_0618 %></td></tr>
  <tr>
	<td height="30" style="padding-top:12px; padding-left:8px; padding-bottom:12px;">
<table width="90%" border="0" align="center" cellspacing="0">
<%
i = 0

do while i < allnum
	if i < allnum then
		Response.Write "<tr><td width='9%' height='24' align='center' class='routdend'><input type='checkbox' name='check" & i & "' id='check" & i & "'></td><td width='41%' class='routdend'><input name='rou" & i & "' id='rou" & i & "' type='text' value='" & rou.Get(i) & "' size='40' readonly maxlength='64' class='jtextbox'></td>" & Chr(13)
		Response.Write "<td id='pwtd_" & i & "' width='50%' class='routdend'><a href='javascript:editpw(" & i & ")'><img src='images/pedit.gif' border=0 title='" & s_lang_0194 & "'></a></td></tr>" & Chr(13)
	end if

	i = i + 1
loop

if is_add = true then
	Response.Write "<tr><td colspan='3' height='24' class='routdend' style='padding-top:10px; padding-bottom:6px;'>&nbsp;" & s_lang_0619 & s_lang_mh & "<input name='rouadd' id='rouadd' type='text' size='40' maxlength='64' class='jtextbox'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	Response.Write s_lang_0614 & s_lang_mh & "<input type='password' id='pwadd' name='pwadd' class='jtextbox'></td></tr>" & Chr(13)
end if
%>
</table>
	</td>
  </tr>
  </table>

  <table align="center" border="0" width="90%" cellspacing="0">
	<tr><td class="block_top_td" style="height:4px;"></td></tr>
	<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:14px; font-weight:bold; color:#093665; padding-left:6px;">
	&nbsp;</td></tr>
	<tr>
	<td height="50" align="right">
	<input type="button" value=" <%=s_lang_add %> " LANGUAGE=javascript onclick="add()" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" <%=s_lang_del %> " LANGUAGE=javascript onclick="del()" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" <%=s_lang_save %> " LANGUAGE=javascript onclick="save()" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" <%=s_lang_return %> " LANGUAGE=javascript onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td></tr>
  </table>
</FORM>
<br>
</BODY>
</HTML>

<%
set rou = nothing
%>
