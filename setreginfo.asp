<!--#include file="passinc.asp" --> 

<%
if isadmin() = false then
	response.redirect "noadmin.asp"
end if
%>

<%
dim ei
set ei = server.createobject("easymail.MoreRegInfo")
'-----------------------------------------
ei.LoadSetting

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ei.RemoveAll_Setting

	errline = -1
	i = 0
	do while i < 99
		s_name = trim(request("name" & i))
		s_name = replace(s_name, """", "'")
		s_name = replace(s_name, Chr(9), "")

		s_sel = trim(request("sel" & i))
		s_len = trim(request("len" & i))

		if s_name <> "" and s_sel <> "" and s_len <> "" and IsNumeric(s_sel) = true and IsNumeric(s_len) = true then
			if errline = -1 then
				errline = ei.Add_Setting(s_name, CLng(s_sel), CLng(s_len))
			end if
		end if 

	    i = i + 1
	loop

	if errline = -1 then
		errline = ei.SaveSetting()
	end if

	set ei = nothing

	if errline = -1 then
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=setreginfo.asp"
	else
		Response.Redirect "err.asp?errstr=��" & errline & "�г���&" & getGRSN() & "&gourl=setreginfo.asp"
	end if
end if

allnum = ei.Count_Setting
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<script LANGUAGE=javascript>
<!--
var maxnumber = 99;
var curnumber = <%=allnum %>;

function moveit(curid, nextid) {
	if (curid == nextid)
		return ;

	if (nextid < 0 || nextid > maxnumber || nextid >= curnumber)
		return ;

	var bfObj = eval("document.all(\"name\" + nextid)");
	var curObj = eval("document.all(\"name\" + curid)");

	var tempstr = "";
	tempstr = bfObj.value;
	bfObj.value = curObj.value;
	curObj.value = tempstr;

	tempstr = "";
	bfObj = eval("document.all(\"len\" + nextid)");
	curObj = eval("document.all(\"len\" + curid)");
	tempstr = bfObj.value;
	bfObj.value = curObj.value;
	curObj.value = tempstr;

	tempstr = "";
	bfObj = eval("document.all(\"sel\" + nextid)");
	curObj = eval("document.all(\"sel\" + curid)");
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
	var curObj = eval("document.all(\"aim\" + curnumber)");
	var dellen = curObj.innerHTML.length;

	var tstr = "<div id=\"" + curnumber + "\"></div>"

	var alllen = document.all("myadd").innerHTML.length;
	document.all("myadd").innerHTML = document.all("myadd").innerHTML.substr(0, alllen - dellen - tstr.length);
}

function add() {
	if (curnumber >= maxnumber)
		return ;

	document.all("myadd").innerHTML = document.all("myadd").innerHTML + "<div id=\"aim" + curnumber + "\">����:&nbsp;<input name=\"name" + curnumber + "\" type=\"text\" class=\"textbox\" maxlength=\"128\">&nbsp;&nbsp;����:&nbsp;<select name=\"sel" + curnumber + "\" class=\"drpdwn\" size=\"1\"><option value=\"0\">С��</option><option value=\"1\">����</option><option value=\"2\" selected>����</option></select>\
	&nbsp;&nbsp;����:&nbsp;<input name=\"len" + curnumber + "\" type=\"text\" class=\"textbox\" maxlength=\"3\" size=\"3\">\
	&nbsp;&nbsp;&nbsp;<a href='javascript:upit(" + curnumber + ")'><img src='images\\arrow_up.gif' border='0' align='absmiddle' alt='����'></a>&nbsp;&nbsp;<a href='javascript:downit(" + curnumber + ")'><img src='images\\arrow_down.gif' border='0' align='absmiddle' alt='����'></a>&nbsp;&nbsp;<a href='javascript:delit(" + curnumber + ")'><img src='images\\del.gif' border='0' align='absmiddle' alt='ɾ��'></a><br><hr size=\"1\" color=\"<%=MY_COLOR_1 %>\"></div>";

	curnumber++;
}

function gosub() {
	var i = 0;
	var tempstr = "";
	var tempint;

	for (; i < curnumber; i++)
	{
		nameObj = eval("document.all(\"name\" + i)");
		selObj = eval("document.all(\"sel\" + i)");
		lenObj = eval("document.all(\"len\" + i)");

		if (nameObj.value.length == 0)
		{
			alert("�������.")
			nameObj.focus();
			return ;
		}

		if (lenObj.value.length == 0 || lenObj.value.length > 3)
		{
			alert("�������.")
			lenObj.focus();
			return ;
		}

		tempint = parseInt(lenObj.value);
		if (tempint < 1 || tempint > 128)
		{
			alert("�������.")
			lenObj.focus();
			return ;
		}

		if ((tempint == 1 && selObj.value == "0") || (tempint == 128 && selObj.value == "2"))
		{
			alert("�������.")
			lenObj.focus();
			return ;
		}
	}

	document.f1.submit();

	return ;
}
//-->
</script>

<BODY>
<br><br><br>
<FORM ACTION="setreginfo.asp" METHOD="POST" NAME="f1">
<input type="hidden" name="allmsgs">
  <table width="85%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td align="center" height="28"><font class="s" color="<%=MY_COLOR_4 %>"><b>����ע����Ϣ</b></font></td>
    </tr>
  </table>
</div>
  <div align="center">
  <table align="center" border="0" width="85%" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
  <tr>
	<td height="30" align="right">
	<input type="button" value=" ��� " LANGUAGE=javascript onclick="add()" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" ���� " LANGUAGE=javascript onclick="gosub()" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" �˳� " LANGUAGE=javascript onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	<br><hr size="1" color="<%=MY_COLOR_1 %>">
	</td>
  </tr>
	<tr><td id="myadd"><%
i = 0

do while i < allnum
	ei.Get_Setting i, s_name, s_sel, s_len

	Response.Write "<div id=""aim" & i & """>����:&nbsp;<input name=""name" & i & """ type=""text"" value=""" & s_name & """ class=""textbox"" maxlength=""128"">&nbsp;&nbsp;"
	Response.Write "����:&nbsp;<select name=""sel" & i & """ class=""drpdwn"" size=""1"">" & Chr(13)

	if s_sel = 0 then
		Response.Write "<option value=""0"" selected>С��</option>" & Chr(13)
	else
		Response.Write "<option value=""0"">С��</option>" & Chr(13)
	end if

	if s_sel = 1 then
		Response.Write "<option value=""1"" selected>����</option>" & Chr(13)
	else
		Response.Write "<option value=""1"">����</option>" & Chr(13)
	end if

	if s_sel = 2 then
		Response.Write "<option value=""2"" selected>����</option>" & Chr(13)
	else
		Response.Write "<option value=""2"">����</option>" & Chr(13)
	end if

	Response.Write "</select>" & Chr(13)
	Response.Write "&nbsp;&nbsp;����:&nbsp;<input name=""len" & i & """ type=""text"" value=""" & s_len & """ class=""textbox"" maxlength=""3"" size=""3"">" & Chr(13)
	Response.Write "&nbsp;&nbsp;&nbsp;<a href='javascript:upit(" & i & ")'><img src='images\arrow_up.gif' border='0' align='absmiddle' alt='����'></a>&nbsp;&nbsp;<a href='javascript:downit(" & i & ")'><img src='images\arrow_down.gif' border='0' align='absmiddle' alt='����'></a>&nbsp;&nbsp;<a href='javascript:delit(" & i & ")'><img src='images\del.gif' border='0' align='absmiddle' alt='ɾ��'></a><br><hr size=""1"" color=""" & MY_COLOR_1 & """></div>" & Chr(13)


	s_name = NULL
	s_sel = NULL
	s_len = NULL

	i = i + 1
loop
%></td></tr>
  </table>
<br>
<div style="position: absolute; left: 60; top: 20;">
<table><tr bgcolor="<%=MY_COLOR_2 %>"><td nowrap style="border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid; border-bottom:1px <%=MY_COLOR_1 %> solid;">
&nbsp;<a href="webadmin.asp?<%=getGRSN() %>#InputMoreInfo" style="text-transform: none; text-decoration: none;"><font class="s" color="<%=MY_COLOR_4 %>"><b>����������</b></font>&nbsp;<img src="images\ugo.gif" border="0" align="absbottom"></a>&nbsp;
</td></tr></table>
</div>
<br><br><br>
  <div align="center">
    <table width="85%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr>
        <td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
		<td width="94%">Ϊ�˻�ø����ע���û���Ϣ(����: ��ʵ����, �Ա������), ����Ա����ͨ������ע����Ϣ���ķ�ʽҪ��ע��������û�������д.
		<br><br>һ�������ַ��ĳ�����2���ֽ�.<br>
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
</FORM>
</BODY>
</HTML>


<%
set ei = nothing
%>
