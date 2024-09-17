<!--#include file="passinc.asp" --> 

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

dim ei
set ei = server.createobject("easymail.SystemMailFilter")
'-----------------------------------------
ei.Load

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	ei.RemoveAll
	ei.Add trim(request("allmsgs"))
	ei.Save
	set ei = nothing

	response.redirect "ok.asp?" & getGRSN() & "&gourl=systemfilter.asp"
end if
%>


<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<script LANGUAGE=javascript>
<!--
function sub()
{
	var tempstr = "";
	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		tempstr = tempstr + document.f1.listall[i].value + "\f";
	}

	document.f1.allmsgs.value = tempstr;
	document.f1.submit();
}

function mdown()
{
	var i = 0;
	var findit = -1;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].selected == true)
		{
			findit = i;
			break;
		}
	}

	if (findit > -1 && findit < document.f1.listall.length - 1)
	{
		var tempstr = document.f1.listall[findit + 1].text;
		document.f1.listall[findit + 1].text = document.f1.listall[findit].text;
		document.f1.listall[findit].text = tempstr;

		tempstr = document.f1.listall[findit + 1].value;
		document.f1.listall[findit + 1].value = document.f1.listall[findit].value;
		document.f1.listall[findit].value = tempstr;

		document.f1.listall[findit + 1].selected = true;
		document.f1.listall[findit].selected = false;
	}
}

function mup()
{
	var i = 0;
	var findit = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].selected == true)
		{
			findit = i;
			break;
		}
	}

	if (findit > 0)
	{
		var tempstr = document.f1.listall[findit - 1].text;
		document.f1.listall[findit - 1].text = document.f1.listall[findit].text;
		document.f1.listall[findit].text = tempstr;

		tempstr = document.f1.listall[findit - 1].value;
		document.f1.listall[findit - 1].value = document.f1.listall[findit].value;
		document.f1.listall[findit].value = tempstr;

		document.f1.listall[findit - 1].selected = true;
		document.f1.listall[findit].selected = false;
	}
}

function delout()
{
	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].selected == true)
		{
			document.f1.listall.remove(i);

			if (i < document.f1.listall.length)
				document.f1.listall[i].selected = true;
			else
			{
				if (i - 1 >= 0)
					document.f1.listall[i - 1].selected = true;
			}

			break;
		}
	}
}

function add()
{
	if (haveit() == false)
	{
		if (document.f1.s_content.value == 3)
		{
			if (isNaN(parseInt(document.f1.s_msg.value)) == true)
			{
				alert("输入错误!");
				return ;
			}
			else
				document.f1.s_msg.value = parseInt(document.f1.s_msg.value);
		}

		var oOption = document.createElement("OPTION");
		oOption.text = getFilterStr(document.f1.s_content.value, document.f1.s_mode.value, document.f1.s_msg.value);
		oOption.value = document.f1.s_content.value + "\t" + document.f1.s_mode.value + "\t" + document.f1.s_msg.value;
<%
if isMSIE = true then
%>
		document.f1.listall.add(oOption);
<%
else
%>
		document.f1.listall.appendChild(oOption);
<%
end if
%>
	}
}

function haveit()
{
	var tempstr = document.f1.s_content.value + "\t" + document.f1.s_mode.value + "\t" + document.f1.s_msg.value;

	var i = 0;
	for (i; i < document.f1.listall.length; i++)
	{
		if (document.f1.listall[i].value == tempstr)
			return true;
	}

	return false;
}

function add_s_mode(a_value, a_text)
{
	var oOption = document.createElement("OPTION");
	oOption.text = a_text;
	oOption.value = a_value;
<%
if isMSIE = true then
%>
	document.f1.s_mode.add(oOption);
<%
else
%>
	document.f1.s_mode.appendChild(oOption);
<%
end if
%>
}

function select_content_onchange()
{
	var i = 0;
	for (i; i < document.f1.s_mode.length; i++)
	{
		document.f1.s_mode.remove(i);
		i--;
	}

	if (document.f1.s_content.value != "3")
	{
		add_s_mode("1", "等于");
		add_s_mode("2", "不等于");
		add_s_mode("3", "包含");
		add_s_mode("4", "不包含");
		add_s_mode("5", "通配符等于");
	}
	else
	{
		add_s_mode("1", "等于");
		add_s_mode("6", "大于");
		add_s_mode("7", "小于");
	}
}

function getFilterStr(f_content, f_mode, f_msg)
{
	var retstr;
	if (f_content == 1)
		retstr = "From";
	else if (f_content == 2)
		retstr = "Sender";
	else if (f_content == 3)
		retstr = "Size";
	else if (f_content == 4)
		retstr = "Subject";
	else if (f_content == 5)
		retstr = "Header";
	else if (f_content == 6)
		retstr = "Body";
	else if (f_content == 7)
		retstr = "To";
	else if (f_content == 8)
		retstr = "Cc";
	else if (f_content == 9)
		retstr = "Reply-To";
	else if (f_content == 10)
		retstr = "Boundary";
	else
		return "";

	if (f_mode == 1)
		retstr = retstr + " 等于 ";
	else if (f_mode == 2)
		retstr = retstr + " 不等于 ";
	else if (f_mode == 3)
		retstr = retstr + " 包含 ";
	else if (f_mode == 4)
		retstr = retstr + " 不包含 ";
	else if (f_mode == 5)
		retstr = retstr + " 通配符等于 ";
	else if (f_mode == 6)
		retstr = retstr + " 大于 ";
	else if (f_mode == 7)
		retstr = retstr + " 小于 ";
	else
		return "";

	if (f_msg.length == 0)
		retstr = retstr + "[Empty]";
	else
		retstr = retstr + f_msg;

	return retstr;
}
//-->
</script>

<BODY>
<br>
<FORM ACTION="#" METHOD="POST" NAME="f1">
<input type="hidden" name="allmsgs">
<div align="center"><br>
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_2 %>" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
    <tr>
      <td width="15%" height="28">&nbsp;</td>
      <td width="32%"><a href="showsysinfo.asp?<%=getGRSN() %>#systemfilter">启动项设置</a></td>
      <td colspan="23"><a href="right.asp?<%=getGRSN() %>">返回</a></td>
      <td width="30%"><font class="s" color="<%=MY_COLOR_4 %>"><b>邮件内容过滤</b></font></td>
    </tr>
  </table>
</div>
  <div align="center">
  <table align="center" border="0" width="90%" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr> 
      <td height="94" rowspan="2" width="45%">
        <div align="center"><br>
          &nbsp;<select name="listall" size="12" class="drpdwn" style="width: 480;">
<%
i = 0
allnum = ei.Count

do while i < allnum
	ei.Get i, as_content, as_mode, as_msg
	Response.Write "<option value=""" & as_content & Chr(9) & as_mode & Chr(9) & server.htmlencode(as_msg) & """>" & server.htmlencode(getFilterStr(as_content, as_mode, as_msg)) & "</option>" & Chr(13)

	as_content = NULL
	as_mode = NULL
	as_msg = NULL

	i = i + 1
loop
%>
          </select>
        </div>
      </td>
    </tr>
    <tr> 
      <td height="47" width="10%"> 
        <div align="center"><br>
          <input type="button" value="上移" class="sbttn" LANGUAGE=javascript onclick="mup()">
        </div>
        <br><br><br>
        <div align="center"> 
          <input type="button" value="删除" class="sbttn" LANGUAGE=javascript onclick="delout()">
        </div>
        <br><br><br>
        <div align="center"> 
          <input type="button" value="下移" class="sbttn" LANGUAGE=javascript onclick="mdown()">
        </div>
      </td>
    </tr>
    <tr>
      <td height="20" colspan="3">
<tr><td height="30" align="left" nowrap>&nbsp;
<select name="s_content" class=drpdwn LANGUAGE=javascript onchange="select_content_onchange()">
<option value="1">From</option>
<option value="2">Sender</option>
<option value="3">Size</option>
<option value="4">Subject</option>
<% if IsEnterpriseVersion = true then %>
<option value="5">Header</option>
<option value="6">Body</option>
<option value="7">To</option>
<option value="8">Cc</option>
<option value="9">Reply-To</option>
<option value="10">Boundary</option>
<% end if %>
</select>
<select name="s_mode" class=drpdwn>
<option value="1">等于</option>
<option value="2">不等于</option>
<option value="3">包含</option>
<option value="4">不包含</option>
<option value="5">通配符等于</option>
</select>
<input type="input" name="s_msg" class="textbox" size="30" maxlength="100">&nbsp;<input type="button" value=" 添加 " class="sbttn" LANGUAGE=javascript onclick="add()">
      </td></tr>
    <tr>
    <tr>
      <td height="20" colspan="3" align="right"><br><hr size="1" color="<%=MY_COLOR_1 %>">
    <input type="button" value=" 保存 " LANGUAGE=javascript onclick="sub()" class="Bsbttn">&nbsp;&nbsp;
    <input type="button" value=" 退出 " LANGUAGE=javascript onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
      </td></tr>
    <tr>
  </table>
<br><br><br>
  <div align="center">
    <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border-top:1px <%=MY_COLOR_1 %> solid;'>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
      <tr> 
        <td width="6%" valign="top">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
        <td width="94%">对所接收到邮件的各项内容进行高级过滤设置.<br>
        </td>
      </tr>
      <tr>
        <td colspan="2" height="10"> 
        </td>
      </tr>
    </table>
  </div>
</FORM>
</div>
<br>
</BODY>
</HTML>


<%
set ei = nothing


function getFilterStr(f_content, f_mode, f_msg)
	if f_content = 1 then
		getFilterStr = "From"
	elseif f_content = 2 then
		getFilterStr = "Sender"
	elseif f_content = 3 then
		getFilterStr = "Size"
	elseif f_content = 4 then
		getFilterStr = "Subject"
	elseif f_content = 5 then
		getFilterStr = "Header"
	elseif f_content = 6 then
		getFilterStr = "Body"
	elseif f_content = 7 then
		getFilterStr = "To"
	elseif f_content = 8 then
		getFilterStr = "Cc"
	elseif f_content = 9 then
		getFilterStr = "Reply-To"
	elseif f_content = 10 then
		getFilterStr = "Boundary"
	else
		Exit Function
	end if

	if f_mode = 1 then
		getFilterStr = getFilterStr & " 等于 "
	elseif f_mode = 2 then
		getFilterStr = getFilterStr & " 不等于 "
	elseif f_mode = 3 then
		getFilterStr = getFilterStr & " 包含 "
	elseif f_mode = 4 then
		getFilterStr = getFilterStr & " 不包含 "
	elseif f_mode = 5 then
		getFilterStr = getFilterStr & " 通配符等于 "
	elseif f_mode = 6 then
		getFilterStr = getFilterStr & " 大于 "
	elseif f_mode = 7 then
		getFilterStr = getFilterStr & " 小于 "
	else
		getFilterStr = ""
		Exit Function
	end if

	if f_msg = "" then
		getFilterStr = getFilterStr & "[Empty]"
	else
		getFilterStr = getFilterStr & f_msg
	end if
end function
%>
