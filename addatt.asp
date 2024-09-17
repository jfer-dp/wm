<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" --> 

<%
	Response.Cookies("attfoldername") = "zatt"
	Response.Cookies("attfoldername").Expires = DateAdd("d", 2, Now())

	dim am
	set am = server.createobject("easymail.Attachments")
	am.Load Session("wem"), Session("tid")
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<script LANGUAGE=javascript>
<!--
function window_onload()
{
<%
zid = trim(request("zid"))

if Len(zid) < 6 then
%>
	var tObj = parent.f3.document.getElementById("zAttName");

	if (tObj == null)
		return ;

	parent.f3.flash_att_div();
<%
else
	Response.Write "parent.f3.add_zatt_select();"
end if
%>
}

function mysub()
{
	if (document.fsa.upfile.value != "")
	{
		var theObj = document.getElementById("esave");
		theObj.style.visibility="visible";

		document.fsa.submit();
	}
}

function test_load()
{
	return 1;
}

function uf_changed()
{
	parent.f3.up_it();
}

function delatt(index_num)
{
	if (index_num > 0)
		location.href = "delatt.asp?id=" + document.fsa.AttName[index_num].value;
	else
	{
		if (document.fsa.AttName.value != "")
			location.href = "delatt.asp?id=" + document.fsa.AttName.value;
	}
}
<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

set ads = server.createobject("easymail.Addresses")
ads.Load Session("wem")

i = 0
ads_allnum = ads.EmailCount

if Session("ac_ads_number") = "" or (IsNumeric(Session("ac_ads_number")) = true and Session("ac_ads_number") <> ads_allnum) then
%>
var i = 0;
var allnum = parent.parent.f1.document.leftval.ads.length;
for (i; i < allnum; i++)
{
	parent.parent.f1.document.leftval.ads.remove(0);
}

<%
do while i < ads_allnum
	ads.MoveTo i
	Response.Write "var oOption = document.createElement(""OPTION"");" & Chr(13)

	if ads.nickname = "" then
		Response.Write "oOption.text = """ & server.htmlencode(Mid(ads.email, 1, InStr(ads.email, "@"))) & " <" & server.htmlencode(ads.email) & ">"";" & Chr(13)
	else
		Response.Write "oOption.text = """ & server.htmlencode(ads.nickname) & " <" & server.htmlencode(ads.email) & ">"";" & Chr(13)
	end if

	Response.Write "oOption.value = """ & ads.email & """;" & Chr(13)

	if isMSIE = true then
		Response.Write "parent.parent.f1.document.leftval.ads.add(oOption);" & Chr(13)
	else
		Response.Write "parent.parent.f1.document.leftval.ads.appendChild(oOption);" & Chr(13)
	end if

    i = i + 1
loop

	Session("ac_ads_number") = ads_allnum
end if

set ads = nothing
%>
//-->
</script>

<BODY LANGUAGE=javascript onload="return window_onload()">
<FORM ENCTYPE="multipart/form-data" ACTION="saveatt.asp" METHOD=POST NAME="fsa">
<div id="esave" style="position:absolute; top:5; left:40; width:800px; z-index:10; visibility:hidden"><TABLE WIDTH=90% BORDER=0 CELLSPACING=0 CELLPADDING=0 align="center"><TR><td width=40%></td>
	<TD bgcolor=<%=MY_COLOR_4 %> width="15%"> 
	<TABLE WIDTH=100% height=22 BORDER=0 CELLSPACING=2 CELLPADDING=0>
	<TR> 
	<td bgcolor=#eeeeee align=center><%=s_lang_0285 %>...</td>
	</tr>
	</table>
	</td>
	<td width=40%> </td>
	</tr></table></div>
  <table width="95%" border="0" bgColor="<%=MY_COLOR_5 %>" style="BORDER-LEFT: <%=MY_COLOR_5 %> 4px solid; BORDER-RIGHT: <%=MY_COLOR_5 %> 4px solid" align="center">
      <td>
		<input type="button" value="<%=s_lang_del %>" onclick="javascript:delatt()" class=sbttn>
		<select name="AttName" id="AttName" class=drpdwn size="1">
<%
allnum = am.Count
i = 0
asize = am.TotalSize

Response.Write "<option value=''>---" & s_lang_0286 & allnum & s_lang_0287
if asize > 0 then
	Response.Write " (" & getShowSize(asize) & ")"
end if
Response.Write "---</option>"

do while i < allnum
	am.GetInfoByIndex i, aid, aname, asize

	if Mid(aid, 15, 2) <> "_z" then
		Response.Write "<option value=""" & aid & """>" & server.htmlencode(aname) & " (" & getShowSize(asize) & ")</option>"
	end if

	aid = NULL
	aname = NULL
	asize = NULL

	i = i + 1
loop
%>
		</select>
        <input name="upfile" id="upfile" type="file" class='textbox' onchange="javascript:uf_changed()">&nbsp;&nbsp;&nbsp;
		<input type="button" name="msub" id="msub" value="<%=s_lang_0288 %>" onclick="javascript:mysub()" class=sbttn>
<%
if trim(request("errcode")) <> "" then
	Response.Write "<br><font color='#FF3333'>*" & s_lang_0289 & ": " & trim(request("errcode")) & "</font>"
end if
%>
      </td>
    </tr>
  </table>
<input name="f4_up_mode" id="f4_up_mode" type="hidden">
<input name="f4_zatt_id" id="f4_zatt_id" type="hidden" value="<%=zid %>">
<input name="f4_zatt_size" id="f4_zatt_size" type="hidden" value="<%=trim(request("zsize")) %>">
<input name="f4_zatt_grsn" id="f4_zatt_grsn" type="hidden" value="<%=trim(request("grsn")) %>">
</FORM>
</BODY>
</HTML>

<%
set am = nothing

function getShowSize(bytesize)
	if bytesize < 1000 then
		getShowSize = "1K"
	else
		if bytesize < 1000000 then
			getShowSize = CLng(bytesize/1000) & "K"
		else
			tmpSize = CStr(CDbl(bytesize/1000000))
			tmpindex = InStr(1, tmpSize, ".")
			if tmpindex = 0 then
				getShowSize = tmpSize & "M"
			else
				getShowSize = Left(tmpSize, tmpindex + 2) & "M"
			end if
		end if
	end if
end function
%>
