<%
Response.CacheControl = "no-cache" 
%>

<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
pw = trim(request("pwhidden"))
oldpw = trim(request("oldpwhidden"))

dim ei
set ei = server.createobject("easymail.UserWorkTimer")
ei.Load_User Session("wem")

Set em = Application("em")

if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	dim rating
	rating = 3
	if IsNumeric(trim(request("rating"))) = true then
		rating = CLng(trim(request("rating")))
	end if

	dim errstr
	errstr = s_lang_0191
	isok = true

	if ei.password_strong = 0 then
		ei.password_strong = 3
	end if

	if rating < ei.password_strong then
		errstr = s_lang_0192
		isok = false
	end if

	if pw = "" then
		errstr = s_lang_0191
		isok = false
	else
		if isok = true then
			pw = strDecode(pw, trim(request("picnum")))

			if em.CheckPassWord(Session("wem"), pw) = true then
				errstr = s_lang_0193
				isok = false
			else
				if pw <> "" then
					em.ChangeUserPassWord Session("wem"), pw
					isok = true
				else
					errstr = s_lang_0191
					isok = false
				end if
			end if
		end if
	end if

	set em = nothing
	set ei = nothing

	if isok = true then
		dim mam
		set mam = server.createobject("easymail.AdminManager")
		mam.Sleep 500
		set mam = nothing

		Response.Redirect "welcome.asp?" & getGRSN()
	else
		Response.Redirect "err.asp?errstr=" & Server.URLEncode(errstr) & "&" & getGRSN() & "&gourl=" & Server.URLEncode("newpw.asp")
	end if
end if
%>

<html>
<head>
<title><%=s_lang_0194 %></title>
<%=s_lang_meta %>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
</head>

<SCRIPT LANGUAGE=javascript>
function getElementLeft(element){
	var actualLeft = element.offsetLeft;
	var current = element.offsetParent;

	while (current !== null){
		actualLeft += current.offsetLeft;
		current = current.offsetParent;
	}
	return actualLeft;
}

function getElementTop(element){
	var actualTop = element.offsetTop;
	var current = element.offsetParent;

	while (current !== null){
		actualTop += current.offsetTop;
		current = current.offsetParent;
	}
	return actualTop;
}

function checkpw(){
	if (now_rt < <%=ei.password_strong %>)
	{
		alert("<%=s_lang_0192 %>! ");
		return ;
	}

	if (document.f1.pw1.value != "" && document.f1.pw2.value != "")
	{
		if (document.f1.pw1.value != document.f1.pw2.value)
			alert("<%=s_lang_0047 %>");
		else
		{
			document.form1.pwhidden.value = encode(document.f1.pw1.value, parseInt(document.form1.picnum.value));
			document.form1.submit();
		}
	}
}

function encode(datastr, bassnum) {
	var tempstr;
	var tchar;
	var newdata = "";

	for (var i = 0; i < datastr.length; i++)
	{
		tchar = 65535 + bassnum - datastr.charCodeAt(i);
		tchar = tchar.toString();

		while(tchar.length < 5)
		{
			tchar = "0" + tchar;
		}

		newdata = newdata + tchar;
	}

	return newdata;
}

	ratingMsgs = new Array(4);
	ratingMsgColors = new Array(4);
	barColors = new Array(4);
	ratingMsgs[0] = "<%=s_lang_0195 %>";
	ratingMsgs[1] = "<%=s_lang_0196 %>";
	ratingMsgs[2] = "<%=s_lang_0197 %>";
	ratingMsgs[3] = "<%=s_lang_0198 %>";
	ratingMsgColors[0] = "#676767";
	ratingMsgColors[1] = "#aa0033";
	ratingMsgColors[2] = "#f5ac00";
	ratingMsgColors[3] = "#008000";
	barColors[0] = "#dddddd";
	barColors[1] = "#aa0033";
	barColors[2] = "#ffcc33";
	barColors[3] = "#008000";
	var now_rt = 0;

function CreateRatePasswdReq(pwd){
	if (!isBrowserCompatible) return;

	if(!pwd) return; 
	passwd=pwd.value;
	var min_passwd_len = 6; 
	if (passwd.length < min_passwd_len){
		if (passwd.length > 0)
			DrawBar(0);
		else
			ResetBar();
	} else {
		rating = checkPasswdRate(passwd);
		DrawBar(rating);
	}
}

function getElement(name){
	if (document.all)
		return document.all(name);

	return document.getElementById(name);
}

function DrawBar(rating){
	now_rt = rating;
	getElement('rating').value = rating;
	var posbar = getElement('posBar');
	var negbar = getElement('negBar');
	var passwdRating = getElement('passwdRating');
	var barLength = getElement('passwdBar').width;
	if (rating >= 0 && rating < 4) {
		posbar.style.width = barLength / 3 * rating + "px";
		negbar.style.width = barLength / 3 * (3 - rating) + "px";
	}

	posbar.style.background = barColors[rating];
 	passwdRating.innerHTML = "<font color='" + ratingMsgColors[rating] + "'>" + ratingMsgs[rating] + "</font>";
}

function ResetBar(){
	var posbar = getElement('posBar');
	var negbar = getElement('negBar');
	var passwdRating = getElement('passwdRating');
	var barLength = getElement('passwdBar').width;
	posbar.style.width = "0px";
	negbar.style.width = barLength + "px";
	passwdRating.innerHTML = "";
}

	var agt = navigator.userAgent.toLowerCase();
	var is_op = (agt.indexOf("opera") != -1);
	var is_ie = (agt.indexOf("msie") != -1) && document.all && !is_op;
	var is_mac = (agt.indexOf("mac") != -1);
	var is_gk = (agt.indexOf("gecko") != -1);
	var is_sf = (agt.indexOf("safari") != -1);

function gff(str, pfx){
	var i = str.indexOf(pfx);
	if (i != -1)
	{
		var v = parseFloat(str.substring(i + pfx.length));
		if (!isNaN(v))
			return v;
	}
	return null;
}

function Compatible(){
	if (is_ie && !is_op && !is_mac)
	{
		var v = gff(agt, "msie ");
		if (v != null)
			return (v >= 6.0);
	}

	if (is_gk && !is_sf)
	{
		var v = gff(agt, "rv:");
		if (v != null)
			return (v >= 1.4);
		else
		{
			v = gff(agt, "galeon/");
			if (v != null)
				return (v >= 1.3);
		}
	}

	if (is_sf)
	{
		var v = gff(agt, "applewebkit/");
		if (v != null)
			return (v >= 124);
	}
	return false;
}

var isBrowserCompatible = Compatible();

function CharMode(iN){
	if (iN>=48 && iN <=57)
		return 1;
	if ((iN>=65 && iN <=90) || (iN>=97 && iN <=122))
		return 2;
	else
		return 4;
}

function bitTotal(num){
	modes=0;
	for (i=0;i<4;i++)
	{
		if (num & 1) modes++;
			num>>>=1;
	}
	return modes; 
}

function checkPasswdRate(sPW){
	if (sPW.length < 6)
		return 0;

	Modes=0;
	for (i=0;i<sPW.length;i++){
		Modes|=CharMode(sPW.charCodeAt(i));
	}

	return bitTotal(Modes);
}
</script>

<body style="padding-top:56px;">
<form name="form1" method="post" action="newpw.asp">
<input type="hidden" name="forget" value="<%=trim(request("forget")) %>">
<input type="hidden" name="pwhidden">
<input type="hidden" name="picnum" value="<%=createRnd() %>">
<input type="hidden" name="rating" id="rating">
</form>
<form name="f1">
<table width="60%" border="0" align="center" cellspacing="0">
<tr><td class="block_top_td" style="height:4px;"></td></tr>
<tr><td style="text-align:center; border-bottom:2px #a7c5e2 solid; font-size:14px; color:#093665; padding-left:6px; padding-bottom:6px;">
<%=s_lang_0199 %>
</td></tr>
<tr><td class="block_top_td" style="height:18px; _height:20px;"></td></tr>
</table>

<table width="50%" border="0" align="center" cellspacing="0">
	<tr><td align="right" vAlign=top width="32%" height="62" style="padding-top:12px;">
	<%=s_lang_0049 & s_lang_mh %>
	</td><td>
	<input type="password" name="pw1" id="pw1" maxlength="32" size="30" class="n_textbox" style="height:24px;" onkeyup="CreateRatePasswdReq(this)"><br>
	<table width="10" cellSpacing=0 cellPadding=0 border=0>
	<tr>
	<td vAlign=top noWrap width="0">&nbsp;<font face="Arial, sans-serif" size=-1><%=s_lang_0200 %>: </font></td>
	<td vAlign=top noWrap style="padding-top:2px;"><font face="Arial, sans-serif" color=#808080 size=-1>
	<strong><div id=passwdRating></div></strong></font>
	</td></tr>
	<tr><td colspan="2">
		<table cellSpacing=0 cellPadding=0 border=0><tr>
		<td width="50">&nbsp;</td>
		<td>
			<table id=passwdBar cellSpacing=0 cellPadding=0 width="150" bgColor=#ffffff border=0>
			<tr>
			<td id=posBar width=0% bgColor=#e0e0e0 height=4></td>
			<td id=negBar width="100%" bgColor=#e0e0e0 height=4></td>
			</tr>
			</table>
		</td></tr>
		</table>
	</td></tr>
	</table>
	</td></tr>
	<tr><td align="right" vAlign=top height="28" style="padding-top:15px; *padding-top:16px;">
	<%=s_lang_0050 & s_lang_mh %>
	</td><td vAlign=top style="padding-top:8px;">
	<input type="password" name="pw2" maxlength="32" size="30" class="n_textbox" style="height:24px;">
	</td></tr>
	</table>

<table width="60%" border="0" align="center" cellspacing="0">
<tr><td class="block_top_td" style="height:1px;"></td></tr>
<tr><td style="border-bottom:2px #a7c5e2 solid; font-size:16px; font-weight:bold; color:#093665; padding-left:6px;">&nbsp;</td></tr>
<tr><td class="block_top_td" style="height:16px; _height:18px;"></td></tr>
<tr><td align="right">
<a class='wwm_btnDownload btn_blue' style="*height:24px;" href="javascript:checkpw();"><%=s_lang_save %></a>&nbsp;&nbsp;
<a class='wwm_btnDownload btn_blue' style="*height:24px;" href="default.asp?logout=true"><%=s_lang_0201 %></a>
	</td></tr>
</table>
</form>
</body>
</html>

<%
set em = nothing
set ei = nothing

function createRnd()
	dim retval
	retval = getGRSN()

	if Len(retval) > 4 then
		retval = Right(retval, 4)
	end if

	if Left(retval, 1) = "0" then
		retval = "5" & Right(retval, 3)
	end if

	createRnd = retval
end function

function strDecode(sd_Data, sd_bassnum)
	dim sd_vChar
	dim sd_NewData
	dim sd_TempChar
	sd_vChar = 1

	do
		if sd_vChar > Len(sd_Data) then
			exit do
		end if

	    sd_TempChar = CLng(Mid(sd_Data, sd_vChar, 5))
		sd_TempChar = ChrW(65535 + sd_bassnum - sd_TempChar)

        sd_NewData = sd_NewData & sd_TempChar
		sd_vChar = sd_vChar + 5
	loop

	strDecode = sd_NewData
end function
%>
