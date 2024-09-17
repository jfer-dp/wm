<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
if IsEnterpriseVersion = false then
	Response.Redirect "noadmin.asp"
end if

Session("SearchStr") = ""

dim march
set march = server.createobject("easymail.MailArchive")
march.Load Session("wem")

md = trim(request("mode"))
mth = trim(request("mth"))

if md = "del" and Len(mth) = 6 and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	march.Del_month mth
end if

if md = "recount" and Request.ServerVariables("REQUEST_METHOD") = "POST" then
	march.Recount true
else
	march.Recount false
end if

allnum = march.Count
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<TITLE>WinWebMail</TITLE>
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

<style>
span {
	text-decoration:underline;
	cursor:pointer;
	font-weight: normal;
	font-size: 12px;
	margin: 10px;
}
span:hover { text-decoration:underline;color: red}

.basic  {
	width: 460px;
	border: 1px solid #999999;
}
.basic a {
	cursor:pointer;
	text-align:center;
	display:block;
	padding:5px;
	margin-top: 0;
	text-decoration: none;
	font-size: 12px;
	color: #333333;
	background-color: #e8e8e8;
	border-top: 1px solid #FFFFFF;
	border-bottom: 1px solid #999;
}
.basic a.selected {
	cursor:default;
	color: white;
	background-color: #113653;
}
table {
	margin-left:12px;
	*margin-top:5px;
	margin-bottom:5px;
}
table:td {padding:2px;}
</style>

<script type="text/javascript" src="images/arc.jquery.js"></script>
<script type="text/javascript" src="images/arc.jquery.accordion.js"></script>

<script type="text/javascript">
<!--
	jQuery().ready(function(){
		jQuery('#listarc').accordion({
			autoheight: false
		});
	});

function showone(vol) {
	location.href = "listarc.asp?date=" + vol + "&<%=getGRSN() %>";
}

var td_old_id = null;
function m_over(tag_obj, tid) {
if (td_old_id != null)
	document.getElementById(td_old_id).style.display = "none";

	document.getElementById(tid).style.display = "inline";
	td_old_id = tid;
}

function del_mon(mvol) {
	var t_y = mvol.substr(0, 4);
	var t_m;

	if (mvol.charAt(4) == "0")
		t_m = mvol.charAt(5);
	else
		t_m = mvol.substr(4, 2);

	if (confirm("<%=s_lang_0580 %>" + t_y + "<%=s_lang_0581 %>" + t_m + "<%=s_lang_0582 %>") == false)
		return ;

	document.f1.mode.value = "del";
	document.f1.mth.value = mvol;
	document.f1.submit();
}

function recount() {
	document.f1.mode.value = "recount";
	document.f1.submit();
}
//-->
</script>
</head>

<body>
<form name="f1" method="post" action="showarchive.asp?<%=getGRSN() %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="mth" value="">
</form>
<table width="96%" border="0" align="center" bgcolor="#EFF7FF" cellspacing="0" style="border:1px solid #8CA5B5; margin-top:4px;">
	<tr><td width="60%" nowrap align="left" height="28" style="padding-left:4px;">
	<a class='wwm_btnDownload btn_blue' href="javascript:recount();"><%=s_lang_0086 %></a>&nbsp;&nbsp;&nbsp;
<%
Response.Write s_lang_0583 & march.All_Archive_Count
%>
	</td>
	<td nowrap align="right" style="padding-right:8px; color:#444444;"><%=s_lang_0570 %>&nbsp;(<%
max_arc = march.Max_Archive

if max_arc < 1 then
	Response.Write s_lang_0584
else
	Response.Write s_lang_0585 & max_arc
end if
%>)</td>
	</tr>
</table>
<br>

<%
if allnum > 0 then
%>
<table width="96%" border="0" align="center" cellspacing="0">
<tr><td>
<div class="basic" style="float:left;" id="listarc">
<%
i = 0
curyearstr = ""
isshowone = false

do while i < allnum
	march.Get i, date_str, date_count

	temp_y = Left(date_str, 4)
	temp_m = get_month(date_str)

	if curyearstr <> temp_y then
		isshowone = true

		if Len(curyearstr) > 0 then
			Response.Write "</table></div>" & Chr(13)
		end if

		Response.Write "<a>" & s_lang_0586 & temp_y & s_lang_0587 & "</a><div>" & Chr(13)
		Response.Write "<table width='100%' border='0' align='center' cellspacing='0'>" & Chr(13)

		curyearstr = temp_y
	end if

	Response.Write "<tr onmouseover='m_over(this, " & date_str & ")'><td>" & Chr(13)
	Response.Write "<span onclick='javascript:showone(" & date_str & ")'>" & temp_y & s_lang_0581 & temp_m & s_lang_0582 &  "&nbsp;(" & date_count & ")</span>" & Chr(13)
	Response.Write "&nbsp;<img id='" & date_str & "' src='images/del.gif' align='absmiddle' border='0' style='cursor:pointer; display:none;' title='" & s_lang_0403 & "' onclick='del_mon(""" & date_str & """)'><br></td></tr>" & Chr(13)

	date_str = NULL
	date_count = NULL

	i = i + 1
loop

if isshowone = true then
	Response.Write "</table></div>" & Chr(13)
end if
%>
</div>
</td></tr>
</table>
<%
end if
%>

</body>
</html>

<%
set march = nothing


function get_month(date_str)
	if Len(date_str) = 6 then
		tmp_month = Mid(date_str, 5, 2)
		if Mid(tmp_month, 1, 1) = "0" then
			tmp_month = Mid(tmp_month, 2, 1)
		else
			tmp_month = Mid(tmp_month, 1, 2)
		end if

		get_month = tmp_month
	else
		get_month = ""
	end if
end function
%>
