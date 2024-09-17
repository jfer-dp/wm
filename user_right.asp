<!--#include file="passinc.asp" -->

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if
%>

<!DOCTYPE html>
<html>
<head>
<TITLE>WinWebMail</TITLE>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
</head>

<style type="text/css">
body {background: #fff; font-size: 11px; width:780px;}
.b_style {display:block; background:transparent url(images/tab_top.gif) no-repeat 0 0; padding:2px 0 0 5px; height:15px; font-size:0.9em; overflow:hidden;}
.span_head {font-family:宋体,MS SONG,SimSun,tahoma,sans-serif; font-size:12px; cursor:default; width:190px; background:#104A7B; color:#fff; padding-left:10px;}
#tabs {padding:0; margin:0 0 0 15px; list-style:none;}
#tabs li {display:inline; padding:0; background:#f8f8f8; float:left; position:relative;}
#tabs li.tb {width:190px; margin:20px 20px 0px;}
#tabs li a.outer-link {background:#f8f8f8; display:block; width:100%; position:relative;}
#tabs table {margin:-1px; border:0px;}
#tabs li div {border:1px solid #888; border-width:0 1px 1px 1px; padding:8px 5px 5px 5px; font-family:宋体,MS SONG,SimSun,tahoma,sans-serif; font-size:9pt; width:190px; cursor:pointer; background:#fff; color:#000; word-break: break-all;}
#tabs li a {color:#000; text-decoration:none;}
#tabs li a.inner-link {color:#c00; text-decoration:none;}
#tabs li a.inner-link:hover {text-decoration:underline; cursor:default;}
#tabs li td {background:#104A7B; border:0px; margin:0px; padding:0px;}
#tabs li.tb:hover, #tabs li.tb a.outer-link:hover {background:#ffc;}
</style>

<script type="text/javascript" src="images/sc_left.js"></script>
<script type="text/javascript" src="images/mglobal.js"></script>

<script language="JavaScript">
<!--
if (ie == 6)
	document.execCommand("BackgroundImageCache", false, true);

function theright(rurl) {
	var mrstr = String(Math.random());

	location.href = rurl + "?GRSN=" + mrstr.substring(2, 10);
}

function therightfol(rurl) {
	var mrstr = String(Math.random());

	location.href = rurl + "&GRSN=" + mrstr.substring(2, 10);
}
//-->
</script>

<body>
<ul id="tabs">

<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">1</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('myreginfo.asp')">
<b>个人资料</b><br>
您的个人信息及登录资料
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">2</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('style.asp')">
<b>邮箱配置</b><br>
邮箱使用中一些常用的配置
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">3</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('userspamguard.asp')">
<b>防止垃圾邮件</b><br>
利用此工具防止“收件箱”收到垃圾邮件
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">4</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('cgfilter.asp')">
<b>邮件分检助理</b><br>
此功能可以从垃圾信中分检出符合条件的邮件放入收件箱中
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">5</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('logon.asp')">
<b>密码修改与帐号保护</b><br>
请定期更改邮箱密码并填写帐号保护信息
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">6</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showusersetup.asp')">
<b>自动回复与转发</b><br>
设定自动回复以及自动转发选项
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">7</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('userfiltermail.asp')">
<b>邮件过滤</b><br>
自动将来信进行分检和处理
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">8</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showautoreplyex.asp')">
<b>增强型自动回复</b><br>
可以支持宏定义的通用自动回复功能
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">9</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showuserkill.asp')">
<b>邮件拒收</b><br>
指定您不想接收邮件的电子邮件地址
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">10</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('trusty.asp')">
<b>信任列表</b><br>
来自该列表中Email的邮件永远不会被过滤
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">11</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('Signature.asp')">
<b>个人签名</b><br>
创建添加到邮件中的签名
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">12</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showuserpop.asp')">
<b>多POP3邮件下载</b><br>
将其它邮箱中的邮件通过POP3协议提取到本系统中
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">13</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('newaddres_1.asp')">
<b>联系人创建程序</b><br>
帮助您创建联系人列表
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">14</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('rpfw.asp')">
<b>回复/转发模板</b><br>
设置可以在回复以及转发邮件时调用的模板
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">15</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('userfeast.asp')">
<b>设置用户节日</b><br>
设置可以在效率手册中显示的节日信息
</div>
</td></tr></table>
</li>

</div>

<div style="float:left; padding-top:<%
if isMSIE = true then
	Response.Write "90px;"
else
	Response.Write "120px;"
end if
%>">
</div>

</ul>
</body>
</html>
