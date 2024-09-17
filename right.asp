<!--#include file="passinc.asp" -->
<!--#include file="language.asp" --> 

<%
if isadmin() = false and Application("em_SpamAdmin") <> LCase(Session("wem")) then
	response.redirect "noadmin.asp"
end if

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
.span_head {font-family:<%=s_lang_font %>; font-size:12px; cursor:default; width:190px; background:#104A7B; color:#fff; padding-left:10px;}
#tabs {padding:0; margin:0 0 0 15px; list-style:none;}
#tabs li {display:inline; padding:0; background:#f8f8f8; float:left; position:relative;}
#tabs li.tb {width:190px; margin:20px 20px 0px;}
#tabs li a.outer-link {background:#f8f8f8; display:block; width:100%; position:relative;}
#tabs table {margin:-1px; border:0px;}
#tabs li div {border:1px solid #888; border-width:0 1px 1px 1px; padding:8px 5px 5px 5px; font-family:<%=s_lang_font %>; font-size:9pt; width:190px; cursor:pointer; background:#fff; color:#000; word-break: break-all;}
#tabs li a {text-decoration:none;}
#tabs li a.inner-link {color:#c00; text-decoration:none;}
#tabs li a.inner-link:hover {text-decoration:underline; cursor:default;}
#tabs li td {background:#104A7B; border:0px; margin:0px; padding:0px;}
#tabs li.tb:hover, #tabs li.tb a.outer-link:hover {background:#ffc;}
</style>

<script type="text/javascript" src="images/sc_left.js"></script>
<script type="text/javascript" src="images/mglobal.js"></script>

<script language="JavaScript">
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
</script>

<body>
<ul id="tabs">

<%
if isadmin() = true then
%>

<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">重要设置</span></b>
<div onmouseover="this.style.background='#fef'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('impset.asp')">
<b>重要设置</b><br>
确保邮件服务器正常运行的重要设置
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">1</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showsysinfo.asp')">
<b>系统设置</b><br>
对服务器运行的各项参数进行设置
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">2</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('webadmin.asp')">
<b>资源使用设置</b><br>
管理系统Web界面下的各项设置
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">3</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showdomain.asp')">
<b>域设置</b><br>
管理邮件系统域名
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">4</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showkill.asp')">
<b>黑名单</b><br>
管理系统拒绝服务范围的信息
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">5</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('logs.asp')">
<b>日志管理</b><br>
管理系统日志信息
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">6</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showadminmsg.asp')">
<b>系统回复设置</b><br>
管理各类型系统邮件内容
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">7</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showstakeout.asp')">
<b>监控管理</b><br>
管理系统监控对象信息
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">8</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('browmailinglist.asp')">
<b>邮件列表管理</b><br>
管理系统邮件列表信息
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">9</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showcondomain.asp')">
<b>域名控制&nbsp;&nbsp;(<a href="JavaScript:theright('s_showcondomain.asp')" style="color:blue;" onmouseover="this.style.textDecoration='underline';this.style.color='red';" onmouseout="this.style.textDecoration='none';this.style.color='blue';">简化</a>,&nbsp;<a href="JavaScript:theright('m_showcondomain.asp')" style="color:blue;" onmouseover="this.style.textDecoration='underline';this.style.color='red';" onmouseout="this.style.textDecoration='none';this.style.color='blue';">快速</a>)</b><br>
域高级属性设置, 包括域管理员及域空间分配等
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">10</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('flysheet.asp')">
<b>广告</b><br>
管理通过浏览器所发邮件尾部追加的信息
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">11</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showmisp.asp')">
<b>多ISP接收设置</b><br>
管理从多个ISP邮箱接收邮件的配置
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">12</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('webkill.asp')">
<b>拒绝Web登录设置</b><br>
设置不允许访问邮件系统Web界面的IP地址
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">13</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('keywords.asp')">
<b>主题关键字过滤</b><br>
设置接收及发送邮件时的过滤字
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">14</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('topsize.asp')">
<b>空间占用管理</b><br>
管理占用空间数最大的前100个用户
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">15</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('pendreg.asp')">
<b>邮箱申请审批</b><br>
对邮箱申请进行审批
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">16</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('signhold.asp')">
<b>保留帐号</b><br>
管理系统保留(不允许申请)的帐号名称
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">17</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('show_dca_domain.asp')">
<b>域邮件Catch All</b><br>
捕获所有发往本域的邮件(无限别名功能)
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">18</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('show_dm_domain.asp')">
<b>域邮件监控</b><br>
监控域用户收、发的邮件
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">19</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:therightfol('wframe.asp?mode=domainlistmail')">
<b>域邮件群发</b><br>
允许向所辖域发送群发邮件
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">20</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showwelcome.asp')">
<b>域欢迎邮件</b><br>
管理申请帐号后发给用户的域欢迎邮件内容
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">21</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showadv.asp')">
<b>域广告</b><br>
管理每个域通过浏览器所发邮件尾部追加的信息
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">22</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showddsize.asp')">
<b>域缺省邮箱大小</b><br>
设置域新建用户的缺省邮箱大小
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">23</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showkillattack.asp')">
<b>接入概率限制</b><br>
为避免某些IP占用过多接入资源, 可以限制它们的接入概率
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">24</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('accessip.asp')">
<b>不受限IP地址</b><br>
系统允许一些IP地址不受限制的连接服务器
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">25</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('trustemail.asp')">
<b><%=s_lang_0023 %></b><br>
<%=s_lang_0024 %>
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">26</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('exnamefilter.asp')">
<b>附件名称或类型过滤</b><br>
系统拒绝接收含有特定附件名称或类型的邮件
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">27</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('systemfilter.asp')">
<b>邮件内容过滤</b><br>
对接收到邮件的各项内容进行高级过滤设置
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">28</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('handpoint2.asp')">
<b>外发地址定向解析</b><br>
系统对外发邮件时的地址做定向解析
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">29</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('setreginfo.asp')">
<b>设置注册信息</b><br>
设置用户公开申请帐号时需要填写的信息内容
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">30</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('syscolor.asp')">
<b>颜色订制</b><br>
由系统管理员订制系统显示颜色
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">31</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('systemcollmail.asp')">
<b>邮件采集</b><br>
将符合条件的邮件收集到指定邮箱
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">32</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('trustuser.asp')">
<b>设置信任帐号</b><br>
信任帐号不受外发邮件数和抄送数的限制
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">33</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('tstyle.asp')">
<b>新用户设置模板</b><br>
可以为新创建帐号使用统一的Web设置内容
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">34</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('syshint.asp')">
<b>设置提醒信息</b><br>
设置在用户登录邮箱后随机显示的提醒信息
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">35</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('trapmail.asp')">
<b>垃圾邮件陷阱</b><br>
虚拟一个诱饵邮件地址做为陷阱, 将所有发往这个邮址的邮件都判为垃圾邮件
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">36</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('relayserver.asp')">
<b>设置中继服务器</b><br>
设置使用中继服务器转发邮件的参数
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">37</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('killhelo.asp')">
<b>拒绝 HELO/EHLO 主机名</b><br>
设置拒绝 HELO/EHLO 的主机名以及域名
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">38</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('feast.asp')">
<b>设置系统节日</b><br>
设置可以在效率手册中显示的自订制节日信息
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">39</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('greylisting.asp')">
<b><%=s_lang_0005 %></b><br>
<%=s_lang_0021 %>
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">40</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('authtrustip.asp')">
<b><%=s_lang_0026 %></b><br>
<%=s_lang_0027 %>
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">41</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('ldapex.asp')">
<b><%=s_lang_0063 %></b><br>
<%=s_lang_0064 %>
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">42</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('groupmail.asp')">
<b><%=s_lang_0078 %></b><br>
<%=s_lang_0079 %>
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">43</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('ss_brow.asp')">
<b><%=s_lang_0111 %></b><br>
<%=s_lang_0112 %>
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">44</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('somrank.asp')">
<b><%=s_lang_0203 %></b><br>
<%=s_lang_0204 %>
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">45</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('checkrbl.asp')">
<b><%=s_lang_0560 %></b><br>
<%=s_lang_0561 %>
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">46</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('setsh.asp')">
<b><%=s_lang_0588 %></b><br>
<%=s_lang_0597 %>
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">47</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('setrou.asp')">
<b><%=s_lang_0615 %></b><br>
<%=s_lang_0616 %>
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">48</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('mdallstyle.asp')">
<b><%=s_lang_0624 %></b><br>
<%=s_lang_0625 %>
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">49</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('banjia.asp')">
<b>邮件搬家</b><br>
将相同帐号名称的邮箱中收件箱内的邮件从其他邮局接收到系统内
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">50</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('wxmanager.asp')">
<b>微信邮件</b><br>
设置微信通知以及微邮件功能
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

<%
else
%>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">1</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('ss_brow.asp')">
<b><%=s_lang_0111 %></b><br>
<%=s_lang_0112 %>
</div>
</td></tr></table>
</li>

<%
end if
%>

</ul>
</body>
</html>
