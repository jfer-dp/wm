<!--#include file="passinc.asp" -->
<!--#include file="language-1.asp" --> 

<%
	Session("cert_imp_type") = ""
	Session("cert_imp_pw") = ""
	Session("cert_imp_save_day") = ""
%>

<!DOCTYPE html>
<HTML<%=s_lang_html %>>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=<%=s_lang_charset %>">
<link rel="stylesheet" type="text/css" href="images/owin.css">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">

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
#tabs li a {color:#000; text-decoration:none;}
#tabs li a.inner-link {color:#c00; text-decoration:none;}
#tabs li a.inner-link:hover {text-decoration:underline; cursor:default;}
#tabs li td {background:#104A7B; border:0px; margin:0px; padding:0px;}
#tabs li.tb:hover, #tabs li.tb a.outer-link:hover {background:#ffc;}
</style>
</head>

<script type="text/javascript" src="images/sc_left.js"></script>
<script type="text/javascript" src="images/mglobal.js"></script>

<script type="text/javascript">
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
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('cert_createca.asp')">
<b><%=a_lang_053 %></b><br><br>
<%=a_lang_054 %>
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">2</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:therightfol('cert_imp.asp?im=sec')">
<b><%=a_lang_055 %></b><br><br>
<%=a_lang_056 %>
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">3</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('cert_mysec.asp')">
<b><%=a_lang_057 %></b><br><br>
<%=a_lang_058 %>
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">4</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('cert_mypub.asp')">
<b><%=a_lang_059 %></b><br><br>
<%=a_lang_060 %>
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">5</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('cert_myothpub.asp')">
<b><%=a_lang_061 %></b><br><br>
<%=a_lang_062 %>
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">6</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('cert_share.asp')">
<b><%=a_lang_063 %></b><br><br>
<%=a_lang_064 %>
</div>
</td></tr></table>
</li>

</div>
<div style="float:left; padding-top:120px; _padding-top:90px;">
</div>

<table width="90%" border="0" align="left" cellspacing="0" bgcolor="#EFF7FF" style='border-top:1px #8CA5B5 solid; margin-top:90px; margin-left:10px; margin-bottom:20px;'>
	<tr>
		<td width="6%" valign="top" style="padding-top:8px;">&nbsp;<img src='images\remind.gif' border='0' align='absmiddle'></td>
		<td width="94%" style="padding:8px 8px 8px 2px; color:#444444;">1. 使用数字证书功能的第一步是先"获取用户数字证书", 即从服务器申请自己的数字证书钥匙环(公共钥匙证书与私人钥匙证书). 已有数字证书的用户可跳过此步骤.
		<br>注意: 尽管您可以申请多个相同名称的数字证书钥匙环, 但是<font color='#901111'>每一个钥匙环的<b>证书指纹</b>以及<b>ID</b>都是不同的</font>, 因此他们是完全不同的数字证书钥匙环.
		<br><br>2. 导入自己包含公、私钥的数字证书.
		<br><br>3. <font color='#901111'>请妥善保管自己的数字证书私人钥匙, 以防别人获得</font>, 因为只有将私人钥匙和私人钥匙密码一起配合使用时才可以解读发给你的加密信息.
		<br>以下是用最快的NFS分解算法攻击相应的密匙长度所耗费的机时年数对应表:
		<br>512位:&nbsp;&nbsp;30,000 年
		<br>768位:&nbsp;&nbsp;200,000,000 年
		<br>1024位:&nbsp;&nbsp;300,000,000,000 年
		<br>2048位:&nbsp;&nbsp;300,000,000,000,000,000,000 年
		<br><br>4. 接下来便是要散布您自己的公共钥匙. 如果您不这样做的话, 将没有人可以发送加密邮件给您. 举例来说, 若用户A想要发送加密邮件给您, 必须要有您的公共钥匙, 
		并使用您的公共钥匙来加密, 这样您在收到加密邮件时才能用您的私人钥匙解开. 若用户A要加密给您时, 是使用A自己或其它人的公共钥匙, 那这封加密邮件便只有用户A自己或相对的某人才能解得开.
		所以让越多人拿到您的公共钥匙越好, 这样也就会有越多人可以送加密邮件给您了.
		<br><br>5. 加入别人的公共钥匙. 当然, 别人也会散播他们的公共钥匙, 所以当您得到别人的公共钥匙后, 便是将其加入您收藏的数字证书, 这样才能给别人发送加密的邮件.
		<br>
		<br><br>注1: 加密邮件只加密邮件的正文部分, 不包括邮件头(主题)部分, 以及附件部分.
		<br><br>注2: 数字签名邮件是利用您的私人数字证书对邮件进行签名, 从而可以证明这封邮件确实是您写的(因为您的私人数字证书应该只有您才会持有), 仅使用"数字签名"时将不会加密邮件的正文内容, 除非您使用"数字签名并加密"的选项.
		<br><br>注3: 为方便表述, 以上的 "公共钥匙" 即为 "公共数字证书", "私人钥匙" 即为 "私人数字证书".
        </td>
	</tr>
</table>
</body>
</html>
