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
		<td width="94%" style="padding:8px 8px 8px 2px; color:#444444;">1. ʹ������֤�鹦�ܵĵ�һ������"��ȡ�û�����֤��", ���ӷ����������Լ�������֤��Կ�׻�(����Կ��֤����˽��Կ��֤��). ��������֤����û��������˲���.
		<br>ע��: ������������������ͬ���Ƶ�����֤��Կ�׻�, ����<font color='#901111'>ÿһ��Կ�׻���<b>֤��ָ��</b>�Լ�<b>ID</b>���ǲ�ͬ��</font>, �����������ȫ��ͬ������֤��Կ�׻�.
		<br><br>2. �����Լ���������˽Կ������֤��.
		<br><br>3. <font color='#901111'>�����Ʊ����Լ�������֤��˽��Կ��, �Է����˻��</font>, ��Ϊֻ�н�˽��Կ�׺�˽��Կ������һ�����ʹ��ʱ�ſ��Խ��������ļ�����Ϣ.
		<br>������������NFS�ֽ��㷨������Ӧ���ܳ׳������ķѵĻ�ʱ������Ӧ��:
		<br>512λ:&nbsp;&nbsp;30,000 ��
		<br>768λ:&nbsp;&nbsp;200,000,000 ��
		<br>1024λ:&nbsp;&nbsp;300,000,000,000 ��
		<br>2048λ:&nbsp;&nbsp;300,000,000,000,000,000,000 ��
		<br><br>4. ����������Ҫɢ�����Լ��Ĺ���Կ��. ��������������Ļ�, ��û���˿��Է��ͼ����ʼ�����. ������˵, ���û�A��Ҫ���ͼ����ʼ�����, ����Ҫ�����Ĺ���Կ��, 
		��ʹ�����Ĺ���Կ��������, ���������յ������ʼ�ʱ����������˽��Կ�׽⿪. ���û�AҪ���ܸ���ʱ, ��ʹ��A�Լ��������˵Ĺ���Կ��, ���������ʼ���ֻ���û�A�Լ�����Ե�ĳ�˲��ܽ�ÿ�.
		������Խ�����õ����Ĺ���Կ��Խ��, ����Ҳ�ͻ���Խ���˿����ͼ����ʼ�������.
		<br><br>5. ������˵Ĺ���Կ��. ��Ȼ, ����Ҳ��ɢ�����ǵĹ���Կ��, ���Ե����õ����˵Ĺ���Կ�׺�, ���ǽ���������ղص�����֤��, �������ܸ����˷��ͼ��ܵ��ʼ�.
		<br>
		<br><br>ע1: �����ʼ�ֻ�����ʼ������Ĳ���, �������ʼ�ͷ(����)����, �Լ���������.
		<br><br>ע2: ����ǩ���ʼ�����������˽������֤����ʼ�����ǩ��, �Ӷ�����֤������ʼ�ȷʵ����д��(��Ϊ����˽������֤��Ӧ��ֻ�����Ż����), ��ʹ��"����ǩ��"ʱ����������ʼ�����������, ������ʹ��"����ǩ��������"��ѡ��.
		<br><br>ע3: Ϊ�������, ���ϵ� "����Կ��" ��Ϊ "��������֤��", "˽��Կ��" ��Ϊ "˽������֤��".
        </td>
	</tr>
</table>
</body>
</html>
