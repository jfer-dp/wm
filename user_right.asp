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
.span_head {font-family:����,MS SONG,SimSun,tahoma,sans-serif; font-size:12px; cursor:default; width:190px; background:#104A7B; color:#fff; padding-left:10px;}
#tabs {padding:0; margin:0 0 0 15px; list-style:none;}
#tabs li {display:inline; padding:0; background:#f8f8f8; float:left; position:relative;}
#tabs li.tb {width:190px; margin:20px 20px 0px;}
#tabs li a.outer-link {background:#f8f8f8; display:block; width:100%; position:relative;}
#tabs table {margin:-1px; border:0px;}
#tabs li div {border:1px solid #888; border-width:0 1px 1px 1px; padding:8px 5px 5px 5px; font-family:����,MS SONG,SimSun,tahoma,sans-serif; font-size:9pt; width:190px; cursor:pointer; background:#fff; color:#000; word-break: break-all;}
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
<b>��������</b><br>
���ĸ�����Ϣ����¼����
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">2</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('style.asp')">
<b>��������</b><br>
����ʹ����һЩ���õ�����
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">3</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('userspamguard.asp')">
<b>��ֹ�����ʼ�</b><br>
���ô˹��߷�ֹ���ռ��䡱�յ������ʼ�
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">4</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('cgfilter.asp')">
<b>�ʼ��ּ�����</b><br>
�˹��ܿ��Դ��������зּ�������������ʼ������ռ�����
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">5</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('logon.asp')">
<b>�����޸����ʺű���</b><br>
�붨�ڸ����������벢��д�ʺű�����Ϣ
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">6</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showusersetup.asp')">
<b>�Զ��ظ���ת��</b><br>
�趨�Զ��ظ��Լ��Զ�ת��ѡ��
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">7</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('userfiltermail.asp')">
<b>�ʼ�����</b><br>
�Զ������Ž��зּ�ʹ���
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">8</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showautoreplyex.asp')">
<b>��ǿ���Զ��ظ�</b><br>
����֧�ֺ궨���ͨ���Զ��ظ�����
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">9</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showuserkill.asp')">
<b>�ʼ�����</b><br>
ָ������������ʼ��ĵ����ʼ���ַ
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">10</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('trusty.asp')">
<b>�����б�</b><br>
���Ը��б���Email���ʼ���Զ���ᱻ����
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">11</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('Signature.asp')">
<b>����ǩ��</b><br>
������ӵ��ʼ��е�ǩ��
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">12</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showuserpop.asp')">
<b>��POP3�ʼ�����</b><br>
�����������е��ʼ�ͨ��POP3Э����ȡ����ϵͳ��
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">13</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('newaddres_1.asp')">
<b>��ϵ�˴�������</b><br>
������������ϵ���б�
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">14</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('rpfw.asp')">
<b>�ظ�/ת��ģ��</b><br>
���ÿ����ڻظ��Լ�ת���ʼ�ʱ���õ�ģ��
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">15</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('userfeast.asp')">
<b>�����û�����</b><br>
���ÿ�����Ч���ֲ�����ʾ�Ľ�����Ϣ
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
