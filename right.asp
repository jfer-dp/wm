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
<b class="b_style"><span class="span_head">��Ҫ����</span></b>
<div onmouseover="this.style.background='#fef'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('impset.asp')">
<b>��Ҫ����</b><br>
ȷ���ʼ��������������е���Ҫ����
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">1</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showsysinfo.asp')">
<b>ϵͳ����</b><br>
�Է��������еĸ��������������
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">2</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('webadmin.asp')">
<b>��Դʹ������</b><br>
����ϵͳWeb�����µĸ�������
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">3</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showdomain.asp')">
<b>������</b><br>
�����ʼ�ϵͳ����
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">4</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showkill.asp')">
<b>������</b><br>
����ϵͳ�ܾ�����Χ����Ϣ
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">5</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('logs.asp')">
<b>��־����</b><br>
����ϵͳ��־��Ϣ
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">6</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showadminmsg.asp')">
<b>ϵͳ�ظ�����</b><br>
���������ϵͳ�ʼ�����
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">7</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showstakeout.asp')">
<b>��ع���</b><br>
����ϵͳ��ض�����Ϣ
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">8</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('browmailinglist.asp')">
<b>�ʼ��б����</b><br>
����ϵͳ�ʼ��б���Ϣ
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">9</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showcondomain.asp')">
<b>��������&nbsp;&nbsp;(<a href="JavaScript:theright('s_showcondomain.asp')" style="color:blue;" onmouseover="this.style.textDecoration='underline';this.style.color='red';" onmouseout="this.style.textDecoration='none';this.style.color='blue';">��</a>,&nbsp;<a href="JavaScript:theright('m_showcondomain.asp')" style="color:blue;" onmouseover="this.style.textDecoration='underline';this.style.color='red';" onmouseout="this.style.textDecoration='none';this.style.color='blue';">����</a>)</b><br>
��߼���������, ���������Ա����ռ�����
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">10</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('flysheet.asp')">
<b>���</b><br>
����ͨ������������ʼ�β��׷�ӵ���Ϣ
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">11</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showmisp.asp')">
<b>��ISP��������</b><br>
����Ӷ��ISP��������ʼ�������
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">12</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('webkill.asp')">
<b>�ܾ�Web��¼����</b><br>
���ò���������ʼ�ϵͳWeb�����IP��ַ
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">13</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('keywords.asp')">
<b>����ؼ��ֹ���</b><br>
���ý��ռ������ʼ�ʱ�Ĺ�����
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">14</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('topsize.asp')">
<b>�ռ�ռ�ù���</b><br>
����ռ�ÿռ�������ǰ100���û�
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">15</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('pendreg.asp')">
<b>������������</b><br>
�����������������
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">16</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('signhold.asp')">
<b>�����ʺ�</b><br>
����ϵͳ����(����������)���ʺ�����
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">17</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('show_dca_domain.asp')">
<b>���ʼ�Catch All</b><br>
�������з���������ʼ�(���ޱ�������)
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">18</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('show_dm_domain.asp')">
<b>���ʼ����</b><br>
������û��ա������ʼ�
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">19</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:therightfol('wframe.asp?mode=domainlistmail')">
<b>���ʼ�Ⱥ��</b><br>
��������Ͻ����Ⱥ���ʼ�
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">20</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showwelcome.asp')">
<b>��ӭ�ʼ�</b><br>
���������ʺź󷢸��û�����ӭ�ʼ�����
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">21</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showadv.asp')">
<b>����</b><br>
����ÿ����ͨ������������ʼ�β��׷�ӵ���Ϣ
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">22</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showddsize.asp')">
<b>��ȱʡ�����С</b><br>
�������½��û���ȱʡ�����С
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">23</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('showkillattack.asp')">
<b>�����������</b><br>
Ϊ����ĳЩIPռ�ù��������Դ, �����������ǵĽ������
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">24</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('accessip.asp')">
<b>������IP��ַ</b><br>
ϵͳ����һЩIP��ַ�������Ƶ����ӷ�����
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
<b>�������ƻ����͹���</b><br>
ϵͳ�ܾ����պ����ض��������ƻ����͵��ʼ�
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">27</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('systemfilter.asp')">
<b>�ʼ����ݹ���</b><br>
�Խ��յ��ʼ��ĸ������ݽ��и߼���������
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">28</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('handpoint2.asp')">
<b>�ⷢ��ַ�������</b><br>
ϵͳ���ⷢ�ʼ�ʱ�ĵ�ַ���������
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">29</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('setreginfo.asp')">
<b>����ע����Ϣ</b><br>
�����û����������ʺ�ʱ��Ҫ��д����Ϣ����
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">30</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('syscolor.asp')">
<b>��ɫ����</b><br>
��ϵͳ����Ա����ϵͳ��ʾ��ɫ
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">31</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('systemcollmail.asp')">
<b>�ʼ��ɼ�</b><br>
�������������ʼ��ռ���ָ������
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">32</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('trustuser.asp')">
<b>���������ʺ�</b><br>
�����ʺŲ����ⷢ�ʼ����ͳ�����������
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">33</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('tstyle.asp')">
<b>���û�����ģ��</b><br>
����Ϊ�´����ʺ�ʹ��ͳһ��Web��������
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">34</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('syshint.asp')">
<b>����������Ϣ</b><br>
�������û���¼����������ʾ��������Ϣ
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">35</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('trapmail.asp')">
<b>�����ʼ�����</b><br>
����һ���ն��ʼ���ַ��Ϊ����, �����з��������ַ���ʼ�����Ϊ�����ʼ�
</div>
</td></tr></table>
</li>

</div>
<div style="float:left;">

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">36</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('relayserver.asp')">
<b>�����м̷�����</b><br>
����ʹ���м̷�����ת���ʼ��Ĳ���
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">37</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('killhelo.asp')">
<b>�ܾ� HELO/EHLO ������</b><br>
���þܾ� HELO/EHLO ���������Լ�����
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">38</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('feast.asp')">
<b>����ϵͳ����</b><br>
���ÿ�����Ч���ֲ�����ʾ���Զ��ƽ�����Ϣ
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
<b>�ʼ����</b><br>
����ͬ�ʺ����Ƶ��������ռ����ڵ��ʼ��������ʾֽ��յ�ϵͳ��
</div>
</td></tr></table>
</li>

<li class="tb">
<table><tr><td>
<b class="b_style"><span class="span_head">50</span></b>
<div onmouseover="this.style.background='#ffc'; this.style.cursor='pointer';" onmouseout="this.style.background='#fff';" onClick="JavaScript:theright('wxmanager.asp')">
<b>΢���ʼ�</b><br>
����΢��֪ͨ�Լ�΢�ʼ�����
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
