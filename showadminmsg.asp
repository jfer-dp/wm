<!--#include file="passinc.asp" --> 
<!--#include file="language.asp" --> 

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
set ei = server.createobject("easymail.adminmsg")
'-----------------------------------------
ei.Load

%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
</HEAD>

<BODY>
<br><br>
<FORM ACTION="saveadminmsg.asp" METHOD=POST NAME="f1">
	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>����ظ��ʼ�</b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24"> 
        <div align="center">�����⣺</div>
      </td>
      <td width="74%">  
        <input name="err_subject" type="text" value="<%=ei.errback_subject %>" size="50" class='textbox'>
        </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center">  
          <textarea name="err_text" cols="<%
if isMSIE = true then
	Response.Write "76"
else
	Response.Write "66"
end if
%>" rows="8" class='textarea'><%=ei.errback_text %></textarea>
          </div>
      </td>
    </tr>
    <tr> 
	<td colspan="2" align="right">
	<br><input type="submit" value=" ���� " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" ȡ�� " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
  </table>
  <p>&nbsp;</p>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>���û���ӭ��</b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24">
        <div align="center">�����⣺</div>
      </td>
      <td width="74%">  
        <input name="welcome_subject" type="text" value="<%=ei.welcome_subject %>" size="50" class='textbox'>
        </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center">  
          <textarea name="welcome_text" cols="<%
if isMSIE = true then
	Response.Write "76"
else
	Response.Write "66"
end if
%>" rows="8" class='textarea'><%=ei.welcome_text %></textarea>
          </div>
      </td>
    </tr>
    <tr> 
	<td colspan="2" align="right">
	<br><input type="submit" value=" ���� " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" ȡ�� " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
  </table>
  <p>&nbsp;</p>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>�ʼ���ȡ(����)ȷ����</b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24">
        <div align="center">�����⣺</div>
      </td>
      <td width="74%">  
        <input name="ReadBack_Subject" type="text" value="<%=ei.ReadBack_Subject %>" size="50" class='textbox'>
        </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center">  
          <textarea name="ReadBack_Text" cols="<%
if isMSIE = true then
	Response.Write "76"
else
	Response.Write "66"
end if
%>" rows="8" class='textarea'><%=ei.ReadBack_Text %></textarea>
          </div>
      </td>
    </tr>
    <tr> 
	<td colspan="2" align="right">
	<br><input type="submit" value=" ���� " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" ȡ�� " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
  </table>
  <p>&nbsp;</p>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>��������������</b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24"> 
        <div align="center">�����⣺</div>
      </td>
      <td width="74%">  
        <input name="Fill_Subject" type="text" value="<%=ei.Fill_Subject %>" size="50" class='textbox'>
        </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center">  
          <textarea name="Fill_Text" cols="<%
if isMSIE = true then
	Response.Write "76"
else
	Response.Write "66"
end if
%>" rows="8" class='textarea'><%=ei.Fill_Text %></textarea>
          </div>
      </td>
    </tr>
    <tr> 
	<td colspan="2" align="right">
	<br><input type="submit" value=" ���� " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" ȡ�� " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
  </table>
  <p>&nbsp;</p>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>����������</b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24"> 
        <div align="center">�����⣺</div>
      </td>
      <td width="74%">  
        <input name="Virus_Subject" type="text" value="<%=ei.Virus_Subject %>" size="50" class='textbox'>
        </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center">  
          <textarea name="Virus_Text" cols="<%
if isMSIE = true then
	Response.Write "76"
else
	Response.Write "66"
end if
%>" rows="8" class='textarea'><%=ei.Virus_Text %></textarea>
          </div>
      </td>
    </tr>
    <tr> 
	<td colspan="2" align="right">
	<br><input type="submit" value=" ���� " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" ȡ�� " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
  </table>
  <p>&nbsp;</p>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>�������ʼ�ȷ����</b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24"> 
        <div align="center">�����⣺</div>
      </td>
      <td width="74%">  
        <input name="NoSpam_Affirm_Subject" type="text" value="<%=ei.NoSpam_Affirm_Subject %>" size="50" class='textbox'>
        </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center">  
          <textarea name="NoSpam_Affirm_Text" cols="<%
if isMSIE = true then
	Response.Write "76"
else
	Response.Write "66"
end if
%>" rows="8" class='textarea'><%=ei.NoSpam_Affirm_Text %></textarea>
          </div>
      </td>
    </tr>
    <tr> 
	<td colspan="2" align="right">
	<br><input type="submit" value=" ���� " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" ȡ�� " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
	<tr><td colspan="2">
	<hr size="1">
 	</td></tr>
	<tr><td colspan="2">
��������"�������ʼ�ȷ����"��ʹ�ú����:<br>
<font color="#FF3333"><b>%question%</b></font> ��ʾ��Ҫ�ش����������&nbsp;(����Ҫ)<br>
<font color="#FF3333">%date%</font> ��ʾ��ǰ����<br>
<font color="#FF3333">%time%</font> ��ʾ��ǰʱ��<br>
<font color="#FF3333">%sendname%</font> ��ʾ���ŵķ���������<br>
<font color="#FF3333">%sendmail%</font> ��ʾ���ŵķ������ʼ���ַ<br>
<font color="#FF3333">%subject%</font> ��ʾ���ŵı���<br>
<font color="#FF3333">%to%</font> ��ʾ���ŵ��ռ����ʼ���ַ<br><br>
 	</td></tr>
  </table>
  <p>&nbsp;</p>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>�ʺŵ��ھ�����</b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24"> 
        <div align="center">�����⣺</div>
      </td>
      <td width="74%">  
        <input name="UserExp_Affirm_Subject" type="text" value="<%=ei.UserExp_Affirm_Subject %>" size="50" class='textbox'>
        </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center">  
          <textarea name="UserExp_Affirm_Text" cols="<%
if isMSIE = true then
	Response.Write "76"
else
	Response.Write "66"
end if
%>" rows="8" class='textarea'><%=ei.UserExp_Affirm_Text %></textarea>
          </div>
      </td>
    </tr>
    <tr> 
	<td colspan="2" align="right">
	<br><input type="submit" value=" ���� " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" ȡ�� " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
	<tr><td colspan="2">
	<hr size="1">
 	</td></tr>
	<tr><td colspan="2">
�����:<br>
<font color="#FF3333"><b>%ExpDate%</b></font> �滻Ϊ�ʺŵ�������(YYYY-MM-DD)<br>
<font color="#FF3333"><b>%RemDays%</b></font> �滻Ϊ�����ʺŵ��ڵ�����<br>
<font color="#FF3333">%ExpAccount%</font> �滻Ϊ��������ʺ�����<br>
 	</td></tr>
  </table>
  <p>&nbsp;</p>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>�������ʼ�ͳ����</b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24"> 
        <div align="center">�����⣺</div>
      </td>
      <td width="74%">  
        <input name="TrashMsg_Subject" type="text" value="<%=ei.TrashMsg_Subject %>" size="50" class='textbox'>
        </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center">  
          <textarea name="TrashMsg_Text" cols="<%
if isMSIE = true then
	Response.Write "76"
else
	Response.Write "66"
end if
%>" rows="8" class='textarea'><%=server.htmlencode(ei.TrashMsg_Text) %></textarea>
          </div>
      </td>
    </tr>
    <tr> 
	<td colspan="2" align="right">
	<br><input type="submit" value=" ���� " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" ȡ�� " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
	<tr><td colspan="2">
	<hr size="1">
 	</td></tr>
	<tr><td colspan="2">
�����:<br>
<font color="#FF3333">%Account%</font> �滻Ϊ�û��ʺ�����<br>
<font color="#FF3333">%Email%</font> �滻Ϊ�û��ʼ���ַ<br>
<font color="#FF3333">%EmailsTotal%</font> �滻Ϊ�������е��ʼ�����<br>
<br>
���ݱ���: (������������ĵ�ǰ����, �� % �ſ�ͷ. <font color="#FF3333">ע��: �����ݲ�����ʾ�ڷ����û����ʼ���</font>)<br><br>
����Ҫ�޸ĵ�һ�� <font color="#FF3333">%URL%=</font> �������Ϊ�ʼ�ϵͳ�������ⲿ���ʵ�http��ַ, ���������Ҫ�� <font color="#FF3333">/trashmsg.asp</font> ��β. ����: http://www.domain.com/mail/trashmsg.asp �� http://mail.domain.com/trashmsg.asp<br>
<font color="#FF3333">��Ҫ</font>: ���������ò���, ��Ӱ���û����������ʼ����й���.</td></tr>
  </table>
  <p>&nbsp;</p>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b><%=s_lang_0165 %></b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24"> 
        <div align="center"><%=s_lang_0166 %></div>
      </td>
      <td width="74%">  
        <input name="Recall_Subject" type="text" value="<%=ei.Recall_Subject %>" size="50" class='textbox'>
        </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center">  
          <textarea name="Recall_Text" cols="<%
if isMSIE = true then
	Response.Write "76"
else
	Response.Write "66"
end if
%>" rows="8" class='textarea'><%=ei.Recall_Text %></textarea>
          </div>
      </td>
    </tr>
    <tr> 
	<td colspan="2" align="right">
	<br><input type="submit" value=" <%=s_lang_save %> " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" <%=s_lang_cancel %> " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
	<tr><td colspan="2">
	<hr size="1">
 	</td></tr>
	<tr><td colspan="2">
<%=s_lang_0167 %>:<br>
<font color="#FF3333">%date%</font> <%=s_lang_0168 %><br>
<font color="#FF3333">%time%</font> <%=s_lang_0169 %><br>
<font color="#FF3333">%sendname%</font> <%=s_lang_0170 %><br>
<font color="#FF3333">%sendmail%</font> <%=s_lang_0171 %><br>
<font color="#FF3333">%subject%</font> <%=s_lang_0172 %><br>
 	</td></tr>
  </table>
</FORM>
<br>
</BODY>
</HTML>

<%
set ei = nothing
%>
