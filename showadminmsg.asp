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
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>错误回复邮件</b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24"> 
        <div align="center">主　题：</div>
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
	<br><input type="submit" value=" 保存 " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
  </table>
  <p>&nbsp;</p>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>新用户欢迎信</b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24">
        <div align="center">主　题：</div>
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
	<br><input type="submit" value=" 保存 " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
  </table>
  <p>&nbsp;</p>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>邮件读取(下载)确认信</b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24">
        <div align="center">主　题：</div>
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
	<br><input type="submit" value=" 保存 " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
  </table>
  <p>&nbsp;</p>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>邮箱容量警告信</b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24"> 
        <div align="center">主　题：</div>
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
	<br><input type="submit" value=" 保存 " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
  </table>
  <p>&nbsp;</p>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>病毒警告信</b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24"> 
        <div align="center">主　题：</div>
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
	<br><input type="submit" value=" 保存 " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
  </table>
  <p>&nbsp;</p>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>非垃圾邮件确认信</b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24"> 
        <div align="center">主　题：</div>
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
	<br><input type="submit" value=" 保存 " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
	<tr><td colspan="2">
	<hr size="1">
 	</td></tr>
	<tr><td colspan="2">
您可以在"非垃圾邮件确认信"中使用宏变量:<br>
<font color="#FF3333"><b>%question%</b></font> 表示需要回答的问题内容&nbsp;(最重要)<br>
<font color="#FF3333">%date%</font> 表示当前日期<br>
<font color="#FF3333">%time%</font> 表示当前时间<br>
<font color="#FF3333">%sendname%</font> 表示来信的发件人名称<br>
<font color="#FF3333">%sendmail%</font> 表示来信的发件人邮件地址<br>
<font color="#FF3333">%subject%</font> 表示来信的标题<br>
<font color="#FF3333">%to%</font> 表示来信的收件人邮件地址<br><br>
 	</td></tr>
  </table>
  <p>&nbsp;</p>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>帐号到期警告信</b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24"> 
        <div align="center">主　题：</div>
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
	<br><input type="submit" value=" 保存 " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
	<tr><td colspan="2">
	<hr size="1">
 	</td></tr>
	<tr><td colspan="2">
宏变量:<br>
<font color="#FF3333"><b>%ExpDate%</b></font> 替换为帐号到期日期(YYYY-MM-DD)<br>
<font color="#FF3333"><b>%RemDays%</b></font> 替换为距离帐号到期的天数<br>
<font color="#FF3333">%ExpAccount%</font> 替换为被警告的帐号名称<br>
 	</td></tr>
  </table>
  <p>&nbsp;</p>

	<table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>">
    <tr bgcolor="<%=MY_COLOR_2 %>">
	<td colspan="2" height="28" style="border-bottom:1px <%=MY_COLOR_1 %> solid; border-top:1px <%=MY_COLOR_1 %> solid; border-left:1px <%=MY_COLOR_1 %> solid; border-right:1px <%=MY_COLOR_1 %> solid;">
        <div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>垃圾箱邮件统计信</b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="26%" height="24"> 
        <div align="center">主　题：</div>
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
	<br><input type="submit" value=" 保存 " name="submit" class="Bsbttn">&nbsp;&nbsp;
	<input type="button" value=" 取消 " onclick="javascript:location.href='right.asp?<%=getGRSN() %>';" class="Bsbttn">
	</td>
	</tr>
	<tr><td colspan="2">
	<hr size="1">
 	</td></tr>
	<tr><td colspan="2">
宏变量:<br>
<font color="#FF3333">%Account%</font> 替换为用户帐号名称<br>
<font color="#FF3333">%Email%</font> 替换为用户邮件地址<br>
<font color="#FF3333">%EmailsTotal%</font> 替换为垃圾箱中的邮件总数<br>
<br>
内容变量: (连续标记于正文的前几行, 以 % 号开头. <font color="#FF3333">注意: 此内容不会显示在发给用户的邮件中</font>)<br><br>
您需要修改第一行 <font color="#FF3333">%URL%=</font> 后的内容为邮件系统可以由外部访问的http地址, 并且最后需要由 <font color="#FF3333">/trashmsg.asp</font> 结尾. 比如: http://www.domain.com/mail/trashmsg.asp 或 http://mail.domain.com/trashmsg.asp<br>
<font color="#FF3333">重要</font>: 此内容设置不当, 会影响用户对垃圾箱邮件进行管理.</td></tr>
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
