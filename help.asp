<!--#include file="passinc.asp" --> 

<html>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<TITLE>WinWebMail</TITLE>
<LINK href="images\hwem.css" rel=stylesheet>
<script language="JavaScript">
<!-- 
function doZoom(size){
	document.getElementById('zoom').style.fontSize=size+'px'
}


var DOM = (document.getElementById) ? 1 : 0;
var NS4 = (document.layers) ? 1 : 0;
var IE4 = 0;
if (document.all)
{
	IE4 = 1;
	DOM = 0;
}

var win = window;   
var n   = 0;

function findIt() {
	if (document.getElementById("searchstr").value != "")
		findInPage(document.getElementById("searchstr").value);
}


function findInPage(str) {
var txt, i, found;

if (str == "")
	return false;

if (DOM)
{
	win.find(str, false, true);
	return true;
}

if (NS4) {
	if (!win.find(str))
		while(win.find(str, false, true))
			n++;
	else
		n++;

	if (n == 0)
		alert("未找到指定内容.");
}

if (IE4) {
	txt = win.document.body.createTextRange();

	for (i = 0; i <= n && (found = txt.findText(str)) != false; i++) {
		txt.moveStart("character", 1);
		txt.moveEnd("textedit");
	}

if (found) {
	txt.moveStart("character", -1);
	txt.findText(str);
	txt.select();
	txt.scrollIntoView();
	n++;
}
else {
	if (n > 0) {
		n = 0;
		findInPage(str);
	}
	else
		alert("未找到指定内容.");
	}
}

return false;
}
// -->
</script>

<BODY>
<br>
<div align="center">
  <table width="90%" border="0" align="center" cellspacing="0" bgcolor="<%=MY_COLOR_3 %>" style='border:1px <%=MY_COLOR_1 %> solid;' id="zoom">
    <tr bgcolor="<%=MY_COLOR_2 %>"> 
      <td height="50" style='border-bottom:1px <%=MY_COLOR_1 %> solid;'>
		<div align="center"><font class="s" color="<%=MY_COLOR_4 %>"><b>帮&nbsp;&nbsp;助</b></div>
		<div align="center">
<table width="98%"><tr><td align="left">
<input type="text" id="searchstr" name="searchstr" class="textbox" size="10">
<input type="button" value="页内查找" onclick="javascript:findIt();" class="sbttn">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</td><td align="right">
查看:[<a href="javascript:doZoom(16)">大字</a> <a href="javascript:doZoom(14)">中字</a> <a href="javascript:doZoom(12)">小字</a>]</font>
</td></tr></table>
</div>
      </td>
    </tr>
    <tr>
      <td> <br>
        <font color="#FF3333">登录系统时的增强安全性</font><br>
        <br>
		用来在共享计算机上增强安全性的登录选项是为那些从图书馆、学校或网吧登录的用户设计的. 该登录选项在您退出帐户时能够使浏览器高速缓存中的页面过期. 这表明一旦您退出, 您所访问的页将不能被共享计算机的其他用户查看.<br>
		注意: 由于页面没有被高速缓存到您的本地磁盘驱动器中, 因此在使用此选项时您会感觉速度变慢了.<br>
        <br>
      </td>
    </tr>
    <tr>
      <td> <br>
        <a name="writemail"></a><font color="#FF3333">查看其他语言的邮件内容</font><br>
		<br>
		当收到其他语言的邮件内容时, 页面可能会显示为乱码, 您可以通过调整IE浏览器的编码来正常阅读邮件内容.
		<br>方法是: 用鼠标右键点击邮件内容页面, 在弹出菜单的"编码"中选择正确的编码.
		<br><br>
      </td>
    </tr>
    <tr>
      <td> <br>
        <font color="#FF3333">报告垃圾邮件</font><br>
        <br>
		我们利用您报告为垃圾邮件的邮件来提高服务器过滤垃圾邮件的性能. 我们还可能将报告的垃圾邮件提交给第三方以一同反击垃圾邮件.<br>
        <br>
      </td>
    </tr>
    <tr>
      <td> <br>
        <a name="writemail"></a><font color="#FF3333">撰写HTML格式邮件</font><br>
        <br>
		如果您的 Web 浏览器是面向 Windows 的 Microsoft Internet Explorer 5.0 或更高版本, 您可以通过启用“个人配置”中的“使用HTML格式写邮件”选项来撰写HTML格式的电子邮件, 您将可以更改字体、更改字体大小以及颜色的选项, 此外, 还有加粗字体、添加下划线以及按照自己的风格来编排邮件的选项.<br>
        <br>
      </td>
    </tr>
    <tr>
      <td> <br>
        <font color="#FF3333">快速地址列表</font><br>
        <br>
		快速地址列表可以帮助您以最快捷的方式输入您所需要的邮件地址.<br>
        <br>
      </td>
    </tr>
    <tr>
      <td> <br>
        <font color="#FF3333">通讯组</font><br>
        <br>
		您可以将经常需要批量发送的邮件地址编辑成为通讯组, 在写邮件时您将可以直接输入通讯组名称, 从而实现将邮件一次发送给指定通讯组内所有成员的功能.<br>
        <br>
      </td>
    </tr>
    <tr> 
      <td> <br>
        <font color="#FF3333">附件上传</font><br>
        <br>
        1. 添加附件: 先按“浏览”, 选取您要添加的附件, 然后按下“上传附件”按钮即可, 此时会上传附件, 附件越大将使用越久的资料传送时间, 请您耐心等候.<br>注意: 单附件上传的最大长度不超过4M.<br>
        <br>
        2. 删除附件: 先选取附件, 再按“删除”即可.<br>
        <br>
      </td>
    </tr>
    <tr>
      <td> <br>
        <font color="#FF3333">存储文件夹</font><br>
        <br>
		1. 您可以上传或是从邮件附件中摘取文件并保存于您邮箱的网络存储空间中, 您将可以在发送邮件时直接将网络存储中的附件加入到待发邮件中.<br>
        <br>
		2. 因为存储文件夹使用的是您的邮箱空间, 所以您需要注意合理分配存储文件夹的数据量.<br>
		<br>
		3. 在您删除存储文件夹下的子文件夹时, 需要先将子文件夹中的文件转移或删除, 否则子文件夹无法被删除.<br>
		<br>
		4. 您可以通过设置密码或是不设置密码的方式共享您的存储文件夹数据.<br>
        <br>
      </td>
    </tr>
    <tr>
      <td> <br>
        <font color="#FF3333">发送系统邮件功能 (只有管理员才有此功能)</font><br>
		<br>
		管理员在选中此项功能后发出的邮件, 系统内用户在通过WebMail浏览此邮件时会看到醒目的标识.<br>
		<br>
      </td>
    </tr>
    <tr>
      <td> <br>
        <font color="#FF3333">私人邮件夹管理</font><br>
        <br>
        1. 新建私人邮件夹时, 邮件夹名称可以是数字、字符和汉字, 并支持长邮件夹名, 但不能使用一些系统保留的名称(如: in, out等).<br>
		<br>
		2. 在您删除私人邮件夹时, 需要先将此私人邮件夹中包含的所有邮件转移或删除, 否则邮件夹无法被删除.<br>
		<br>
		3. 您可以通过设置密码或是不设置密码的方式共享自己的私人邮件夹, 从而让系统中的其他用户可以共享您的资源.<br>
        <br>
      </td>
    </tr>
    <tr> 
      <td> <br>
        <font color="#FF3333">读取确认</font><br>
        <br>
        选中此功能后, 在对本系统内用户写信后, 当此收件人通过WebMail看到信或是通过POP3下载了您写给他的这封邮件时, 系统将会发一封回执给您. 
        但需注意的是: 此项功能只当收件人是系统内用户时才有效.<br>
        <br>
      </td>
    </tr>
	<tr>
	<td><a name="delivermail"></a><br>
		<font color="#FF3333">转递邮件</font><br>
		<br>
		转递邮件是将邮件原封不动的转交至指定地址的邮件传送方式.<br>
		<br>
	</td>
	</tr>
    <tr> 
      <td> <br>
        <font color="#FF3333">邮件查找</font><br>
        <br>
        本系统能按照时间或其它条件来协助您查找所有邮箱中的邮件.<br>
        <br>
        1. 在主题(发信地址、发信人)查找中可以按照通配符的方式来进行邮件查找, *: 代表任意长字符串, ?: 代表一个字符(一个中文字符需用两个??来代替).<br>
        <br>
        2. 日期查找时请注意缺省查找将只查找今天以前收到的邮件.<br>
        <br>
      </td>
    </tr>
	<tr>
		<td><a name="sharefolder"></a><br>
		<font color="#FF3333">文件夹共享设置 (将自己的私人信箱共享给其他用户)</font><br>
		<br>
		您可以使用此项功能将自己创建的私人信箱共享给其他用户(允许其他用户读取此文件夹中的内容).<br>
		共享方式共分为三种:<br>
		<b>1、不共享:</b> 此私人信箱不共享给其他任何用户.<br>
		<b>2、加密码后共享:</b> 此私人信箱允许共享, 但需要其他用户输入正确的密码后才可以查看此私人信箱中的内容.<br>
		<b>3、无密码共享:</b> 此私人信箱允许共享, 并且无需密码其他任何用户即可查看此私人信箱中的内容.<br>
		<br>
		<b>是否允许系统显示:</b> 选中该选项后, 此私人信箱将会被加入到"系统共享文件夹"列表中(您可以通过点击左边的"共享文件夹"来显示"系统共享文件夹"列表), 所有的用户都可以在此列表中看到您共享了一个私人信箱.<br>
		<br>
		<b>密码设定:</b> 当您只想修改私人信箱的共享方式、显示方式或直接共享给指定用户, 而不想修改原先已设定的密码时, 您只要不输入密码即可.<br>
		<br>
		<b>直接共享给指定用户:</b> 您可以将私人信箱直接共享给系统内的其他用户. (注: 建议您将共享方式设置为加密码后共享, 如果您设置为无密码共享时, 除了您指定的用户外, 其他用户也可以查看您共享的私人信箱)<br>
		<br>
		</td>
	</tr>
	<tr>
	<td><a name="ff_showall"></a><br>
		<font color="#FF3333">收藏其他用户共享出来的文件夹</font><br>
		<br>
		您可以使用此功能将其他用户共享出来的文件夹加入您收藏的共享文件夹列表(显示在上方的"我收藏的共享文件夹")中.<br>
		<br>
		加入方法分为两种:<br>
		<b>1、从显示在"系统共享文件夹"(显示在下方)中的列表里直接收藏:</b> 您可以直接点击任一显示在系统共享文件夹列表中的文件夹最后的图标来进行收藏.<br>
		注意: 如果此共享文件夹是加密码后共享时, 您必须要输入正确的共享密码时才可以正常查看其他用户共享出来的文件夹(具体的密码您可以询问此文件夹的所属用户).<br>
		<br>
		<b>2、直接收藏:</b> 当系统内的某一个用户共享了他自己的私人信箱后, 如设置为不允许系统显示时, 您将无法在"系统共享文件夹"中看到, 这时您必须使用直接收藏的方式才可以将他共享出来的私人信箱加入到您的共享文件夹列表中. 不过, 您将需要直接输入文件夹所属人的用户名、文件夹名以及密码.<br>
		注意: 以上两种加入方式, 当共享文件夹是无密码共享时, 您将无需输入任何密码.<br>
		<br>
		<b>修改:</b> 当共享文件夹已经被原共享人修改了共享密码时, 您可以使用修改功能更新此文件夹的密码, 修改成功后才可以继续查看此共享文件夹.
		<br>
	</td>
	</tr>
    <tr> 
      <td><a name="showuserpop"></a><br>
        <font color="#FF3333">多POP3接收代理</font><br>
        <br>
        如果你以前已经有了其它Email地址, 并且你的朋友都在向那些地址发信; 你就可以设置"POP3接收"功能, 让系统把你在其它地方的Email通过POP3协议提取到本系统中. 
        请在"服务器地址"中填写你的POP3服务器名称或地址, 如"pop.21cn.com", 然后填写你收取该服务器上邮件的帐号名称和口令, 如果你不知道你的服务器使用什么端口, 
        请使用缺省设置"110".<br>
        <br>
      </td>
    </tr>
	<tr>
	<td><a name="showuserkill"></a><br>
		<font color="#FF3333">用户拒收邮件地址</font><br>
		<br>
		对于您不想接收的邮件地址, 您可以将其加入到拒收列表中.<br>
		<br>
	</td>
	</tr>
	<tr>
	<td><a name="userfiltermail"></a><br>
		<font color="#FF3333">高级邮件过滤功能</font><br>
		<br>
		高级邮件过滤功能, 可以让系统帮助您自动将符合指定条件(“邮件地址”、“发件人”、“邮件大小”或“主题”)的邮件进行删除、自动回复或是移到垃圾箱的操作.<br><br>
		我们可以使用此功能对付日益增多的垃圾邮件和其他不受欢迎的邮件. 每一个过滤器的排列顺序是很重要的, 当用户接收到一封符合某一过滤器的邮件后, 如果此过滤器的"符合条件后的处理"是中止的话, 那么此邮件将不会使用余下的过滤器进行过滤.<br>
		<br>
	</td>
	</tr>
    <tr> 
      <td><a name="showusersetup"></a><br>
        <font color="#FF3333">启动相关功能</font><br>
        <br>
        系统中的POP3接收功能、邮件拒收功能、自动回复功能、自动转发功能将需要在启动该功能后, 才会被正式启用.<br>
        <br>
      </td>
    </tr>
	<tr>
	<td><a name="style"></a><br>
		<font color="#FF3333">帐号保护</font><br>
		<br>
		为防止您忘记密码后无法进入系统, 您需要填写帐号保护信息. 当您因为忘记密码而通过回答帐号保护问题进入邮箱后, 请立即修改您的密码.<br>
		<br>
	</td>
	</tr>
	<tr>
	<td><a name="calshare"></a><br>
		<font color="#FF3333">“私人”、“公开”和“显示正忙”的区别</font><br>
		<br>
		效率手册事件中的“私人”、“公开”和“显示正忙”许可, 决定他人查看您的公开效率手册时的事件显示方式. 如果要让他人看到事件的标题、说明和时间等信息, 则可将事件设置为“公开”. 如果要让其他人知道您在此事件期间正忙, 但他们又没有必要知道您在干什么, 则可以将事件设置为“显示正忙”. 如果不想让其他人看到您此期间计划的事件, 则将事件设置为“私人”.<br>
		<br>
	</td>
	</tr>
    <tr> 
      <td> <br>
        <font color="#FF3333">退出电子邮箱</font><br>
        <br>
        请不要以直接关闭浏览器的方法退出邮箱, 建议使用点击"退出"的方式, 然后再关闭所有浏览器, 这样将可确保您的信箱安全.<br>
        <br>
      </td>
    </tr>
  </table>
</div>
<br><br>
</BODY>
</html>