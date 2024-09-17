<%
wid = replace(trim(request("wid")), " ", "+")

dim wx
set wx = server.createobject("easymail.WXSet")
wx.GetInfo wid, wx_user, filename
is_expires = wx.IsExpires(filename, 3)
set wx = nothing

if is_expires = true then
	Response.End
end if

dim a
set a = server.createobject("easymail.emmail")

i = trim(request("count"))
pt = trim(request("pt"))

	if pt <> "" then
		a.LoadAll1 wx_user, filename, CDbl(pt)
	else
		a.LoadAll wx_user, filename
	end if


ishtml = false
if trim(request("ishtml")) = "1" then
	ishtml = true
end if

if ishtml = false then
	Response.ContentType = a.GetContentType(cint(i))
else
	Response.Charset = a.GetCharSet(cint(i))
	a.IsRemoveScript = true
	Response.Buffer = FALSE
%>
<script type="text/javascript">
function load(){
var fm=document.getElementsByTagName("iframe");
for(var i=0;i<fm.length;i++)
{
	if (window.stop)
		window.stop();
	else
		document.execCommand("Stop");

	var fobj=fm[i];
	fobj.src='about:blank';
	fobj.parentNode.removeChild(fobj);
	fobj.outerHTML='';
}

var cfob=document.getElementsByTagName("object");
for(var i=0;i<cfob.length;i++)
{
	if (window.stop)
		window.stop();
	else
		document.execCommand("Stop");

	var fobj=cfob[i];
	fobj.src='about:blank';
	fobj.parentNode.removeChild(fobj);
	fobj.outerHTML='';
}

var obj=document.getElementsByTagName("img");
for(var i=0;i<obj.length;i++)
{
	var objVal=obj[i];
	var src = objVal.getAttribute('src');  

	if (src.toLowerCase().indexOf("showatt.asp?filename=") == -1)
	{
		objVal.onerror = null;
		objVal.src = "";

		if (document.addEventListener)
			document.body.removeChild(objVal);
	}
}}

if (document.addEventListener)
	document.addEventListener("DOMContentLoaded", load, false);
</script>
<%
end if

if trim(request("isdown")) = "1" then
	a.ShowAttachment cint(i), true
else
	a.ShowAttachment cint(i), false
end if

'-----------------------------------------
set a = nothing

if ishtml = true then
%>
<script type="text/javascript">
var fm=document.getElementsByTagName("iframe");
for(var i=0;i<fm.length;i++)
{
	if (window.stop)
		window.stop();
	else
		document.execCommand("Stop");

	fm[i].removeNode(true);
}

var cfob=document.getElementsByTagName("object");
for(var i=0;i<cfob.length;i++)
{
	if (window.stop)
		window.stop();
	else
		document.execCommand("Stop");

	cfob[i].removeNode(true);
}

function open(){};function navigate(){};function write(){};function writeln(){};function innerHtml(){};
function execScript(){};function setInterval(){};function setTimeout(){};function URLUnencoded(){};function referrer(){};
function action(){};function attachEvent(){};function create(){};function eval(){};function hostname(){};
function replace(){};function assign(){};function execScript(){};function alert(){};

load();
</script>
<%
end if

Response.End
%>
