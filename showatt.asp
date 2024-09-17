<%
mode = trim(request("mode"))
sname = trim(request("sname"))
sfname = trim(request("sfname"))

if trim(request("isattfolder")) <> "att" or sfname = "att" then
	isatt = false
else
	isatt = true
end if


dim a
set a = server.createobject("easymail.emmail")
'-----------------------------------------

if sname <> "" and sfname <> "" then
	if Session("wem") = "" then
		if Application("em").IsLogin(trim(request.Cookies("name")), Request.ServerVariables("REMOTE_ADDR")) = true then
			openresult = a.OpenFriendFolder(trim(request.Cookies("name")), sname, sfname, isatt)
		else
			openresult = -1
		end if
	else
		openresult = a.OpenFriendFolder(Session("wem"), sname, sfname, isatt)
	end if

	if openresult = -1 then
		set a = nothing
		Response.Redirect "err.asp?errstr=失败"
	elseif  openresult = 1 then
		set a = nothing
		Response.Redirect "err.asp?errstr=密码错误"
	elseif  openresult = 2 then
		set a = nothing
		Response.Redirect "err.asp?errstr=文件夹不存在或不允许访问"
	end if
end if



if mode = "post" then
	a.IsInPublicFolder = true
end if


i = trim(request("count"))

filename = trim(request("filename"))

pt = trim(request("pt"))

if Session("wem") = "" then
	if Application("em").IsLogin(trim(request.Cookies("name")), Request.ServerVariables("REMOTE_ADDR")) = true then
		if pt <> "" then
			a.LoadAll1 trim(request.Cookies("name")), filename, CDbl(pt)
		else
			a.LoadAll trim(request.Cookies("name")), filename
		end if
	else
		set a = nothing
		Response.Redirect "err.asp"
	end if
else
	if pt <> "" then
		a.LoadAll1 Session("wem"), filename, CDbl(pt)
	else
		a.LoadAll Session("wem"), filename
	end if
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
			objVal.parentElement.removeChild(objVal);
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
