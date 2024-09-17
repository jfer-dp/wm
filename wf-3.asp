<!doctype html>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0,user-scalable=no">
<link rel="stylesheet" type="text/css" href="images/hrefbt.css">
<style>
body {
overflow-x:hidden; 
overflow-y:hidden; 
}

html, body {
    height: 30px;
    width:100%;
    margin: 0px;
    padding: 0px;
}
</style>

<script>
var showmail = true;

function kd() {
	if (showmail == true)
	{
		document.getElementById("div_show").style.display = "none";
		document.getElementById("div_reply").style.display = "inline";
		parent.document.getElementById("pageframe").rows="0,*,60";
		showmail = false;
	}
	else
	{
		document.getElementById("div_show").style.display = "inline";
		document.getElementById("div_reply").style.display = "none";
		parent.document.getElementById("pageframe").rows="*,0,60";
		showmail = true;
	}
}

function send() {
	document.getElementById("bt_send").href = "#";
	parent.f2.send();
}

function sendend() {
	document.getElementById("bt_send").href = "javascript:send();";
}

function show_back() {
	document.getElementById("div_show").style.display = "none";
	document.getElementById("div_reply").style.display = "none";
	document.getElementById("div_back").style.display = "inline";
}

function f2_back() {
	parent.f2.back2write();
	document.getElementById("div_show").style.display = "none";
	document.getElementById("div_reply").style.display = "inline";
	document.getElementById("div_back").style.display = "none";
}
</script>
</head>

<body>
<div style="text-align:center; margin-top:14px;">
<hr style="height:1px; border:none; border-top:1px dashed #0066CC;">
<div id="div_show">
<a id="bt_show" class='wwm_btnDownload btn_gray' href="javascript:kd();" style="font-size:14px; font-weight:bold;">&nbsp;回 复&nbsp;</a><input type="submit" value="" onclick="javascript:kd();" style="filter:alpha(opacity=0); opacity:0; font-size:0pt; height:0px; width:0px; border:0px;">
</div>
<div id="div_reply" style="display:none;">
<a id="bt_show" class='wwm_btnDownload btn_gray' href="javascript:kd();" style="font-size:14px; font-weight:bold;">&nbsp;原 信&nbsp;</a><input type="submit" value="" onclick="javascript:kd();" style="filter:alpha(opacity=0); opacity:0; font-size:0pt; height:0px; width:0px; border:0px;">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<a id="bt_send" class='wwm_btnDownload btn_gray' href="javascript:send();" style="font-size:14px; font-weight:bold;">&nbsp;发 送&nbsp;</a><input type="submit" value="" onclick="javascript:send();" style="filter:alpha(opacity=0); opacity:0; font-size:0pt; height:0px; width:0px; border:0px;">
</div>
<div id="div_back" style="display:none;">
<a class='wwm_btnDownload btn_gray' href="javascript:f2_back();" style="font-size:14px; font-weight:bold;">&nbsp;返 回&nbsp;</a><input type="submit" value="" onclick="javascript:f2_back();" style="filter:alpha(opacity=0); opacity:0; font-size:0pt; height:0px; width:0px; border:0px;">
</div>
</div>
</body>
</html>
