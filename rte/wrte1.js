var isRichText = false;
var rng;
var currentRTE;
var allRTEs = "";

var isIE;
var isGecko;
var isSafari;
var isKonqueror;

var imagesPath;
var includesPath;
var cssFile;
var generateXHTML;

var lang = "zh-cn";
var encoding = "GB2312";


var ieversion = 0;
var iepst2 = navigator.userAgent.toLowerCase().indexOf("msie");
if (iepst2 > 0)
	ieversion = parseFloat(navigator.userAgent.substr(iepst2 + 5, 3));


function initRTE(imgPath, incPath, css, genXHTML) {
	var ua = navigator.userAgent.toLowerCase();
	isIE = ((ua.indexOf("msie") != -1) && (ua.indexOf("opera") == -1) && (ua.indexOf("webtv") == -1)); 
	isGecko = (ua.indexOf("gecko") != -1);
	isSafari = (ua.indexOf("safari") != -1);
	isKonqueror = (ua.indexOf("konqueror") != -1);
	isChrome = (ua.indexOf("chrome") != -1);

	generateXHTML = genXHTML;

	if ((document.getElementById && document.designMode) && ((!isSafari && !isKonqueror) || isChrome == true)) {
		isRichText = true;
	}

	if (isIE) {
		document.onmouseover = raiseButton;
		document.onmouseout  = normalButton;
		document.onmousedown = lowerButton;
		document.onmouseup   = raiseButton;
	}

	imagesPath = imgPath;
	includesPath = incPath;
	cssFile = css;

	if (isRichText) document.writeln('<style type="text/css">@import "' + includesPath + 'rte.css";</style>');
}

function writeRichText(rte, html, width, height, buttons, readOnly) {
	if (isRichText) {
		if (allRTEs.length > 0) allRTEs += ";";
		allRTEs += rte;

		if (readOnly) buttons = false;

		//		if (isIE) {
		//			if (buttons && (width < 540)) width = 540;
		//			var tablewidth = width;
		//		} else {
		//			if (buttons && (width < 540)) width = 540;
		//			var tablewidth = width + 4;
		//		}
		var tablewidth = "100%";
		width = "100%";

		document.writeln('<div class="rteDiv">');
		if (buttons == true) {
			document.writeln('<table class="rteBack" cellpadding=2 cellspacing=0 id="Buttons1_' + rte + '" width="' + tablewidth + '">');
			document.writeln('	<tr>');
			document.writeln('		<td>');
			document.writeln('			<select id="formatblock_' + rte + '" onchange="selectFont(\'' + rte + '\', this.id);">');
			document.writeln('				<option value="">[格式]</option>');
			document.writeln('				<option value="<p>">普通 &lt;p&gt;</option>');
			document.writeln('				<option value="<h1>">标题 1 &lt;h1&gt;</option>');
			document.writeln('				<option value="<h2>">标题 2 &lt;h2&gt;</option>');
			document.writeln('				<option value="<h3>">标题 3 &lt;h3&gt;</option>');
			document.writeln('				<option value="<h4>">标题 4 &lt;h4&gt;</option>');
			document.writeln('				<option value="<h5>">标题 5 &lt;h5&gt;</option>');
			document.writeln('				<option value="<h6>">标题 6 &lt;h6&gt;</option>');
			document.writeln('				<option value="<address>">地址 &lt;ADDR&gt;</option>');
			document.writeln('				<option value="<pre>">格式化文本 &lt;pre&gt;</option>');
			document.writeln('			</select>');
			document.writeln('		</td>');
			document.writeln('		<td>');
			document.writeln('			<select id="fontname_' + rte + '" onchange="selectFont(\'' + rte + '\', this.id)">');
			document.writeln('				<option value="Font" selected>[字体]</option>');
			document.writeln('				<option value="宋体, Helvetica, sans-serif">宋体</option>');
			document.writeln('				<option value="仿宋, Helvetica, sans-serif">仿宋</option>');
			document.writeln('				<option value="黑体, Helvetica, sans-serif">黑体</option>');
			document.writeln('				<option value="隶书, Helvetica, sans-serif">隶书</option>');
			document.writeln('				<option value="幼圆, Helvetica, sans-serif">幼圆</option>');
			document.writeln('				<option value="Arial, Helvetica, sans-serif">Arial</option>');
			document.writeln('				<option value="Arial Narrow, Helvetica, sans-serif">Arial Narrow</option>');
			document.writeln('				<option value="Arial Black, Helvetica, sans-serif">Arial Black</option>');
			document.writeln('				<option value="Comic Sans MS, Courier, mono">Comic Sans MS</option>');
			document.writeln('				<option value="Courier, Courier, mono">Courier</option>');
			document.writeln('				<option value="Courier New, Courier, mono">Courier New</option>');
			document.writeln('				<option value="System, Times, serif">System</option>');
			document.writeln('				<option value="Times New Roman, Times, serif">Times New Roman</option>');
			document.writeln('				<option value="Verdana, Arial, Helvetica, sans-serif">Verdana</option>');
			document.writeln('				<option value="Wingdings, Arial, Helvetica, sans-serif">Wingdings</option>');
			document.writeln('			</select>');
			document.writeln('		</td>');
			document.writeln('		<td>');
			document.writeln('			<select unselectable="on" id="fontsize_' + rte + '" onchange="selectFont(\'' + rte + '\', this.id);">');
			document.writeln('				<option value="Size">[字号]</option>');
			document.writeln('				<option value="1">1</option>');
			document.writeln('				<option value="2">2</option>');
			document.writeln('				<option value="3">3</option>');
			document.writeln('				<option value="4">4</option>');
			document.writeln('				<option value="5">5</option>');
			document.writeln('				<option value="6">6</option>');
			document.writeln('				<option value="7">7</option>');
			document.writeln('			</select>');
			document.writeln('		</td>');
			document.writeln('		<td width="100%">');
			document.writeln('		</td>');
			document.writeln('	</tr>');
			document.writeln('</table>');
			document.writeln('<table class="rteBack" cellpadding="0" cellspacing="0" id="Buttons2_' + rte + '" width="' + tablewidth + '">');
			document.writeln('	<tr>');
			document.writeln('		<td><img id="bold" class="rteImage" src="' + imagesPath + 'bold.gif" width="25" height="24" alt="粗体" title="粗体" onClick="rteCommand(\'' + rte + '\', \'bold\', \'\')"></td>');
			document.writeln('		<td><img class="rteImage" src="' + imagesPath + 'italic.gif" width="25" height="24" alt="斜体" title="斜体" onClick="rteCommand(\'' + rte + '\', \'italic\', \'\')"></td>');
			document.writeln('		<td><img class="rteImage" src="' + imagesPath + 'underline.gif" width="25" height="24" alt="下划线" title="下划线" onClick="rteCommand(\'' + rte + '\', \'underline\', \'\')"></td>');
			document.writeln('		<td><img class="rteImage" src="' + imagesPath + 'left_just.gif" width="25" height="24" alt="左对齐" title="左对齐" onClick="rteCommand(\'' + rte + '\', \'justifyleft\', \'\')"></td>');
			document.writeln('		<td><img class="rteImage" src="' + imagesPath + 'centre.gif" width="25" height="24" alt="居中" title="居中" onClick="rteCommand(\'' + rte + '\', \'justifycenter\', \'\')"></td>');
			document.writeln('		<td><img class="rteImage" src="' + imagesPath + 'right_just.gif" width="25" height="24" alt="右对齐" title="右对齐" onClick="rteCommand(\'' + rte + '\', \'justifyright\', \'\')"></td>');
			document.writeln('		<td><img class="rteImage" src="' + imagesPath + 'hr.gif" width="25" height="24" alt="水平线" title="水平线" onClick="rteCommand(\'' + rte + '\', \'inserthorizontalrule\', \'\')"></td>');
			document.writeln('		<td><img class="rteImage" src="' + imagesPath + 'numbered_list.gif" width="25" height="24" alt="编号" title="编号" onClick="rteCommand(\'' + rte + '\', \'insertorderedlist\', \'\')"></td>');
			document.writeln('		<td><img class="rteImage" src="' + imagesPath + 'list.gif" width="25" height="24" alt="项目符号" title="项目符号" onClick="rteCommand(\'' + rte + '\', \'insertunorderedlist\', \'\')"></td>');
			document.writeln('		<td><img class="rteImage" src="' + imagesPath + 'outdent.gif" width="25" height="24" alt="减小缩进" title="减小缩进" onClick="rteCommand(\'' + rte + '\', \'outdent\', \'\')"></td>');
			document.writeln('		<td><img class="rteImage" src="' + imagesPath + 'indent.gif" width="25" height="24" alt="增大缩进" title="增大缩进" onClick="rteCommand(\'' + rte + '\', \'indent\', \'\')"></td>');
			document.writeln('		<td><div id="forecolor_' + rte + '"><img class="rteImage" src="' + imagesPath + 'textcolor.gif" width="25" height="24" alt="字体颜色" title="字体颜色" onClick="dlgColorPalette(\'' + rte + '\', \'forecolor\', \'\')"></div></td>');
			document.writeln('		<td><div id="hilitecolor_' + rte + '"><img class="rteImage" src="' + imagesPath + 'bgcolor.gif" width="25" height="24" alt="背景颜色" title="背景颜色" onClick="dlgColorPalette(\'' + rte + '\', \'hilitecolor\', \'\')"></div></td>');
			document.writeln('		<td><img class="rteImage" src="' + imagesPath + 'hyperlink.gif" width="25" height="24" alt="插入链接" title="插入链接" onClick="dlgInsertLink(\'' + rte + '\', \'link\')"></td>');
			document.writeln('		<td><img class="rteImage" src="' + imagesPath + 'image.gif" width="25" height="24" alt="插入图片" title="插入图片" onClick="addImage(\'' + rte + '\')"></td>');
			document.writeln('		<td><div id="table_' + rte + '"><img class="rteImage" src="' + imagesPath + 'insert_table.gif" width="25" height="24" alt="插入表格" title="插入表格" onClick="dlgInsertTable(\'' + rte + '\', \'table\', \'\')"></div></td>');
			document.writeln('		<td><img class="rteImage" src="' + imagesPath + 'undo.gif" width="25" height="24" alt="撤消" title="撤消" onClick="rteCommand(\'' + rte + '\', \'undo\')"></td>');
			document.writeln('		<td><img class="rteImage" src="' + imagesPath + 'redo.gif" width="25" height="24" alt="恢复" title="恢复" onClick="rteCommand(\'' + rte + '\', \'redo\')"></td>');
			document.writeln('		<td width="100%"></td>');
			document.writeln('	</tr>');
			document.writeln('</table>');
		}
		if (isIE)
			document.writeln('<iframe id="' + rte + '" name="' + rte + '" width="' + width + 'px" height="' + height + 'px" src="' + includesPath + 'blank.htm"></iframe>');
		else
			document.writeln('<iframe style="color:#777777; BORDER-BOTTOM:1px solid; BORDER-LEFT:1px solid; BORDER-RIGHT:1px solid; BORDER-TOP:0px solid;" id="' + rte + '" name="' + rte + '" width="' + width + 'px" height="' + height + 'px" src="' + includesPath + 'blank.htm"></iframe>');

		document.writeln('<iframe width="154" height="104" id="cp' + rte + '" src="' + includesPath + 'palette.htm" marginwidth="0" marginheight="0" scrolling="no" style="visibility:hidden; position: absolute; top:10; left:10;"></iframe>');
		document.writeln('<input type="hidden" id="hdn' + rte + '" name="' + rte + '" value="">');
		document.writeln('</div>');
		
		document.getElementById('hdn' + rte).value = html;
		enableDesignMode(rte, html, readOnly);
	} else {
		if (!readOnly) {
			document.writeln('<textarea name="' + rte + '" id="' + rte + '" style="width: ' + width + 'px; height: ' + height + 'px;">' + html + '</textarea>');
		} else {
			document.writeln('<textarea name="' + rte + '" id="' + rte + '" style="width: ' + width + 'px; height: ' + height + 'px;" readonly>' + html + '</textarea>');
		}
	}
}

function enableDesignMode(rte, html, readOnly) {
	var frameHtml = "<html id=\"" + rte + "\">\n";
	frameHtml += "<head>\n";
	if (cssFile.length > 0) {
		frameHtml += "<link media=\"all\" type=\"text/css\" href=\"" + cssFile + "\" rel=\"stylesheet\">\n";
	} else {
		frameHtml += "<style>\n";
		frameHtml += "body {\n";
		frameHtml += "	background: #FFFFFF;\n";
		frameHtml += "	margin: 0px;\n";
		frameHtml += "	padding: 0px;\n";
		frameHtml += "}\n";
		frameHtml += "</style>\n";
	}
	frameHtml += "</head>\n";

	if (iepst2 > 0)
	{
		frameHtml += "<body><div>\n";
		frameHtml += html + "\n</div>";
	}
	else
	{
		frameHtml += "<body>\n";
		frameHtml += html + "\n";
	}

	frameHtml += "</body>\n";
	frameHtml += "</html>";
	
	if (document.all) {
		var oRTE = frames[rte].document;
		oRTE.open();
		oRTE.write(frameHtml);
		oRTE.close();
		if (!readOnly) {
			oRTE.designMode = "On";
			frames[rte].document.attachEvent("onkeypress", function evt_ie_keypress(event) {ieKeyPress(event, rte);});
		}
	} else {
		try {
			if (!readOnly) document.getElementById(rte).contentDocument.designMode = "on";
			try {
				var oRTE = document.getElementById(rte).contentWindow.document;
				oRTE.open();
				oRTE.write(frameHtml);
				oRTE.close();
				if (isGecko && !readOnly) {
					oRTE.addEventListener("keypress", geckoKeyPress, true);
				}
			} catch (e) {
				alert("Error preloading content.");
			}
		} catch (e) {
			if (isGecko) {
				setTimeout("enableDesignMode('" + rte + "', '" + html + "', " + readOnly + ");", 10);
			} else {
				return false;
			}
		}
	}
}

function updateRTE(rte) {
	if (!isRichText) return;

	var readOnly = false;
	if (document.all) {
		if (frames[rte].document.designMode != "On") readOnly = true;
	} else {
		if (document.getElementById(rte).contentDocument.designMode != "on") readOnly = true;
	}

	if (isRichText && !readOnly)
		setHiddenVal(rte);
}

function setHiddenVal(rte) {
	var oHdnField = document.getElementById('hdn' + rte);

	if (oHdnField.value == null) oHdnField.value = "";
	if (document.all) {
		if (generateXHTML) {
			oHdnField.value = get_xhtml(frames[rte].document.body, lang, encoding);
		} else {
			oHdnField.value = frames[rte].document.body.innerHTML;
		}
	} else {
		if (generateXHTML) {
			oHdnField.value = get_xhtml(document.getElementById(rte).contentWindow.document.body, lang, encoding);
		} else {
			oHdnField.value = document.getElementById(rte).contentWindow.document.body.innerHTML;
		}
	}

	if (stripHTML(oHdnField.value.replace("&nbsp;", " ")) == "" &&
		oHdnField.value.toLowerCase().search("<hr") == -1 &&
		oHdnField.value.toLowerCase().search("<img") == -1) oHdnField.value = "";
}

function updateRTEs() {
	var vRTEs = allRTEs.split(";");
	for (var i = 0; i < vRTEs.length; i++) {
		updateRTE(vRTEs[i]);
	}
}

function rteCommand(rte, command, option) {
	var oRTE;
	if (document.all) {
		oRTE = frames[rte];
	} else {
		oRTE = document.getElementById(rte).contentWindow;
	}
	
	try {
		oRTE.focus();
	  	oRTE.document.execCommand(command, false, option);
		oRTE.focus();
	} catch (e) {
	}
}

function dlgColorPalette(rte, command) {
	setRange(rte);

	var oDialog = document.getElementById('cp' + rte);
	var buttonElement = document.getElementById(command + '_' + rte);
	var iLeftPos = getOffsetLeft(buttonElement);
	var iTopPos = getOffsetTop(buttonElement) + (buttonElement.offsetHeight + 4);
	if (isIE && ieversion < 5.5)
		iTopPos += 20;
	oDialog.style.left = (iLeftPos) + "px";
	oDialog.style.top = (iTopPos) + "px";

	if ((command == parent.command) && (rte == currentRTE)) {
		if (oDialog.style.visibility == "hidden") {
			showHideElement(oDialog, 'show');
		} else {
			showHideElement(oDialog, 'hide');
		}
	} else {
		var vRTEs = allRTEs.split(";");
		for (var i = 0; i < vRTEs.length; i++) {
			showHideElement('cp' + vRTEs[i], 'hide');
		}
		showHideElement(oDialog, 'show');
	}

	parent.command = command;
	currentRTE = rte;
}

function dlgInsertTable(rte, command) {
	parent.command = command;
	currentRTE = rte;
	InsertTable = popUpWin(includesPath + 'insert_table.htm', 'InsertTable', 360, 180, 'resizable=yes,status=no,');
}

function dlgInsertLink(rte, command) {
	parent.command = command;
	currentRTE = rte;
	InsertLink = popUpWin(includesPath + 'insert_link.htm', 'InsertLink', 360, 180, 'resizable=yes,status=no,');

	setRange(rte);
	var linkText = '';
	if (isIE) {
		linkText = stripHTML(rng.htmlText);
	} else {
		linkText = stripHTML(rng.toString());
	}
	setLinkText(linkText);
}

function setLinkText(linkText) {
	try {
		window.InsertLink.document.linkForm.linkText.value = linkText;
	} catch (e) {
		setTimeout("setLinkText('" + linkText + "');", 10);
	}
}

function popUpWin (url, win, width, height, options) {
	var leftPos = (screen.availWidth - width) / 2;
	var topPos = (screen.availHeight - height) / 2;
	options += 'width=' + width + ',height=' + height + ',left=' + leftPos + ',top=' + topPos;
	return window.open(url, win, options);
}

function setColor(color) {
	var rte = currentRTE;
	var parentCommand = parent.command;
	
	if (document.all) {
		if (parentCommand == "hilitecolor") parentCommand = "backcolor";

		rng.select();
	}
	
	rteCommand(rte, parentCommand, color);
	showHideElement('cp' + rte, "hide");
}

function addImage(rte) {
	imagePath = prompt('请输入图片地址:', 'http://');				
	if ((imagePath != null) && (imagePath != "")) {
		rteCommand(rte, 'InsertImage', imagePath);
	}
}

function getOffsetTop(elm) {
	var mOffsetTop = elm.offsetTop;
	var mOffsetParent = elm.offsetParent;
	var parents_up = 2;
	
	while(parents_up > 0) {
		mOffsetTop += mOffsetParent.offsetTop;
		mOffsetParent = mOffsetParent.offsetParent;
		parents_up--;
	}
	
	return mOffsetTop;
}

function getOffsetLeft(elm) {
	var mOffsetLeft = elm.offsetLeft;
	var mOffsetParent = elm.offsetParent;
	var parents_up = 2;
	
	while(parents_up > 0) {
		mOffsetLeft += mOffsetParent.offsetLeft;
		mOffsetParent = mOffsetParent.offsetParent;
		parents_up--;
	}
	
	return mOffsetLeft;
}

function selectFont(rte, selectname) {
	var idx = document.getElementById(selectname).selectedIndex;
	if (idx != 0) {
		var selected = document.getElementById(selectname).options[idx].value;
		var cmd = selectname.replace('_' + rte, '');
		rteCommand(rte, cmd, selected);
		document.getElementById(selectname).selectedIndex = 0;
	}
}

function insertHTML(html) {
	var rte = currentRTE;

	var oRTE;
	if (document.all) {
		oRTE = frames[rte];
	} else {
		oRTE = document.getElementById(rte).contentWindow;
	}
	
	oRTE.focus();
	if (document.all) {
		var oRng = oRTE.document.selection.createRange();
		oRng.pasteHTML(html);
		oRng.collapse(false);
		oRng.select();
	} else {
		oRTE.document.execCommand('insertHTML', false, html);
	}
}

function showHideElement(element, showHide) {
	if (document.getElementById(element)) {
		element = document.getElementById(element);
	}

	if (showHide == "show") {
		element.style.visibility = "visible";
	} else if (showHide == "hide") {
		element.style.visibility = "hidden";
	}
}

function setRange(rte) {
	var oRTE;
	if (document.all) {
		oRTE = frames[rte];
		var selection = oRTE.document.selection; 
		if (selection != null) rng = selection.createRange();
	} else {
		oRTE = document.getElementById(rte).contentWindow;
		var selection = oRTE.getSelection();
		rng = selection.getRangeAt(selection.rangeCount - 1).cloneRange();
	}
	return rng;
}

function stripHTML(oldString) {
	var newString = oldString.replace(/(<([^>]+)>)/ig,"");

	newString = newString.replace(/\r\n/g," ");
	newString = newString.replace(/\n/g," ");
	newString = newString.replace(/\r/g," ");

	newString = trim(newString);

	return newString;
}

function trim(inputString) {
   if (typeof inputString != "string") return inputString;
   var retValue = inputString;
   var ch = retValue.substring(0, 1);
	
   while (ch == " ") {
      retValue = retValue.substring(1, retValue.length);
      ch = retValue.substring(0, 1);
   }
   ch = retValue.substring(retValue.length - 1, retValue.length);
	
   while (ch == " ") {
      retValue = retValue.substring(0, retValue.length - 1);
      ch = retValue.substring(retValue.length - 1, retValue.length);
   }

   while (retValue.indexOf("  ") != -1) {
      retValue = retValue.substring(0, retValue.indexOf("  ")) + retValue.substring(retValue.indexOf("  ") + 1, retValue.length);
   }
   return retValue;
}

function geckoKeyPress(evt) {
	var rte = evt.target.id;

	if (evt.ctrlKey) {
		var key = String.fromCharCode(evt.charCode).toLowerCase();
		var cmd = '';
		switch (key) {
			case 'b': cmd = "bold"; break;
			case 'i': cmd = "italic"; break;
			case 'u': cmd = "underline"; break;
		};

		if (cmd) {
			rteCommand(rte, cmd, null);

			evt.preventDefault();
			evt.stopPropagation();
		}
 	}
}

function ieKeyPress(evt, rte) {
	var key = (evt.which || evt.charCode || evt.keyCode);
	var stringKey = String.fromCharCode(key).toLowerCase();
}

function raiseButton(e) {
	var el = window.event.srcElement;
	
	className = el.className;
	if (className == 'rteImage' || className == 'rteImageLowered') {
		el.className = 'rteImageRaised';
	}
}

function normalButton(e) {
	var el = window.event.srcElement;
	
	className = el.className;
	if (className == 'rteImageRaised' || className == 'rteImageLowered') {
		el.className = 'rteImage';
	}
}

function lowerButton(e) {
	var el = window.event.srcElement;
	
	className = el.className;
	if (className == 'rteImage' || className == 'rteImageRaised') {
		el.className = 'rteImageLowered';
	}
}
