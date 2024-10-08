var iepst = navigator.userAgent.toLowerCase().indexOf("msie");
if (iepst > 0)
{
	if (parseFloat(navigator.userAgent.substr(iepst + 5, 3)) < 5.5)
		document.writeln("<script language=\"JavaScript\" type=\"text/javascript\" src=\"rte/wrte4.js\"></script>");
	else
		document.writeln("<script language=\"JavaScript\" type=\"text/javascript\" src=\"rte/wrte3.js\"></script>");
}
else
	document.writeln("<script language=\"JavaScript\" type=\"text/javascript\" src=\"rte/wrte5.js\"></script>");


var need_nl_before = '|div|p|table|tbody|tr|td|th|title|head|body|script|comment|li|meta|h1|h2|h3|h4|h5|h6|hr|ul|ol|option|';
var need_nl_after = '|html|head|body|p|th|style|';

var re_comment = new RegExp();
re_comment.compile("^<!--(.*)-->$");

var re_hyphen = new RegExp();
re_hyphen.compile("-$");

function get_xhtml(node, lang, encoding, need_nl, inside_pre) {
	var i;
	var text = '';
	var children = node.childNodes;
	var child_length = children.length;
	var tag_name;
	var do_nl = need_nl ? true : false;
	var page_mode = true;

	for (i = 0; i < child_length; i++) {
		var child = children[i];

		switch (child.nodeType) {
			case 1: {
				var tag_name = String(child.tagName).toLowerCase();

				if (tag_name == '') break;

				if (tag_name == 'meta') {
					var meta_name = String(child.name).toLowerCase();
					if (meta_name == 'generator') break;
				}

				if (!need_nl && tag_name == 'body') {
					page_mode = false;
				}

				if (tag_name == '!') {
					var parts = re_comment.exec(child.text);

					if (parts) {
						var inner_text = parts[1];
						text += fix_comment(inner_text);
					}
				} else {
					if (tag_name == 'html') {
						text = '<?xml version="1.0" encoding="'+encoding+'"?>\n<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">\n';
					}

					if (need_nl_before.indexOf('|'+tag_name+'|') != -1) {
						if ((do_nl || text != '') && !inside_pre) text += '\n';
					} else {
						do_nl = true;
					}

					text += '<'+tag_name;

					var attr = child.attributes;
					var attr_length = attr.length;
					var attr_value;

					var attr_lang = false;
					var attr_xml_lang = false;
					var attr_xmlns = false;

					var is_alt_attr = false;

					for (j = 0; j < attr_length; j++) {
						var attr_name = attr[j].nodeName.toLowerCase();

						if (!attr[j].specified && 
							(attr_name != 'selected' || !child.selected) && 
							(attr_name != 'style' || child.style.cssText == '') && 
							attr_name != 'value') continue;

						if (attr_name == '_moz_dirty' || 
							attr_name == '_moz_resizing' || 
							tag_name == 'br' && 
							attr_name == 'type' && 
							child.getAttribute('type') == '_moz') continue;

						var valid_attr = true;

						switch (attr_name) {
							case "style":
								attr_value = child.style.cssText;
								break;
							case "class":
								attr_value = child.className;
								break;
							case "http-equiv":
								attr_value = child.httpEquiv;
								break;
							case "noshade": break;
							case "checked": break;
							case "selected": break;
							case "multiple": break;
							case "nowrap": break;
							case "disabled": break;
								attr_value = attr_name;
								break;
							default:
								try {
									attr_value = child.getAttribute(attr_name, 2);
								} catch (e) {
									valid_attr = false;
								}
								break;
						}

						if (attr_name == 'lang') {
							attr_lang = true;
							attr_value = lang;
						}
						if (attr_name == 'xml:lang') {
							attr_xml_lang = true;
							attr_value = lang;
						}
						if (attr_name == 'xmlns') attr_xmlns = true;
						if (valid_attr) {
							if (!(tag_name == 'li' && attr_name == 'value')) {
								text += ' '+attr_name+'="'+fix_attribute(attr_value)+'"';
							}
						}

						if (attr_name == 'alt') is_alt_attr = true;
					}

					if (tag_name == 'img' && !is_alt_attr) {
						text += ' alt=""';
					}

					if (tag_name == 'html') {
						if (!attr_lang) text += ' lang="'+lang+'"';
						if (!attr_xml_lang) text += ' xml:lang="'+lang+'"';
						if (!attr_xmlns) text += ' xmlns="http://www.w3.org/1999/xhtml"';
					}

					if (child.canHaveChildren || child.hasChildNodes()){
						text += '>';
						text += get_xhtml(child, lang, encoding, true, inside_pre || tag_name == 'pre' ? true : false);
						text += '</'+tag_name+'>';
					} else {
						if (tag_name == 'style' || tag_name == 'title' || tag_name == 'script') {
							text += '>';
							var inner_text;
							if (tag_name == 'script') {
								inner_text = child.text;
							} else {
								inner_text = child.innerHTML;
							}

							if (tag_name == 'style') {
								inner_text = String(inner_text).replace(/[\n]+/g,'\n');
							}

							text += inner_text+'</'+tag_name+'>';
						} else {
							text += ' />';
						}
					}
				}
				break;
			}
			case 3: {
				if (!inside_pre) {
					if (child.nodeValue != '\n') {
						text += fix_text(child.nodeValue);
					}
				} else {
					text += child.nodeValue;
				}
				break;
			}
			case 8: {
				text += fix_comment(child.nodeValue);
				break;
			}
			default:
				break;
		}
	}

	if (!need_nl && !page_mode) {
		text = text.replace(/<\/?head>[\n]*/gi, "");
		text = text.replace(/<head \/>[\n]*/gi, "");
		text = text.replace(/<\/?body>[\n]*/gi, "");
	}

	return text;
}

function fix_comment(text) {
	text = text.replace(/--/g, "__");

	if(re_hyphen.exec(text)) {
		text += " ";
	}

	return "<!--"+text+"-->";
}

function fix_text(text) {
	return String(text).replace(/\n{2,}/g, "\n").replace(/\&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\u00A0/g, "&nbsp;");
}

function fix_attribute(text) {
	return String(text).replace(/\&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\"/g, "&quot;");
}
