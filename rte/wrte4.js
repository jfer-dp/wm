function getText(strHtml)
{
	var cv_sh = strHtml.replace(/[\r\n]*/g, "");
	cv_sh = cv_sh.replace(/<[\s\/]*br(\s[^>]*>|[^>]*>)/gi, "\r\n");

	cv_sh = cv_sh.replace(/^<[\s]*p[\s]*>/i, "");
	cv_sh = cv_sh.replace(/<[\s]*\/+[\s]*p(\s[^>]*>|[^>]*>)/gi, "\r\n");

	cv_sh = cv_sh.replace(/<[\s]*li(\s[^>]*>|[^>]*>)/gi, "\r\n");
	cv_sh = cv_sh.replace(/<[\s\/]*blockquote(\s[^>]*>|[^>]*>)/gi, "\r\n");
	cv_sh = cv_sh.replace(/<[\s]*hr(\s[^>]*>|[^>]*>)/gi, "\r\n");
	cv_sh = cv_sh.replace(/<[\s]*\/+[\s]*tr(\s[^>]*>|[^>]*>)/gi, "\r\n");
	cv_sh = cv_sh.replace(/<[\s]*\/+[\s]*div(\s[^>]*>|[^>]*>)/gi, "\r\n");

	cv_sh = cv_sh.replace(/<[\s]*\/+[\s]*td(\s[^>]*>|[^>]*>)/gi, " ");

	cv_sh = cv_sh.replace(/<[^>]*>/g, "");

	cv_sh = cv_sh.replace(/&amp;/g, "&");
	cv_sh = cv_sh.replace(/&#38;/g, "&");

	cv_sh = cv_sh.replace(/&lt;/g, "<");
	cv_sh = cv_sh.replace(/&#60;/g, "<");

	cv_sh = cv_sh.replace(/&gt;/g, ">");
	cv_sh = cv_sh.replace(/&#62;/g, ">");

	cv_sh = cv_sh.replace(/&quot;/g, "\"");
	cv_sh = cv_sh.replace(/&nbsp;/g, " ");

	cv_sh = cv_sh.replace(/^[\r\n]{0,2}|[\r\n]{0,2}$/g, "");

	return cv_sh;
}

function RemoveScript(strHtml)
{
	var cv_sh = strHtml.replace(/&lt;script[^>]*>[\s\S]*&lt;\/script>/gi, "");
	cv_sh = cv_sh.replace(/&lt;/g, "<");
	return cv_sh.replace(/&#11;/g, "&lt;");
}
