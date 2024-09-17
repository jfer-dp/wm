<%
Response.CacheControl = "no-cache"
wid = replace(trim(request("wid")), " ", "+")
%>

<!doctype html>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312-80">
<meta name="viewport" content="width=device-width,initial-scale=1.0,user-scalable=no">
</head>

<frameset name="pageframe" id="pageframe" rows="*,0,60" frameborder="NO" border="0" framespacing="0"> 
  <frame id="f1" name="f1" noresize src="wf-1.asp?wid=<%=Server.URLEncode(wid) %>&<%=getGRSN() %>">
  <frame id="f2" name="f2" noresize src="wf-2.asp?wid=<%=Server.URLEncode(wid) %>&<%=getGRSN() %>">
  <frame id="f3" name="f3" noresize src="wf-3.asp?<%=getGRSN() %>">
</frameset>
</html>

<%
function getGRSN()
	dim theGRSN
	Randomize
	theGRSN = Int((9999999 * Rnd) + 1)

	getGRSN = "GRSN=" & CStr(theGRSN)
end function
%>
