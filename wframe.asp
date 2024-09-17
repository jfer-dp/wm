<!--#include file="passinc.asp" -->

<%
isMSIE = false
if InStr(Request.ServerVariables("HTTP_User-Agent"), "MSIE") > 0 then
	isMSIE = true
end if

dim am
set am = server.createobject("easymail.Attachments")
am.Load Session("wem"), Session("tid")
am.RemoveAll
set am = nothing


dim upurl
dim downurl

downurl = "addatt.asp?" & getGRSN()

if trim(request("mode")) = "reply" then
	upurl = "replymail.asp?" & Request.QueryString
elseif trim(request("mode")) = "replyall" then
	upurl = "replymail.asp?" & Request.QueryString
elseif trim(request("mode")) = "post" then
	upurl = "post.asp?" & Request.QueryString
elseif trim(request("mode")) = "editpost" then
	upurl = "editpfmail.asp?" & Request.QueryString
elseif trim(request("mode")) = "forward" then
	upurl = "forwardmail.asp?" & Request.QueryString
elseif trim(request("mode")) = "domainlistmail" then
	dim dm
	set dm = server.createobject("easymail.Domain")
	dm.Load

	if dm.GetUserManagerDomainCount(Session("wem")) < 1 then
		if isadmin() = false then
			set dm = nothing
			response.redirect "noadmin.asp"
		end if
	end if

	set dm = nothing
	upurl = "domainlistmail.asp?" & Request.QueryString
else
	upurl = "writemail.asp?" & Request.QueryString
end if
%>

<HTML>
<frameset rows="100%, *" border="0" framespacing="0" frameborder="0" name="wf">
  <frame id=f3 name="f3" scrolling="AUTO" noresize src="<%=upurl %>" marginHeight=20>
  <frame id=f4 name="f4" src="<%=downurl %>" marginHeight=10>
</frameset>
</HTML>
