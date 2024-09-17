<%
Response.Charset="GB2312"
%>

<!--#include file="passinc.asp" --> 

<%
dim ads
set ads = server.createobject("easymail.Addresses")
ads.Load Session("wem")

allnum = ads.EmailCount
i = 0
do while i < allnum
	ads.MoveTo i
	Response.Write ads.email & Chr(12) & ads.nickname & Chr(9)
	i = i + 1
loop

allnum = ads.GroupCount
i = 0
do while i < allnum
	ads.GetGroupInfo i, nickname, emails
	Response.Write Chr(12) & nickname & Chr(9)

	nickname = NULL
	emails = NULL

	i = i + 1
loop

set ads = nothing
%>
