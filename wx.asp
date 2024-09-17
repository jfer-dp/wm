<%
if Len(trim(Request("echostr"))) > 0 then
	Response.Write trim(Request("echostr"))
else
	Set xml = Server.CreateObject ("msxml2.DOMDocument") 
	xml.async = False 
	xml.Load Request

	if xml.getElementsByTagName("Event").length > 0 then
		if xml.getElementsByTagName("Event").item(0).text = "subscribe" then
			set wx = server.createobject("easymail.WXSet")
			wx.load
			wx.Subscribe xml.getElementsByTagName("FromUserName").item(0).text
			set wx = nothing
		end if

		if xml.getElementsByTagName("Event").item(0).text = "unsubscribe" then
			set wx = server.createobject("easymail.WXSet")
			wx.load
			wx.UnSubscribe xml.getElementsByTagName("FromUserName").item(0).text
			set wx = nothing
		end if
	end if
	set xml = nothing
end if
%>
