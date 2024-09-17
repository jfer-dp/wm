<%
Response.Charset="GB2312"
%>

<!--#include file="passinc.asp" --> 

<%
dim eads
set eads = server.createobject("easymail.EntAddress")
eads.Load

allnum = eads.Count

i = 0
fi = 0
tmplayer = -1

do while i < allnum
	eads.Get i, isFolder, fid, layer, name, email, comment

	if isFolder = true then
		if layer <= tmplayer then
			li = layer
			do while li <= tmplayer
				Response.Write "</div>" & Chr(13)
				li = li + 1
			loop
		end if

		Response.Write "<table><tr><td height=""2""></td></tr></table>" & Chr(13)
		Response.Write "<img src=""images/b.gif"" width=""" & 12 * layer & """ height=1 border=0>" & Chr(13)

		Response.Write "<a class=""clsNode"" href=""javascript:hide_folder(" & fi & ")""><span id=""folder_" & fi & """></span></a>&nbsp;<a class=""clsNode"" href=""javascript:choose_folder(" & fi & "," & i & ")"">" & server.htmlencode(name) & "</a>"
		if Len(comment) > 0 then
			Response.Write "&nbsp;&nbsp;<font style='color:#444444'>[" & server.htmlencode(comment) & "]</font>"
		end if
		Response.Write "<div id=""folder_" & fi & "_div""><br>" & Chr(13)

		tmplayer = layer
		fi = fi + 1
	else
		if Len(name) < 1 then
			name = email
		end if

		Response.Write "<img src=""images/b.gif"" width=""" & 30 + tmplayer * 12 & """ height=1 border=0><input id=""cbx_" & i & """ name=""cbx_" & i & """ value=""" & server.htmlencode(name & " <" & email & ">") & """ type=""checkbox"" onclick=""javascript:checkit(" & i & ", false)"">&nbsp;<a class=""clsNodeAds"" title=""" & server.htmlencode(email) & """ href=""javascript:checkit(" & i & ", true)"">" & server.htmlencode(name) & "</a>"
		if Len(comment) > 0 then
			Response.Write "&nbsp;&nbsp;<font style='color:#444444'>[" & server.htmlencode(comment) & "]</font>"
		end if
		Response.Write "<br>" & Chr(13)
	end if

	isFolder = NULL
	fid = NULL
	layer = NULL
	name = NULL
	email = NULL
	comment = NULL

	i = i + 1
loop

li = 0
do while li <= tmplayer
	Response.Write "</div>" & Chr(13)
	li = li + 1
loop

if allnum < 1 then
%>
<table width="96%" border="0" align="center" cellspacing="0" bgcolor="white">
<tr><td class="block_top_td" style="height:12px;"></td></tr>
<tr><td align="center">
<div class="wwm_line_msg">未发现有效数据</div>
</td></tr></table>
<%
end if

Response.Write "<input id=""ef_count"" type=""hidden"" value=""" & eads.FolderCount & """>"
Response.Write "<input id=""eads_count"" type=""hidden"" value=""" & allnum & """>"

Response.Write "<select id=""eads_all"" style=""display:none;"">"
i = 0
do while i < allnum
	eads.Get i, isFolder, fid, layer, name, email, comment

	if isFolder = false then
		if Len(name) < 1 then
			name = email
		end if

		Response.Write "<option value=""" & email & """>" & name & "</option>" & Chr(9)
	end if

	isFolder = NULL
	fid = NULL
	layer = NULL
	name = NULL
	email = NULL
	comment = NULL

	i = i + 1
loop
Response.Write "</select>"

set eads = nothing
%>
