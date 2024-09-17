<!--#include file="passinc.asp" --> 

<%
if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	mode = trim(Request.Form("mode"))
	sname = trim(Request.Form("sname"))
	sfname = trim(Request.Form("sfname"))
	fid = trim(Request.Form("fid"))
	sortstr = trim(Request.Form("sortstr"))
	sortmode = trim(Request.Form("sortmode"))
	tgname = trim(Request.Form("tgname"))
	ret_value = ""

	if Len(fid) > 6 then
		dim ei
		set ei = server.createobject("easymail.InfoList")

		if sortmode = "1" then
			sortmode = true
		else
			sortmode = false
		end if

		if sortstr <> "" then
			ei.SetSort sortstr, sortmode
		end if

		dim open_isok
		open_isok = true

		if mode = "0" and sname <> "" and sfname <> "" then
			if Application("em_Enable_ShareFolder") = true then
				openresult = ei.LoadFriendMailBox(Session("wem"), sname, sfname, false)
			else
				openresult = -1
			end if

			if openresult = -1 then
				open_isok = false
			elseif  openresult = 1 then
				open_isok = false
			elseif  openresult = 2 then
				open_isok = false
			end if
		elseif mode = "1" then
			ei.LoadMailBox Session("wem"), tgname
		elseif mode = "2" then
			ei.LoadSession Session("wem")
		elseif mode = "3" then
			ei.LoadLabel Session("wem"), tgname
		elseif mode = "4" then
			ei.searchstring = Session("SearchStr")
			ei.LoadMailBox Session("wem"), "empty"
		end if

		if open_isok = true then
			dim get_num
			get_num = 20
			allnum = ei.getMailsCount
			i = 0
			min_i = -1
			max_i = -1

			if mode <> "2" then
				do while i < allnum
					ei.getMailInfoEx allnum - i - 1, idname, isread, priority, sendMail, sendName, subject, size, etime, mstate

					if min_i > -1 and max_i > -1 then
						if i >= min_i and i <= max_i then
							if i = 0 then
								if ret_value = "" then
									ret_value = "|"
								end if
							end if

							if ret_value = "" then
								ret_value = idname
							else
								ret_value = ret_value & Chr(9) & idname
							end if

							if i = allnum - 1 then
								ret_value = ret_value & Chr(9) & "|"
							end if
						else
							Exit Do
						end if
					end if

					if idname = fid and min_i = -1 and max_i = -1 then
						min_i = i - get_num
						max_i = i + get_num
	
						if min_i < 0 then
							min_i = 0
						end if

						if max_i > allnum then
							max_i = allnum
						end if

						i = min_i - 1
					end if

					idname = NULL
					isread = NULL
					priority = NULL
					sendMail = NULL
					sendName = NULL
					subject = NULL
					size = NULL
					etime = NULL
					mstate = NULL

					i = i + 1
				loop
			else
				do while i < allnum
					ei.getMailSessionInfo allnum - i - 1, idname, isread, priority, sendMail, sendName, subject, size, etime, mstate, msnum, msid

					if min_i > -1 and max_i > -1 then
						if i >= min_i and i <= max_i then
							if i = 0 then
								if ret_value = "" then
									ret_value = "|"
								end if
							end if

							if ret_value = "" then
								ret_value = msid
							else
								ret_value = ret_value & Chr(9) & msid
							end if

							if i = allnum - 1 then
								ret_value = ret_value & Chr(9) & "|"
							end if
						else
							Exit Do
						end if
					end if

					if msid = fid and min_i = -1 and max_i = -1 then
						min_i = i - get_num
						max_i = i + get_num
	
						if min_i < 0 then
							min_i = 0
						end if

						if max_i > allnum then
							max_i = allnum
						end if

						i = min_i - 1
					end if

					idname = NULL
					isread = NULL
					priority = NULL
					sendMail = NULL
					sendName = NULL
					subject = NULL
					size = NULL
					etime = NULL
					mstate = NULL
					msnum = NULL
					msid = NULL

					i = i + 1
				loop
			end if
		end if

		set ei = nothing
	end if

	Response.Write ret_value
end if
%>
