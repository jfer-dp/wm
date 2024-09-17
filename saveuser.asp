<!--#include file="passinc.asp" -->

<%
if isadmin() = false and isAccountsAdmin() = false then
	Response.Redirect "noadmin.asp"
end if
%>

<%
mode = trim(request("mode"))

dim ei
set ei = Application("em")

if mode = "apptmp" or mode = "apptmptoall" then
	gourl = trim(request("gourl"))

	dim uwt
	set uwt = server.createobject("easymail.UserWorkTimer")
	uwt.Load_Templet

	if mode = "apptmptoall" then
		allnum = ei.GetUsersCount
		i = 0

		do while i < allnum
			ei.GetUserByIndex3 i, name, domain, comment, forbid, lasttime, accessmode, limitout, expiresday, monitor

			uwt.Set_Templet_To_User name

			if uwt.is_update_disabled_user = true then
				if forbid = false then
					if uwt.disabled_user_over = "1" or Len(uwt.disabled_user_over) = 8 then
						ei.ForbidUserByName name, true
					end if
				else
					if uwt.disabled_user_over = "0" then
						ei.ForbidUserByName name, false
					end if
				end if
			end if

			if uwt.is_update_limitout = true then
				if limitout = false then
					if uwt.limitout_over = "1" or Len(uwt.limitout_over) = 8 then
						ei.SetLimitOut name, true
					end if
				else
					if uwt.limitout_over = "0" then
						ei.SetLimitOut name, false
					end if
				end if
			end if

			name = NULL
			domain = NULL
			comment = NULL
			forbid = NULL
			lasttime = NULL
			accessmode = NULL
			limitout = NULL
			expiresday = NULL
			monitor = NULL

			i = i + 1
		loop

		set uwt = nothing
		set ei = nothing
		Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
	end if

	dim mt
	set mt = server.createobject("easymail.WMethod")

	dim msg
	msg = trim(request("mulusers"))
	dim item
	dim ss
	dim se

	if Len(msg) > 0 then
		ss = 1
		se = 1

		Do While 1
			se = InStr(ss, msg, Chr(12))

			If se <> 0 Then
				item = Mid(msg, ss, se - ss)
				uwt.Set_Templet_To_User item

				if uwt.is_update_disabled_user = true then
					if mt.Is_Disabled_User(item) = false then
						if uwt.disabled_user_over = "1" or Len(uwt.disabled_user_over) = 8 then
							ei.ForbidUserByName item, true
						end if
					else
						if uwt.disabled_user_over = "0" then
							ei.ForbidUserByName item, false
						end if
					end if
				end if

				if uwt.is_update_limitout = true then
					if mt.Is_Limitout_User(item) = false then
						if uwt.limitout_over = "1" or Len(uwt.limitout_over) = 8 then
							ei.SetLimitOut item, true
						end if
					else
						if uwt.limitout_over = "0" then
							ei.SetLimitOut item, false
						end if
					end if
				end if
			Else
				Exit Do
			End If

			ss = se + 1
		Loop
	end if

	set mt = nothing
	set uwt = nothing
	set ei = nothing
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=" & Server.URLEncode(gourl)
end if


min_index = 0
max_index = 0

if IsNumeric(trim(request("minindex"))) = true then
	min_index = CLng(trim(request("minindex")))
end if

if IsNumeric(trim(request("maxindex"))) = true then
	max_index = CLng(trim(request("maxindex")))
end if

i = max_index

dim cdomain
cdomain = trim(request("cdomain"))

if mode = "del" then

	do while i >= min_index
		if trim(request("check" & i)) <> "" then
			ei.DelUserByIndex i
		end if 

	    i = i - 1
	loop

elseif mode = "forbid" then

	do while i >= min_index
		if trim(request("check" & i)) <> "" then
			ei.ForbidUserByIndex i, TRUE
		end if 

	    i = i - 1
	loop

elseif mode = "clear" then

	do while i >= min_index
		if trim(request("check" & i)) <> "" then
			ei.ForbidUserByIndex i, FALSE
		end if 

	    i = i - 1
	loop

end if

set ei = nothing

searchtext = trim(request("searchtext"))
page = trim(request("page"))
sortby = trim(request("sortby"))

Response.Redirect "showuser.asp?" & getGRSN() & "&sortby=" & sortby & "&cdomain=" & cdomain & "&page=" & page & "&searchtext=" & searchtext
%>
