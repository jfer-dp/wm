<!--#include file="passinc.asp" -->

<%
dim dm
set dm = server.createobject("easymail.Domain")
dm.Load

if dm.GetUserManagerDomainCount(Session("wem")) < 1 then
	if isadmin() = false then
		set dm = nothing
		response.redirect "noadmin.asp"
	end if
end if

dim ManagerDomainString
i = 0
if isadmin() = false then
	ManagerDomainString = Chr(9)
	allnum = dm.GetUserManagerDomainCount(Session("wem"))

	do while i < allnum
		domain = dm.GetUserManagerDomain(Session("wem"), i)
		ManagerDomainString = ManagerDomainString & LCase(domain) & Chr(9)
		domain = NULL

		i = i + 1
	loop
end if

set dm = nothing


dim ei
set ei = server.createobject("easymail.mailinglist")
'-----------------------------------------

i = 0
if trim(request("mode")) = "mdel" then
	ei.LoadLists

	allnum = ei.MailingListCount + 1
	do while i < allnum
		tmp_ck_name = trim(request("check" & i))
		if tmp_ck_name <> "" then
			if isadmin() = true or canDelete(tmp_ck_name) = true then
				ei.DeleteList tmp_ck_name
			end if
		end if 

	    i = i + 1
	loop

	ei.Save
elseif trim(request("mode")) = "addnew" then
	dim msg
	msg = trim(request("addlist"))
	addname = trim(request("addname"))

	if Len(msg) > 0 or Len(addname) > 0 then
		ei.LoadOne addname
		ei.RemoveAllItem()
		ei.RemoveAllAccreditItem()

		dim ss
		dim se
		ss = 1
		se = 1

	    Do While 1
	        se = InStr(ss, msg, Chr(9))

	        If se <> 0 Then
    	        item = Mid(msg, ss, se - ss)
    	        ei.AddItem item
			Else
	            Exit Do
    	    End If

	        ss = se + 1
	    Loop


		msg = trim(request("accredit_list"))
		if Len(msg) > 0 then
			ss = 1
			se = 1

		    Do While 1
		        se = InStr(ss, msg, Chr(9))

		        If se <> 0 Then
	    	        item = Mid(msg, ss, se - ss)
	    	        ei.AddAccreditItem item
				Else
		            Exit Do
	    	    End If

		        ss = se + 1
		    Loop
		end if


		if trim(request("isSendWithMailingList")) <> "" then
			ei.isSendWithMailingList = true
		else
			ei.isSendWithMailingList = false
		end if

		if trim(request("isPrivate")) <> "" then
			ei.isPrivate = true
		else
			ei.isPrivate = false
		end if

		if trim(request("isShowToCc")) <> "" then
			ei.isShowToCc = true
		else
			ei.isShowToCc = false
		end if

		if trim(request("isDisabled")) <> "" then
			ei.isDisabled = true
		else
			ei.isDisabled = false
		end if

		if isadmin() = true then
			ei.dManagerDomain = trim(request("dManagerDomain"))
		else
			ei.dManagerDomain = Session("mail")
		end if

		ei.Save()
	end if
end if

delnm = trim(request("del"))
if delnm <> "" then
	ei.LoadLists

	if isadmin() = true or canDelete(delnm) = true then
		ei.DeleteList delnm
		ei.Save
	end if
end if

set ei = nothing

if err.number = 0 then
	Response.Redirect "ok.asp?" & getGRSN() & "&gourl=browmailinglist.asp"
else
	Response.Redirect "err.asp?" & getGRSN()
end if



function canDelete(ml_name)
	canDelete = false

	dim sd_i
	dim sd_allnum
	dim sd_findp

	sd_i = 0
	sd_allnum = ei.MailingListCount

	do while sd_i < sd_allnum
		ei.Get sd_i, sd_name, sd_isSendWithMailingList, sd_isPrivate, sd_dManagerDomain

		if LCase(sd_name) = LCase(ml_name) then
			if Len(sd_dManagerDomain) > 0 then
				sd_findp = InStr(1, sd_dManagerDomain, "@")
				if sd_findp > 0 then
					if InStr(1, ManagerDomainString, Chr(9) & LCase(Mid(sd_dManagerDomain, sd_findp + 1)) & Chr(9)) > 0 then
						canDelete = true
					end if
				end if
			end if

			sd_name = NULL
			sd_isSendWithMailingList = NULL
			sd_isPrivate = NULL
			sd_dManagerDomain = NULL

			exit do
		end if

		sd_name = NULL
		sd_isSendWithMailingList = NULL
		sd_isPrivate = NULL
		sd_dManagerDomain = NULL

	    sd_i = sd_i + 1
	loop
end function
%>
