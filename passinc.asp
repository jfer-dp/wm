<%
if Session("wem") = "" then
	Response.Redirect "default.asp"
end if

if Session("SecEx") = "1" then
	Response.CacheControl = "no-cache"
end if

if Application("em").HaveKick() = true then
	if Application("em").IsKick(Session("wem"), Request.ServerVariables("REMOTE_ADDR")) = true then
		Application("em").DelKick Session("wem"), Request.ServerVariables("REMOTE_ADDR")

		Session("wem") = ""
		Session("mail") = ""
		Session("tid") = ""
		Session("SecEx") = ""
		Session("scpw") = ""
		Session("cert_ca") = ""

		Response.Redirect "default.asp"
	end if
end if

Application("em").SessionHeart Session("wem"), Session("tid")

pageline = Session("pl")

IsEnterpriseVersion = true


csbi_color_str_default = "8CA5B5DBEAF5EFF7FF104A7B93BEE2336699FFCE00458090FF8102458090F0F0F0"

if Session("my_Show_Color") <> "" then
	MY_COLOR_1 = "#" & getColorStringByIndex(Session("my_Show_Color"), 1)
	MY_COLOR_2 = "#" & getColorStringByIndex(Session("my_Show_Color"), 2)
	MY_COLOR_3 = "#" & getColorStringByIndex(Session("my_Show_Color"), 3)
	MY_COLOR_4 = "#" & getColorStringByIndex(Session("my_Show_Color"), 4)
	MY_COLOR_5 = "#" & getColorStringByIndex(Session("my_Show_Color"), 5)
	MY_COLOR_6 = "#" & getColorStringByIndex(Session("my_Show_Color"), 6)
	MY_COLOR_7 = "#" & getColorStringByIndex(Session("my_Show_Color"), 7)
	MY_COLOR_8 = "#" & getColorStringByIndex(Session("my_Show_Color"), 8)
	MY_COLOR_9 = "#" & getColorStringByIndex(Session("my_Show_Color"), 9)
	MY_COLOR_10 = "#" & getColorStringByIndex(Session("my_Show_Color"), 10)
	MY_COLOR_11 = "#" & getColorStringByIndex(Session("my_Show_Color"), 11)
else
	MY_COLOR_1 = "#" & getColorStringByIndex(csbi_color_str_default, 1)
	MY_COLOR_2 = "#" & getColorStringByIndex(csbi_color_str_default, 2)
	MY_COLOR_3 = "#" & getColorStringByIndex(csbi_color_str_default, 3)
	MY_COLOR_4 = "#" & getColorStringByIndex(csbi_color_str_default, 4)
	MY_COLOR_5 = "#" & getColorStringByIndex(csbi_color_str_default, 5)
	MY_COLOR_6 = "#" & getColorStringByIndex(csbi_color_str_default, 6)
	MY_COLOR_7 = "#" & getColorStringByIndex(csbi_color_str_default, 7)
	MY_COLOR_8 = "#" & getColorStringByIndex(csbi_color_str_default, 8)
	MY_COLOR_9 = "#" & getColorStringByIndex(csbi_color_str_default, 9)
	MY_COLOR_10 = "#" & getColorStringByIndex(csbi_color_str_default, 10)
	MY_COLOR_11 = "#" & getColorStringByIndex(csbi_color_str_default, 11)
end if


function isadmin()
	isadmin = false
	if Session("wem") = Application("em_SystemAdmin") then
		isadmin = true
	end if
end function


function isAccountsAdmin()
	isAccountsAdmin = false
	if Session("wem") = Application("em_AccountsAdmin") then
		isAccountsAdmin = true
	end if
end function


function getGRSN()
	dim theGRSN
	Randomize
	theGRSN = Int((9999999 * Rnd) + 1)

	getGRSN = "GRSN=" & CStr(theGRSN)
end function


function getColorStringByIndex(csbi_color_str, csbi_color_index)
	getColorStringByIndex = ""

	if IsNumeric(csbi_color_index) = true and csbi_color_index > 0 then
		getColorStringByIndex = Mid(csbi_color_str, ((csbi_color_index - 1) * 6) + 1, 6)
	end if
end function
%>
