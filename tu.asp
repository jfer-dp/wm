<%
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.cachecontrol = "no-cache"
Response.ContentType = "Image/Bmp"

Call Com_CreatValidCode()

Sub Com_CreatValidCode()
Randomize

Dim i, ii, iii
Const cAmount = 10
Const cCode = "0123456789"

Dim vColorData(2)
vColorData(0) = ""
vColorData(1) = ChrB(255) & ChrB(255) & ChrB(255)

Dim vCode(4), vCodes

For i = 0 To 3
	vCode(i) = Int(Rnd * cAmount)
	if i=1 then vCode(i)="11"
	if i=3 then vCode(i)="10"
	vCodes=vCodes&Mid(cCode,vCode(i)+1,1)
Next

Response.Cookies("zatt_checkcode") = int(Mid(vCodes,1,1)) + int(Mid(vCodes,2,1))

Dim vNumberData(36)
vNumberData(0) = "1110000111110111101111011110111101111011110111101111011110111101111011110111101111011110111110000111"
vNumberData(1) = "1111011111110001111111110111111111011111111101111111110111111111011111111101111111110111111100000111"
vNumberData(2) = "1110000111110111101111011110111111111011111111011111111011111111011111111011111111011110111100000011"
vNumberData(3) = "1110000111110111101111011110111111110111111100111111111101111111111011110111101111011110111110000111"
vNumberData(4) = "1111101111111110111111110011111110101111110110111111011011111100000011111110111111111011111111000011"
vNumberData(5) = "1100000011110111111111011111111101000111110011101111111110111111111011110111101111011110111110000111"
vNumberData(6) = "1111000111111011101111011111111101111111110100011111001110111101111011110111101111011110111110000111"
vNumberData(7) = "1100000011110111011111011101111111101111111110111111110111111111011111111101111111110111111111011111"
vNumberData(8) = "1110000111110111101111011110111101111011111000011111101101111101111011110111101111011110111110000111"
vNumberData(9) = "1110001111110111011111011110111101111011110111001111100010111111111011111111101111011101111110001111"

vNumberData(10) = "1111111111111111111111111111111000000001111111111111111111111000000001111111111111111111111111111111"
vNumberData(11) = "1111111111111100111111110011111111001111100000000110000000011111001111111100111111110011111111111111"

Response.BinaryWrite ChrB(66) & ChrB(77) & ChrB(230) & ChrB(4) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(54) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(40) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(40) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(10) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(1) & ChrB(0)
Response.BinaryWrite ChrB(24) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(176) & ChrB(4) & ChrB(0) & ChrB(0) & ChrB(18) & ChrB(11) & ChrB(0) & ChrB(0) & ChrB(18) & ChrB(11) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0)

For i = 9 To 0 Step -1
	For ii = 0 To 3
		For iii = 1 To 10
			if Mid(vNumberData(vCode(ii)), i * 10 + iii , 1) = "0" then
				dim a,b,c
				a=abs(Rnd * 256-60)
				b=abs(Rnd * 256-128)
				c=abs(Rnd * 256-60)
				vColorData(0) = ChrB(a) & ChrB(b) & ChrB(c)
				Response.BinaryWrite vColorData(Mid(vNumberData(vCode(ii)), i * 10 + iii , 1))
			else
				dim d,e,f
				d=abs(Rnd * 255)
				e=abs(Rnd * 255)
				f=abs(Rnd * 255)
				if d+e+f>640 then
					vColorData(1) = ChrB(d) & ChrB(e) & ChrB(f)
					Response.BinaryWrite vColorData(Mid(vNumberData(vCode(ii)), i * 10 + iii , 1))
				else
					Response.BinaryWrite vColorData(Mid(vNumberData(vCode(ii)), i * 10 + iii , 1))
				end if
			end if
		Next
	Next
Next

End Sub
%>