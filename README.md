<div align="center">

## Pure ASP Barcode Generator


</div>

### Description

This script generates a .bmp barcode from scratch with no COM+ object required. Supports only a few types, but the common ones (UPC-A, code128b, code39, EAN-13).
 
### More Info
 
<img src="http://www.yoursite.com/barcode.asp?code=YourBarCode012345&height=20&width=1&mode=code39">

code = bar code value

height = height of barcode in pixels.

width = width MULTIPLIER in pixels.

mode = type of barcode (Currently supported barcode types: code39, code128b, UPC-A, EAN-13)

a barcode :-)

none...please notify me if any.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mark Kahn](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mark-kahn.md)
**Level**          |Advanced
**User Rating**    |5.0 (85 globes from 17 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Graphics/ Sound](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics-sound__4-15.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mark-kahn-pure-asp-barcode-generator__4-8383/archive/master.zip)





### Source Code

```
<%
OPTION EXPLICIT
response.contenttype	=	"image/bmp"
'img src="http://www.yoursite.com/barcode.asp?code=YourBarCode012345&height=20&width=1&mode=code39"
'
' code = bar code value
' height = height of barcode in pixels.
' width = width MULTIPLIER in pixels.
' mode = type of barcode (Currently supported barcode types: code39, code128b, UPC-A, EAN-13)
'
' NOTE: If you prefer, you can also set the mode to 'raw' and create the barcode yourself by setting the code to 1s and 0s representing the barcode, ie: 11001100001010... In this case, 1s are black, 0s are white.
'
' NOTE: Maximum width & height values are 65536 pixels. Values larger than this will cause errors in the bmp file. This is a limitation of the bmp file format (why would you WANT an barcode this large anyway?)
'
' Additional code types are very easy to implement.
'
' Images generated are very small. For instance, an ean-13 barcode at a height of 50 pixels is a mere 662 bytes (less than 1kb). The largest realistic barcodes I've generated were less than 2kb.
'
' I added support for code caching. Note that the image is NOT cached, only the final set of 1s and 0s that represent the bars.
'
' If anyone adds additional codes, please send me the source, thanks :-)
' cwolves@cwolves.com
dim code, origcode, height, width, mode, caching, FontKey, FontCN10, FontCN12
caching	= True	' turn this on to cache barcodes in '10101010' format. Might speed things up on busy servers, although this script doesn't take many resources to begin with. An EAN-13 or UPC barcode will take less than 100 bytes of memory space. Other types will take more or less depending on the length of the barcode created.
' DO NOT EDIT BELOW THIS LINE!
code		= request.querystring("code")
height	= request.querystring("height")
width		= request.querystring("width")
mode		= request.querystring("mode")
origcode	= code
if not IsNumeric(height) or height	= "" then	height	= 1 else height	= numeric(height)
if not IsNumeric(width) or width		= "" then	width		= 1 else width		= numeric(width)
if caching AND application("cache" & origcode & mode & height & width) <> "" then
	code	= application("cache" & origcode & mode & height & width)
else
	select case lcase(mode)
		case "raw"			' do nothing. non-0 chars are automatically 1s
		case "code39":		code	= code39(code)
		case "code128b":	code	= code128b(code)
		case "upc-a":		code	= codeean13("0" & code, "AAAAAA")
		case "ean-13":		code	= codeean13(code, eanflag(left(code, 1)))
	end select
	if caching then
		Application.Lock
		Application("cache" & origcode & mode & height & width)	= code
		Application.UnLock
	end if
end if
Function stb(String)
 Dim I, B
 For I=1 to len(String)
 B	= B & ChrB(Asc(Mid(String,I,1)))
 Next
 stb	= B
End Function
function tstr(data, width)
	dim tchar, total, tpos, i, j, x
	tchar	= 0
	total	= ""
	tpos	= 8
	for i	= 1 to len(data)
		for j	= 1 to width
			tpos		= tpos - 1
			if mid(data, i, 1) <> "0" then tchar	= tchar + 2^tpos
			if tpos	= 0 then
				total	= total & chr(tchar)
				tpos	= 8
				tchar	= 0
			end if
		next
	next
	if tpos <> 8 then
		total	= total & chr(tchar)
	end if
	x		= len(total) mod 4
	if x	= 0 then x	= 4
	for i	= x to 3
		total	= total & chr(0)
	next
	tstr	= total
end function
function numeric(num)
	dim numb, valid, i
	numb	= ""
	valid	= "0123456789"
	for i	= 1 to len(num)
		if InStr(valid, mid(num, i, 1)) > 0 then numb	= numb & mid(num, i, 1)
	next
	num		= left(num, 30)
	numeric	= cint(num)
end function
function size(lngth)
	lngth	= cdbl(lngth)
	if lngth	> 255 then
		if lngth > 65535 then lngth	= 65535
		size	= chr(lngth mod 256) & chr(int(lngth/256))
	else
		size	= chr(lngth) & chr(0)
	end if
end function
function code39(code)
	dim output, i, clet
	output	= ""
	code		= "*" & replace(code, "*", "") & "*"
	for i	= 1 to len(code)
		clet	= ""
		select case ucase(mid(code, i, 1))
			case "1": clet	= "111010001010111"
			case "2": clet	= "101110001010111"
			case "3": clet	= "111011100010101"
			case "4": clet	= "101000111010111"
			case "5": clet	= "111010001110101"
			case "6": clet	= "101110001110101"
			case "7": clet	= "101000101110111"
			case "8": clet	= "111010001011101"
			case "9": clet	= "101110001011101"
			case "0": clet	= "101000111011101"
			case "A": clet	= "111010100010111"
			case "B": clet	= "101110100010111"
			case "C": clet	= "111011101000101"
			case "D": clet	= "101011100010111"
			case "E": clet	= "111010111000101"
			case "F": clet	= "101110111000101"
			case "G": clet	= "101010001110111"
			case "H": clet	= "111010100011101"
			case "I": clet	= "101110100011101"
			case "J": clet	= "101011100011101"
			case "K": clet	= "111010101000111"
			case "L": clet	= "101110101000111"
			case "M": clet	= "111011101010001"
			case "N": clet	= "101011101000111"
			case "O": clet	= "111010111010001"
			case "P": clet	= "101110111010001"
			case "Q": clet	= "101010111000111"
			case "R": clet	= "111010101110001"
			case "S": clet	= "101110101110001"
			case "T": clet	= "101011101110001"
			case "U": clet	= "111000101010111"
			case "V": clet	= "100011101010111"
			case "W": clet	= "111000111010101"
			case "X": clet	= "100010111010111"
			case "Y": clet	= "111000101110101"
			case "Z": clet	= "100011101110101"
			case "-": clet	= "100010101110111"
			case ".": clet	= "111000101011101"
			case " ": clet	= "100011101011101"
			case "*": clet	= "100010111011101"
			case "$": clet	= "100010001000101"
			case "/": clet	= "100010001010001"
			case "+": clet	= "100010100010001"
			case "%": clet	= "101000100010001"
		end select
		output	= output & clet & "0"
	next
	code39		= left(output, len(output)-1)
end function
Function code128b(ByVal InputString)
	Const MinValidAscii	= 32
	Const MaxValidAscii	= 126
	Dim CharValue(255)
	Dim i
	for i	= 0 to 94
		CharValue(i+32)	= i
	next
	for i	= 95 to 106
		CharValue(i+100)	= i
	next
	' Encode the input string
	InputString	= Trim(InputString)
	Dim CheckDigitValue, CharPos, CharAscii, InvalidCharsFound
	InvalidCharsFound	= false
	CheckDigitValue	= CharValue(204)
	For CharPos	= 1 To Len(InputString)
		CharAscii		= Asc(Mid(InputString, CharPos, 1))
		if (CharAscii < MinValidAscii) OR (CharAscii > MaxValidAscii) then
			CharAscii			= Asc("?")
			InvalidCharsFound	= true
		end if
		CheckDigitValue	= CheckDigitValue + (CharValue(CharAscii) * CharPos)
	Next
	CheckDigitValue		= (CheckDigitValue Mod 103)
	Dim CheckDigitAscii
	if CheckDigitValue < 95 then
		CheckDigitAscii	= CheckDigitValue + 32
	else
		CheckDigitAscii	= CheckDigitValue + 100
	end if
	Dim OutputString
	OutputString			= Chr(204) & InputString & Chr(CheckDigitAscii) & Chr(206)
	Dim BarcodePattern(255)
	BarcodePattern(32) 	= "212222"		' <SPACE>
	BarcodePattern(33) 	= "222122"		' !
	BarcodePattern(34) 	= "222221"		' "
	BarcodePattern(35) 	= "121223"		' #
	BarcodePattern(36) 	= "121322"		' $
	BarcodePattern(37) 	= "131222"		' %
	BarcodePattern(38) 	= "122213"		' &
	BarcodePattern(39) 	= "122312"		' '
	BarcodePattern(40) 	= "132212"		' (
	BarcodePattern(41) 	= "221213"		' )
	BarcodePattern(42) 	= "221312"		' *
	BarcodePattern(43) 	= "231212"		' +
	BarcodePattern(44) 	= "112232"		' ,
	BarcodePattern(45) 	= "122132"		' -
	BarcodePattern(46) 	= "122231"		' .
	BarcodePattern(47) 	= "113222"		' /
	BarcodePattern(48) 	= "123122"		' 0
	BarcodePattern(49) 	= "123221"		' 1
	BarcodePattern(50) 	= "223211"		' 2
	BarcodePattern(51) 	= "221132"		' 3
	BarcodePattern(52) 	= "221231"		' 4
	BarcodePattern(53) 	= "213212"		' 5
	BarcodePattern(54) 	= "223112"		' 6
	BarcodePattern(55) 	= "312131"		' 7
	BarcodePattern(56) 	= "311222"		' 8
	BarcodePattern(57) 	= "321122"		' 9
	BarcodePattern(58) 	= "321221"		' :
	BarcodePattern(59) 	= "312212"		' ;
	BarcodePattern(60) 	= "322112"		' <
	BarcodePattern(61) 	= "322211"		' =
	BarcodePattern(62) 	= "212123"		' >
	BarcodePattern(63) 	= "212321"		' ?
	BarcodePattern(64) 	= "232121"		' @
	BarcodePattern(65) 	= "111323"		' A
	BarcodePattern(66) 	= "131123"		' B
	BarcodePattern(67) 	= "131321"		' C
	BarcodePattern(68) 	= "112313"		' D
	BarcodePattern(69) 	= "132113"		' E
	BarcodePattern(70) 	= "132311"		' F
	BarcodePattern(71) 	= "211313"		' G
	BarcodePattern(72) 	= "231113"		' H
	BarcodePattern(73) 	= "231311"		' I
	BarcodePattern(74) 	= "112133"		' J
	BarcodePattern(75) 	= "112331"		' K
	BarcodePattern(76) 	= "132131"		' L
	BarcodePattern(77) 	= "113123"		' M
	BarcodePattern(78) 	= "113321"		' N
	BarcodePattern(79) 	= "133121"		' O
	BarcodePattern(80) 	= "313121"		' P
	BarcodePattern(81) 	= "211331"		' Q
	BarcodePattern(82) 	= "231131"		' R
	BarcodePattern(83) 	= "213113"		' S
	BarcodePattern(84) 	= "213311"		' T
	BarcodePattern(85) 	= "213131"		' U
	BarcodePattern(86) 	= "311123"		' V
	BarcodePattern(87) 	= "311321"		' W
	BarcodePattern(88) 	= "331121"		' X
	BarcodePattern(89) 	= "312113"		' Y
	BarcodePattern(90) 	= "312311"		' Z
	BarcodePattern(91) 	= "332111"		' [
	BarcodePattern(92) 	= "314111"		' /
	BarcodePattern(93) 	= "221411"		' ]
	BarcodePattern(94) 	= "431111"		' ^
	BarcodePattern(95) 	= "111224"		' _
	BarcodePattern(96) 	= "111422"		' `
	BarcodePattern(97) 	= "121124"		' a
	BarcodePattern(98) 	= "121421"		' b
	BarcodePattern(99) 	= "141122"		' c
	BarcodePattern(100)	= "141221"		' d
	BarcodePattern(101)	= "112214"		' e
	BarcodePattern(102)	= "112412"		' f
	BarcodePattern(103)	= "122114"		' g
	BarcodePattern(104)	= "122411"		' h
	BarcodePattern(105)	= "142112"		' i
	BarcodePattern(106)	= "142211"		' j
	BarcodePattern(107)	= "241211"		' k
	BarcodePattern(108)	= "221114"		' l
	BarcodePattern(109)	= "413111"		' m
	BarcodePattern(110)	= "241112"		' n
	BarcodePattern(111)	= "134111"		' o
	BarcodePattern(112)	= "111242"		' p
	BarcodePattern(113)	= "121142"		' q
	BarcodePattern(114)	= "121241"		' r
	BarcodePattern(115)	= "114212"		' s
	BarcodePattern(116)	= "124112"		' t
	BarcodePattern(117)	= "124211"		' u
	BarcodePattern(118)	= "411212"		' v
	BarcodePattern(119)	= "421112"		' w
	BarcodePattern(120)	= "421211"		' x
	BarcodePattern(121)	= "212141"		' y
	BarcodePattern(122)	= "214121"		' z
	BarcodePattern(123)	= "412121"		' {
	BarcodePattern(124)	= "111143"		' |
	BarcodePattern(125)	= "111341"		' }
	BarcodePattern(126)	= "131141"		' ~
	BarcodePattern(195)	= "114113"
	BarcodePattern(196)	= "114311"
	BarcodePattern(197)	= "411113"
	BarcodePattern(198)	= "411311"
	BarcodePattern(199)	= "113141"
	BarcodePattern(200)	= "114131"
	BarcodePattern(201)	= "311141"
	BarcodePattern(202)	= "411131"
	BarcodePattern(203)	= "211412"
	BarcodePattern(204)	= "211214"
	BarcodePattern(205)	= "211232"
	BarcodePattern(206)	= "2331112"
	Dim OutputPattern, ThisPattern, thischar
	OutputPattern	= ""
	for CharPos		= 1 to Len(OutputString)
		ThisPattern	= BarcodePattern(Asc(Mid(OutputString, CharPos, 1)))
		for i = 1 to len(ThisPattern)
			if i mod 2 = 1 then thischar	= "1" else thischar	= "0"
			OutputPattern	= OutputPattern & replace(space(int(mid(ThisPattern, i, 1))), " ", thischar)
		next
	next
	code128b	= OutputPattern
End Function
Function CodeEAN13(code, encoding)
	Dim leftA, leftB, rght, OutputPattern, i
	if len(code) = 13 then
		LeftA	= Array("0001101", "0011001", "0010011", "0111101", "0100011", "0110001", "0101111", "0111011", "0110111", "0001011")
		LeftB	= Array("0100111", "0110011", "0011011", "0100001", "0011101", "0111001", "0000101", "0010001", "0001001", "0010111")
		Rght	= Array("1110010", "1100110", "1101100", "1000010", "1011100", "1001110", "1010000", "1000100", "1001000", "1110100")
		OutputPattern	= "101"
		for i = 1 to 6
			if mid(ucase(encoding), i, 1)	= "A" then
				OutputPattern	= OutputPattern & LeftA(cint(mid(code, i+1, 1)))
			else
				OutputPattern	= OutputPattern & LeftB(cint(mid(code, i+1, 1)))
			end if
		next
		OutputPattern		= OutputPattern & "01010"
		for i = 1 to 6
			OutputPattern	= OutputPattern & Rght(cint(mid(code, i+7, 1)))
		next
		OutputPattern		= OutputPattern & "101"
		CodeEAN13			= OutputPattern
	end if
End Function
Function eanflag(num)
	select case num
		case 0:	eanflag	= "AAAAAA"
		case 1:	eanflag	= "AABABB"
		case 2:	eanflag	= "AABBAB"
		case 3:	eanflag	= "AABBBA"
		case 4:	eanflag	= "ABAABB"
		case 5:	eanflag	= "ABBAAB"
		case 6:	eanflag	= "ABBBAA"
		case 7:	eanflag	= "ABABAB"
		case 8:	eanflag	= "ABABBA"
		case 9:	eanflag	= "ABBABA"
	end select
End Function
dim dataout, i
if code <> "" then
	dataout	= tstr(code, width)
	response.binarywrite stb(chr(66) & chr(77) & size(62+(len(dataout)*height)) & chr(0) & chr(0) & chr(0) & chr(0) & chr(0) & chr(0) & chr(62) & chr(0) & chr(0) & chr(0) & chr(40) & chr(0) & chr(0) & chr(0) & size(len(code)*width) & chr(0) & chr(0) & size(height) & chr(0) & chr(0) & chr(1) & chr(0) & chr(1) & chr(0) & chr(0) & chr(0) & chr(0) & chr(0) & chr(0) & chr(0) & chr(0) & chr(0) & chr(37) & chr(14) & chr(0) & chr(0) & chr(37) & chr(14) & chr(0) & chr(0) & chr(0) & chr(0) & chr(0) & chr(0) & chr(0) & chr(0) & chr(0) & chr(0) & chr(255) & chr(255) & chr(255) & chr(0) & chr(0) & chr(0) & chr(0) & chr(0))
	for i	= 1 to height
		response.binarywrite stb(dataout)
	next
end if
%>
```

