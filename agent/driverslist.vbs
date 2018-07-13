'----------------------------------------------------------
' Plugin for OCS Inventory NG 2.x
' Script :	Retrieve list of installed drivers
' Version :	1.00
' Date :	13/07/2018
' Author :	Stéphane PAUTREL (acb78.com)
'----------------------------------------------------------
' OS checked [X] on	32b	64b	(Professionnal edition)
' Windows XP		[X]
' Windows Vista		[X]	[X]
' Windows 7			[X]	[X]
' Windows 8.1		[X]	[X]	
' Windows 10		[X]	[X]
' Windows 2k8R2			[X]
' Windows 2k12R2		[X]
' Windows 2k16			[X]
' ---------------------------------------------------------
' NOTE : No checked on Windows 8
' ---------------------------------------------------------
On Error Resume Next

Set objWMIService = GetObject("winmgmts:root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPSignedDriver")

For Each objItem in colItems

	If Not IsNull(objItem.DeviceClass) Then
		Result = "<DRIVERSLIST>" & VbCrLf
		Result = Result & "<DESCRIPTION>" & replaceSpecialCar(objItem.Description) & "</DESCRIPTION>" & VbCrLf
		Result = Result & "<CLASS>" & objItem.DeviceClass & "</CLASS>" & VbCrLf
		Result = Result & "<DEVICEID>" & objItem.DeviceID & "</DEVICEID>" & VbCrLf
		Result = Result & "<NAME>" & replaceSpecialCar(objItem.DeviceName) & "</NAME>" & VbCrLf

		dtmWMIDate = objItem.DriverDate
		If Not IsNull(Left(objItem.DriverDate, 4)) Then
			DriverDate = Left(objItem.DriverDate, 4) & "/" & Mid(objItem.DriverDate, 5, 2) & "/" & Mid(objItem.DriverDate, 7, 2)
		Else
			DriverDate = "NC"
		End If

		Result = Result & "<DATE>" & DriverDate & "</DATE>" & VbCrLf
		Result = Result & "<PROVIDERNAME>" & objItem.DriverProviderName & "</PROVIDERNAME>" & VbCrLf
		Result = Result & "<VERSION>" & objItem.DriverVersion & "</VERSION>" & VbCrLf
		Result = Result & "<FRIENDLYNAME>" & replaceSpecialCar(objItem.FriendlyName) & "</FRIENDLYNAME>" & VbCrLf
		Result = Result & "<INFNAME>" & objItem.InfName & "</INFNAME>" & VbCrLf
		Result = Result & "<SIGNED>" & objItem.IsSigned & "</SIGNED>" & VbCrLf
		Result = Result & "</DRIVERSLIST>" & VbCrLf
		Wscript.Echo Result
	End If

Next

Function replaceSpecialCar(sText)
	If IsNull(sText) Then Exit Function
 
	Dim arrSpecCar(94,2)
	'arrSpecCar(0,0) = "À"
	arrSpecCar(0,0) = Chr(192)
		arrSpecCar(0,1) = "A"
	arrSpecCar(1,0) = "Á"
		arrSpecCar(1,1) = "A"
	arrSpecCar(2,0) = "Â"
		arrSpecCar(2,1) = "A"
	arrSpecCar(3,0) = "Ã"
		arrSpecCar(3,1) = "A"
	arrSpecCar(4,0) = "Ä"
		arrSpecCar(4,1) = "A"
	arrSpecCar(5,0) = "Å"
		arrSpecCar(5,1) = "A"
	'arrSpecCar(6,0) = "à"
	arrSpecCar(6,0)= Chr(224)
		arrSpecCar(6,1) = "a"
	arrSpecCar(7,0) = "á"
		arrSpecCar(7,1) = "a"
	'arrSpecCar(8,0) = "â"
	arrSpecCar(8,0) = Chr(226)
		arrSpecCar(8,1) = "a"
	arrSpecCar(9,0) = "ã"
		arrSpecCar(9,1) = "a"
	arrSpecCar(10,0)= "ä"
		arrSpecCar(10,1)= "a"
	arrSpecCar(11,0)= "å"
		arrSpecCar(11,1)= "a"
	arrSpecCar(12,0)= "Ò"
		arrSpecCar(12,1)= "O"
	arrSpecCar(13,0)= "Ó"
		arrSpecCar(13,1)= "O"
	arrSpecCar(14,0)= "Ô"
		arrSpecCar(14,1)= "O"
	arrSpecCar(15,0)= "Õ"
		arrSpecCar(15,1)= "O"
	arrSpecCar(16,0)= "Ö"
		arrSpecCar(16,1)= "O"
	arrSpecCar(17,0)= "Ø"
		arrSpecCar(17,1)= "O"
	'arrSpecCar(18,0)= "ò"
	arrSpecCar(18,0)= Chr(242)
		arrSpecCar(18,1)= "o"
	'arrSpecCar(19,0)= "ó"
	arrSpecCar(19,0)= Chr(243)
		arrSpecCar(19,1)= "o"
	'arrSpecCar(20,0)= "ô"
	arrSpecCar(20,0)= Chr(244)
		arrSpecCar(20,1)= "o"
	arrSpecCar(21,0)= "õ"
		arrSpecCar(21,1)= "o"
	arrSpecCar(22,0)= "ö"
		arrSpecCar(22,1)= "o"
	arrSpecCar(23,0)= "ø"
		arrSpecCar(23,1)= "o"
	arrSpecCar(24,0)= "È"
		arrSpecCar(24,1)= "E"
	'arrSpecCar(25,0)= "É"
	arrSpecCar(25,0)= Chr(201)
		arrSpecCar(25,1)= "E"
	'arrSpecCar(26,0)= "Ê"
	arrSpecCar(26,0)= Chr(202)
		arrSpecCar(26,1)= "E"
	'arrSpecCar(27,0)= "Ë"
	arrSpecCar(27,0)= Chr(203)
		arrSpecCar(27,1)= "E"
	'arrSpecCar(28,0)= "è"
	arrSpecCar(28,0)= Chr(232)
		arrSpecCar(28,1)= "e"
	'arrSpecCar(29,0)= "é"
	arrSpecCar(29,0)= Chr(233)
		arrSpecCar(29,1)= "e"
	'arrSpecCar(30,0)= "ê"
	arrSpecCar(30,0)= Chr(234)
		arrSpecCar(30,1)= "e"
	'arrSpecCar(31,0)= "ë"
	arrSpecCar(31,0)= Chr(235)
		arrSpecCar(31,1)= "e"
	'arrSpecCar(32,0)= "Ç"
	arrSpecCar(32,0)= Chr(199)
		arrSpecCar(32,1)= "C"
	'arrSpecCar(33,0)= "ç"
	arrSpecCar(33,0)= Chr(231)
		arrSpecCar(33,1)= "c"
	arrSpecCar(34,0)= "Ì"
		arrSpecCar(34,1)= "I"
	arrSpecCar(35,0)= "Í"
		arrSpecCar(35,1)= "I"
	arrSpecCar(36,0)= Chr(206)
		arrSpecCar(36,1)= "I"
	arrSpecCar(37,0)= "Ï"
		arrSpecCar(37,1)= "I"
	'arrSpecCar(38,0)= "ì"
	arrSpecCar(38,0)= Chr(236)
		arrSpecCar(38,1)= "i"
	'arrSpecCar(39,0)= "í"
	arrSpecCar(39,0)= Chr(237)
		arrSpecCar(39,1)= "i"
	'arrSpecCar(40,0)= "î"
	arrSpecCar(40,0)= Chr(238)
		arrSpecCar(40,1)= "i"
	'arrSpecCar(41,0)= "ï"
	arrSpecCar(41,0)= Chr(239)
		arrSpecCar(41,1)= "i"
	arrSpecCar(42,0)= "Ù"
		arrSpecCar(42,1)= "U"
	arrSpecCar(43,0)= "Ú"
		arrSpecCar(43,1)= "U"
	arrSpecCar(44,0)= "Û"
		arrSpecCar(44,1)= "U"
	arrSpecCar(45,0)= "Ü"
		arrSpecCar(45,1)= "U"
	'arrSpecCar(46,0)= "ù"
	arrSpecCar(46,0)= Chr(249)
		arrSpecCar(46,1)= "u"
	arrSpecCar(47,0)= "ú"
		arrSpecCar(47,1)= "u"
	arrSpecCar(48,0)= "û"
		arrSpecCar(48,1)= "u"
	arrSpecCar(49,0)= "ü"
		arrSpecCar(49,1)= "u"
	arrSpecCar(50,0)= "ÿ"
		arrSpecCar(50,1)= "y"
	arrSpecCar(51,0)= "Ñ"
		arrSpecCar(51,1)= "N"
	arrSpecCar(52,0)= "ñ"
		arrSpecCar(52,1)= "n"
	arrSpecCar(53,0)= "¢"
		arrSpecCar(53,1)= "cent"
	arrSpecCar(54,0)= "ß"
		arrSpecCar(54,1)= "beta"
	arrSpecCar(55,0)= "¥"
		arrSpecCar(55,1)= "yen"
	arrSpecCar(56,0)= "€"
		arrSpecCar(56,1)= "euro"
	arrSpecCar(57,0)= "£"
		arrSpecCar(57,1)= "pounds"
	arrSpecCar(58,0)= "™"
		arrSpecCar(58,1)= "trademark"
	arrSpecCar(59,0)= "©"
		arrSpecCar(59,1)= "copyright"
	arrSpecCar(60,0)= "®"
		arrSpecCar(60,1)= "registered"
	arrSpecCar(61,0)= "ª"
		arrSpecCar(61,1)= "exp(a)"
	arrSpecCar(62,0)= "×"
		arrSpecCar(62,1)= "x"
	arrSpecCar(63,0)= "÷"
		arrSpecCar(63,1)= "/"
	arrSpecCar(64,0)= "±"
		arrSpecCar(64,1)= "+/-"
	arrSpecCar(65,0)= "²"
		arrSpecCar(65,1)= "exp(2)"
	arrSpecCar(66,0)= "³"
		arrSpecCar(66,1)= "exp(3)"
	arrSpecCar(67,0)= "¼"
		arrSpecCar(67,1)= "1/4"
	arrSpecCar(68,0)= "½"
		arrSpecCar(68,1)= "1/2"
	arrSpecCar(69,0)= "¾"
		arrSpecCar(69,1)= "3/4"
	arrSpecCar(70,0)= "µ"
		arrSpecCar(70,1)= "u"
	arrSpecCar(71,0)= "¿"
		arrSpecCar(71,1)= "?"
	arrSpecCar(72,0)= "¶"
		arrSpecCar(72,1)= "paragraph"
	arrSpecCar(73,0)= "·"
		arrSpecCar(73,1)= "."
	arrSpecCar(74,0)= "¸"
		arrSpecCar(74,1)= ","
	arrSpecCar(75,0)= "º"
		arrSpecCar(75,1)= "o"
	arrSpecCar(76,0)= "°"
		arrSpecCar(76,1)= "degre"
	arrSpecCar(77,0)= "¯"
		arrSpecCar(77,1)= "_"
	arrSpecCar(78,0)= "§"
		arrSpecCar(78,1)= "paragraph"
	'arrSpecCar(79,0)= "…"
	arrSpecCar(79,0)= chr(133)
		arrSpecCar(79,1)= "..."
	arrSpecCar(80,0)= "¤"
		arrSpecCar(80,1)= "¤"
	'arrSpecCar(81,0)= "¦"
	arrSpecCar(81,0)= chr(166)
		arrSpecCar(81,1)= "pipe"
	arrSpecCar(82,0)= "‡"
		arrSpecCar(82,1)= "dagger"
	arrSpecCar(83,0)= "¬"
		arrSpecCar(83,1)= "-"
	arrSpecCar(84,0)= "ˆ"
		arrSpecCar(84,1)= " "
	arrSpecCar(85,0)= "¨"
		arrSpecCar(85,1)= " "
	arrSpecCar(86,0)= "‰"
		arrSpecCar(86,1)= "0/00"
	'arrSpecCar(87,0)= "œ"
	arrSpecCar(87,0)= chr(156)
		arrSpecCar(87,1)= "oe"
	arrSpecCar(88,0)= chr(171)
		arrSpecCar(88,1)= chr(34)
	arrSpecCar(89,0)= chr(187)
		arrSpecCar(89,1)= chr(34)
	arrSpecCar(90,0)= chr(32)
		arrSpecCar(90,1)= " "
	arrSpecCar(91,0)= chr(160)
		arrSpecCar(91,1)= ""
	arrSpecCar(92,0)= chr(60)
		arrSpecCar(92,1)= ""
	arrSpecCar(93,0)= chr(62)
		arrSpecCar(93,1)= ""
 
	' pour chaque element du tableau
	For i=0 To UBound(arrSpecCar,1)
		' on remplace
		sText = Replace(sText, arrSpecCar(i,0), arrSpecCar(i,1))
	Next
	' puis on affecte le retour
	replaceSpecialCar = sText
End Function