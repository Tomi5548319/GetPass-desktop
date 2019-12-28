version = "2.5"

nameMsg = "Enter the name of Website, app, ... you want to generate a new password for, i.e. " & Chr(34) & "Google" & Chr(34)
keyMsg = "Enter your key (remember this one) (can be the same for all sites)"
Title = "Password Generator v" & version & " by Tomi5548319"
prolog = "This is an application designed to generate a login password for your website, app ... using actual date, time, number from 1 to 9 and a key." & Chr(13) & Chr(13) & "Do you want to view changelog?"
useOldDataMsg = "Name was found in the database, do you want to copy information from database? (If you enter the same key as before, output password will be the same"
saveNewDataMsg = "Do you want to save your data into database to be able to access it later? (if you save it and enter the same name next time, app will ask you if you want to use this data. If you answer yes, app will ask you only for your key. If you enter the same key, you will get the same password.)"

Dim hv1						'Help variable 1 - used in reading functions and for encryption
Dim hv2						'Help variable 2 - used for encryption
Dim hv3						'Help variable 3 - used for encryption
Dim hv4()					'Help variable 4 - used for encryption
Dim hv5						'Help variable 5 - used for encryption
Dim outputPassword
Dim errorText
Dim restartAfterChangelog	'Variable used for restarting the app after viewing changelog
Dim name
Dim d
Dim t
Dim number
Dim password
Dim errorBox
Dim errorCode
Dim answer
Dim fso
Dim f
Dim line
Dim lineNumber
Dim found
Dim str
Dim saveMode

hv3 = Split(version, ".")
fileName = "passwordsv" & hv3(0) & hv3(1) & ".txt"

Set fso = CreateObject("Scripting.FileSystemObject")

If hv3(1) <> "0" Then
	If fso.FileExists("passwordsv" & hv3(0) & (hv3(1)-1) & ".txt") Then
		fso.DeleteFile("passwordsv" & hv3(0) & (hv3(1)-1) & ".txt")
	End If
End If

Do
	errorText= "Error " & Chr(13) & Chr(13) & "Do you want to restart the app or exit? (retry/exit)"
	errorCode = 0
	errorBox = 0
	restartAfterChangelog = 0
	found = False
	changelog = Msgbox(prolog,259,Title)
	If changelog = 7 Then
		restartAfterChangelog = 2
		Call getName()
		If Len(name) > 0 And errorCode = 0 Then			'If name = entered Then
			Call readAndRun()
		ElseIf errorCode = 0 Then
			errorCode = 1
			Call error(errorCode)
		End If
	ElseIf changelog = 6 Then
		Call Changeloglist()
	End If
Loop While restartAfterChangelog = 1 Or errorBox = 4



Function Changeloglist()
	changelogTitle = "Changelog of Password Generator by Tomi5548319"
	Changelog = ""
	Changelog = Changelog & "v2.5"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Code was simplified"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Fixed some bugs"
	Changelog = Changelog & Chr(13) & Chr(9) & Chr(9) & "-RNG was generating the same number"
	Changelog = Changelog & Chr(13) & Chr(9) & Chr(9) & "-Reading from database gave different output then when the password was generated for the first time"
	Changelog = Changelog & Chr(13) & "v2.4"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Data needed for generating a password is generated" & Chr(13) & Chr(9) & " automatically"
	Changelog = Changelog & Chr(13) & "v2.3"
	Changelog = Changelog & Chr(13) & Chr(9) & "*App name changed (Encrypter -> Password Generator)"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Added database for saving data needed to generate" & Chr(13) & Chr(9) & " the same password"
	Changelog = Changelog & Chr(13) & Chr(9) & "     -Only thing that you have to remember to generate" & Chr(13) & Chr(9) & "       the same password is the name of the website, app..." & Chr(13) & Chr(9) & "       and your key"
	Changelog = Changelog & Chr(13) & "v2.2"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Method of sending output changed"
	Changelog = Changelog & Chr(13) & "v2.1"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Code was simplified"
	Changelog = Changelog & Chr(13) & "v2.0"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Encoding method changed"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Added more functions (code)"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Added time"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Added choose number from 1 to 9"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Fixed some bugs"
	Changelog = Changelog & Chr(13) & Chr(9) & Chr(9) & "-main screen won't let you exit after restart"
	Changelog = Changelog & Chr(13) & Chr(9) & Chr(9) & "-Error message " & Chr(34) & "Date was not entered" & Chr(34) & " didn't show"
	Changelog = Changelog & Chr(13) & "v1.4"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Code was simplified"
	Changelog = Changelog & Chr(13) & "v1.3"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Application was optimized for Windows 8, 8.1 and 10"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Code was simplified"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Added error messages"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Added Automatic date output"
	Changelog = Changelog & Chr(13) & "v1.2"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Added changelog"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Added buttons"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Added Automatic date"
	Changelog = Changelog & Chr(13) & "v1.1"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Code was simplified"
	Changelog = Changelog & Chr(13) & "v1.0"
	Changelog = Changelog & Chr(13) & Chr(9) & "*Encrypter was born on the 6th of May 2018"
	Changelog = Changelog & Chr(13) & Chr(9) & "*First functional version"
	restartAfterChangelog = Msgbox(Changelog,1,changelogTitle)
End Function


Function readAndRun()
	Set f = fso.OpenTextFile(fileName,1,True)
	Do Until f.AtEndOfStream
		Line = f.ReadLine
		If Line = name Then
			found = True
			Exit Do
		End If 
		LineNumber = LineNumber + 1
	Loop
	f.Close
	If found = True Then
		Set f = fso.OpenTextFile(fileName,1,True)
		str = f.ReadAll
		f.Close
		str = Split(str,Chr(13) & Chr(10))
		If errorCode = 0 Then
			answer = Msgbox(useOldDataMsg,3,Title)
			If answer = 6 Then					'Use old data from database, ask for password and encrypt
				d = str(LineNumber+1)
				t = str(LineNumber+2)
				number = CInt(str(LineNumber+3))
				Call getPassword()
				If Len(password) > 0 And errorCode = 0 Then
					outputPassword = inputbox("Your password for " & name & ": ",Title,encode)
				ElseIf errorCode = 0 Then
					errorCode = 2
					Call error(errorCode)
				End If
			ElseIf answer = 7 Then				'Start normally
				Call run()
				If errorCode = 0 Then
					answer = Msgbox(saveNewDataMsg,3,Title)
					If answer = 6 And errorCode = 0 Then				'Save (modify) old password
						Call save("Old")
					End If
				End If
			End If
		End If
	Else									'Start normally
		Call run()
		If errorCode = 0 Then
			answer = Msgbox(saveNewDataMsg,3,Title)
			If answer = 6 And errorCode = 0 Then				'Save new password (add at end)
				Call save("New")
			End If
		End If
	End If
End Function


Function run()
	Call getDate()
	If Len(d) > 0 And errorCode = 0 Then
		Call getTime()
		If Len(t) > 0 And errorCode = 0 Then
			Call getNumber()
			If Len(number) > 0 And errorCode = 0 Then
				Call getPassword()
				If Len(password) > 0 And errorCode = 0 Then
					outputPassword = inputbox("Your password for " & name & ": ",Title,encode)
				ElseIf errorCode = 0 Then
					errorCode = 2
					Call error(errorCode)
				End If
			End If
		End If
	End If
End Function


Function save(saveMode)
	If saveMode = "New" Then
		Set f = fso.OpenTextFile(fileName,8,True)
		f.WriteLine(name)
		f.WriteLine(d)
		f.WriteLine(t)
		f.WriteLine(number)
		f.Close
	ElseIf saveMode = "Old" Then
		str(LineNumber+1) = d
		str(LineNumber+2) = t
		str(LineNumber+3) = number
		Set f = fso.OpenTextFile(fileName,2,True)
		For i = 0 To UBound(str)
			f.WriteLine(str(i))
		Next
	End If
End Function


Function encode()
	hv1 = StrReverse(name)
	Call nameToNumbers()
	Call addPassword()
	Call numbersToName()
	Call swap()
	Call nameLengthChange()
	hv1 = StrReverse(hv1)
	Call swap()
	Call nameToNumbers()
	Call addNumber()
	Call addPasswordTimesTime()
	Call timesTime()
	Call numbersToName()
	
	encode = hv1
End Function


Function nameToNumbers()
	hv2 = ""						'Turns hv2 into an empty string
	Redim hv3(Len(hv1)-1)			'Changes hv3 array size
		
	For i = 1 To Len(hv1)
		hv2 = Mid(hv1, i, 1)
		hv2 = Asc(hv2)
		Select Case true
			Case (hv2 > 47 And hv2 < 58)
				hv2 = hv2 - 48
			Case (hv2 > 96 And hv2 < 123)
				hv2 = hv2 - 87
			Case (hv2 > 64 And hv2 < 91)
				hv2 = hv2 - 29
		End Select
		hv3(i-1) = hv2
		If Len(hv3(i-1)) = 1 Then
			hv3(i-1) = 0 & hv3(i-1)
		End If
	Next
	hv2 = Len(hv1)
	hv1 = ""
	For i = 1 To hv2
		hv1 = hv1 & hv3(i-1)
	Next
End Function


Function numbersToName()
	Redim hv3((Len(hv1)/2)-1)
	For i = 1 To Len(hv1)-1 Step 2
		hv2 = Mid(hv1, i, 2)
		Select Case true
			Case (hv2 >= 0 And hv2 <= 9)
				hv2 = hv2 + 48
			Case (hv2 >= 10 And hv2 <= 35)
				hv2 = hv2 + 87
			Case (hv2 >= 36 And hv2 <= 61)
				hv2 = hv2 + 29
		End Select
		hv3((i-1)/2) = Chr(hv2)
	Next
	hv2 = Len(hv1)/2
	hv1 = ""
	For i = 1 To hv2
		hv1 = hv1 & hv3(i-1)
	Next
End Function


Function addPassword()
	Call add("password")
End Function


Function add(hv5)
	Redim hv3((Len(hv1)/2)-1)
	For i = 0 To UBound(hv3)
		hv3(i) = Mid(hv1, (i*2)+1, 2)
	Next
	Redim hv4(Len(password)-1)
	For i = 1 To Len(password)
		hv2 = Mid(password, i, 1)
		hv2 = Asc(hv2)
		Select Case true
			Case (hv2 > 47 And hv2 < 58)
				hv2 = hv2 - 48
			Case (hv2 > 96 And hv2 < 123)
				hv2 = hv2 - 87
			Case (hv2 > 64 And hv2 < 91)
				hv2 = hv2 - 29
		End Select
		hv4(i-1) = hv2
	Next
	If hv5 = "passwordTimesTime" Then
		Redim hv3(UBound(hv3))
		For i = 0 To UBound(hv3)
			hv3(i) = hv4(i Mod (UBound(hv4)+1))
		Next
		Redim hv4(UBound(hv3))
		For i = 0 To UBound(hv4)
			hv4(i) = hv3(i)
		Next
		For i = 0 To UBound(hv4)
			If (i+1) Mod Len(t) > 0 Then
				If Asc(Mid(t, ((i+1) Mod Len(t)), 1))-48 <> 0 Then
					hv4(i) = hv4(i) * (Asc(Mid(t, (i+1) Mod Len(t), 1))-48)
				Else
					hv4(i) = hv4(i) * 7
				End If
			Else
				If Asc(Mid(t, Len(t), 1))-48 <> 0 Then
					hv4(i) = hv4(i) * (Asc(Mid(t, Len(t), 1))-48)
				Else
					hv4(i) = hv4(i) * 7
				End If
			End If
		Next
		
		Redim hv3((Len(hv1)/2)-1)
		For i = 0 To UBound(hv3)
			hv3(i) = Mid(hv1, (i*2)+1, 2)
		Next
		For i = 0 To UBound(hv3)
			hv3(i) = hv3(i) + hv4(i)
			Do While hv3(i) > 61
				hv3(i) = hv3(i) - 62
			Loop
			If hv3(i) < 10 Then
				hv3(i) = 0 & hv3(i)
			End If
		Next
	ElseIf hv5 = "password" Then
		For i = 1 To (UBound(hv3)+1)
			If i Mod (UBound(hv4)+1) > 0 Then
				hv3(i-1) = hv3(i-1) + hv4((i-1) Mod (UBound(hv4)+1))
			Else
				hv3(i-1) = hv3(i-1) + hv4(UBound(hv4))
			End If
			Do While hv3(i-1) > 61
				hv3(i-1) = hv3(i-1) - 62
			Loop
			If hv3(i-1) < 10 Then
				hv3(i-1) = 0 & hv3(i-1)
			End If
		Next
	End If
	
	hv2 = Len(hv1)/2
	hv1 = ""
	For i = 1 to hv2
		hv1 = hv1 & hv3(i-1)
	Next
End Function


Function swap()
	hv2 = Replace(d, "0", number)
	Redim hv3(Len(hv2)-1)
	For i = 1 To Len(hv2)
		hv3(i-1) = Asc(Mid(hv2, i, 1))-48
		Do While hv3(i-1) > Len(hv1)
			hv3(i-1) = hv3(i-1) - Len(hv1)
		Loop
	Next
	hv2 = ""
	For i = 0 To UBound(hv3)
		hv2 = hv2 & hv3(i)
	Next
	Redim hv3(Len(hv1)-1)
	For i = 1 To Len(hv1)
		hv3(i-1) = Mid(hv1, i, 1)
	Next
	For i = 1 To Len(hv2)-1 Step 2
		hv1 = hv3(Mid(hv2, i, 1)-1)
		hv3(Mid(hv2, i, 1)-1) = hv3(Mid(hv2, i+1, 1)-1)
		hv3(Mid(hv2, i+1, 1)-1) = hv1
	Next
	hv1 = ""
	For i = 0 To UBound(hv3)
		hv1 = hv1 & hv3(i)
	Next
End Function


Function nameLengthChange()
	Do While UBound(hv3)+1 < 10
		Redim hv3(UBound(hv3)+number)
	Loop
	For i = 1 To UBound(hv3)+1
		If i Mod Len(hv1) > 0 Then
			hv3(i-1) = Mid(hv1, i Mod Len(hv1), 1)
		Else
			hv3(i-1) = Mid(hv1, Len(hv1), 1)
		End If
	Next
	hv1 = ""
	For i = 1 To UBound(hv3)+1
		hv1 = hv1 & hv3(i-1)
	Next
End Function


Function addNumber()
	Redim hv3((Len(hv1)/2)-1)
	For i = 1 To Len(hv1)-1 Step 2
		hv3((i-1)/2) = Mid(hv1, i, 2) + number
		Do While hv3((i-1)/2) > 61
			hv3((i-1)/2) = hv3((i-1)/2) - 62
		Loop
		If hv3((i-1)/2) < 10 Then
			hv3((i-1)/2) = 0 & hv3((i-1)/2)
		End If
	Next
	hv2 = Len(hv1)/2
	hv1 = ""
	For i = 1 to hv2
		hv1 = hv1 & hv3(i-1)
	Next
End Function


Function addPasswordTimesTime()
	Call add("passwordTimesTime")
End Function


Function timesTime()
	Redim hv3((Len(hv1)/2)-1)
	For i = 1 To Len(hv1)-1 Step 2
		hv3((i-1)/2) = Mid(hv1, i, 2)
	Next
	For i = 1 To UBound(hv3)+1
		If i Mod Len(t) > 0 Then
			hv2 = Mid(t, i Mod Len(t), 1)
		Else
			hv2 = Mid(t, Len(t), 1)
		End If
		hv2 = Asc(hv2) - 48
		If hv2 = 0 Then
			hv2 = 7
		End If
		hv3(i-1) = hv3(i-1) * hv2
		Do While hv3(i-1) > 61
			hv3(i-1) = hv3(i-1) - 62
		Loop
		If hv3(i-1) < 10 Then
			hv3(i-1) = 0 & hv3(i-1)
		End If
	Next
	hv1 = ""
	For i = 0 To UBound(hv3)
		hv1 = hv1 & hv3(i)
	Next
End Function


Function getName()
	name = inputbox(nameMsg,Title)
	name = Replace(name, Chr(34), "")
	For i = 1 To Len(name)
		hv1 = Asc(Mid(name, i, 1))
		Select Case false
			Case (hv1 > 47 And hv1 < 58) Or (hv1 > 96 And hv1 < 123) Or (hv1 > 64 And hv1 < 91) Or errorCode <> 0
				errorCode = 3
				Call error(errorCode)
		End Select
	Next
End Function


Function getDate()
	d = DATE
	If Asc(Mid(d, 2, 1)) = 46 Then
		d = "0" & d
	End If
	d = Replace(d, ".", "")
	d = Replace(d, " ", "")
	d = Replace(d, "/", "")
	If Len(d) = 7 Then
		d = Mid(d, 1, 2) & "0" & Mid(d, 3, 5)
	End If
End Function


Function getTime()
	t = Time
	If Mid(t,2,1) = ":" Then
		t = 0 & t
	End If
	If Mid(t,5,1) = ":" Then
		t = Mid(t,1,3) & 0 & Mid(t,4,Len(t)-3)
	End If
	t = Replace(t, ":", "")
	If Len(t) = 5 Then
		t = Mid(t,1,4) & 0 & Mid(t,5,1)
	End If
End Function


Function getNumber()
	min = 1
	max = 9
	Randomize
	number = Int((max-min+1)*Rnd()+min)
End Function


Function getPassword()
	password = inputbox(keyMsg,Title)
	For i = 1 To Len(password)
		hv1 = Asc(Mid(password, i, 1))
		Select Case false
			Case (hv1 > 47 And hv1 < 58) Or (hv1 > 96 And hv1 < 123) Or (hv1 > 64 And hv1 < 91) Or errorCode <> 0
				errorCode = 4
				Call error(errorCode)
		End Select
	Next
End Function


Function error(errorCode)
	If errorCode <> 0 Then
		errorText = Mid(errorText,1,6) & errorCode & Mid(errorText,7,Len(errorText)-6)
		Select Case errorCode
			Case 1
				errorText = Mid(errorText,1,Len(errorCode)+6) & ":" & Chr(13) & "Name not entered" & Mid(errorText,7+Len(errorCode),Len(errorText)-Len(errorCode)-6)
			Case 2
				errorText = Mid(errorText,1,Len(errorCode)+6) & ":" & Chr(13) & "Key not entered" & Mid(errorText,7+Len(errorCode),Len(errorText)-Len(errorCode)-6)
			Case 3
				errorText = Mid(errorText,1,Len(errorCode)+6) & ":" & Chr(13) & "Incorrect name format" & Mid(errorText,7+Len(errorCode),Len(errorText)-Len(errorCode)-6)
			Case 4
				errorText = Mid(errorText,1,Len(errorCode)+6) & ":" & Chr(13) & "Incorrect key format" & Mid(errorText,7+Len(errorCode),Len(errorText)-Len(errorCode)-6)
		End Select
		errorBox = Msgbox(errorText,53,Title)
	End If
End Function