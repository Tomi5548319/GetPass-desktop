Dim onlyOnce
onlyOnce = 1

	'If onlyOnce = 1 Then
	'	onlyOnce = 0
		
	'	Dim outputText
	'	outputText = ""
	
	'	For i = 0 To 15
	'		outputText = outputText & expandedKeys(i) & vbCrLf
	'	Next
	
	'	msgbox(outputText)
	'End If

version = "1.0"

'& Chr(34)
keyMsg = "Enter your access key"
nameMsg = "Enter the name"
Title = "GetPass v" & version & " by Tomi5548319"
prolog = "This"
useOldDataMsg = "Name"
saveNewDataMsg = "Do"

'TODO remove global variables
Dim name					'Website, app...
Dim hv1						'Used for checking, if name contains valid characters
Dim splitVersion			'1.0 -> array(1, 0)
Dim errorBox				'Return value of error message (restart/quit app)
Dim errorCode				'Code of the error which occured
Dim errorText				'String, which will be shown to user when an error occurs
Dim line					'Actual line in the text file
Dim lineNumber				'Line number in the text file, where the name can be found
Dim fso						'File System Object
Dim f						'File
Dim nameFound				'Variable, which is true if the name was found in the text file

Dim length
Dim smallLetters
Dim bigLetters
Dim numbers
Dim basicChars
Dim advancedChars
Dim customChars				'Saved as ASCII codes, I.E 'a' is saved as 97
Dim seed					'Saved as ASCII codes, I.E 'a' is saved as 97

Dim str						'String
Dim accessKey				'String
Dim password				'String

Dim nameCharArray			'char[]
Dim keyCharArray			'char[]
Dim seedCharArray			'char[]
Dim customCharArray			'char[]

Dim ACTUAL_ALPHABET			'int[]
Dim actualAlphabetLength	'int
Dim SMALL_ALPHABET
Dim BIG_ALPHABET
Dim NUMBERS_ARRAY
Dim BASIC_CHARS
Dim ADVANCED_CHARS

ACTUAL_ALPHABET = array()											'10													'20						'25
SMALL_ALPHABET = array(97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122) '25
BIG_ALPHABET = array(65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90) '51
NUMBERS_ARRAY = array(48, 49, 50, 51, 52, 53, 54, 55, 56, 57) '61
BASIC_CHARS = array(46, 44, 45, 63, 58, 95, 33, 47, 59, 42, 43, 64, 40, 41, 37) '76
ADVANCED_CHARS = array(124, 36, 91, 93, 123, 125) '82

Dim SUB_TABLE
Dim RCON
Dim MUL2
Dim MUL3

SUB_TABLE = array(&H63,	&H7c, &H77, &H7b, &Hf2, &H6b, &H6f, &Hc5, &H30, &H01, &H67, &H2b, &Hfe, &Hd7, &Hab, &H76, &Hca, &H82, &Hc9, &H7d, &Hfa, &H59, &H47, &Hf0, &Had, &Hd4, &Ha2, &Haf, &H9c, &Ha4, &H72, &Hc0, &Hb7, &Hfd, &H93, &H26, &H36, &H3f, &Hf7, &Hcc, &H34, &Ha5, &He5, &Hf1, &H71, &Hd8, &H31, &H15, &H04, &Hc7, &H23, &Hc3, &H18, &H96, &H05, &H9a, &H07, &H12, &H80, &He2, &Heb, &H27, &Hb2, &H75, &H09, &H83, &H2c, &H1a, &H1b, &H6e, &H5a, &Ha0, &H52, &H3b, &Hd6, &Hb3, &H29, &He3, &H2f, &H84, &H53, &Hd1, &H00, &Hed, &H20, &Hfc, &Hb1, &H5b, &H6a, &Hcb, &Hbe, &H39, &H4a, &H4c, &H58, &Hcf, &Hd0, &Hef, &Haa, &Hfb, &H43, &H4d, &H33, &H85, &H45, &Hf9, &H02, &H7f, &H50, &H3c, &H9f, &Ha8, &H51, &Ha3, &H40, &H8f, &H92, &H9d, &H38, &Hf5, &Hbc, &Hb6, &Hda, &H21, &H10, &Hff, &Hf3, &Hd2, &Hcd, &H0c, &H13, &Hec, &H5f, &H97, &H44, &H17, &Hc4, &Ha7, &H7e, &H3d, &H64, &H5d, &H19, &H73, &H60, &H81, &H4f, &Hdc, &H22, &H2a, &H90, &H88, &H46, &Hee, &Hb8, &H14, &Hde, &H5e, &H0b, &Hdb, &He0, &H32, &H3a, &H0a, &H49, &H06, &H24, &H5c, &Hc2, &Hd3, &Hac, &H62, &H91, &H95, &He4, &H79, &He7, &Hc8, &H37, &H6d, &H8d, &Hd5, &H4e, &Ha9, &H6c, &H56, &Hf4, &Hea, &H65, &H7a, &Hae, &H08, &Hba, &H78, &H25, &H2e, &H1c, &Ha6, &Hb4, &Hc6, &He8, &Hdd, &H74, &H1f, &H4b, &Hbd, &H8b, &H8a, &H70, &H3e, &Hb5, &H66, &H48, &H03, &Hf6, &H0e, &H61, &H35, &H57, &Hb9, &H86, &Hc1, &H1d, &H9e, &He1, &Hf8, &H98, &H11, &H69, &Hd9, &H8e, &H94, &H9b, &H1e, &H87, &He9, &Hce, &H55, &H28, &Hdf, &H8c, &Ha1, &H89, &H0d, &Hbf, &He6, &H42, &H68, &H41, &H99, &H2d, &H0f, &Hb0, &H54, &Hbb, &H16)
RCON = array(&H8d, &H01, &H02, &H04, &H08, &H10, &H20, &H40, &H80, &H1b, &H36, &H6c, &Hd8, &Hab, &H4d, &H9a, &H2f, &H5e, &Hbc, &H63, &Hc6, &H97, &H35, &H6a, &Hd4, &Hb3, &H7d, &Hfa, &Hef, &Hc5, &H91, &H39, &H72, &He4, &Hd3, &Hbd, &H61, &Hc2, &H9f, &H25, &H4a, &H94, &H33, &H66, &Hcc, &H83, &H1d, &H3a, &H74, &He8, &Hcb, &H8d, &H01, &H02, &H04, &H08, &H10, &H20, &H40, &H80, &H1b, &H36, &H6c, &Hd8, &Hab, &H4d, &H9a, &H2f, &H5e, &Hbc, &H63, &Hc6, &H97, &H35, &H6a, &Hd4, &Hb3, &H7d, &Hfa, &Hef, &Hc5, &H91, &H39, &H72, &He4, &Hd3, &Hbd, &H61, &Hc2, &H9f, &H25, &H4a, &H94, &H33, &H66, &Hcc, &H83, &H1d, &H3a, &H74, &He8, &Hcb, &H8d, &H01, &H02, &H04, &H08, &H10, &H20, &H40, &H80, &H1b, &H36, &H6c, &Hd8, &Hab, &H4d, &H9a, &H2f, &H5e, &Hbc, &H63, &Hc6, &H97, &H35, &H6a, &Hd4, &Hb3, &H7d, &Hfa, &Hef, &Hc5, &H91, &H39, &H72, &He4, &Hd3, &Hbd, &H61, &Hc2, &H9f, &H25, &H4a, &H94, &H33, &H66, &Hcc, &H83, &H1d, &H3a, &H74, &He8, &Hcb, &H8d, &H01, &H02, &H04, &H08, &H10, &H20, &H40, &H80, &H1b, &H36, &H6c, &Hd8, &Hab, &H4d, &H9a, &H2f, &H5e, &Hbc, &H63, &Hc6, &H97, &H35, &H6a, &Hd4, &Hb3, &H7d, &Hfa, &Hef, &Hc5, &H91, &H39, &H72, &He4, &Hd3, &Hbd, &H61, &Hc2, &H9f, &H25, &H4a, &H94, &H33, &H66, &Hcc, &H83, &H1d, &H3a, &H74, &He8, &Hcb, &H8d, &H01, &H02, &H04, &H08, &H10, &H20, &H40, &H80, &H1b, &H36, &H6c, &Hd8, &Hab, &H4d, &H9a, &H2f, &H5e, &Hbc, &H63, &Hc6, &H97, &H35, &H6a, &Hd4, &Hb3, &H7d, &Hfa, &Hef, &Hc5, &H91, &H39, &H72, &He4, &Hd3, &Hbd, &H61, &Hc2, &H9f, &H25, &H4a, &H94, &H33, &H66, &Hcc, &H83, &H1d, &H3a, &H74, &He8, &Hcb, &H8d)
MUL2 = array(&H00,&H02,&H04,&H06,&H08,&H0a,&H0c,&H0e,&H10,&H12,&H14,&H16,&H18,&H1a,&H1c,&H1e,&H20,&H22,&H24,&H26,&H28,&H2a,&H2c,&H2e,&H30,&H32,&H34,&H36,&H38,&H3a,&H3c,&H3e,&H40,&H42,&H44,&H46,&H48,&H4a,&H4c,&H4e,&H50,&H52,&H54,&H56,&H58,&H5a,&H5c,&H5e,&H60,&H62,&H64,&H66,&H68,&H6a,&H6c,&H6e,&H70,&H72,&H74,&H76,&H78,&H7a,&H7c,&H7e,&H80,&H82,&H84,&H86,&H88,&H8a,&H8c,&H8e,&H90,&H92,&H94,&H96,&H98,&H9a,&H9c,&H9e,&Ha0,&Ha2,&Ha4,&Ha6,&Ha8,&Haa,&Hac,&Hae,&Hb0,&Hb2,&Hb4,&Hb6,&Hb8,&Hba,&Hbc,&Hbe,&Hc0,&Hc2,&Hc4,&Hc6,&Hc8,&Hca,&Hcc,&Hce,&Hd0,&Hd2,&Hd4,&Hd6,&Hd8,&Hda,&Hdc,&Hde,&He0,&He2,&He4,&He6,&He8,&Hea,&Hec,&Hee,&Hf0,&Hf2,&Hf4,&Hf6,&Hf8,&Hfa,&Hfc,&Hfe,&H1b,&H19,&H1f,&H1d,&H13,&H11,&H17,&H15,&H0b,&H09,&H0f,&H0d,&H03,&H01,&H07,&H05,&H3b,&H39,&H3f,&H3d,&H33,&H31,&H37,&H35,&H2b,&H29,&H2f,&H2d,&H23,&H21,&H27,&H25,&H5b,&H59,&H5f,&H5d,&H53,&H51,&H57,&H55,&H4b,&H49,&H4f,&H4d,&H43,&H41,&H47,&H45,&H7b,&H79,&H7f,&H7d,&H73,&H71,&H77,&H75,&H6b,&H69,&H6f,&H6d,&H63,&H61,&H67,&H65,&H9b,&H99,&H9f,&H9d,&H93,&H91,&H97,&H95,&H8b,&H89,&H8f,&H8d,&H83,&H81,&H87,&H85,&Hbb,&Hb9,&Hbf,&Hbd,&Hb3,&Hb1,&Hb7,&Hb5,&Hab,&Ha9,&Haf,&Had,&Ha3,&Ha1,&Ha7,&Ha5,&Hdb,&Hd9,&Hdf,&Hdd,&Hd3,&Hd1,&Hd7,&Hd5,&Hcb,&Hc9,&Hcf,&Hcd,&Hc3,&Hc1,&Hc7,&Hc5,&Hfb,&Hf9,&Hff,&Hfd,&Hf3,&Hf1,&Hf7,&Hf5,&Heb,&He9,&Hef,&Hed,&He3,&He1,&He7,&He5)
MUL3 = array(&H00,&H03,&H06,&H05,&H0c,&H0f,&H0a,&H09,&H18,&H1b,&H1e,&H1d,&H14,&H17,&H12,&H11,&H30,&H33,&H36,&H35,&H3c,&H3f,&H3a,&H39,&H28,&H2b,&H2e,&H2d,&H24,&H27,&H22,&H21,&H60,&H63,&H66,&H65,&H6c,&H6f,&H6a,&H69,&H78,&H7b,&H7e,&H7d,&H74,&H77,&H72,&H71,&H50,&H53,&H56,&H55,&H5c,&H5f,&H5a,&H59,&H48,&H4b,&H4e,&H4d,&H44,&H47,&H42,&H41,&Hc0,&Hc3,&Hc6,&Hc5,&Hcc,&Hcf,&Hca,&Hc9,&Hd8,&Hdb,&Hde,&Hdd,&Hd4,&Hd7,&Hd2,&Hd1,&Hf0,&Hf3,&Hf6,&Hf5,&Hfc,&Hff,&Hfa,&Hf9,&He8,&Heb,&Hee,&Hed,&He4,&He7,&He2,&He1,&Ha0,&Ha3,&Ha6,&Ha5,&Hac,&Haf,&Haa,&Ha9,&Hb8,&Hbb,&Hbe,&Hbd,&Hb4,&Hb7,&Hb2,&Hb1,&H90,&H93,&H96,&H95,&H9c,&H9f,&H9a,&H99,&H88,&H8b,&H8e,&H8d,&H84,&H87,&H82,&H81,&H9b,&H98,&H9d,&H9e,&H97,&H94,&H91,&H92,&H83,&H80,&H85,&H86,&H8f,&H8c,&H89,&H8a,&Hab,&Ha8,&Had,&Hae,&Ha7,&Ha4,&Ha1,&Ha2,&Hb3,&Hb0,&Hb5,&Hb6,&Hbf,&Hbc,&Hb9,&Hba,&Hfb,&Hf8,&Hfd,&Hfe,&Hf7,&Hf4,&Hf1,&Hf2,&He3,&He0,&He5,&He6,&Hef,&Hec,&He9,&Hea,&Hcb,&Hc8,&Hcd,&Hce,&Hc7,&Hc4,&Hc1,&Hc2,&Hd3,&Hd0,&Hd5,&Hd6,&Hdf,&Hdc,&Hd9,&Hda,&H5b,&H58,&H5d,&H5e,&H57,&H54,&H51,&H52,&H43,&H40,&H45,&H46,&H4f,&H4c,&H49,&H4a,&H6b,&H68,&H6d,&H6e,&H67,&H64,&H61,&H62,&H73,&H70,&H75,&H76,&H7f,&H7c,&H79,&H7a,&H3b,&H38,&H3d,&H3e,&H37,&H34,&H31,&H32,&H23,&H20,&H25,&H26,&H2f,&H2c,&H29,&H2a,&H0b,&H08,&H0d,&H0e,&H07,&H04,&H01,&H02,&H13,&H10,&H15,&H16,&H1f,&H1c,&H19,&H1a)

'Set file system object
splitVersion = Split(version, ".")
fileName = "getpassv" & splitVersion(0) & splitVersion(1) & ".txt"
Set fso = CreateObject("Scripting.FileSystemObject")

Do
	errorText= "Error " & Chr(13) & Chr(13) & "Do you want to restart the app or exit?"
	errorBox = 0
	nameFound = False
	
	Call getName()
	
	If Len(name) > 0 Then
		Call readAndRun()
	Else
		errorCode = 1
		Call error(errorCode)
	End If
	
Loop While errorBox = vbRetry 'App restart required

'TODO check later
Function getName()
	name = inputbox(nameMsg,Title)
	name = Replace(name, Chr(34), "") 'Deletes double quotes (") from the name
	For i = 1 To Len(name)
		hv1 = Asc(Mid(name, i, 1))
		Select Case false
			Case (hv1 > 47 And hv1 < 58) Or (hv1 > 96 And hv1 < 123) Or (hv1 > 64 And hv1 < 91) Or errorCode <> 0
				errorCode = 3
				Call error(errorCode)
		End Select
	Next
End Function


Function readAndRun()
	Set f = fso.OpenTextFile(fileName,1,True)
	Do Until f.AtEndOfStream
		Line = f.ReadLine
		If Line = name Then
			nameFound = True
			Exit Do
		Else
			For i=1 To 9
				Line = f.ReadLine
				lineNumber = lineNumber + 1
			Next
		End If 
		lineNumber = lineNumber + 1
	Loop
	f.Close
	
	If nameFound = True Then
		Set f = fso.OpenTextFile(fileName,1,True) '(filename, [ iomode, [ create, [ format ]]]) 1=ForReading
		str = f.ReadAll 'Read the whole file and save it as a string
		f.Close
		str = Split(str, vbCrlf) '"Name true true..." => array("Name", "true", "true")
		
		If errorCode = 0 Then
			length = str(LineNumber + 1)
			smallLetters = str(LineNumber + 2)
			bigLetters = str(LineNumber + 3)
			numbers = str(LineNumber + 4)
			basicChars = str(LineNumber + 5)
			advancedChars = str(LineNumber + 6)
			customChars = str(LineNumber + 7)
			seed = str(LineNumber + 8)
			Call getPassword()
			
			'TODO rename password to key
			If Len(accessKey) > 0 And errorCode = 0 Then 'Key was entered
				
				Call encode()
				
				'create a new text file and put the password there
				Set f = fso.OpenTextFile("clip.txt",2,True)
				f.Write password
				f.Close
				
				'read the text file and save it to the clipboard using Windows cmd
				winCommand = "cmd /K type clip.txt | clip & exit"
				
				Dim oShell
				Set oShell = WScript.CreateObject ("WScript.Shell")
				oShell.run winCommand, 7, true
				Set oShell = Nothing
				
				'delete the text file
				fso.DeleteFile "clip.txt", true
				
				msgbox("Password copied to clipboard")
				
				'inputbox "Password", Title, password
				
			ElseIf errorCode = 0 Then 'Key was not entered
				errorCode = 2
				Call error(errorCode)
			End If
		End If
	Else 'Name was not found
	
	End If
End Function

'TODO check later
Function getPassword()
	accessKey = inputbox(keyMsg,Title)
	For i = 1 To Len(accessKey)
		hv1 = Asc(Mid(accessKey, i, 1))
		Select Case false
			Case (hv1 > 47 And hv1 < 58) Or (hv1 > 96 And hv1 < 123) Or (hv1 > 64 And hv1 < 91) Or errorCode <> 0
				errorCode = 4
				Call error(errorCode)
		End Select
	Next
End Function


Function encode()
	
	Dim i
	nameCharArray = modify(name)
	keyCharArray = modify(accessKey)
	seedCharArray = Split(seed, ",")
	
	'nameCharArray, keyCharArray, seedCharArray = array(132, 220...)
	
	'ReDim Preserve nameCharArray(15)
	'ReDim Preserve keyCharArray(15)
	'ReDim Preserve seedCharArray(15)

	nameCharArray = AES_Encrypt(nameCharArray, keyCharArray)
	nameCharArray = AES_Encrypt(nameCharArray, seedCharArray)
	
	Call changeAlphabet()
	
	For i = 1 To length
		nameCharArray(i-1) = ACTUAL_ALPHABET(nameCharArray(i-1) Mod actualAlphabetLength)
	Next
	
	password = ""
	
	For i = 0 To length-1
		password = password & Chr(nameCharArray(i))
	Next
End Function


Function modify(textStream) 'name / accessKey
	
	charLength = Len(textStream)
	Dim charArray()
	Redim Preserve charArray(charLength - 1)
	
	For i = 0 To charLength - 1
		charArray(i) = Asc(Mid(textStream, i + 1 ,1))
	Next
	
	Dim textLength
	textLength = length
	textLength = modifyLength(textLength) 'length 13 => 16
	msgbox(length)
	
	i = 0
	
	For charLength = Len(textStream) To textLength-1
		
		ReDim Preserve charArray(charLength)	
		charArray(charLength) = charArray(i)
			
		i = i + 1
	Next
	
	modify = charArray
End Function


Function modifyLength(textLength)

	If textLength Mod 16 <> 0 Then
		
		i = 0
		While textLength Mod 16 <> 0
			textLength = textLength-1
			i = i + 1
		Wend
		
		textLength = textLength + 16
		
	End If
	
	modifyLength = textLength
End Function


Function AES_Encrypt(c_name, c_key)
	
	'msgbox("name" & vbCrLf & c_name(0) & vbCrLf & c_name(1) & vbCrLf & c_name(2) & vbCrLf & c_name(3) & vbCrLf & c_name(4) & vbCrLf & c_name(5) & vbCrLf & c_name(6) & vbCrLf & c_name(7) & vbCrLf & c_name(8) & vbCrLf & c_name(9) & vbCrLf & c_name(10) & vbCrLf & c_name(11) & vbCrLf & c_name(12) & vbCrLf & c_name(13) & vbCrLf & c_name(14) & vbCrLf & c_name(15))
	
	numberOfRounds = 9
	
	For j = 0 To ((UBound(c_name) + 1) / 16) - 1 'Encode 128 byte blocks
		Dim temp
		Dim tempNew
		
		ReDim temp(15)
		
		For i = 0 To 15
			temp(i) = c_name(j * 16 + i)
		Next
		
		'Expand the keys
		expandedKey = KeyExpansion(array(c_key(j*16), c_key(j*16+1), c_key(j*16+2), c_key(j*16+3), c_key(j*16+4), c_key(j*16+5), c_key(j*16+6), c_key(j*16+7), c_key(j*16+8), c_key(j*16+9), c_key(j*16+10), c_key(j*16+11), c_key(j*16+12), c_key(j*16+13), c_key(j*16+14), c_key(j*16+15)))
		
		'Initial round
		tempNew = AddRoundKey(temp, array(c_key(j*16), c_key(j*16+1), c_key(j*16+2), c_key(j*16+3), c_key(j*16+4), c_key(j*16+5), c_key(j*16+6), c_key(j*16+7), c_key(j*16+8), c_key(j*16+9), c_key(j*16+10), c_key(j*16+11), c_key(j*16+12), c_key(j*16+13), c_key(j*16+14), c_key(j*16+15)))
		
		'Rounds
		For i = 1 To numberOfRounds
			tempNew = SubBytes(tempNew)
			tempNew = ShiftRows(tempNew)
			tempNew = MixColumns(tempNew)
			tempNew = AddRoundKey(tempNew, array(expandedKey(16*i), expandedKey(16*i+1), expandedKey(16*i+2), expandedKey(16*i+3), expandedKey(16*i+4), expandedKey(16*i+5), expandedKey(16*i+6), expandedKey(16*i+7), expandedKey(16*i+8), expandedKey(16*i+9), expandedKey(16*i+10), expandedKey(16*i+11), expandedKey(16*i+12), expandedKey(16*i+13), expandedKey(16*i+14), expandedKey(16*i+15)))
		Next
		
		'Final round
		tempNew = SubBytes(tempNew)
		tempNew = ShiftRows(tempNew)
		tempNew = AddRoundKey(tempNew, array(expandedKey(160), expandedKey(161), expandedKey(162), expandedKey(163), expandedKey(164), expandedKey(165), expandedKey(166), expandedKey(167), expandedKey(168), expandedKey(169), expandedKey(170), expandedKey(171), expandedKey(172), expandedKey(173), expandedKey(174), expandedKey(175)))
		
		For i = 0 To UBound(tempNew)
			c_name(j * 16 + i) = tempNew(i)
		Next
	Next
	
	AES_Encrypt = c_name 'return value of this function
End Function

'TODO refaktoring needed
Function KeyExpansion(c_key)
	
	Dim expandedKeys(175)
	
	'The first 16 bytes are the original key
	For i = 0 To 15
		expandedKeys(i) = c_key(i)
	Next

    bytesGenerated = 16
	rconIteration = 1
	Dim temp(3)
	Dim prev(3)
	Dim tempNew
	
	
	While bytesGenerated < 176
		
		'Read 4 bytes for the core
		For i = 0 To 3
			temp(i) = expandedKeys(bytesGenerated - 4 + i)
			prev(i) = expandedKeys(bytesGenerated - 16 + i)
		Next
		
		'Perform the core once for each 16 byte key
		If bytesGenerated Mod 16 = 0 Then
			
			tempNew = KeyExpansionCore(temp, prev, rconIteration)
			
			rconIteration = rconIteration + 1
		Else
			tempNew = KeyExpansionCoreTwo(temp, prev)
		End If
		
		'Save the generated bytes (4 bytes)
		expandedKeys(bytesGenerated) = tempNew(0)
		expandedKeys(bytesGenerated + 1) = tempNew(1)
		expandedKeys(bytesGenerated + 2) = tempNew(2)
		expandedKeys(bytesGenerated + 3) = tempNew(3)
		
		bytesGenerated = bytesGenerated + 4
	Wend
	
	KeyExpansion = expandedKeys
End Function


Function KeyExpansionCore(temp, prev, rconIteration)
	
	'Rotate left
	h = temp(0)
	temp(0) = temp(1)
	temp(1) = temp(2)
	temp(2) = temp(3)
	temp(3) = h
	
	'S-Box four bytes
	temp(0) = SUB_TABLE(temp(0))
	temp(1) = SUB_TABLE(temp(1))
	temp(2) = SUB_TABLE(temp(2))
	temp(3) = SUB_TABLE(temp(3))
	
	'RCon
    temp(0) = temp(0) Xor RCON(rconIteration) Xor prev(0)
	temp(1) = temp(1) Xor prev(1)
	temp(2) = temp(2) Xor prev(2)
	temp(3) = temp(3) Xor prev(3)
	
	KeyExpansionCore = temp
End Function


Function KeyExpansionCoreTwo(temp, prev)
	'Xor
    temp(0) = temp(0) Xor prev(0)
	temp(1) = temp(1) Xor prev(1)
	temp(2) = temp(2) Xor prev(2)
	temp(3) = temp(3) Xor prev(3)
	
	KeyExpansionCoreTwo = temp
End Function


Function AddRoundKey(c_name, c_key)
	
	For i = 0 To 15
		c_name(i) = c_name(i) Xor c_key(i)
	Next
	
	AddRoundKey = c_name
End Function


Function SubBytes(c_name)
	
	For i = 0 To 15
		c_name(i) = SUB_TABLE(c_name(i))
	Next
	
	SubBytes = c_name
End Function


Function ShiftRows(c_name)
	
	Dim tmp(15)
	
	tmp(0) = c_name(0)
	tmp(1) = c_name(5)
	tmp(2) = c_name(10)
	tmp(3) = c_name(15)
	
	tmp(4) = c_name(4)
	tmp(5) = c_name(9)
	tmp(6) = c_name(14)
	tmp(7) = c_name(3)
	
	tmp(8) = c_name(8)
	tmp(9) = c_name(13)
	tmp(10) = c_name(2)
	tmp(11) = c_name(7)
	
	tmp(12) = c_name(12)
	tmp(13) = c_name(1)
	tmp(14) = c_name(6)
	tmp(15) = c_name(11)
	
	For i = 0 To 15
		c_name(i) = tmp(i)
	Next
	
	ShiftRows = c_name
End Function


Function MixColumns(c_name)
	
	Dim tmp(15)
	
	tmp(0) = MUL2(c_name(0)) Xor MUL3(c_name(1)) Xor c_name(2) Xor c_name(3)
	tmp(1) = c_name(0) Xor MUL2(c_name(1)) Xor MUL3(c_name(2)) Xor c_name(3)
	tmp(2) = c_name(0) Xor c_name(1) Xor MUL2(c_name(2)) Xor MUL3(c_name(3))
	tmp(3) = MUL3(c_name(0)) Xor c_name(1) Xor c_name(2) Xor MUL2(c_name(3))
	
	tmp(4) = MUL2(c_name(4)) Xor MUL3(c_name(5)) Xor c_name(6) Xor c_name(7)
	tmp(5) = c_name(4) Xor MUL2(c_name(5)) Xor MUL3(c_name(6)) Xor c_name(7)
	tmp(6) = c_name(4) Xor c_name(5) Xor MUL2(c_name(6)) Xor MUL3(c_name(7))
	tmp(7) = MUL3(c_name(4)) Xor c_name(5) Xor c_name(6) Xor MUL2(c_name(7))
	
	tmp(8) = MUL2(c_name(8)) Xor MUL3(c_name(9)) Xor c_name(10) Xor c_name(11)
	tmp(9) = c_name(8) Xor MUL2(c_name(9)) Xor MUL3(c_name(10)) Xor c_name(11)
	tmp(10) = c_name(8) Xor c_name(9) Xor MUL2(c_name(10)) Xor MUL3(c_name(11))
	tmp(11) = MUL3(c_name(8)) Xor c_name(9) Xor c_name(10) Xor MUL2(c_name(11))
	
	tmp(12) = MUL2(c_name(12)) Xor MUL3(c_name(13)) Xor c_name(14) Xor c_name(15)
	tmp(13) = c_name(12) Xor MUL2(c_name(13)) Xor MUL3(c_name(14)) Xor c_name(15)
	tmp(14) = c_name(12) Xor c_name(13) Xor MUL2(c_name(14)) Xor MUL3(c_name(15))
	tmp(15) = MUL3(c_name(12)) Xor c_name(13) Xor c_name(14) Xor MUL2(c_name(15))
	
	For i = 0 To 15
		c_name(i) = tmp(i)
	Next
	
	MixColumns = c_name
End Function


Function changeAlphabet()
	actualAlphabetLength = 0
	
	If smallLetters = 1 Then
		ReDim Preserve ACTUAL_ALPHABET(UBound(ACTUAL_ALPHABET) + UBound(SMALL_ALPHABET) + 1)
		
		For i = actualAlphabetLength To UBound(ACTUAL_ALPHABET)
			ACTUAL_ALPHABET(i) = SMALL_ALPHABET(i - actualAlphabetLength)
		Next
		
		actualAlphabetLength = UBound(ACTUAL_ALPHABET) + 1
	End If
	
	If bigLetters = 1 Then
		ReDim Preserve ACTUAL_ALPHABET(UBound(ACTUAL_ALPHABET) + UBound(BIG_ALPHABET) + 1)
		
		For i = actualAlphabetLength To UBound(ACTUAL_ALPHABET)
			ACTUAL_ALPHABET(i) = BIG_ALPHABET(i - actualAlphabetLength)
		Next
		
		actualAlphabetLength = UBound(ACTUAL_ALPHABET) + 1
	End If
	
	If numbers = 1 Then
		ReDim Preserve ACTUAL_ALPHABET(UBound(ACTUAL_ALPHABET) + UBound(NUMBERS_ARRAY) + 1)
		
		For i = actualAlphabetLength To UBound(ACTUAL_ALPHABET)
			ACTUAL_ALPHABET(i) = NUMBERS_ARRAY(i - actualAlphabetLength)
		Next
		
		actualAlphabetLength = UBound(ACTUAL_ALPHABET) + 1
	End If
	
	If basicChars = 1 Then
		ReDim Preserve ACTUAL_ALPHABET(UBound(ACTUAL_ALPHABET) + UBound(BASIC_CHARS) + 1)
		
		For i = actualAlphabetLength To UBound(ACTUAL_ALPHABET)
			ACTUAL_ALPHABET(i) = BASIC_CHARS(i - actualAlphabetLength)
		Next
		
		actualAlphabetLength = UBound(ACTUAL_ALPHABET) + 1
	End If
	
	If advancedChars = 1 Then
		ReDim Preserve ACTUAL_ALPHABET(UBound(ACTUAL_ALPHABET) + UBound(ADVANCED_CHARS) + 1)
		
		For i = actualAlphabetLength To UBound(ACTUAL_ALPHABET)
			ACTUAL_ALPHABET(i) = ADVANCED_CHARS(i - actualAlphabetLength)
		Next
		
		actualAlphabetLength = UBound(ACTUAL_ALPHABET) + 1
	End If
	
	If customchars <> "" Then
		customCharArray = Split(customChars, ",")
		ReDim Preserve ACTUAL_ALPHABET(UBound(ACTUAL_ALPHABET) + UBound(customCharArray) + 1)
		
		For i = actualAlphabetLength To UBound(ACTUAL_ALPHABET)
			ACTUAL_ALPHABET(i) = customCharArray(i - actualAlphabetLength)
		Next
		
		actualAlphabetLength = UBound(ACTUAL_ALPHABET) + 1
	End If
		
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