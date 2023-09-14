<%
Function getSalt(intLen)
' Function takes a given length x and generates a random hex value of x digits.
' Salt can be used to help protect passwords.  When a password is first stored in a
' database generate a salt value also.  Concatenate the salt value with the password, 
' and then encrypt it using the HashEncode function below.  Store both the salt value,
' and the encrypted value in the database.  When a password needs to be verified, take 
' the password concatenate the salt from the database.  Encode it using the HashEncode 
' function below.  If the result matches the the encrypted password stored in the
' database, then it is a match.  If not then the password is invalid.
'
'
' Note: Passwords become case sensitive when using this encryption.
' For more information on Password HASH Encoding, and SALT visit: http://local.15seconds.com/issue/000217.htm
'
' Call this function if you wish to generate a random hex value of any given length
'
' Written By: Mark G. Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact

	Dim strSalt
	Dim intIndex, intRand

	If Not IsNumeric(intLen) Then
		getSalt = "00000000"
		exit function
	ElseIf CInt(intLen) <> CDbl(intLen) Or CInt(intLen) < 1 Then
		getSalt = "00000000"
		exit function
	End If

	Randomize

	For intIndex = 1 to CInt(intLen)
		intRand = CInt(Rnd * 1000) Mod 16
		strSalt = strSalt & getDecHex(intRand)
	Next
	
	getSalt = strSalt

End Function


Function HashEncode(strSecret)
' Function takes an ASCII string less than 2^61 characters long and 
' one way hash encrypts it using 160 bit encryption into a 40 digit hex value.
' The encoded hex value cannot be decoded to the original string value.
'
' This is the only function that you need to call for encryption.
'
' Written By: Mark G. Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'
' The author makes no warranties as to the validity, and/or authenticity of this code.
' You may use any code found herein at your own risk.
' This code was written to follow as closely as possible the standards found in
' Federal Information Processing Standards Publication (FIPS PUB 180-1)
' http://csrc.nist.gov/fips/fip180-1.txt -- Secure Hash Standard SHA-1
'
' This code is for private use only, and the security and/or encryption of the resulting
' hexadecimal value is not warrented or gaurenteed in any way.
'
    Dim strEncode, strH(4)
    Dim intPos
    
    
    If len(strSecret) = 0 or len(strSecret) >= 2^61 then
		HashEncode = "0000000000000000000000000000000000000000"
		exit function
    end if
    
    
    'Initial Hex words are used for encoding Digest.  
    'These can be any valid 8-digit hex value (0 to F)
    strH(0) = "FB0C14C2"
    strH(1) = "9F00AB2E"
    strH(2) = "991FFA67"
    strH(3) = "76FA2C3F"
    strH(4) = "ADE426FA"
    
    For intPos = 1 to len(strSecret) step 56
		
		strEncode = Mid(strSecret, intPos, 56) 'get 56 character chunks
		strEncode = WordToBinary(strEncode) 'convert to binary
		strEncode = PadBinary(strEncode) 'make it 512 bites
		strEncode = BlockToHex(strEncode) 'convert to hex value
		
		'Encode the hex value using the previous runs digest
		'If it is the first run then use the initial values above
		strEncode = DigestHex(strEncode, strH(0), strH(1), strH(2), strH(3), strH(4))

		'Combine the old digest with the new digest
		strH(0) = HexAdd(left(strEncode, 8), strH(0))
		strH(1) = HexAdd(mid(strEncode, 9, 8), strH(1))
		strH(2) = HexAdd(mid(strEncode, 17, 8), strH(2))
		strH(3) = HexAdd(mid(strEncode, 25, 8), strH(3))
		strH(4) = HexAdd(right(strEncode, 8), strH(4))
		
    Next
    
    'This is the final Hex Digest
    HashEncode = strH(0) & strH(1) & strH(2) & strH(3) & strH(4)
    
End Function



Function HexToBinary(btHex)

' Function Converts a single hex value into it's binary equivalent
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'

    Select Case btHex
    Case "0"
        HexToBinary = "0000"
    Case "1"
        HexToBinary = "0001"
    Case "2"
        HexToBinary = "0010"
    Case "3"
        HexToBinary = "0011"
    Case "4"
        HexToBinary = "0100"
    Case "5"
        HexToBinary = "0101"
    Case "6"
        HexToBinary = "0110"
    Case "7"
        HexToBinary = "0111"
    Case "8"
        HexToBinary = "1000"
    Case "9"
        HexToBinary = "1001"
    Case "A"
        HexToBinary = "1010"
    Case "B"
        HexToBinary = "1011"
    Case "C"
        HexToBinary = "1100"
    Case "D"
        HexToBinary = "1101"
    Case "E"
        HexToBinary = "1110"
    Case "F"
        HexToBinary = "1111"
    Case Else
        HexToBinary = "2222"
    End Select
End Function

Function BinaryToHex(strBinary)

' Function Converts a 4 bit binary value into it's hex equivalent
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'
    Select Case strBinary
    Case "0000"
        BinaryToHex = "0"
    Case "0001"
        BinaryToHex = "1"
    Case "0010"
        BinaryToHex = "2"
    Case "0011"
        BinaryToHex = "3"
    Case "0100"
        BinaryToHex = "4"
    Case "0101"
        BinaryToHex = "5"
    Case "0110"
        BinaryToHex = "6"
    Case "0111"
        BinaryToHex = "7"
    Case "1000"
        BinaryToHex = "8"
    Case "1001"
        BinaryToHex = "9"
    Case "1010"
        BinaryToHex = "A"
    Case "1011"
        BinaryToHex = "B"
    Case "1100"
        BinaryToHex = "C"
    Case "1101"
        BinaryToHex = "D"
    Case "1110"
        BinaryToHex = "E"
    Case "1111"
        BinaryToHex = "F"
    Case Else
        BinaryToHex = "Z"
    End Select
End Function

Function WordToBinary(strWord)
' Function Converts a 8 digit hex value into it's 32 bit binary equivalent
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header kept intact
'
	Dim strTemp, strBinary 
	Dim intPos

	For intPos = 1 To Len(strWord)
	    strTemp = Mid(strWord, cint(intPos), 1)
	    strBinary = strBinary & IntToBinary(Asc(strTemp))
	Next

	WordToBinary = strBinary

End Function

Function HexToInt(strHex)
' Function Converts a hex word to its base 10(decimal) equivalent
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'

	Dim intNew, intPos, intLen

	intNew = 0
	intLen = CDbl(len(strHex)) - 1
	
	For intPos = CDbl(intLen) to 0 step -1
	    Select Case Mid(strHex, CDbl(intPos) + 1, 1)       
	    Case "0"
			intNew = CDbl(intNew) + (0 * 16^CDbl(intLen - intPos))
	    Case "1"
	        intNew = CDbl(intNew) + (1 * 16^CDbl(intLen - intPos))
	    Case "2"
	        intNew = CDbl(intNew) + (2 * 16^CDbl(intLen - intPos))
	    Case "3"
	        intNew = CDbl(intNew) + (3 * 16^CDbl(intLen - intPos))
	    Case "4"
	        intNew = CDbl(intNew) + (4 * 16^CDbl(intLen - intPos))
	    Case "5"
	        intNew = CDbl(intNew) + (5 * 16^CDbl(intLen - intPos))
	    Case "6"
	        intNew = CDbl(intNew) + (6 * 16^CDbl(intLen - intPos))
	    Case "7"
	        intNew = CDbl(intNew) + (7 * 16^CDbl(intLen - intPos))
	    Case "8"
	        intNew = CDbl(intNew) + (8 * 16^CDbl(intLen - intPos))
	    Case "9"
	        intNew = CDbl(intNew) + (9 * 16^CDbl(intLen - intPos))
	    Case "A"
	        intNew = CDbl(intNew) + (10 * 16^CDbl(intLen - intPos))
	    Case "B"
	        intNew = CDbl(intNew) + (11 * 16^CDbl(intLen - intPos))
	    Case "C"
	        intNew = CDbl(intNew) + (12 * 16^CDbl(intLen - intPos))
	    Case "D"
	        intNew = CDbl(intNew) + (13 * 16^CDbl(intLen - intPos))
	    Case "E"
	        intNew = CDbl(intNew) + (14 * 16^CDbl(intLen - intPos))
	    Case "F"
	        intNew = CDbl(intNew) + (15 * 16^CDbl(intLen - intPos))
		End Select

	Next

	HexToInt = CDbl(intNew)
	
End Function

Function IntToBinary(intNum)

' Function Converts an integer number to it's binary equivalent
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'
    Dim strBinary, strTemp
    Dim intNew, intTemp
    Dim dblNew
    
    intNew = intNum
    
    Do While intNew > 1
        dblNew = CDbl(intNew) / 2
        intNew = Round(CDbl(dblNew) - 0.1, 0)
        If CDbl(dblNew) = CDbl(intNew) Then
            strBinary = "0" & strBinary
        Else
            strBinary = "1" & strBinary
        End If

    Loop
    
    strBinary = intNew & strBinary
    
    intTemp = Len(strBinary) mod 8
    
    For intNew = intTemp To 7
        strBinary = "0" & strBinary
    Next
    
    IntToBinary = strBinary
    
End Function

Function PadBinary(strBinary)

' Function adds 0's to a binary string until it reaches 448 bits.
' The lenghth of the original string is incoded into the last 16 bits.
' The end result is a binary string 512 bits long
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'

	Dim intPos, intLen
	Dim strTemp
	    
	intLen = Len(strBinary)
	    
	strBinary = strBinary & "1"
	    
	For intPos = Len(strBinary) To 447
	    strBinary = strBinary & "0"
	Next
	    
	strTemp = IntToBinary(intLen)
	    
	For intPos = Len(strTemp) To 63
	    strTemp = "0" & strTemp
	Next
	    
	strBinary = strBinary & strTemp
	    
	PadBinary = strBinary
	    
End Function

Function BlockToHex(strBinary)

' Function Converts a 32 bit binary string into it's 8 digit hex equivalent
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'
	Dim intPos
	Dim strHex

	For intPos = 1 To Len(strBinary) Step 4
	    strHex = strHex & BinaryToHex(Mid(strBinary, intPos, 4))
	Next

	BlockToHex = strHex

End Function

Function DigestHex(strHex, strH0, strH1, strH2, strH3, strH4)

' Main encoding function.  Takes a 128 digit/512 bit hex value and one way encrypts it into
' a 40 digit/160 bit hex value.
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'
	Dim strWords(79), adoConst(4), strTemp, strTemp1, strTemp2, strTemp3, strTemp4
	Dim intPos
	Dim strH(4), strA(4), strK(3)

	'Constant hex words are used for encryption, these can be any valid 8 digit hex value
    strK(0) = "5A827999"
    strK(1) = "6ED9EBA1"
    strK(2) = "8F1BBCDC"
    strK(3) = "CA62C1D6"
    
    'Hex words are used in the encryption process, these can be any valid 8 digit hex value
    strH(0) = strH0
    strH(1) = strH1
    strH(2) = strH2
    strH(3) = strH3
    strH(4) = strH4
    
    'divide the Hex block into 16 hex words
	For intPos = 0 To (len(strHex) / 8) - 1
	    strWords(cint(intPos)) = Mid(strHex, (cint(intPos)*8) + 1, 8)
	Next


    'encode the Hex words using the constants above
    'innitialize 80 hex word positions
	For intPos = 16 To 79
	    strTemp = strWords(cint(intPos) - 3)
	    strTemp1 = HexBlockToBinary(strTemp)
	    strTemp = strWords(cint(intPos) - 8)
	    strTemp2 = HexBlockToBinary(strTemp)
	    strTemp = strWords(cint(intPos) - 14)
	    strTemp3 = HexBlockToBinary(strTemp)
	    strTemp = strWords(cint(intPos) - 16)
	    strTemp4 = HexBlockToBinary(strTemp)
	    strTemp = BinaryXOR(strTemp1, strTemp2)
	    strTemp = BinaryXOR(strTemp, strTemp3)
	    strTemp = BinaryXOR(strTemp, strTemp4)
	    strWords(cint(intPos)) = BlockToHex(BinaryShift(strTemp, 1))
	Next

    'initialize the changing word variables with the initial word variables
	strA(0) = strH(0)
	strA(1) = strH(1)
	strA(2) = strH(2)
	strA(3) = strH(3)
	strA(4) = strH(4)

	'Main encryption loop on all 80 hex word positions
	For intPos = 0 To 79
	    strTemp = BinaryShift(HexBlockToBinary(strA(0)), 5)
	    strTemp1 = HexBlockToBinary(strA(3))
	    strTemp2 = HexBlockToBinary(strWords(cint(intPos)))
	    
	    Select Case intPos
	    
	    Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19
	        strTemp3 = HexBlockToBinary(strK(0))
	        strTemp4 = BinaryOR(BinaryAND(HexBlockToBinary(strA(1)), _
				HexBlockToBinary(strA(2))), BinaryAND(BinaryNOT(HexBlockToBinary(strA(1))), _
				HexBlockToBinary(strA(3))))
	    Case 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39
	        strTemp3 = HexBlockToBinary(strK(1))
	        strTemp4 = BinaryXOR(BinaryXOR(HexBlockToBinary(strA(1)), _
				HexBlockToBinary(strA(2))), HexBlockToBinary(strA(3)))
	    Case 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59
	        strTemp3 = HexBlockToBinary(strK(2))
	        strTemp4 = BinaryOR(BinaryOR(BinaryAND(HexBlockToBinary(strA(1)), _
				HexBlockToBinary(strA(2))), BinaryAND(HexBlockToBinary(strA(1)), _
				HexBlockToBinary(strA(3)))), BinaryAND(HexBlockToBinary(strA(2)), _
				HexBlockToBinary(strA(3))))
	    Case 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79
	        strTemp3 = HexBlockToBinary(strK(3))
	        strTemp4 = BinaryXOR(BinaryXOR(HexBlockToBinary(strA(1)), _
				HexBlockToBinary(strA(2))), HexBlockToBinary(strA(3)))
	    End Select
	    
	    strTemp = BlockToHex(strTemp)
	    strTemp1 = BlockToHex(strTemp1)
	    strTemp2 = BlockToHex(strTemp2)
	    strTemp3 = BlockToHex(strTemp3)
	    strTemp4 = BlockToHex(strTemp4)
	    
	    strTemp = HexAdd(strTemp, strTemp1)
	    strTemp = HexAdd(strTemp, strTemp2)
	    strTemp = HexAdd(strTemp, strTemp3)
	    strTemp = HexAdd(strTemp, strTemp4)
	    
	    strA(4) = strA(3)
	    strA(3) = strA(2)
	    strA(2) = BlockToHex(BinaryShift(HexBlockToBinary(strA(1)), 30))
	    strA(1) = strA(0)
	    strA(0) = strTemp
	    
	Next

	'Concatenate the final Hex Digest
	DigestHex = strA(0) & strA(1) & strA(2) & strA(3) & strA(4)

End Function

Function HexAdd(strHex1, strHex2)
' Function adds to 8 digit/32 bit hex values together Mod 2^32
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'
    Dim intCalc
    Dim strNew
    
    intCalc = 0
    intCalc = CDbl(CDbl(HexToInt(strHex1)) + CDbl(HexToInt(strHex2)))
    Do While CDbl(intCalc) > 2^32
		intCalc = CDbl(intCalc) - 2^32
    Loop
       
    strNew = IntToBinary(CDbl(intCalc))
    Do While Len(strNew) < 32
        strNew = "0" & strNew
    Loop
    strNew = BlockToHex(strNew)
    
    if InStr(strNew, "00") = 1 and len(strNew) = 10 then
		strNew = right(strNew, 8)
    end if
    
    HexAdd = strNew

End Function

Function getHexDec(strHex)
' Function Converts a single hex value into it's decimal equivalent
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'
    Select Case strHex
    Case "0"
        getHexDec = 0
    Case "1"
        getHexDec = 1
    Case "2"
        getHexDec = 2
    Case "3"
        getHexDec = 3
    Case "4"
        getHexDec = 4
    Case "5"
        getHexDec = 5
    Case "6"
        getHexDec = 6
    Case "7"
        getHexDec = 7
    Case "8"
        getHexDec = 8
    Case "9"
        getHexDec = 9
    Case "A"
        getHexDec = 10
    Case "B"
        getHexDec = 11
    Case "C"
        getHexDec = 12
    Case "D"
        getHexDec = 13
    Case "E"
        getHexDec = 14
    Case "F"
        getHexDec = 15
    Case Else
        getHexDec = -1
    End Select
End Function

Function getDecHex(strHex)
' Function Converts a single decimal value(0 - 15) into it's hex equivalent
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'
    Select Case CInt(strHex)
    Case 0
       getDecHex = "0"
    Case 1
       getDecHex = "1"
    Case 2
       getDecHex = "2"
    Case 3
       getDecHex = "3"
    Case 4
       getDecHex = "4"
    Case 5
       getDecHex = "5"
    Case 6
       getDecHex = "6"
    Case 7
       getDecHex = "7"
    Case 8
       getDecHex = "8"
    Case 9
       getDecHex = "9"
    Case 10
       getDecHex = "A"
    Case 11
       getDecHex = "B"
    Case 12
       getDecHex = "C"
    Case 13
       getDecHex = "D"
    Case 14
       getDecHex = "E"
    Case 15
       getDecHex = "F"
    Case Else
       getDecHex = "Z"
    End Select
End Function

Function BinaryShift(strBinary, intPos)
' Function circular left shifts a binary value n places
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'
    BinaryShift = Right(strBinary, Len(strBinary) - cint(intPos)) & _
		Left(strBinary, cint(intPos))

End Function

Function BinaryXOR(strBin1, strBin2)
' Function performs an exclusive or function on each position of two binary values
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'
    Dim strBinaryFinal
    Dim intPos
    
    For intPos = 1 To Len(strBin1)
        Select Case Mid(strBin1, cint(intPos), 1)
        
        Case Mid(strBin2, cint(intPos), 1)
            strBinaryFinal = strBinaryFinal & "0"
        Case Else
            strBinaryFinal = strBinaryFinal & "1"
        End Select
    Next
    
    BinaryXOR = strBinaryFinal
    
End Function

Function BinaryOR(strBin1, strBin2)
' Function performs an inclusive or function on each position of two binary values
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'
    Dim strBinaryFinal
    Dim intPos
    
    For intPos = 1 To Len(strBin1)
        If Mid(strBin1, cint(intPos), 1) = "1" Or Mid(strBin2, cint(intPos), 1) = "1" Then
            strBinaryFinal = strBinaryFinal & "1"
        Else
            strBinaryFinal = strBinaryFinal & "0"
        End If
    Next
    
    BinaryOR = strBinaryFinal
End Function

Function BinaryAND(strBin1, strBin2)
' Function performs an AND function on each position of two binary values
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'
    Dim strBinaryFinal
    Dim intPos
    
    For intPos = 1 To Len(strBin1)
        If Mid(strBin1, cint(intPos), 1) = "1" And Mid(strBin2, cint(intPos), 1) = "1" Then
            strBinaryFinal = strBinaryFinal & "1"
        Else
            strBinaryFinal = strBinaryFinal & "0"
        End If
    Next
    
    BinaryAND = strBinaryFinal
End Function

Function BinaryNOT(strBinary)
' Function makes each position of a binary value from 1 to 0 and 0 to 1
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'
    Dim strBinaryFinal
    Dim intPos
    
    For intPos = 1 To Len(strBinary)
        If Mid(strBinary, cint(intPos), 1) = "1" Then
            strBinaryFinal = strBinaryFinal & "0"
        Else
            strBinaryFinal = strBinaryFinal & "1"
        End If
    Next
    
    BinaryNOT = strBinaryFinal
    
End Function

Function HexBlockToBinary(strHex)
' Function Converts a 8 digit/32 bit hex value to its 32 bit binary equivalent
'
' Written By: Mark Jager
' Written Date: 8/10/2000
'
' Free to distribute as long as code is not modified, and header is kept intact
'
    Dim intPos
    Dim strTemp
    
    For intPos = 1 To Len(strHex)
        strTemp = strTemp & HexToBinary(Mid(strHex, cint(intPos), 1))
    Next
    
    HexBlockToBinary = strTemp
    
End Function
%>
