Attribute VB_Name = "basBCREAD"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' BARCODE READER module                                                    ''
'' By Paul Bahlawan Oct 2004                                                ''
''                                                                          ''
'' Usage:                                                                   ''
'' value$ = bcRead(pb, x, y, bcType, retries, verbos)                       ''
''   pb = name of the picturebox with the barcode                           ''
''   x,y = coordanates in pixels of start of barcode                        ''
''   bcType: (see bcTypes)                                                  ''
''   retries = times to retry read (1 - 64)(optional)(default 5)            ''
''   verbos = what to return in case of a no-read (error): (optional)       ''
''            0= "" [empty string] (default)                                ''
''            1= "Error"                                                    ''
''            2= Full error message (form the LAST retry)                   ''
''                                                                          ''
''                                                                          ''
'' -Image must be bitonal, that is: black and white only (monochrome/1-bit) ''
'' -Picturebox scalemode must be pixels                                     ''
''                                                                          ''
''                                                                          ''
'' Updates:                                                                 ''
''  Nov.2011    - Use GetPixel; faster than Point                           ''
''              - Much code re-arranging                                    ''
''              - Add CodaBar                                               ''
''              - Add i2of5                                                 ''
''                                                                          ''
'' Based on specs from:                                                     ''
'' www.adams1.com/pub/russadam/39code.html                                  ''
'' www.barcodeman.com/info/barspec.php                                      ''
'' en.wikipedia.org                                                         ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Enum bcTypes
    bc39 = 0
    bcCodabar = 1
    bci25 = 2
End Enum

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long


'----------------------------------------------------------------------------'
'Program Interface for reading a barcode                                     '
'----------------------------------------------------------------------------'
Public Function bcRead(pb As PictureBox, ByVal rX As Long, ByVal rY As Long, ByVal bcType As Long, Optional ByVal retries As Long = 5, Optional ByVal verbos As Long = 0) As String
    If retries < 1 Then retries = 1
    If retries > 64 Then retries = 64
    
    'Sometimes it takes a few tries to get a read
    Do
        Select Case bcType
            Case bc39
                bcRead = bcRead39(pb, rX, rY)
            Case bcCodabar
                bcRead = bcReadCodabar(pb, rX, rY)
            Case bci25
                bcRead = bcReadi25(pb, rX, rY)
        End Select
        If Left$(bcRead, 4) <> "ERR:" Then Exit Do
        retries = retries - 1
        rY = rY + 1 'drop down a line and try the read again
    Loop While retries
    
    'Verbos level for an Error
    If Left$(bcRead, 4) = "ERR:" Then
        Select Case verbos
            Case 0
                bcRead = ""
            Case 1
                bcRead = "Error"
            'case else, just return the error string as is
        End Select
    End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' READ 3 of 9                                                                '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function bcRead39(pb As PictureBox, ByVal xBC As Long, ByVal yBC As Long) As String
Dim tmpStr As String
Dim i As Long
Dim j As Long
Dim BC(43) As String

    '3 of the 9 elements are wide: 1=narrow, 2=wide
    'note: the additional (10th) element is the intercharactor gap (1 narrow space)
    BC(0) = "1112212111" '0
    BC(1) = "2112111121" '1
    BC(2) = "1122111121" '2
    BC(3) = "2122111111" '3
    BC(4) = "1112211121" '4
    BC(5) = "2112211111" '5
    BC(6) = "1122211111" '6
    BC(7) = "1112112121" '7
    BC(8) = "2112112111" '8
    BC(9) = "1122112111" '9
    BC(10) = "2111121121" 'A
    BC(11) = "1121121121" 'B
    BC(12) = "2121121111" 'C
    BC(13) = "1111221121" 'D
    BC(14) = "2111221111" 'E
    BC(15) = "1121221111" 'F
    BC(16) = "1111122121" 'G
    BC(17) = "2111122111" 'H
    BC(18) = "1121122111" 'I
    BC(19) = "1111222111" 'J
    BC(20) = "2111111221" 'K
    BC(21) = "1121111221" 'L
    BC(22) = "2121111211" 'M
    BC(23) = "1111211221" 'N
    BC(24) = "2111211211" 'O
    BC(25) = "1121211211" 'P
    BC(26) = "1111112221" 'Q
    BC(27) = "2111112211" 'R
    BC(28) = "1121112211" 'S
    BC(29) = "1111212211" 'T
    BC(30) = "2211111121" 'U
    BC(31) = "1221111121" 'V
    BC(32) = "2221111111" 'W
    BC(33) = "1211211121" 'X
    BC(34) = "2211211111" 'Y
    BC(35) = "1221211111" 'Z
    BC(36) = "1211112121" '-
    BC(37) = "2211112111" '.
    BC(38) = "1221112111" '<SPC>
    BC(39) = "1212121111" '$
    BC(40) = "1212111211" '/
    BC(41) = "1211121211" '+
    BC(42) = "1112121211" '%
    BC(43) = "1211212111" '*  (used for start/stop character only)

'Scan the barcode image
    tmpStr = bcScan(pb, xBC, yBC)
    
    If Left$(tmpStr, 4) = "ERR:" Then
        bcRead39 = tmpStr
        Exit Function
    End If

    tmpStr = tmpStr & "1"
    If Len(tmpStr) Mod 10 Then
        bcRead39 = "ERR: BC parity"
        Exit Function
    End If
    
'Decode 1's and 2's string into characters
    For j = 0 To Len(tmpStr) / 10 - 1
        For i = 0 To 43
            If Mid$(tmpStr, 1 + j * 10, 10) = BC(i) Then
                bcRead39 = bcRead39 & Mid$("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%*", i + 1, 1)
                Exit For
            End If
            If i = 43 Then
                bcRead39 = "ERR: BC unrecognized"
                Exit Function
            End If
        Next i
    Next j
    
'Valid 3 of 9 starts & ends with a *
    If Left$(bcRead39, 1) <> "*" Or Right$(bcRead39, 1) <> "*" Or Len(bcRead39) < 2 Then
        bcRead39 = "ERR: BC invalid"
        Exit Function
    End If
    
'Finally, trim off the *'s
    bcRead39 = Mid$(bcRead39, 2, Len(bcRead39) - 2)
    
'if check character is used, verify before returning (to be done)
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' READ CODABAR (aka NW-7 and 2 of 7)                                         '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function bcReadCodabar(pb As PictureBox, ByVal xBC As Long, ByVal yBC As Long) As String
Dim tmpStr As String
Dim i As Long
Dim j As Long
Dim BC(43) As String

    '2 of 7 elements are wide; 1=narrow, 2=wide
    'note: the additional (8th) element is the intercharactor gap (1 narrow space)
    BC(0) = "11111221" '0
    BC(1) = "11112211" '1
    BC(2) = "11121121" '2
    BC(3) = "22111111" '3
    BC(4) = "11211211" '4
    BC(5) = "21111211" '5
    BC(6) = "12111121" '6
    BC(7) = "12112111" '7
    BC(8) = "12211111" '8
    BC(9) = "21121111" '9
    BC(10) = "11122111" '-
    BC(11) = "11221111" '$
    BC(12) = "21112121" ':
    BC(13) = "21211121" '/
    BC(14) = "21212111" '.
    BC(15) = "11212121" '+
    BC(16) = "11221211" 'start/stop A
    BC(17) = "12121121" 'start/stop B
    BC(18) = "11121221" 'start/stop C
    BC(19) = "11122211" 'start/stop D

'Scan the barcode image
    tmpStr = bcScan(pb, xBC, yBC)
    
    If Left$(tmpStr, 4) = "ERR:" Then
        bcReadCodabar = tmpStr
        Exit Function
    End If
    
    tmpStr = tmpStr & "1"
    If Len(tmpStr) Mod 8 Then
        bcReadCodabar = "ERR: BC parity"
        Exit Function
    End If
    
'Decode 1's and 2's string into characters
    For j = 0 To Len(tmpStr) / 8 - 1
        For i = 0 To 19
            If Mid$(tmpStr, 1 + j * 8, 8) = BC(i) Then
                bcReadCodabar = bcReadCodabar & Mid$("0123456789-$:/.+ABCD", i + 1, 1)
                Exit For
            End If
            If i = 19 Then
                bcReadCodabar = "ERR: BC unrecognized"
                Exit Function
            End If
        Next i
    Next j
    
'Valid Codabar starts & ends with an A,B,C or D start/stop chr
    If InStr("ABCD", Left$(bcReadCodabar, 1)) = 0 Or InStr("ABCD", Right$(bcReadCodabar, 1)) = 0 Or Len(bcReadCodabar) < 2 Then
        bcReadCodabar = "ERR: BC invalid"
        Exit Function
    End If
    
'Finally, trim off the start/stop chracters
    bcReadCodabar = Mid$(bcReadCodabar, 2, Len(bcReadCodabar) - 2)
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' READ INTERLEAVED 2 of 5                                                    '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function bcReadi25(pb As PictureBox, ByVal xBC As Long, ByVal yBC As Long) As String
Dim tmpStr As String
Dim i As Long
Dim j As Long
Dim BC(43) As String
Dim DEi As String
Dim ch1 As String
Dim ch2 As String

    '2 of the 5 elements are wide: 1=narrow, 2=wide
    BC(0) = "11221" '0
    BC(1) = "21112" '1
    BC(2) = "12112" '2
    BC(3) = "22111" '3
    BC(4) = "11212" '4
    BC(5) = "21211" '5
    BC(6) = "12211" '6
    BC(7) = "11122" '7
    BC(8) = "21121" '8
    BC(9) = "12121" '9
    BC(10) = "11111" 'Start chr 1111 *modified for convenience
    BC(11) = "21111" 'Stop chr 211 *modified for convenience

'Scan the barcode image
    tmpStr = bcScan(pb, xBC, yBC)
    
    If Left$(tmpStr, 4) = "ERR:" Then
        bcReadi25 = tmpStr
        Exit Function
    End If
    
    tmpStr = "1" & tmpStr & "11" ' modify start & stop chr
    
    If Len(tmpStr) Mod 5 Then
        bcReadi25 = "ERR: BC parity"
        Exit Function
    End If
    
'DE-interleave the code string
    DEi = Left$(tmpStr, 5)
    For i = 6 To Len(tmpStr) - 5 Step 10
        ch1 = ""
        ch2 = ""
        For j = 0 To 9 Step 2
            ch1 = ch1 & Mid$(tmpStr, i + j, 1)
            ch2 = ch2 & Mid$(tmpStr, i + j + 1, 1)
        Next j
        DEi = DEi & ch1 & ch2
    Next i
    DEi = DEi & Right$(tmpStr, 5)
    tmpStr = DEi

'Decode 1's and 2's string into characters
    For j = 0 To Len(tmpStr) / 5 - 1
        For i = 0 To 11
            If Mid$(tmpStr, 1 + j * 5, 5) = BC(i) Then
                bcReadi25 = bcReadi25 & Mid$("0123456789AB", i + 1, 1)
                Exit For
            End If
            If i = 11 Then
                bcReadi25 = "ERR: BC unrecognized"
                Exit Function
            End If
        Next i
    Next j
    
'Valid i25 have a start and stop character (using A and B for convenience)
    If Left$(bcReadi25, 1) <> "A" Or Right$(bcReadi25, 1) <> "B" Or Len(bcReadi25) < 2 Then
        bcReadi25 = "ERR: BC invalid"
        Exit Function
    End If
    
'Finally, trim off the start/stop chracters
    bcReadi25 = Mid$(bcReadi25, 2, Len(bcReadi25) - 2)
End Function



'............................................................................'
'Make a string representing the bars and spaces from the image               '
'............................................................................'
Private Function bcScan(pb As PictureBox, ByVal xBC As Long, ByVal yBC As Long) As String

Dim XbcStart As Long
Dim sample As Long
Dim refsample As Long
Dim i As Long
Dim nSpace As Long
Dim wSpace As Long
Dim nBar As Long
Dim wBar As Long
    
'Find the first black pixel (ie. start of barcode)'
    XbcStart = xBC - 1
    Do
        XbcStart = XbcStart + 1
        sample = GetPixel(pb.hdc, XbcStart, yBC)
        If XbcStart > xBC + 75 Or sample = -1 Then
            bcScan = "ERR: BC not seen"
            Exit Function
        End If
    Loop While sample
    
'Scan to find narrowest and widest bars and spaces
    nSpace = 100
    wSpace = 0
    nBar = 100
    wBar = 0
    xBC = XbcStart
    
    Do
        refsample = GetPixel(pb.hdc, xBC, yBC)
        i = 0
        Do While GetPixel(pb.hdc, xBC + i, yBC) = refsample
            i = i + 1
            If i > 22 Then Exit Do
        Loop
        
        If i > 22 Or pb.Point(xBC + i, yBC) = -1 Then Exit Do
        
        If refsample Then
            If i < nSpace Then nSpace = i
            If i > wSpace Then wSpace = i
        Else
            If i < nBar Then nBar = i
            If i > wBar Then wBar = i
        End If
           
        xBC = xBC + i
    Loop
    
    If nSpace >= wSpace Or nBar >= wBar Then
        bcScan = "ERR: BC not readable"
        Exit Function
    End If
    
'Rescan and build temp string; 1 = narrow, 2 = wide
    xBC = XbcStart
    Do
        refsample = GetPixel(pb.hdc, xBC, yBC)
        i = 0
        
        Do While GetPixel(pb.hdc, xBC + i, yBC) = refsample
            i = i + 1
            If i > wSpace * 2 Then Exit Do
        Loop
        
        If i > wSpace * 2 Then Exit Do
        
        If refsample Then
            If i * 2 < nSpace + wSpace Then
                    bcScan = bcScan & "1"
            Else
                    bcScan = bcScan & "2"
            End If
        Else
            If i * 2 < nBar + wBar Then
                    bcScan = bcScan & "1"
            Else
                    bcScan = bcScan & "2"
            End If
        End If
            
        xBC = xBC + i
    Loop
End Function

