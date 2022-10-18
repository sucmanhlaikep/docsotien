'Ham chinh la DocSo
Function DocSo(ByVal numIn)
    nghin = "nghìn"
    trieu = "tri" & ChrW(7879) & "u"
    ty = "t" & ChrW(7927)
    phay = "ph" & ChrW(7849) & "y"
    
    Dim LSide, RSide, Temp, DecPlace, Count, oNum
    oNum = numIn
    ReDim Place(9) As String
    Place(2) = " " & nghin & " "
    Place(3) = " " & trieu & " "
    Place(4) = " " & ty & " "
    numIn = Trim(Str(numIn))
    DecPlace = InStr(numIn, ".")
    If DecPlace > 0 Then
        RSide = Doc_Chuc(Left(Mid(numIn, DecPlace + 1) & "00", 2))
        numIn = Trim(Left(numIn, DecPlace - 1))
    End If
    RSide = numIn
    Count = 1
    Do While numIn <> ""
        Temp = Doc_Tram(Right(numIn, 3))
        If Temp <> "" Then LSide = Temp & Place(Count) & LSide
        If Len(numIn) > 3 Then
            numIn = Left(numIn, Len(numIn) - 3)
        Else
            numIn = ""
        End If
        Count = Count + 1
    Loop

    DocSo = LSide
    If InStr(oNum, Application.DecimalSeparator) > 0 Then
        DocSo = DocSo & " " & phay & " " & Doc_ThapPhan(oNum)
    End If

End Function

Function Doc_Tram(ByVal numIn) 'Chuyen so tu 100-999 sang chu
    tram = "tr" & ChrW(259) & "m"
    
    Dim w As String
    If Val(numIn) = 0 Then Exit Function
    numIn = Right("000" & numIn, 3)
    If Mid(numIn, 1, 1) <> "0" Then
        w = Doc_Donvi(Mid(numIn, 1, 1)) & " " & tram & " "
    End If
    If Mid(numIn, 2, 1) <> "0" Then
        w = w & Doc_Chuc(Mid(numIn, 2))
    Else
        w = w & Doc_Donvi(Mid(numIn, 3))
    End If
    Doc_Tram = w
End Function

Function Doc_Chuc(TensText)  'Chuyen so tu 10-99 sang chu
    mot = "m" & ChrW(7897) & "t"
    hai = "hai"
    ba = "ba"
    bon = "b" & ChrW(7889) & "n"
    nam = "n" & ChrW(259) & "m"
    sau = "sáu"
    bay = "b" & ChrW(7843) & "y"
    tam = "tám"
    chin = "chín"
    muoi = "m" & ChrW(432) & ChrW(7901) & "i"
    muoii = "m" & ChrW(432) & ChrW(417) & "i"
    mott = "m" & ChrW(7889) & "t"
    
    Dim w As String
    w = ""
    If Val(Left(TensText, 1)) = 1 Then   'Neu gia tri tu 10-19
        Select Case Val(TensText)
            Case 10: w = muoi
            Case 11: w = muoi & " " & mot
            Case 12: w = muoi & " " & hai
            Case 13: w = muoi & " " & ba
            Case 14: w = muoi & " " & bon
            Case 15: w = muoi & " " & lam
            Case 16: w = muoi & " " & sau
            Case 17: w = muoi & " " & bay
            Case 18: w = muoi & " " & tam
            Case 19: w = muoi & " " & chin
            Case Else
        End Select
    Else      'Neu gia tri tu 20-99
        Select Case Val(Left(TensText, 1))
            Case 2: w = hai & " " & muoii
            Case 3: w = ba & " " & muoii
            Case 4: w = bon & " " & muoii
            Case 5: w = nam & " " & muoii
            Case 6: w = sau & " " & muoii
            Case 7: w = bay & " " & muoii
            Case 8: w = tam & " " & muoii
            Case 9: w = chin & " " & muoii
            Case Else
        End Select
        If Val(Right(TensText, 1)) = 1 Then
            w = w & " " & mott
        ElseIf Val(Right(TensText, 1)) <> 0 Then
            w = w & " " & Doc_Donvi(Right(TensText, 1))
        End If
    End If
    Doc_Chuc = w
End Function

Function Doc_Donvi(Digit) 'Chuyen so tu 1-9 sang chu
    mot = "m" & ChrW(7897) & "t"
    hai = "hai"
    ba = "ba"
    bon = "b" & ChrW(7889) & "n"
    nam = "n" & ChrW(259) & "m"
    sau = "sáu"
    bay = "b" & ChrW(7843) & "y"
    tam = "tám"
    chin = "chín"
    Select Case Val(Digit)
        Case 1: Doc_Donvi = mot
        Case 2: Doc_Donvi = hai
        Case 3: Doc_Donvi = ba
        Case 4: Doc_Donvi = bon
        Case 5: Doc_Donvi = nam
        Case 6: Doc_Donvi = sau
        Case 7: Doc_Donvi = bay
        Case 8: Doc_Donvi = tam
        Case 9: Doc_Donvi = chin
        Case Else: Doc_Donvi = ""
    End Select
End Function

Function Doc_ThapPhan(n) As String
    Dim fraction As String, x As Long
    fraction = Split(n, Application.DecimalSeparator)(1)
    For x = 1 To Len(fraction)
        If Doc_ThapPhan <> "" Then Doc_ThapPhan = Doc_ThapPhan & " "
        Doc_ThapPhan = Doc_ThapPhan & Doc_Donvi(Mid(fraction, x, 1))
    Next x
End Function
