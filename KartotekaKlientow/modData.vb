Module modData

    'zmienia format daty z "RRRRMMDD" na "RRRR-MM-DD"
    Public Function Data2TXT(ByVal DataStr As String) As String
        Return (Mid(DataStr, 1, 4) & "-" & Mid(DataStr, 5, 2) & "-" & Mid(DataStr, 7, 2))
    End Function

    'zmienia format daty z "RRRR-MM-DD" na "RRRRMMDD" 
    Public Function Data2DBF(ByVal DataStr As String) As String
        Return (Mid(DataStr, 1, 4) & Mid(DataStr, 6, 2) & Mid(DataStr, 9, 2))
    End Function

    'wylicz ilość dni między datami (rrrr-mm-dd) od DATA1 do DATA2
    Public Function IleDniMiedzyDatami(ByVal Data1 As String, ByVal Data2 As String) As Integer
        Return (IleDniOd20000101(Val(Mid(Data2, 1, 4)), Val(Mid(Data2, 6, 2)), Val(Mid(Data2, 9, 2))) - _
                IleDniOd20000101(Val(Mid(Data1, 1, 4)), Val(Mid(Data1, 6, 2)), Val(Mid(Data1, 9, 2))))
    End Function

    'termin zaplaty
    Public Function TerminZaplaty(ByVal Data1 As String, ByVal IleDni As Long) As String
        Dim RRRR As Integer = Val(Mid(Data1, 1, 4))
        Dim MM As Integer = Val(Mid(Data1, 6, 2))
        Dim DD As Integer = Val(Mid(Data1, 9, 2))
        Return (JakiToDzien(IleDniOd20000101(RRRR, MM, DD) + IleDni))
    End Function

    ' wylicza ile dni minelo do 2000.01.01
    Private Function IleDniOd20000101(ByVal YY As Integer, ByVal MM As Integer, ByVal DD As Integer) As Long
        Dim IleDni As Long = 0
        Dim ii As Integer

        'rok przestępny
        'jest podzielny przez 4, ale nie jest podzielny przez 100, lub jest podzielny przez 400
        If YY > 2000 Then
            For ii = 2000 To YY - 1
                If ii / 400 = Int(ii / 400) Then
                    IleDni += 366   'przestepny
                ElseIf (ii / 4 = Int(ii / 4)) And (ii / 100 <> Int(ii / 100)) Then
                    IleDni += 366   'przestepny
                Else
                    IleDni += 365   'ZWYKŁY
                End If
            Next ii
        End If

        If MM > 1 Then
            For ii = 1 To MM - 1
                Select Case ii
                    Case 1
                        IleDni += 31
                    Case 2
                        If YY / 100 = Int(YY / 100) Then
                            IleDni += 29   'przestepny
                        ElseIf YY / 4 = Int(YY / 4) Then
                            IleDni += 29   'przestepny
                        Else
                            IleDni += 28
                        End If
                    Case 3
                        IleDni += 31
                    Case 4
                        IleDni += 30
                    Case 5
                        IleDni += 31
                    Case 6
                        IleDni += 30
                    Case 7
                        IleDni += 31
                    Case 8
                        IleDni += 31
                    Case 9
                        IleDni += 30
                    Case 10
                        IleDni += 31
                    Case 11
                        IleDni += 30
                    Case 12
                        IleDni += 31
                End Select
            Next ii
        End If

        IleDni += DD
        Return (IleDni)
    End Function

    'wylicza date na podstawie ilosci dni od 2000.01.01
    Private Function JakiToDzien(ByVal IleDniOd As Long) As String
        Dim RRRR As String
        Dim MM As String
        Dim DD As String

        Dim rok As Integer
        Dim miesiac As Integer
        Dim dzien As Integer
        Dim iledni As Integer

        'wyliczamy rok...
        rok = 2000
        iledni = 366
        Do While IleDniOd > iledni
            IleDniOd -= iledni
            rok += 1
            If rok / 100 = Int(rok / 100) Then
                iledni = 366   'przestepny
            ElseIf rok / 4 = Int(rok / 4) Then
                iledni = 366   'przestepny
            Else
                iledni = 365
            End If
        Loop

        'wyliczamy miesiac...
        miesiac = 1
        iledni = 31
        Do While IleDniOd > iledni
            IleDniOd -= iledni
            miesiac += 1
            Select Case miesiac
                Case 1
                    iledni = 31
                Case 2
                    If rok / 100 = Int(rok / 100) Then
                        iledni = 29   'przestepny
                    ElseIf rok / 4 = Int(rok / 4) Then
                        iledni = 29   'przestepny
                    Else
                        iledni = 28
                    End If
                Case 3
                    iledni = 31
                Case 4
                    iledni = 30
                Case 5
                    iledni = 31
                Case 6
                    iledni = 30
                Case 7
                    iledni = 31
                Case 8
                    iledni = 31
                Case 9
                    iledni = 30
                Case 10
                    iledni = 31
                Case 11
                    iledni = 30
                Case 12
                    iledni = 31
            End Select
        Loop

        'wyliczam dzien...
        dzien = IleDniOd

        RRRR = "0000" & Trim(rok.ToString)
        MM = "00" & Trim(miesiac.ToString)
        DD = "00" & Trim(dzien.ToString)
        Return (Mid(RRRR, Len(RRRR) - 4 + 1, 4) & "-" & Mid(MM, Len(MM) - 2 + 1, 2) & "-" & Mid(DD, Len(DD) - 2 + 1, 2))
    End Function


    Public Function PierwszyDzienPoprzedniegoMiesiaca(ByVal BiezData As String) As String
        Dim RRRR As String = Mid(BiezData, 1, 4)
        Dim RRRRnum As Integer = GetNumFromText(Mid(BiezData, 1, 4))
        Dim MM As String = Mid(BiezData, 6, 2)
        Dim RetData As String = ""

        Select Case MM
            Case "01"
                '1 grudzien poprzedniego roku
                RRRR = Trim((RRRRnum - 1).ToString)
                RetData = RRRR & "-12-01"
            Case "02"
                RetData = RRRR & "-01-01"
            Case "03"
                RetData = RRRR & "-02-01"
            Case "04"
                RetData = RRRR & "-03-01"
            Case "05"
                RetData = RRRR & "-04-01"
            Case "06"
                RetData = RRRR & "-05-01"
            Case "07"
                RetData = RRRR & "-06-01"
            Case "08"
                RetData = RRRR & "-07-01"
            Case "09"
                RetData = RRRR & "-08-01"
            Case "10"
                RetData = RRRR & "-09-01"
            Case "11"
                RetData = RRRR & "-10-01"
            Case "12"
                RetData = RRRR & "-11-01"
        End Select
        Return RetData
    End Function

    Public Function OstatniDzienPoprzedniegoMiesiaca(ByVal BiezData As String) As String
        Dim RRRR As String = Mid(BiezData, 1, 4)
        Dim RRRRnum As Integer = GetNumFromText(Mid(BiezData, 1, 4))
        Dim MM As String = Mid(BiezData, 6, 2)
        Dim RetData As String = ""

        Select Case MM
            Case "01"
                '1 grudzien poprzedniego roku
                RRRR = Trim((RRRRnum - 1).ToString)
                RetData = RRRR & "-12-31"
            Case "02"
                RetData = RRRR & "-01-31"
            Case "03"
                If RRRRnum / 100 = Int(RRRRnum / 100) Then
                    RetData = RRRR & "-02-29"
                ElseIf RRRRnum / 4 = Int(RRRRnum / 4) Then
                    RetData = RRRR & "-02-29"
                Else
                    RetData = RRRR & "-02-28"
                End If
            Case "04"
                RetData = RRRR & "-03-31"
            Case "05"
                RetData = RRRR & "-04-30"
            Case "06"
                RetData = RRRR & "-05-31"
            Case "07"
                RetData = RRRR & "-06-30"
            Case "08"
                RetData = RRRR & "-07-31"
            Case "09"
                RetData = RRRR & "-08-31"
            Case "10"
                RetData = RRRR & "-09-30"
            Case "11"
                RetData = RRRR & "-10-31"
            Case "12"
                RetData = RRRR & "-11-30"
        End Select
        Return RetData
    End Function




    Public Function PierwszyDzienWybranegoMiesiaca(ByVal BiezData As String, _
                                               ByVal RoznicaMiesiecy As Integer) As String
        Dim RRRRnum As Integer = GetNumFromText(Mid(BiezData, 1, 4))
        Dim MMnum As Integer = GetNumFromText(Mid(BiezData, 6, 2))

        MMnum = MMnum + RoznicaMiesiecy
        Do While MMnum < 1
            MMnum += 12
            RRRRnum -= 1
        Loop
        Do While MMnum > 12
            MMnum -= 12
            RRRRnum += 1
        Loop
        'data w formacie YYYY-MM-DD
        Return Right("0000" & Trim(RRRRnum.ToString()), 4) & "-" & _
               Right("00" & Trim(MMnum.ToString()), 2) & "-01"
    End Function



    Public Function OstatniDzienWybranegoMiesiaca(ByVal BiezData As String, _
                                               ByVal RoznicaMiesiecy As Integer) As String
        Dim RRRRnum As Integer = GetNumFromText(Mid(BiezData, 1, 4))
        Dim MMnum As Integer = GetNumFromText(Mid(BiezData, 6, 2))
        Dim DDnum As Integer = 30

        MMnum = MMnum + RoznicaMiesiecy
        Do While MMnum < 1
            MMnum += 12
            RRRRnum -= 1
        Loop
        Do While MMnum > 12
            MMnum -= 12
            RRRRnum += 1
        Loop

        Select Case MMnum
            Case 1, 3, 5, 7, 8, 10, 12
                DDnum = 31
            Case 4, 6, 9, 11
                DDnum = 30
            Case 2
                If RRRRnum / 100 = Int(RRRRnum / 100) Then
                    DDnum = 29
                ElseIf RRRRnum / 4 = Int(RRRRnum / 4) Then
                    DDnum = 29
                Else
                    DDnum = 28
                End If
        End Select

        'data w formacie YYYY-MM-DD
        Return Right("0000" & Trim(RRRRnum.ToString()), 4) & "-" & _
               Right("00" & Trim(MMnum.ToString()), 2) & "-" & _
               Right("00" & Trim(DDnum.ToString()), 2)
    End Function






End Module
