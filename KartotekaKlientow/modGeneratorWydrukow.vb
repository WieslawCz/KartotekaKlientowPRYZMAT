Module modGeneratorWydrukow

    Public Function GWNaglowek(ByRef GenStr As String) As String
        Dim pos As Integer = InStr(GenStr, vbCrLf)
        'naglowek wydruku jest na poczatku
        If pos = 0 Then
            Return ("")
        Else
            Return (Mid(GenStr, 1, pos - 1))
        End If
    End Function

    Public Function GWIleKolumn(ByRef GenStr As String) As Integer
        Dim pos As Integer = InStr(GenStr, vbCrLf)
        Dim txt As String = GenStr
        Dim ileKol As Integer = 0
        'naglowek wydruku jest na poczatku- pomin
        txt = Mid(txt, pos + 2)
        pos = InStr(txt, vbCrLf)
        Do While pos > 0
            ileKol += 1
            txt = Mid(txt, pos + 2)
            pos = InStr(txt, vbCrLf)
        Loop
        Return (ileKol)
    End Function


    Public Function GWKolNazwa(ByRef GenStr As String, _
                                 ByVal NrKol As Integer) As String
        Dim pos As Integer = InStr(GenStr, vbCrLf)
        Dim txt As String = GenStr
        Dim kol As String = ""
        Dim ileKol As Integer = 0

        Dim N As String = ""
        Dim T As String = ""
        Dim R As Integer = 0
        Dim H As String = ""
        Dim W As String = ""
        Dim I As Integer = 0

        'naglowek wydruku jest na poczatku- pomin
        txt = Mid(txt, pos + 2)
        pos = InStr(txt, vbCrLf)
        Do While pos > 0
            ileKol += 1
            kol = Mid(txt, 1, pos - 1)
            txt = Mid(txt, pos + 2)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK
            N = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK
            T = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            R = CInt(Mid(kol, 1, pos - 1))
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            H = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            W = Mid(kol, 1, pos - 1)
            I = CInt(Mid(kol, pos + 1))

            If ileKol = NrKol Then
                Exit Do
            End If

            pos = InStr(txt, vbCrLf)
        Loop
        Return (N)
    End Function

    Public Function GWKolTyp(ByRef GenStr As String, _
                               ByVal NrKol As Integer) As String
        Dim pos As Integer = InStr(GenStr, vbCrLf)
        Dim txt As String = GenStr
        Dim kol As String = ""
        Dim ileKol As Integer = 0

        Dim N As String = ""
        Dim T As String = ""
        Dim R As Integer = 0
        Dim H As String = ""
        Dim W As String = ""
        Dim I As Integer = 0

        'naglowek wydruku jest na poczatku- pomin
        txt = Mid(txt, pos + 2)
        pos = InStr(txt, vbCrLf)
        Do While pos > 0
            ileKol += 1
            kol = Mid(txt, 1, pos - 1)
            txt = Mid(txt, pos + 2)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK
            N = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK
            T = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            R = CInt(Mid(kol, 1, pos - 1))
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            H = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            W = Mid(kol, 1, pos - 1)
            I = CInt(Mid(kol, pos + 1))

            If ileKol = NrKol Then
                Exit Do
            End If

            pos = InStr(txt, vbCrLf)
        Loop
        Return (T)
    End Function

    Public Function GWKolRozmiar(ByRef GenStr As String, _
                               ByVal NrKol As Integer) As Integer
        Dim pos As Integer = InStr(GenStr, vbCrLf)
        Dim txt As String = GenStr
        Dim kol As String = ""
        Dim ileKol As Integer = 0

        Dim N As String = ""
        Dim T As String = ""
        Dim R As Integer = 0
        Dim H As String = ""
        Dim W As String = ""
        Dim I As Integer = 0

        'naglowek wydruku jest na poczatku- pomin
        txt = Mid(txt, pos + 2)
        pos = InStr(txt, vbCrLf)
        Do While pos > 0
            ileKol += 1
            kol = Mid(txt, 1, pos - 1)
            txt = Mid(txt, pos + 2)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK
            N = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK
            T = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            R = CInt(Mid(kol, 1, pos - 1))
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            H = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            W = Mid(kol, 1, pos - 1)
            I = CInt(Mid(kol, pos + 1))

            If ileKol = NrKol Then
                Exit Do
            End If

            pos = InStr(txt, vbCrLf)
        Loop
        Return (R)
    End Function

    Public Function GWKolNaglowek(ByRef GenStr As String, _
                           ByVal NrKol As Integer) As String
        Dim pos As Integer = InStr(GenStr, vbCrLf)
        Dim txt As String = GenStr
        Dim kol As String = ""
        Dim ileKol As Integer = 0

        Dim N As String = ""
        Dim T As String = ""
        Dim R As Integer = 0
        Dim H As String = ""
        Dim W As String = ""
        Dim I As Integer = 0

        'naglowek wydruku jest na poczatku- pomin
        txt = Mid(txt, pos + 2)
        pos = InStr(txt, vbCrLf)
        Do While pos > 0
            ileKol += 1
            kol = Mid(txt, 1, pos - 1)
            txt = Mid(txt, pos + 2)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            N = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            T = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            R = CInt(Mid(kol, 1, pos - 1))
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            H = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            W = Mid(kol, 1, pos - 1)
            I = CInt(Mid(kol, pos + 1))

            If ileKol = NrKol Then
                Exit Do
            End If

            pos = InStr(txt, vbCrLf)
        Loop
        Return (H)
    End Function


    Public Function GWKolWyrownanie(ByRef GenStr As String, _
                       ByVal NrKol As Integer) As String
        Dim pos As Integer = InStr(GenStr, vbCrLf)
        Dim txt As String = GenStr
        Dim kol As String = ""
        Dim ileKol As Integer = 0
        Dim wyr As String = ""
        Dim ileWi As Integer = 0

        Dim N As String = ""
        Dim T As String = ""
        Dim R As Integer = 0
        Dim H As String = ""
        Dim W As String = ""
        Dim I As Integer = 0

        'naglowek wydruku jest na poczatku- pomin
        txt = Mid(txt, pos + 2)
        pos = InStr(txt, vbCrLf)
        Do While pos > 0
            ileKol += 1
            kol = Mid(txt, 1, pos - 1)
            txt = Mid(txt, pos + 2)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            N = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            T = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            R = CInt(Mid(kol, 1, pos - 1))
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            H = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            W = Mid(kol, 1, pos - 1)
            I = CInt(Mid(kol, pos + 1))

            If ileKol = NrKol Then
                Exit Do
            End If

            pos = InStr(txt, vbCrLf)
        Loop
        Return (W)
    End Function

    Public Function GWKolIleWierszy(ByRef GenStr As String, _
                   ByVal NrKol As Integer) As Integer
        Dim pos As Integer = InStr(GenStr, vbCrLf)
        Dim txt As String = GenStr
        Dim kol As String = GenStr
        Dim ileKol As Integer = 0

        Dim N As String = ""
        Dim T As String = ""
        Dim R As Integer = 0
        Dim H As String = ""
        Dim W As String = ""
        Dim I As Integer = 0

        'naglowek wydruku jest na poczatku- pomin
        txt = Mid(txt, pos + 2)
        pos = InStr(txt, vbCrLf)
        Do While pos > 0
            ileKol += 1
            kol = Mid(txt, 1, pos - 1)
            txt = Mid(txt, pos + 2)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            N = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            T = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            R = CInt(Mid(kol, 1, pos - 1))
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            H = Mid(kol, 1, pos - 1)
            kol = Mid(kol, pos + 1)

            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
            W = Mid(kol, 1, pos - 1)
            I = CInt(Mid(kol, pos + 1))

            If ileKol = NrKol Then
                Exit Do
            End If

            pos = InStr(txt, vbCrLf)
        Loop
        Return (I)
    End Function


    '**********************************************************
    ' obsluga wierszy wielowierszowych ??...
    '**********************************************************

    Public Sub GWDopiszNastWiersz(ByRef WiStr As String, _
                   ByVal NrKol As Integer, ByVal IleWi As Integer)
        Dim pos As Integer = 0
        Dim txt As String = WiStr
        Dim Pozo As String = WiStr
        Dim i As Integer = 0

        txt = WiStr
        Pozo = ""
        For i = 1 To NrKol
            pos = InStr(txt, "|")
            If i = NrKol Then
                If IleWi > 1 Then
                    Pozo &= "1," & Trim(Str(IleWi)) & "|"
                Else
                    Pozo &= "0,0|"
                End If
            Else
                Pozo &= Mid(txt, 1, pos)
            End If
            txt = Mid(txt, pos + 1)
        Next
        'posk³adaj
        WiStr = Pozo & txt
    End Sub

    Public Function GWNrNastWiersz(ByRef WiStr As String, _
               ByVal NrKol As Integer) As Integer
        Dim pos As Integer = 0
        Dim pop As Integer = 0
        Dim txt As String = WiStr
        Dim txt2 As String = ""
        Dim Pozo As String = WiStr
        Dim i As Integer = 0
        Dim wi1 As Integer
        Dim wi2 As Integer
        Dim ret As Integer = 0

        txt = WiStr
        Pozo = ""
        For i = 1 To NrKol
            pos = InStr(txt, "|")
            txt2 = Mid(txt, 1, pos - 1)

            'analizuj pozycje 9,9 (ile wydrukowano, ile wszystkich linii)
            pop = InStr(txt2, ",")
            wi1 = CInt(Mid(txt2, 1, pop - 1))
            wi2 = CInt(Mid(txt2, pop + 1))

            If i = NrKol Then
                If wi1 < wi2 Then
                    wi1 += 1
                    ret = wi1
                Else
                    ret = wi1
                End If
                If wi1 = wi2 Then
                    wi1 = 0
                    wi2 = 0
                End If
            End If
            'posk³adaj
            Pozo &= Trim(Str(wi1)) & "," & Trim(Str(wi2)) & "|"
            txt = Mid(txt, pos + 1)
        Next
        WiStr = Pozo & txt
        Return (ret)
    End Function

    Public Function GWCzyJestNastWiersz(ByRef WiStr As String) As Boolean
        Dim pos As Integer = InStr(WiStr, "|")
        Dim txt As String = WiStr
        Dim Pozo As String = WiStr
        Dim i As Integer = 0
        Dim ret As Boolean = False  'nic nie ma do roboty

        txt = WiStr
        pos = InStr(txt, "|")
        Do While pos > 0
            If Mid(txt, 1, pos - 1) = "0,0" Then
                'nie trzeba drukowac kolejnego wiersz
            Else
                'OK - kolejny wiersz z tego rekordu
                ret = True
                Exit Do
            End If
            txt = Mid(txt, pos + 1)
            pos = InStr(txt, "|")
        Loop
        Return (ret)
    End Function

End Module
