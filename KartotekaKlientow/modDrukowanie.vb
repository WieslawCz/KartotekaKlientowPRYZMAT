Imports System.Drawing.Printing

Module modDrukowanie

    Public Function IleLiniiPotrzebaNaText(ByVal txt As String, _
                     ByVal DlugoscPola As Long, _
                     ByVal FontPola As Font, _
                     ByVal pp As PrintPageEventArgs) As Integer
        Dim DrawPen As New Pen(Color.Black, 1)
        Dim eTxt As String = txt
        Dim eTxt1 As String
        Dim CrLfPos As Integer
        Dim ileLi As Integer = 0

        If Len(eTxt) > 0 Then
            ileLi = 0
            CrLfPos = InStr(eTxt, vbCrLf)
            Do While CrLfPos > 0
                eTxt1 = Mid(eTxt, 1, CrLfPos - 1)
                eTxt = Mid(eTxt, CrLfPos + 2)
                ileLi += WIluLiniachTrzebaPisacWierszTekstu(eTxt1, DlugoscPola, FontPola, pp)
                CrLfPos = InStr(eTxt, vbCrLf)
            Loop
            If Len(eTxt) > 0 Then
                ileLi += WIluLiniachTrzebaPisacWierszTekstu(eTxt, DlugoscPola, FontPola, pp)
            End If
        End If
        Return (ileLi)
    End Function

    Public Function WIluLiniachTrzebaPisacWierszTekstu(ByVal txt As String, _
                     ByVal DlugoscPola As Long, _
                     ByVal FontPola As Font, _
                     ByVal pp As PrintPageEventArgs) As Integer
        Dim eTxt1 As String = txt
        Dim LinTxt As String
        Dim DlugoscTxt As Long
        Dim DlugoscLinTxt As Long

        'sprawdz dopasowanie
        Dim ileLi As Integer = 0
        DlugoscTxt = pp.Graphics.MeasureString(eTxt1, FontPola).Width
        Do While DlugoscTxt > DlugoscPola
            'pobierz kolejna linie
            LinTxt = eTxt1
            DlugoscLinTxt = pp.Graphics.MeasureString(LinTxt, FontPola).Width
            Do While DlugoscLinTxt > DlugoscPola
                LinTxt = Mid(LinTxt, 1, Len(LinTxt) - 1)
                DlugoscLinTxt = pp.Graphics.MeasureString(LinTxt, FontPola).Width
            Loop
            '----------------------------------
            'linTxt dopasowany - sprawdz ostatni znak-czy mozna na nim przeniesc
            Dim i As Integer = 0
            For i = Len(LinTxt) To 1 Step -1
                If InStr(" .,/\-=", Mid(LinTxt, i, 1)) > 0 Then
                    Exit For
                Else
                End If
            Next
            If i > 1 Then
                'przenies na znalezionym znaku
                LinTxt = Mid(LinTxt, 1, i)
            Else
                'pozostaw tak jak jest...
            End If
            '----------------------------------
            'odejmij ta linie od tekstu - to jest kolejna linia
            eTxt1 = Mid(eTxt1, Len(LinTxt) + 1)
            DlugoscTxt = pp.Graphics.MeasureString(eTxt1, FontPola).Width
            ileLi += 1
        Loop
        'ostatni kawa³ek mniejszy ni¿ pole - nie analizuj ostatniego znaku bo za nim CRLF
        If Len(eTxt1) > 0 Then
            ileLi += 1
        End If
        Return (ileLi)
    End Function


    Public Function DajLinieTextuNr(ByVal txt As String, _
                     ByVal NumerLi As Integer, _
                     ByVal DlugoscPola As Long, _
                     ByVal FontPola As Font, _
                     ByVal pp As PrintPageEventArgs) As String

        Dim DrawPen As New Pen(Color.Black, 1)
        Dim eTxt As String = txt
        Dim eTxt1 As String
        Dim CrLfPos As Integer
        Dim ileLi As Integer = 0
        Dim RetTxt As String = ""
        Dim LinTxt As String = ""
        Dim DlugoscTxt As Long = 0
        Dim DlugoscLinTxt As Long = 0

        If Len(eTxt) > 0 Then
            ileLi = 0
            CrLfPos = InStr(eTxt, vbCrLf)
            Do While CrLfPos > 0
                eTxt1 = Mid(eTxt, 1, CrLfPos - 1)   'linia do analizy
                eTxt = Mid(eTxt, CrLfPos + 2)       'to co pozostaje

                'sprawdz dopasowanie
                DlugoscTxt = pp.Graphics.MeasureString(eTxt1, FontPola).Width
                Do While DlugoscTxt > DlugoscPola
                    'pobierz kolejna linie od poczatku aTxt1
                    LinTxt = eTxt1
                    DlugoscLinTxt = pp.Graphics.MeasureString(LinTxt, FontPola).Width
                    Do While DlugoscLinTxt > DlugoscPola
                        LinTxt = Mid(LinTxt, 1, Len(LinTxt) - 1)
                        DlugoscLinTxt = pp.Graphics.MeasureString(LinTxt, FontPola).Width
                    Loop
                    '----------------------------------
                    'linTxt dopasowany - sprawdz ostatni znak-czy mozna na nim przeniesc
                    Dim i As Integer = 0
                    For i = Len(LinTxt) To 1 Step -1
                        If InStr(" .,/\-=", Mid(LinTxt, i, 1)) > 0 Then
                            Exit For
                        Else
                        End If
                    Next
                    If i > 1 Then
                        'przenies na znalezionym znaku
                        LinTxt = Mid(LinTxt, 1, i)
                    Else
                        'pozostaw tak jak jest...
                    End If
                    '----------------------------------
                    ileLi += 1
                    If ileLi = NumerLi Then
                        RetTxt = LinTxt
                    End If
                    'odejmij ta linie od tekstu - to jest kolejna linia
                    eTxt1 = Mid(eTxt1, Len(LinTxt) + 1)
                    DlugoscTxt = pp.Graphics.MeasureString(eTxt1, FontPola).Width
                Loop
                'ostatni kawa³ek mniejszy ni¿ pole - nie analizuj ostatniego znaku bo za nim CRLF
                If Len(eTxt1) > 0 Then
                    ileLi += 1
                    If ileLi = NumerLi Then
                        RetTxt = eTxt1
                    End If
                End If
                'analizuj kolejny wiersz
                CrLfPos = InStr(eTxt, vbCrLf)
            Loop

            'ostatni wiersz bez CRLF
            If Len(eTxt) > 0 Then
                eTxt1 = eTxt
                'sprawdz dopasowanie
                DlugoscTxt = pp.Graphics.MeasureString(eTxt1, FontPola).Width
                Do While DlugoscTxt > DlugoscPola
                    'pobierz kolejna linie
                    LinTxt = eTxt1
                    DlugoscLinTxt = pp.Graphics.MeasureString(LinTxt, FontPola).Width
                    Do While DlugoscLinTxt > DlugoscPola
                        LinTxt = Mid(LinTxt, 1, Len(LinTxt) - 1)
                        DlugoscLinTxt = pp.Graphics.MeasureString(LinTxt, FontPola).Width
                    Loop
                    '----------------------------------
                    'linTxt dopasowany - sprawdz ostatni znak-czy mozna na nim przeniesc
                    Dim i As Integer = 0
                    For i = Len(LinTxt) To 1 Step -1
                        If InStr(" .,/\-=", Mid(LinTxt, i, 1)) > 0 Then
                            Exit For
                        Else
                        End If
                    Next
                    If i > 1 Then
                        'przenies na znalezionym znaku
                        LinTxt = Mid(LinTxt, 1, i)
                    Else
                        'pozostaw tak jak jest...
                    End If
                    '----------------------------------
                    ileLi += 1
                    If ileLi = NumerLi Then
                        RetTxt = LinTxt
                    End If
                    'odejmij ta linie od tekstu - to jest kolejna linia
                    eTxt1 = Mid(eTxt1, Len(LinTxt) + 1)
                    DlugoscTxt = pp.Graphics.MeasureString(eTxt1, FontPola).Width
                Loop
                'ostatni kawa³ek mniejszy ni¿ pole - nie analizuj ostatniego znaku bo za nim CRLF
                If Len(eTxt1) > 0 Then
                    ileLi += 1
                    If ileLi = NumerLi Then
                        RetTxt = eTxt1
                    End If
                End If
            End If
        End If
        Return (RetTxt)
    End Function
























    Public Function IleLiniiMaText(ByVal txt As String) As Integer
        Dim eTxt As String = txt
        Dim ileLi As Integer = 0
        Dim CrLfPos As Integer

        If Len(eTxt) > 0 Then
            ileLi = 1
            CrLfPos = InStr(eTxt, vbCrLf)
            Do While CrLfPos > 0
                ileLi += 1
                eTxt = Mid(eTxt, CrLfPos + 2)
                CrLfPos = InStr(eTxt, vbCrLf)
            Loop
        End If
        Return (ileLi)
    End Function

    Public Function IleLiniiMaText(ByVal txt As String, _
                                   ByVal DlugoscPola As Long, _
                                   ByVal FontPola As Font, _
                                   ByVal pp As PrintPageEventArgs) As Integer
        Dim eTxt As String = txt & vbCrLf
        Dim eWiersz As String = ""
        Dim eLinia As String = ""
        Dim eNowaLinia As String = ""
        Dim LiTXT As String = ""
        Dim ileLi As Integer = 0
        Dim CrLfPos As Integer = 0

        Dim i As Integer = 0

        If Len(eTxt) > 0 Then
            CrLfPos = InStr(eTxt, vbCrLf)
            Do While CrLfPos > 0
                eWiersz = Mid(eTxt, 1, CrLfPos - 1)
                eTxt = Mid(eTxt, CrLfPos + 2)
                eLinia = ""
                '---------------------
                'analizuj ten wiersz eWiersz
                Do While Len(eWiersz) > 0
                    If pp.Graphics.MeasureString(eLinia & Mid(eWiersz, 1, 1), FontPola).Width > DlugoscPola Then
                        'czy ostatni znak to znak na którym mo¿na przenieœæ
                        If InStr(" .,/\-=", Mid(eLinia, 1, 1)) > 0 Then
                            Exit Do
                        Else
                            'cofaj sie i sprawdz czy jest znak na którym mo¿na przenieœæ do nowego wiersza
                            If Len(eLinia) > 0 Then
                                For i = Len(eLinia) To 1 Step -1
                                    If InStr(" .,/\-=", Mid(eLinia, i, 1)) > 0 Then
                                        'ok - na tym znaku mozna przeniesc
                                        If i = 1 Then
                                        Else
                                            eWiersz = Mid(eLinia, i, 2) & eWiersz
                                            eLinia = Mid(eLinia, 1, i - 1)
                                        End If
                                        Exit For
                                    Else
                                    End If
                                Next
                            End If
                            ileLi += 1
                            eLinia = ""
                        End If
                    Else
                        eLinia &= Mid(eWiersz, 1, 1)
                        eWiersz = Mid(eWiersz, 2)
                    End If
                Loop
                '---------------------
                CrLfPos = InStr(eTxt, vbCrLf)
            Loop

            'analizuj wiersz bez CRLF
            eWiersz = eTxt
            eLinia = ""
            'analizuj ten wiersz eWiersz
            Do While Len(eWiersz) > 0
                If pp.Graphics.MeasureString(eLinia & Mid(eWiersz, 1, 1), FontPola).Width > DlugoscPola Then
                    'czy ostatni znak to znak na którym mo¿na przenieœæ
                    If InStr(" .,/\-=", Mid(eLinia, 1, 1)) > 0 Then
                        Exit Do
                    Else
                        'cofaj sie i sprawdz czy jest znak na którym mo¿na przenieœæ do nowego wiersza
                        If Len(eLinia) > 0 Then
                            For i = Len(eLinia) To 1 Step -1
                                If InStr(" .,/\-=", Mid(eLinia, i, 1)) > 0 Then
                                    'ok - na tym znaku mozna przeniesc
                                    If i = 1 Then
                                    Else
                                        eWiersz = Mid(eLinia, i, 2) & eWiersz
                                        eLinia = Mid(eLinia, 1, i - 1)
                                    End If
                                    Exit For
                                Else
                                End If
                            Next
                        End If
                        ileLi += 1
                        eLinia = ""
                    End If
                Else
                    eLinia &= Mid(eWiersz, 1, 1)
                    eWiersz = Mid(eWiersz, 2)
                End If
            Loop
            '---------------------
        End If
        Return (ileLi)
    End Function








    Public Function LiniaTxtNr(ByVal txt As String, _
                               ByVal NrLi As Integer) As String
        Dim eTxt As String = txt & vbCrLf
        Dim LiTXT As String = ""
        Dim ileLi As Integer = 0
        Dim CrLfPos As Integer = 0

        If Len(eTxt) > 0 Then
            CrLfPos = InStr(eTxt, vbCrLf)
            If CrLfPos > 0 Then
                Do While CrLfPos > 0
                    ileLi += 1
                    If ileLi = NrLi Then
                        LiTXT = Mid(eTxt, 1, CrLfPos - 1)
                    Else
                        eTxt = Mid(eTxt, CrLfPos + 2)
                        CrLfPos = InStr(eTxt, vbCrLf)
                    End If
                Loop
            Else
            End If
        End If
        Return (LiTXT)
    End Function

    Public Function LiniaTxtNr(ByVal txt As String, _
                               ByVal NrLi As Integer, _
                               ByVal DlugoscPola As Long, _
                               ByVal FontPola As Font, _
                               ByVal pp As PrintPageEventArgs) As String
        Dim eTxt As String = txt & vbCrLf
        Dim eWiersz As String = ""
        Dim eLinia As String = ""
        Dim eNowaLinia As String = ""
        Dim LiTXT As String = ""
        Dim ileLi As Integer = 0
        Dim CrLfPos As Integer = 0

        Dim i As Integer = 0

        If Len(eTxt) > 0 Then
            CrLfPos = InStr(eTxt, vbCrLf)
            Do While CrLfPos > 0
                eWiersz = Mid(eTxt, 1, CrLfPos - 1)
                eTxt = Mid(eTxt, CrLfPos + 2)
                eLinia = ""
                '---------------------
                'analizuj ten wiersz eWiersz
                Do While Len(eWiersz) > 0
                    If pp.Graphics.MeasureString(eLinia & Mid(eWiersz, 1, 1), FontPola).Width > DlugoscPola Then
                        'czy ostatni znak to znak na którym mo¿na przenieœæ
                        If InStr(" .,/\-=", Mid(eLinia, 1, 1)) > 0 Then
                            Exit Do
                        Else
                            'cofaj sie i sprawdz czy jest znak na którym mo¿na przenieœæ do nowego wiersza
                            If Len(eLinia) > 0 Then
                                For i = Len(eLinia) To 1 Step -1
                                    If InStr(" .,/\-=", Mid(eLinia, i, 1)) > 0 Then
                                        'ok - na tym znaku mozna przeniesc
                                        If i = 1 Then
                                        Else
                                            eWiersz = Mid(eLinia, i, 2) & eWiersz
                                            eLinia = Mid(eLinia, 1, i - 1)
                                        End If
                                        Exit For
                                    Else
                                    End If
                                Next
                            End If
                            ileLi += 1
                            If ileLi = NrLi Then
                                LiTXT = eLinia
                            End If
                            eLinia = ""
                        End If
                    Else
                        eLinia &= Mid(eWiersz, 1, 1)
                        eWiersz = Mid(eWiersz, 2)
                    End If
                Loop
                '---------------------
                CrLfPos = InStr(eTxt, vbCrLf)
            Loop

            'analizuj wiersz bez CRLF
            eWiersz = eTxt
            eLinia = ""
            'analizuj ten wiersz eWiersz
            Do While Len(eWiersz) > 0
                If pp.Graphics.MeasureString(eLinia & Mid(eWiersz, 1, 1), FontPola).Width > DlugoscPola Then
                    'czy ostatni znak to znak na którym mo¿na przenieœæ
                    If InStr(" .,/\-=", Mid(eLinia, 1, 1)) > 0 Then
                        Exit Do
                    Else
                        'cofaj sie i sprawdz czy jest znak na którym mo¿na przenieœæ do nowego wiersza
                        If Len(eLinia) > 0 Then
                            For i = Len(eLinia) To 1 Step -1
                                If InStr(" .,/\-=", Mid(eLinia, i, 1)) > 0 Then
                                    'ok - na tym znaku mozna przeniesc
                                    If i = 1 Then
                                    Else
                                        eWiersz = Mid(eLinia, i, 2) & eWiersz
                                        eLinia = Mid(eLinia, 1, i - 1)
                                    End If
                                    Exit For
                                Else
                                End If
                            Next
                        End If
                        ileLi += 1
                        If ileLi = NrLi Then
                            LiTXT = eLinia
                        End If
                        eLinia = ""
                    End If
                Else
                    eLinia &= Mid(eWiersz, 1, 1)
                    eWiersz = Mid(eWiersz, 2)
                End If
            Loop
            '---------------------
        End If
        Return (LiTXT)
    End Function



















    Public Sub CentrTxt(ByVal TrescPola As String, _
                         ByVal posX As Single, _
                         ByVal posY As Single, _
                         ByVal DlugoscPola As Long, _
                         ByVal FontPola As Font, _
                         ByVal pp As PrintPageEventArgs)

        Dim DrawPen As New Pen(Color.Black, 1)
        Dim DlugoscTekstu As Long = pp.Graphics.MeasureString(TrescPola, FontPola).Width
        If DlugoscTekstu >= DlugoscPola Then
            Do While pp.Graphics.MeasureString(TrescPola, FontPola).Width > DlugoscPola
                TrescPola = Mid(TrescPola, 1, Len(TrescPola) - 1)
            Loop
            pp.Graphics.DrawString(TrescPola, FontPola, Brushes.Black, posX, posY)
        Else
            pp.Graphics.DrawString(TrescPola, FontPola, Brushes.Black, posX + Int((DlugoscPola - DlugoscTekstu) / 2), posY)
        End If
    End Sub

    Public Sub LeftTxt(ByVal TrescPola As String, _
                        ByVal posX As Single, _
                        ByVal posY As Single, _
                        ByVal DlugoscPola As Long, _
                        ByVal FontPola As Font, _
                        ByVal pp As PrintPageEventArgs)

        Dim DrawPen As New Pen(Color.Black, 1)
        Dim DlugoscTekstu As Long = pp.Graphics.MeasureString(TrescPola, FontPola).Width
        If DlugoscTekstu >= DlugoscPola Then
            Do While pp.Graphics.MeasureString(TrescPola, FontPola).Width > DlugoscPola
                TrescPola = Mid(TrescPola, 1, Len(TrescPola) - 1)
            Loop
            pp.Graphics.DrawString(TrescPola, FontPola, Brushes.Black, posX, posY)
        Else
            pp.Graphics.DrawString(TrescPola, FontPola, Brushes.Black, posX, posY)
        End If
    End Sub

    Public Sub RightTxt(ByVal TrescPola As String, _
                         ByVal posX As Single, _
                         ByVal posY As Single, _
                         ByVal DlugoscPola As Long, _
                         ByVal FontPola As Font, _
                         ByVal pp As PrintPageEventArgs)

        Dim DrawPen As New Pen(Color.Black, 1)
        Dim DlugoscTekstu As Long = pp.Graphics.MeasureString(TrescPola, FontPola).Width
        If DlugoscTekstu >= DlugoscPola Then
            Do While pp.Graphics.MeasureString(TrescPola, FontPola).Width > DlugoscPola
                TrescPola = Mid(TrescPola, 1, Len(TrescPola) - 1)
            Loop
            pp.Graphics.DrawString(TrescPola, FontPola, Brushes.Black, posX, posY)
        Else
            pp.Graphics.DrawString(TrescPola, FontPola, Brushes.Black, posX + Int(DlugoscPola - DlugoscTekstu), posY)
        End If
    End Sub

End Module
