Module modFiltrowanie
    Public SoftartFilterBackColor As Color = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))


    Public Sub InicjujFiltryColumn(ByVal myForm As Control, _
                                   ByVal myDataGrid As DataGrid, _
                                   ByVal myColumnNo As Integer, _
                                   ByVal myStatusBar As StatusBar, _
                                   ByVal FormLoad As Boolean)
        PokazFiltryColumn(myForm, myDataGrid, myColumnNo, myStatusBar, FormLoad)
        'inicjowanie wartoœci filtrow...
        Dim tBox As TextBox = Nothing
        Dim tCmd As Button = Nothing
        Dim LiCol As Integer = 0
        For LiCol = 0 To myColumnNo - 1
            tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
            tBox.BackColor = System.Drawing.SystemColors.Window
            tBox.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
            tBox.ForeColor = System.Drawing.Color.Navy
            tBox.BorderStyle = BorderStyle.Fixed3D
            tBox.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or _
                                 System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
            tBox.Text = ""
        Next
    End Sub











    '*************************************************
    ' Pokaz filtry
    '*************************************************

    Public Sub PokazFiltryColumn(ByVal myForm As Control, _
                         ByVal myDataGrid As DataGrid, _
                         ByVal myColumnNo As Integer, _
                         ByVal myStatusBar As StatusBar, _
                         ByVal FormLoad As Boolean)
        PokazFiltryColumn(myForm, myDataGrid, myColumnNo, myStatusBar, FormLoad, 0, 0)
    End Sub


    Public Sub PokazFiltryColumn(ByVal myForm As Control, _
                         ByVal myDataGrid As DataGrid, _
                         ByVal myColumnNo As Integer, _
                         ByVal myStatusBar As StatusBar, _
                         ByVal FormLoad As Boolean, _
                         ByVal offsetX As Integer, _
                         ByVal offsetY As Integer)
        If Not FormLoad Then
            'Dim IleCol As Integer = DataViewCennik.Table.Columns.Count
            Dim IleCol As Integer = myColumnNo
            Dim LiCol As Integer = 0
            Dim ThisColWidth As Integer = 0
            Dim ThisColLeft As Integer = 0
            Dim ThisColRight As Integer = 0
            Dim SumColWidth As Integer = 0
            Dim tBox As TextBox = Nothing
            Dim tCmd As Button = Nothing

            Dim LeftVArea As Integer = myStatusBar.Location.X + myStatusBar.Panels(0).Width
            Dim RightVArea As Integer = myStatusBar.Location.X + myStatusBar.Width - 20


            Dim dv As DataView = myDataGrid.DataSource
            If dv.Count = 0 Then
                'tabela jest pusta...
                For LiCol = 0 To myColumnNo - 1
                    ThisColWidth = myDataGrid.TableStyles(0).GridColumnStyles(LiCol).Width
                    tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                    tBox.Location = New Point(LeftVArea + offsetX, myStatusBar.Location.Y + 22 + offsetY)
                    tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
                    tBox.Visible = False
                    tBox.BringToFront()

                    tCmd = SzukajButtonONazwie(myForm, "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                    tCmd.Location = New Point(LeftVArea + ThisColWidth - 16 + offsetX, myStatusBar.Location.Y + offsetY)
                    tCmd.Size = New System.Drawing.Size(16, 22)
                    tCmd.Visible = False
                    tCmd.BringToFront()
                Next
            Else
                'analizuj kolumny przed obszarem widzialnym
                If myDataGrid.FirstVisibleColumn > 0 Then
                    For LiCol = 0 To myDataGrid.FirstVisibleColumn - 1
                        ThisColWidth = myDataGrid.TableStyles(0).GridColumnStyles(LiCol).Width

                        'ThisColWidth = myDataGrid.GetCellBounds(0, LiCol).Width
                        tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                        tBox.Location = New Point(LeftVArea + offsetX, myStatusBar.Location.Y + 22 + offsetY)
                        tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
                        tBox.Visible = False
                        tBox.BringToFront()

                        tCmd = SzukajButtonONazwie(myForm, "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                        tCmd.Location = New Point(LeftVArea + ThisColWidth - 16 + offsetX, myStatusBar.Location.Y + 22 + offsetY)
                        tCmd.Size = New System.Drawing.Size(16, 22)
                        tCmd.Visible = False
                        tCmd.BringToFront()
                    Next
                End If


                'analizuj kolumny widzialne
                For LiCol = myDataGrid.FirstVisibleColumn To myDataGrid.FirstVisibleColumn + myDataGrid.VisibleColumnCount - 1
                    ThisColWidth = myDataGrid.TableStyles(0).GridColumnStyles(LiCol).Width - 1
                    ThisColLeft = myDataGrid.GetCellBounds(0, LiCol).Left
                    ThisColRight = myDataGrid.GetCellBounds(0, LiCol).Right

                    tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))

                    If myDataGrid.Left + ThisColRight <= LeftVArea Or _
                       myDataGrid.Left + ThisColLeft >= RightVArea Then
                        'jesteœmy poza obszarem widzialnym 
                        tBox.Location = New Point(LeftVArea + offsetX, _
                                                  myStatusBar.Location.Y + 22 + offsetY)
                        tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
                        tBox.Visible = False
                        tBox.BringToFront()

                        tCmd = SzukajButtonONazwie(myForm, "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                        tCmd.Location = New Point(LeftVArea + ThisColWidth - 16 + offsetX, _
                                                    myStatusBar.Location.Y + 22 + offsetY)
                        tCmd.Size = New System.Drawing.Size(16, 22)
                        tCmd.Visible = False
                        tCmd.BringToFront()
                    Else
                        'czy kolumna zaczyna sie przed czescia widzialn¹
                        If myDataGrid.Left + ThisColLeft < LeftVArea Then
                            ThisColWidth -= LeftVArea - ThisColLeft - myDataGrid.Left
                        End If
                        'czy kolumna koñczy siê za czêœci¹ wudzialn¹
                        If myDataGrid.Left + ThisColRight > RightVArea Then
                            ThisColWidth -= ThisColRight + myDataGrid.Left - RightVArea
                        End If

                        tBox.Location = New Point(LeftVArea + SumColWidth + offsetX, _
                                                  myStatusBar.Location.Y + 1 + offsetY)
                        tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
                        tBox.Visible = True
                        tBox.BringToFront()

                        tCmd = SzukajButtonONazwie(myForm, "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                        tCmd.Location = New Point(LeftVArea + SumColWidth + ThisColWidth - 16 + offsetX, _
                                                 myStatusBar.Location.Y + offsetY)
                        tCmd.Size = New System.Drawing.Size(16, 22)
                        If tBox.Width = 0 Then
                            tCmd.Visible = False
                        Else
                            tCmd.Visible = True
                        End If
                        tCmd.BringToFront()

                        SumColWidth += ThisColWidth + 1
                        'System.Windows.Forms.Application.DoEvents()
                    End If
                Next

                'analizuj kolumny za obszarem widzialnym
                If myDataGrid.FirstVisibleColumn + myDataGrid.VisibleColumnCount < IleCol - 1 Then
                    For LiCol = myDataGrid.FirstVisibleColumn + myDataGrid.VisibleColumnCount To IleCol - 1
                        ThisColWidth = myDataGrid.GetCellBounds(0, LiCol).Width
                        tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                        tBox.Location = New Point(LeftVArea + offsetX, _
                                                myStatusBar.Location.Y + 22 + offsetY)
                        tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
                        tBox.Visible = False
                        tBox.BringToFront()

                        tCmd = SzukajButtonONazwie(myForm, "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                        tCmd.Location = New Point(LeftVArea + ThisColWidth - 16 + offsetX, _
                                                myStatusBar.Location.Y + 22 + offsetY)
                        tCmd.Size = New System.Drawing.Size(16, 22)
                        tCmd.Visible = False
                        tCmd.BringToFront()
                    Next
                End If
            End If
        End If
        'System.Windows.Forms.Application.DoEvents()
    End Sub



    'Public Sub PokazFiltryColumn(ByVal myForm As Control, _
    '                     ByVal myDataGrid As DataGrid, _
    '                     ByVal myColumnNo As Integer, _
    '                     ByVal myStatusBar As StatusBar, _
    '                     ByVal FormLoad As Boolean)
    '    If Not FormLoad Then
    '        'Dim IleCol As Integer = DataViewCennik.Table.Columns.Count
    '        Dim IleCol As Integer = myColumnNo
    '        Dim LiCol As Integer = 0
    '        Dim ThisColWidth As Integer = 0
    '        Dim ThisColLeft As Integer = 0
    '        Dim ThisColRight As Integer = 0
    '        Dim SumColWidth As Integer = 0
    '        Dim tBox As TextBox = Nothing
    '        Dim tCmd As Button = Nothing

    '        Dim LeftVArea As Integer = myStatusBar.Location.X + myStatusBar.Panels(0).Width
    '        Dim RightVArea As Integer = myStatusBar.Location.X + myStatusBar.Width - 20


    '        Dim dv As DataView = myDataGrid.DataSource
    '        If dv.Count = 0 Then
    '            '    'tabela jest pusta...
    '            '    For LiCol = 0 To myColumnNo - 1
    '            '        ThisColWidth = myDataGrid.TableStyles(0).GridColumnStyles(LiCol).Width
    '            '        tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
    '            '        tBox.Location = New Point(LeftVArea, myStatusBar.Location.Y + 22)
    '            '        tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
    '            '        tBox.Visible = False
    '            '        tBox.BringToFront()

    '            '        tCmd = SzukajButtonONazwie(myForm, "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
    '            '        tCmd.Location = New Point(LeftVArea + ThisColWidth - 16, myStatusBar.Location.Y)
    '            '        tCmd.Size = New System.Drawing.Size(16, 22)
    '            '        tCmd.Visible = False
    '            '        tCmd.BringToFront()
    '            '    Next
    '        Else
    '            'analizuj kolumny przed obszarem widzialnym
    '            If myDataGrid.FirstVisibleColumn > 0 Then
    '                For LiCol = 0 To myDataGrid.FirstVisibleColumn - 1
    '                    ThisColWidth = myDataGrid.TableStyles(0).GridColumnStyles(LiCol).Width

    '                    'ThisColWidth = myDataGrid.GetCellBounds(0, LiCol).Width
    '                    tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
    '                    tBox.Location = New Point(LeftVArea, myStatusBar.Location.Y + 22)
    '                    tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
    '                    tBox.Visible = False
    '                    tBox.BringToFront()

    '                    tCmd = SzukajButtonONazwie(myForm, "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
    '                    tCmd.Location = New Point(LeftVArea + ThisColWidth - 16, myStatusBar.Location.Y)
    '                    tCmd.Size = New System.Drawing.Size(16, 22)
    '                    tCmd.Visible = False
    '                    tCmd.BringToFront()
    '                Next
    '            End If


    '            'analizuj kolumny widzialne
    '            For LiCol = myDataGrid.FirstVisibleColumn To myDataGrid.FirstVisibleColumn + myDataGrid.VisibleColumnCount - 1
    '                ThisColWidth = myDataGrid.TableStyles(0).GridColumnStyles(LiCol).Width - 1
    '                ThisColLeft = myDataGrid.GetCellBounds(0, LiCol).Left
    '                ThisColRight = myDataGrid.GetCellBounds(0, LiCol).Right

    '                tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))

    '                If myDataGrid.Left + ThisColRight <= LeftVArea Or _
    '                   myDataGrid.Left + ThisColLeft >= RightVArea Then
    '                    'jesteœmy poza obszarem widzialnym 
    '                    tBox.Location = New Point(LeftVArea, _
    '                                              myStatusBar.Location.Y + 22)
    '                    tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
    '                    tBox.Visible = False
    '                    tBox.BringToFront()

    '                    tCmd = SzukajButtonONazwie(myForm, "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
    '                    tCmd.Location = New Point(LeftVArea + ThisColWidth - 16, myStatusBar.Location.Y)
    '                    tCmd.Size = New System.Drawing.Size(16, 22)
    '                    tCmd.Visible = False
    '                    tCmd.BringToFront()
    '                Else
    '                    'czy kolumna zaczyna sie przed czescia widzialn¹
    '                    If myDataGrid.Left + ThisColLeft < LeftVArea Then
    '                        ThisColWidth -= LeftVArea - ThisColLeft - myDataGrid.Left
    '                    End If
    '                    'czy kolumna koñczy siê za czêœci¹ wudzialn¹
    '                    If myDataGrid.Left + ThisColRight > RightVArea Then
    '                        ThisColWidth -= ThisColRight + myDataGrid.Left - RightVArea
    '                    End If

    '                    tBox.Location = New Point(LeftVArea + SumColWidth, _
    '                                              myStatusBar.Location.Y + 2)
    '                    tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
    '                    tBox.Visible = True
    '                    tBox.BringToFront()

    '                    tCmd = SzukajButtonONazwie(myForm, "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
    '                    tCmd.Location = New Point(LeftVArea + SumColWidth + ThisColWidth - 16, myStatusBar.Location.Y)
    '                    tCmd.Size = New System.Drawing.Size(16, 22)
    '                    If tBox.Width = 0 Then
    '                        tCmd.Visible = False
    '                    Else
    '                        tCmd.Visible = True
    '                    End If
    '                    tCmd.BringToFront()

    '                    SumColWidth += ThisColWidth + 1
    '                End If
    '            Next

    '            'analizuj kolumny za obszarem widzialnym
    '            If myDataGrid.FirstVisibleColumn + myDataGrid.VisibleColumnCount < IleCol - 1 Then
    '                For LiCol = myDataGrid.FirstVisibleColumn + myDataGrid.VisibleColumnCount To IleCol - 1
    '                    ThisColWidth = myDataGrid.GetCellBounds(0, LiCol).Width
    '                    tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
    '                    tBox.Location = New Point(LeftVArea, myStatusBar.Location.Y + 22)
    '                    tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
    '                    tBox.Visible = False
    '                    tBox.BringToFront()

    '                    tCmd = SzukajButtonONazwie(myForm, "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
    '                    tCmd.Location = New Point(LeftVArea + ThisColWidth - 16, myStatusBar.Location.Y)
    '                    tCmd.Size = New System.Drawing.Size(16, 22)
    '                    tCmd.Visible = False
    '                    tCmd.BringToFront()
    '                Next
    '            End If
    '        End If
    '    End If
    'End Sub








    '*************************************************
    ' Szukaj kontrolek na formie
    '*************************************************

    ' Visual Basic 2005
    Public Function SzukajTextBoxONazwie(ByVal container As Control, _
                                     ByVal Nazwa As String) As Windows.Forms.TextBox
        Dim RetCtrl As Windows.Forms.TextBox = Nothing
        Dim ctrl As Control
        For Each ctrl In container.Controls
            If TypeOf (ctrl) Is TextBox Then
                If ctrl.Name = Nazwa Then
                    RetCtrl = ctrl
                    Exit For
                End If
            End If
        Next
        Return RetCtrl
    End Function

    Public Function SzukajButtonONazwie(ByVal container As Control, _
                                 ByVal Nazwa As String) As Button
        Dim RetCtrl As Button = Nothing
        Dim ctrl As Control
        For Each ctrl In container.Controls
            If TypeOf (ctrl) Is Button Then
                If ctrl.Name = Nazwa Then
                    RetCtrl = ctrl
                End If
            End If
        Next
        Return RetCtrl
    End Function


    '*************************************************
    ' edycja pola TxtBox
    '*************************************************

    Public Sub StartEdycjiFiltraTxt(ByVal Pole As TextBox)
        Pole.BackColor = SoftartCursorBackColor
        Pole.ForeColor = SoftartCursorForeColor
    End Sub
    Public Sub KoniecEdycjiFiltraTxt(ByVal Pole As TextBox)
        If Len(Trim(Pole.Text)) > 0 Then
            Pole.BackColor = SoftartFilterBackColor
            Pole.ForeColor = SoftartEditableForeColor
            'RysujGradientTextBox(Pole)
        Else
            Pole.BackColor = SoftartEditableBackColor
            Pole.ForeColor = SoftartEditableForeColor
        End If
    End Sub


    '*************************************************
    ' Edycja Button'a
    '*************************************************

    Public Sub StartEdycjiFiltraCmd(ByVal myForm As Control, _
                                    ByVal Pole As Button)
        Dim NazwaTxt As String = ""
        Dim tBox As TextBox = Nothing
        NazwaTxt = "txtFi" & Mid(Pole.Name, 6, 3)
        tBox = SzukajTextBoxONazwie(myForm, NazwaTxt)

        tBox.BackColor = SoftartCursorBackColor
        tBox.ForeColor = SoftartCursorForeColor
    End Sub
    Public Sub KoniecEdycjiFiltraCmd(ByVal myForm As Control, _
                                     ByVal Pole As Button)
        Dim NazwaTxt As String = ""
        Dim tBox As TextBox = Nothing
        NazwaTxt = "txtFi" & Mid(Pole.Name, 6, 3)
        tBox = SzukajTextBoxONazwie(myForm, NazwaTxt)

        If Len(Trim(tBox.Text)) > 0 Then
            tBox.BackColor = SoftartFilterBackColor
            tBox.ForeColor = SoftartEditableForeColor
            'RysujGradientTextBox(tBox)
        Else
            tBox.BackColor = SoftartEditableBackColor
            tBox.ForeColor = SoftartEditableForeColor
        End If
    End Sub






    'Public Sub RysujGradientTextBox(ByVal meTxtBx As TextBox)
    '    ' gradient w tle formy
    '    Dim G As Graphics
    '    G = meTxtBx.CreateGraphics
    '    Dim R As New RectangleF(0, 0, meTxtBx.Width, meTxtBx.Height)
    '    Dim startColor As Color = System.Drawing.SystemColors.Control
    '    Dim EndColor As Color = System.Drawing.SystemColors.ControlDark

    '    Dim LGBrush As New System.Drawing.Drawing2D.LinearGradientBrush _
    '            (R, startColor, EndColor, System.Drawing.Drawing2D.LinearGradientMode.Horizontal)
    '    G.FillRectangle(LGBrush, New Rectangle(0, 0, meTxtBx.Width, meTxtBx.Height))
    'End Sub


    '*************************************************
    ' Wartoœæ filtra
    '*************************************************

    Public Sub PobierzWartoscFiltra(ByVal myForm As Control, _
                                    ByVal myDataGrid As DataGrid, _
                                    ByVal myColumnNo As Integer, _
                                    ByVal DBTableName As String)

        'wyczysc filtr jesli istnieje
        Dim tBox As TextBox = Nothing
        tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(myColumnNo)), 3))
        tBox.Text = ""

        Dim NazwaPola As String = myDataGrid.TableStyles(0).GridColumnStyles(myColumnNo).MappingName
        Dim NazwaDV As DataView = myDataGrid.DataSource
        Dim FiltrDV As String = NazwaDV.RowFilter

        'mamy aliasy pol - trzeba przetlumaczyc na rzeczywiste nazyw pol
        Select Case DBTableName
            Case "Klienci"
                NazwaPola = OryginalnePolaTabeliKlienci(NazwaPola)
            Case Else
        End Select


        'trzeba powymieniac aliasy w filtrze....
        TestAliasy(FiltrDV, "NRKARTY", "IDENTKLIENTA")
        TestAliasy(FiltrDV, "KATA", "KATKLIENTAA")
        TestAliasy(FiltrDV, "KATB", "KATKLIENTAB")
        TestAliasy(FiltrDV, "KATC", "KATKLIENTAC")
        TestAliasy(FiltrDV, "NAZWAKLIENTA", "NAZWA1")
        TestAliasy(FiltrDV, "REJONKLIENTA", "REJKLIENTA")
        TestAliasy(FiltrDV, "STRONAWWW", "WWW")
        TestAliasy(FiltrDV, "SPRZETILOSC", "ILOSPRZETU2")
        TestAliasy(FiltrDV, "SKUPTYDZIEN", "SKUPWARTOSC")
        TestAliasy(FiltrDV, "PROMOTOR", "SKUPOPIEKUN")
        TestAliasy(FiltrDV, "PROMOTOROSTKONTAKTROK", "SKUPOKONTAKTR")
        TestAliasy(FiltrDV, "PROMOTOROSTKONTAKTTYDZ", "SKUPOKONTAKTT")
        TestAliasy(FiltrDV, "PROMOTOROSTKONTAKTDZIEN", "SKUPOKONTAKTD")
        TestAliasy(FiltrDV, "PROMOTORKOLKONTAKTROK", "SKUPNKONTAKTR")
        TestAliasy(FiltrDV, "PROMOTORKOLKONTAKTTYDZ", "SKUPNKONTAKTT")
        TestAliasy(FiltrDV, "PROMOTORKOLKONTAKTDZIEN", "SKUPNKONTAKTD")
        TestAliasy(FiltrDV, "PROMOTORCODALEJ", "SKUPPLAN")
        TestAliasy(FiltrDV, "PROMOTORUWAGI", "SKUPUWAGI")
        TestAliasy(FiltrDV, "OPIEKUN", "SPRZEDOPIEKUN")
        TestAliasy(FiltrDV, "SPRZEDAZ", "SPRZEDWARTOSC")
        TestAliasy(FiltrDV, "OPIEKUNOSTKONTAKTROK", "SPRZEDOKONTAKTR")
        TestAliasy(FiltrDV, "OPIEKUNOSTKONTAKTTYDZ", "SPRZEDOKONTAKTT")
        TestAliasy(FiltrDV, "OPIEKUNOSTKONTAKTDZIEN", "SPRZEDOKONTAKTD")
        TestAliasy(FiltrDV, "OPIEKUNKOLKONTAKTROK", "SPRZEDNKONTAKTR")
        TestAliasy(FiltrDV, "OPIEKUNKOLKONTAKTTYDZ", "SPRZEDNKONTAKTT")
        TestAliasy(FiltrDV, "OPIEKUNKOLKONTAKTDZIEN", "SPRZEDNKONTAKTD")
        TestAliasy(FiltrDV, "OPIEKUNUWAGI", "SPRZEDUWAGI")
        TestAliasy(FiltrDV, "WAGAKLIENTA", "SPRZEDWAGA")
        TestAliasy(FiltrDV, "TEST", "SPRZEDTEST")


        oWybralem = ""
        Dim Form As New FiltrowanieWybierzWartoscPola(NazwaPola, DBTableName, FiltrDV)
        Form.ShowDialog()
        If Len(oWybralem) > 0 Then
            'wpisz wybrany tekst do filtra...
            tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(myColumnNo)), 3))
            tBox.Text = RTrim(oWybralem.ToString)
        End If
    End Sub


    '***********************************************************
    '** test aliasu w filtrze
    '***********************************************************

    Public Sub TestAliasy(ByRef Filtr As String, _
                          ByVal AliasName As String, _
                          ByVal RealName As String)
        Dim pos As Integer
        pos = InStr(Filtr, AliasName)
        'Do While pos > 0
        '    If pos = 1 Then
        '        Filtr = RealName & Mid(Filtr, pos + Len(AliasName))
        '    Else
        '        Filtr = Mid(Filtr, 1, pos - 1) & RealName & Mid(Filtr, pos + Len(AliasName))
        '    End If
        '    pos = InStr(Filtr, AliasName)
        'Loop

        If pos > 0 Then
            If pos = 1 Then
                Filtr = RealName & Mid(Filtr, pos + Len(AliasName))
            Else
                Filtr = Mid(Filtr, 1, pos - 1) & RealName & Mid(Filtr, pos + Len(AliasName))
            End If
            'pos = InStr(Filtr, AliasName)
        End If
    End Sub




    '***********************************************************
    '** Czyszczenie filtra
    '***********************************************************

    Public Sub CzyscFiltry(ByVal myForm As Control, _
                               ByVal myDataGrid As DataGrid, _
                               ByVal myColumnNo As Integer, _
                               ByVal myDataView As DataView, _
                               ByVal myStatusBar As StatusBar)

        Dim tBox As TextBox = Nothing
        Dim LiCol As Integer = 0

        oLokalnyFiltr = ""
        myDataView.RowFilter = oLokalnyFiltr
        If myDataView.Count > 0 Then
            For LiCol = 0 To myColumnNo - 1
                tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                tBox.Text = ""
                tBox.BackColor = SoftartEditableBackColor
                tBox.ForeColor = SoftartEditableForeColor
            Next

            'Try
            '    myDataView.RowFilter = oLokalnyFiltr
            '    myStatusBar.Panels(0).Text = "Il.rec.: " & myDataView.Count.ToString
            'Catch Ex As System.Exception
            '    MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
            '        System.Windows.Forms.MessageBoxButtons.OK, _
            '        MessageBoxIcon.Information, _
            '        MessageBoxDefaultButton.Button1)
            'Finally
            'End Try
        End If
    End Sub


    Public Sub CzyscFiltry(ByVal myForm As Control, _
                           ByVal myDataGrid As DataGrid, _
                           ByVal myColumnNo As Integer, _
                           ByVal myDataView As DataView, _
                           ByVal myStatusBar As StatusBar, _
                           ByVal myFilter As String)

        Dim tBox As TextBox = Nothing
        Dim LiCol As Integer = 0

        If Len(myFilter) = 0 Then
            oLokalnyFiltr = ""
        Else
            oLokalnyFiltr = myFilter
        End If
        myDataView.RowFilter = oLokalnyFiltr

        If myDataView.Count > 0 Then
            For LiCol = 0 To myColumnNo - 1
                tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                tBox.Text = ""
                tBox.BackColor = SoftartEditableBackColor
                tBox.ForeColor = SoftartEditableForeColor
            Next
        End If
    End Sub



    '***********************************************************
    '** FILTROWANIE
    '***********************************************************

    Public Sub FiltrujDataView(ByVal pForm As Control, _
                               ByVal pDataGrid As DataGrid, _
                               ByVal pColumnNo As Integer, _
                               ByVal pDataView As DataView, _
                               ByVal pStatusBar As StatusBar)
        FiltrujDV(pForm, pDataGrid, pColumnNo, pDataView, pStatusBar, "")
    End Sub

    Public Sub FiltrujDataView(ByVal pForm As Control, _
                               ByVal pDataGrid As DataGrid, _
                               ByVal pColumnNo As Integer, _
                               ByVal pDataView As DataView, _
                               ByVal pStatusBar As StatusBar, _
                               ByVal pFilter As String)
        FiltrujDV(pForm, pDataGrid, pColumnNo, pDataView, pStatusBar, pFilter)
    End Sub


    Private Sub FiltrujDV(ByVal myForm As Control, _
                          ByVal myDataGrid As DataGrid, _
                          ByVal myColumnNo As Integer, _
                          ByVal myDataView As DataView, _
                          ByVal myStatusBar As StatusBar, _
                          ByVal myFilter As String)
        Dim tBox As TextBox = Nothing
        Dim LiCol As Integer = 0
        Dim Separ As String = ""

        Dim words() As String
        Dim splitChars() As Char = {"|"c}
        Dim iwords As Integer = 0
        Dim SepPoz As Integer = 0
        Dim TxtOd As String = ""
        Dim TxtDo As String = ""

        If Len(myFilter) = 0 Then
            oLokalnyFiltr = ""
            Separ = ""
        Else
            oLokalnyFiltr = myFilter
            Separ = " AND "
        End If

        myDataView.RowFilter = oLokalnyFiltr
        If myDataView.Count > 0 Then
            For LiCol = 0 To myColumnNo - 1
                tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))

                If Len(Trim(tBox.Text)) > 0 Then
                    tBox.BackColor = SoftartFilterBackColor
                    'RysujGradientTextBox(tBox)

                    If TypeOf myDataView.Item(0).Item(LiCol) Is String Then
                        'podziel linie na poszczególne czêœci
                        'words = tBox.Text.Split(splitChars, StringSplitOptions.RemoveEmptyEntries)
                        'analiza filtrow alternatywnych
                        words = tBox.Text.Split(splitChars, StringSplitOptions.None)
                        If words.Length = 1 Then

                            'jeden warunek...
                            If Left(tBox.Text, 3) = "***" Then
                                'pole puste
                                oLokalnyFiltr &= Separ & _
                                        "(" & myDataView.Table.Columns(LiCol).ColumnName & "='')"

                            ElseIf Left(tBox.Text, 2) = "**" Then
                                'nie zawiera tekstu
                                oLokalnyFiltr &= Separ & _
                                        "(NOT " & myDataView.Table.Columns(LiCol).ColumnName & " LIKE '*" & Mid(tBox.Text, 3) & "*') "

                            Else
                                'sprawdz czy to nie jest przedzia³ danych OD..DO
                                SepPoz = InStr(tBox.Text, "..")
                                If SepPoz > 0 Then
                                    'drukuj przedzia³
                                    TxtOd = Mid(tBox.Text, 1, SepPoz - 1)
                                    TxtDo = Mid(tBox.Text, SepPoz + 2)
                                    If Len(TxtOd) > 0 And Len(TxtDo) > 0 Then
                                        'fuiltruj przedzial
                                        oLokalnyFiltr &= Separ & _
                                                "(" & _
                                                myDataView.Table.Columns(LiCol).ColumnName & ">='" & TxtOd & "' and " & _
                                                myDataView.Table.Columns(LiCol).ColumnName & "<='" & TxtDo & "'" & _
                                                ") "
                                    Else
                                        If Len(TxtOd) > 0 Then
                                            'filtruj od
                                            oLokalnyFiltr &= Separ & _
                                                    "(" & myDataView.Table.Columns(LiCol).ColumnName & ">='" & TxtOd & "')"
                                        Else
                                            'filtruj Do
                                            oLokalnyFiltr &= Separ & _
                                                    "(" & myDataView.Table.Columns(LiCol).ColumnName & "<='" & TxtDo & "')"
                                        End If
                                    End If
                                Else
                                    'drukuj pojedynczy rekord
                                    oLokalnyFiltr &= Separ & _
                                            "(" & myDataView.Table.Columns(LiCol).ColumnName & _
                                            " LIKE '" & tBox.Text & "*')"
                                End If
                            End If
                        Else
                            'wiêcej ni¿ jeden warunek - musi byæspe³niony TEN i TEN i TEN...


                            oLokalnyFiltr &= Separ & "("
                            For iwords = 0 To words.Length - 1
                                'sprawdzamy kolejne warunki (i ten) 
                                If Len(words(iwords)) > 0 Then
                                    If Left(words(iwords), 3) = "***" Then
                                        'pole puste
                                        oLokalnyFiltr &= "(" & myDataView.Table.Columns(LiCol).ColumnName & "='') OR "

                                    ElseIf Left(words(iwords), 2) = "**" Then
                                        'nie zawiera tekstu
                                        oLokalnyFiltr &= " (NOT " & myDataView.Table.Columns(LiCol).ColumnName & _
                                                        " LIKE '*" & Mid(words(iwords), 3) & "*') OR "

                                    Else
                                        'sprawdz czy to nie jest przedzia³ danych OD..DO
                                        SepPoz = InStr(words(iwords), "..")
                                        If SepPoz > 0 Then
                                            'drukuj przedzia³
                                            TxtOd = Mid(words(iwords), 1, SepPoz - 1)
                                            TxtDo = Mid(words(iwords), SepPoz + 2)
                                            If Len(TxtOd) > 0 And Len(TxtDo) > 0 Then
                                                'fuiltruj przedzial
                                                oLokalnyFiltr &= "(" & _
                                                        myDataView.Table.Columns(LiCol).ColumnName & ">='" & TxtOd & "' AND " & _
                                                        myDataView.Table.Columns(LiCol).ColumnName & "<='" & TxtDo & "') OR "
                                            Else
                                                If Len(TxtOd) > 0 Then
                                                    'filtruj od
                                                    oLokalnyFiltr &= "(" & myDataView.Table.Columns(LiCol).ColumnName & ">='" & TxtOd & "') OR "
                                                Else
                                                    'filtruj Do
                                                    oLokalnyFiltr &= "(" & myDataView.Table.Columns(LiCol).ColumnName & "<='" & TxtDo & "') OR "
                                                End If
                                            End If
                                        Else
                                            'drukuj pojedynczy rekord
                                            oLokalnyFiltr &= "(" & myDataView.Table.Columns(LiCol).ColumnName & _
                                                            " LIKE '" & words(iwords) & "*') OR "
                                        End If
                                    End If
                                End If
                            Next
                            'usun ostatnie ' AND '
                            oLokalnyFiltr = Mid(oLokalnyFiltr, 1, Len(oLokalnyFiltr) - 4)
                            oLokalnyFiltr &= ")"
                        End If
                        Separ = " AND "


                    ElseIf TypeOf myDataView.Item(0).Item(LiCol) Is Integer Or _
                           TypeOf myDataView.Item(0).Item(LiCol) Is Int16 Or _
                           TypeOf myDataView.Item(0).Item(LiCol) Is Int32 Or _
                           TypeOf myDataView.Item(0).Item(LiCol) Is Int64 Then
                        'podziel linie na poszczególne czêœci
                        'words = tBox.Text.Split(splitChars, StringSplitOptions.RemoveEmptyEntries)
                        'analiza filtrow alternatywnych
                        words = tBox.Text.Split(splitChars, StringSplitOptions.None)
                        If words.Length = 1 Then
                            If IsNumeric(tBox.Text) Then
                                oLokalnyFiltr &= Separ & _
                                        "(" & myDataView.Table.Columns(LiCol).ColumnName & _
                                        "=" & Trim(Str(tBox.Text)) & ")"
                                Separ = " AND "
                            Else
                            End If
                        Else
                            oLokalnyFiltr &= Separ & "("
                            Separ = ""
                            'analizuj wszystkie alternatywy
                            For iwords = 0 To words.Length - 1
                                If Len(words(iwords)) > 0 Then
                                    If IsNumeric(words(iwords)) Then
                                        oLokalnyFiltr &= Separ & _
                                                "(" & myDataView.Table.Columns(LiCol).ColumnName & _
                                                "=" & Trim(Str(words(iwords))) & ")"
                                        Separ = " OR "
                                    Else
                                    End If
                                End If
                            Next
                            'usun ostatnie OR
                            'oLokalnyFiltr = Mid(oLokalnyFiltr, 1, Len(oLokalnyFiltr) - 4)
                            oLokalnyFiltr &= ")"
                        End If
                        Separ = " AND "












                    ElseIf TypeOf myDataView.Item(0).Item(LiCol) Is DBNull Then
                        oLokalnyFiltr &= Separ & _
                                "(" & myDataView.Table.Columns(LiCol).ColumnName & _
                                "=" & Trim(Str(tBox.Text)) & ")"
                        Separ = " AND "


                    ElseIf TypeOf myDataView.Item(0).Item(LiCol) Is Double Then
                        If IsNumeric(tBox.Text) Then
                            oLokalnyFiltr &= Separ & _
                                    "(" & myDataView.Table.Columns(LiCol).ColumnName & _
                                    "=" & Trim(Str(tBox.Text)) & ")"
                            Separ = " AND "
                        Else
                        End If



                    ElseIf TypeOf myDataView.Item(0).Item(LiCol) Is Boolean Then
                        Select Case UCase(tBox.Text)
                            Case "T", "Y", "TAK", "YES", "TRUE"
                                oLokalnyFiltr &= Separ & _
                                                "(" & myDataView.Table.Columns(LiCol).ColumnName & _
                                                "=1)"
                                Separ = " AND "
                            Case "N", "NIE", "NO", "FALSE"
                                oLokalnyFiltr &= Separ & _
                                                "(" & myDataView.Table.Columns(LiCol).ColumnName & _
                                                "=0)"
                                Separ = " AND "
                            Case Else
                        End Select


                    ElseIf TypeOf myDataView.Item(0).Item(LiCol) Is Guid Then
                        oLokalnyFiltr &= Separ & _
                                "(" & myDataView.Table.Columns(LiCol).ColumnName & _
                                " LIKE '" & tBox.Text.ToString & "*')"
                        Separ = " AND "
                    End If



                Else
                    If tBox.ContainsFocus Then
                        tBox.BackColor = SoftartCursorBackColor
                    Else
                        tBox.BackColor = SoftartEditableBackColor
                    End If
                End If
            Next

            Try
                myDataView.RowFilter = oLokalnyFiltr
                'myDataGrid.DataSource = myDataView
                myStatusBar.Panels(0).Text = "Il.rec.: " & myDataView.Count.ToString
            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            Finally
            End Try

            'If myDataView.Count > 0 Then
            '    'ustaw sie na poczatku DataGrid
            '    myDataGrid.Select(0)
            '    myDataGrid.CurrentCell = New DataGridCell(0, 0)
            'End If
        End If
    End Sub


End Module
