Module _modFiltrowanieDataGridView

    Public oFiltrZlozony As String = ""


    '*************************************************
    ' Wartość filtra w określonej kolumnie
    '*************************************************

    Public Sub PobierzWartoscFiltraDGV(ByVal myForm As Control, _
                                    ByRef myDataGridView As DataGridView, _
                                    ByRef myColumnNo As Integer)
        Dim NazwaPola As String = myDataGridView.Columns(myColumnNo).DataPropertyName
        'Dim NazwaDV As DataView = myDataGridView.DataSource
        'Dim FiltrDV As String = NazwaDV.RowFilter




        oWybralem = ""
        Dim Form As New _frmFiltrowanieWybierzWartoscPola(myDataGridView.DataSource, NazwaPola)
        Using Form
            Form.ShowDialog()
        End Using
        Form = Nothing

        If Len(oWybralem) > 0 Then
            'wpisz wybrany tekst do filtra...
            Dim tBox As System.Windows.Forms.TextBox = Nothing

            'tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(myColumnNo)), 3))
            tBox = myForm.Controls("txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(myColumnNo)), 3))

            If tBox Is Nothing Then
            Else
                tBox.Text = RTrim(oWybralem.ToString)
            End If
        End If
    End Sub

    Public Sub PobierzWartoscFiltraDGV(ByVal myForm As Control, _
                                    ByRef myDataGridView As DataGridView, _
                                    ByRef myColumnNo As Integer, _
                                    ByRef DBTableName As String)
        Dim NazwaPola As String = myDataGridView.Columns(myColumnNo).DataPropertyName

        oWybralem = ""
        Dim Form As New _frmFiltrowanieWybierzWartoscPola(myDataGridView.DataSource, NazwaPola)
        Using Form
            Form.ShowDialog()
        End Using
        Form = Nothing

        If Len(oWybralem) > 0 Then
            'wpisz wybrany tekst do filtra...
            Dim tBox as system.windows.forms.textbox = Nothing

            'tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(myColumnNo)), 3))
            tBox = myForm.Controls("txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(myColumnNo)), 3))

            If tBox Is Nothing Then
            Else
                tBox.Text = RTrim(oWybralem.ToString)
            End If
        End If
    End Sub

    ' dla Bazy Danych innej niż ACCESS i DBF
    Public Sub PobierzWartoscFiltraDGV(ByVal myForm As Control, _
                                ByRef myDataGridView As DataGridView, _
                                ByRef myColumnNo As Integer, _
                                ByRef DBTableName As String, _
                                ByRef DBConnString As String)
        Dim NazwaPola As String = myDataGridView.Columns(myColumnNo).DataPropertyName

        oWybralem = ""
        Dim Form As New _frmFiltrowanieWybierzWartoscPola(myDataGridView.DataSource, NazwaPola)
        Using Form
            Form.ShowDialog()
        End Using
        Form = Nothing

        If Len(oWybralem) > 0 Then
            'wpisz wybrany tekst do filtra...
            Dim tBox as system.windows.forms.textbox = Nothing

            'tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(myColumnNo)), 3))
            tBox = myForm.Controls("txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(myColumnNo)), 3))

            If tBox Is Nothing Then
            Else
                tBox.Text = RTrim(oWybralem.ToString)
            End If
        End If
    End Sub







    '*************************************************
    ' Inicjuj filtry
    '*************************************************

    Public Sub InicjujFiltryColumnDGV(ByVal myForm As Control, _
                                   ByRef myDataGridView As DataGridView, _
                                   ByRef myColumnNo As Integer, _
                                   ByRef myStatusBar As StatusBar, _
                                   ByRef FormLoad As Boolean, _
                                   ByRef offsetX As Integer, _
                                   ByRef offsetY As Integer)
        PokazFiltryColumnDGV(myForm, myDataGridView, myColumnNo, myStatusBar, FormLoad, offsetX, offsetY)
        '__InicjujFiltryDGV(myForm, myColumnNo)
    End Sub

    Public Sub InicjujFiltryColumnDGV(ByVal myForm As Control, _
                                   ByRef myDataGridView As DataGridView, _
                                   ByRef myColumnNo As Integer, _
                                   ByRef myStatusBar As StatusBar, _
                                   ByRef FormLoad As Boolean)
        PokazFiltryColumnDGV(myForm, myDataGridView, myColumnNo, myStatusBar, FormLoad, 0, 0)
        '__InicjujFiltryDGV(myForm, myColumnNo)
    End Sub



    'inicjowanie wartości filtrow...
    Private Sub __InicjujFiltryDGV(ByVal myForm As Control, _
                                   ByRef myColumnNo As Integer)
        Dim tBox as system.windows.forms.textbox = Nothing
        Dim tCmd as system.windows.forms.button = Nothing
        Dim LiCol As Integer = 0
        For LiCol = 0 To myColumnNo - 1
            'tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
            tBox = myForm.Controls("txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))

            If tBox Is Nothing Then
            Else
                tBox.BackColor = System.Drawing.SystemColors.Window
                tBox.Font = New System.Drawing.Font(Softart_FormFontName, _
                                                       Softart_FormFontSize, _
                                                        System.Drawing.FontStyle.Regular, _
                                                        System.Drawing.GraphicsUnit.Point, _
                                                        CType(238, Byte))
                tBox.ForeColor = System.Drawing.Color.Navy
                tBox.BorderStyle = BorderStyle.Fixed3D
                tBox.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or _
                                         System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
                tBox.Text = ""
            End If
        Next
    End Sub







    '*************************************************
    ' Pokaz filtry
    '*************************************************

    Public Sub PokazFiltryColumnDGV(ByVal myForm As Control, _
                         ByRef myDataGridView As DataGridView, _
                         ByRef myColumnNo As Integer, _
                         ByRef myStatusBar As StatusBar, _
                         ByRef FormLoad As Boolean, _
                         ByRef offsetX As Integer, _
                         ByRef offsetY As Integer)
        __PokazFiltryDGV2(myForm, myDataGridView, myColumnNo, myStatusBar, FormLoad, offsetX, offsetY)
    End Sub

    Public Sub PokazFiltryColumnDGV(ByVal myForm As Control, _
                         ByRef myDataGridView As DataGridView, _
                         ByRef myColumnNo As Integer, _
                         ByRef myStatusBar As StatusBar, _
                         ByRef FormLoad As Boolean)
        __PokazFiltryDGV2(myForm, myDataGridView, myColumnNo, myStatusBar, FormLoad, 0, 0)
    End Sub

    ''pokaż filtry na ekranie
    'Public Sub __PokazFiltryDGV(ByVal myForm As Control, _
    '                     ByRef myDataGridView As DataGridView, _
    '                     ByRef myColumnNo As Integer, _
    '                     ByRef myStatusBar As StatusBar, _
    '                     ByRef FormLoad As Boolean, _
    '                     ByRef offsetX As Integer, _
    '                     ByRef offsetY As Integer)
    '    If Not FormLoad Then
    '        'Dim IleCol As Integer = DataViewCennik.Table.Columns.Count
    '        Dim IleCol As Integer = myColumnNo
    '        Dim LiCol As Integer = 0

    '        Dim ThisColWidth As Integer = 0
    '        Dim ThisColLeft As Integer = 0
    '        Dim ThisColRight As Integer = 0

    '        Dim NextColWidth As Integer = 0
    '        Dim NextColLeft As Integer = 0
    '        Dim NextColRight As Integer = 0

    '        Dim SumColWidth As Integer = 0
    '        Dim tBox as system.windows.forms.textbox = Nothing
    '        Dim tCmd as system.windows.forms.button = Nothing

    '        Dim LeftVArea As Integer = myStatusBar.Location.X + myStatusBar.Panels(0).Width
    '        Dim RightVArea As Integer = myStatusBar.Location.X + myStatusBar.Width


    '        Dim dv As DataView = myDataGridView.DataSource
    '        If dv.Count = 0 Then
    '            'tabela jest pusta - nie pokazuj filtra...ma być niewidoczny
    '            For LiCol = 0 To myColumnNo - 1
    '                If myDataGridView.Columns(LiCol).Visible Then
    '                    ThisColWidth = myDataGridView.Columns(LiCol).Width

    '                    tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
    '                    If tBox Is Nothing Then
    '                    Else
    '                        'tBox.Location = New system.drawing.Point(LeftVArea + offsetX, myStatusBar.Location.Y + 1 + offsetY)
    '                        'tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
    '                        tBox.Visible = False
    '                        'tBox.BringToFront()
    '                    End If

    '                    tCmd = SzukajButtonONazwie(myForm, "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
    '                    If tCmd Is Nothing Then
    '                    Else
    '                        'tCmd.Location = New system.drawing.Point(LeftVArea + ThisColWidth - 16 + offsetX, myStatusBar.Location.Y + 1 + offsetY)
    '                        'tCmd.Size = New System.Drawing.Size(16, 22)
    '                        tBox.Visible = False
    '                        'tCmd.BringToFront()
    '                    End If
    '                End If
    '            Next
    '        Else
    '            'sprawdz czy są widzialne komumny...
    '            If (myDataGridView.FirstDisplayedScrollingColumnIndex >= 0) And (myDataGridView.FirstDisplayedScrollingColumnIndex <= myColumnNo - 1) Then
    '                'analizuj kolumny przed obszarem widzialnym
    '                If myDataGridView.FirstDisplayedScrollingColumnIndex > 0 Then
    '                    'wyswietlamy od dalszej kolumny....
    '                    For LiCol = 0 To myDataGridView.FirstDisplayedScrollingColumnIndex - 1
    '                        If myDataGridView.Columns(LiCol).Visible Then
    '                            ThisColWidth = myDataGridView.Columns(LiCol).Width

    '                            tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
    '                            If tBox Is Nothing Then
    '                            Else
    '                                'tBox.Location = New system.drawing.Point(LeftVArea + offsetX, myStatusBar.Location.Y + 1 + offsetY)
    '                                'tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
    '                                tBox.Visible = False
    '                                'tBox.BringToFront()
    '                            End If

    '                            tCmd = SzukajButtonONazwie(myForm, "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
    '                            If tCmd Is Nothing Then
    '                            Else
    '                                'tCmd.Location = New system.drawing.Point(LeftVArea + ThisColWidth - 16 + offsetX, myStatusBar.Location.Y + 1 + offsetY)
    '                                'tCmd.Size = New System.Drawing.Size(16, 22)
    '                                tCmd.Visible = False
    '                                'tCmd.BringToFront()
    '                            End If
    '                        End If
    '                    Next
    '                End If



    '                'MsgBox("po obsludze PRZED widzialnymi")



    '                'LiCol = myDataGridView.FirstDisplayedScrollingColumnIndex
    '                'LiCol = myDataGridView.DisplayedColumnCount(True)

    '                'analizuj kolumny widzialne
    '                'For LiCol = myDataGridView.FirstDisplayedScrollingColumnIndex To myDataGridView.FirstDisplayedScrollingColumnIndex + myDataGridView.Columns.GetColumnCount(DataGridViewElementStates.Displayed) - 1

    '                For LiCol = myDataGridView.FirstDisplayedScrollingColumnIndex To myColumnNo - 1

    '                    'For LiCol = myDataGridView.FirstDisplayedScrollingColumnIndex To myDataGridView.FirstDisplayedScrollingColumnIndex + myDataGridView.DisplayedColumnCount(True) - 1
    '                    'For LiCol = myDataGridView.FirstDisplayedScrollingColumnIndex To myDataGridView.FirstDisplayedScrollingColumnIndex + myDataGridView.DisplayedColumnCount(True) - 1
    '                    'ThisColWidth = myDataGridView.GetCellDisplayRectangle(LiCol, 0, True).Width - myDataGridView.FirstDisplayedScrollingColumnHiddenWidth
    '                    'ThisColLeft = myDataGridView.GetCellDisplayRectangle(LiCol, 0, True).Left
    '                    'ThisColRight = myDataGridView.GetCellDisplayRectangle(LiCol, 0, True).Right

    '                    If myDataGridView.ColumnHeadersVisible Then
    '                        If myDataGridView.Columns(LiCol).Visible Then
    '                            If LiCol = myDataGridView.FirstDisplayedScrollingColumnIndex Then
    '                                ThisColWidth = myDataGridView.Columns(LiCol).Width - myDataGridView.FirstDisplayedScrollingColumnHiddenWidth
    '                                ThisColLeft = myDataGridView.GetColumnDisplayRectangle(LiCol, True).Location.X
    '                                NextColLeft = ThisColLeft + ThisColWidth
    '                            Else
    '                                ThisColWidth = myDataGridView.Columns(LiCol).Width - 1
    '                                ThisColLeft = NextColLeft
    '                                NextColLeft = ThisColLeft + ThisColWidth
    '                            End If
    '                            ThisColRight = ThisColLeft + ThisColWidth


    '                            If myDataGridView.Columns(LiCol).Visible AndAlso (ThisColWidth > 0) Then
    '                                If ThisColLeft < RightVArea Then
    '                                    tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
    '                                    If tBox Is Nothing Then
    '                                    Else
    '                                        If (myDataGridView.Left + ThisColRight <= LeftVArea) Or _
    '                                           (myDataGridView.Left + ThisColLeft >= RightVArea) Then
    '                                            'jesteśmy poza obszarem widzialnym 
    '                                            Exit For
    '                                        Else
    '                                            ''czy kolumna zaczyna sie przed czescia widzialną
    '                                            'If myDataGridView.Left + ThisColLeft < LeftVArea Then
    '                                            '    ThisColWidth -= LeftVArea - ThisColLeft - myDataGridView.Left
    '                                            'End If

    '                                            'czy kolumna kończy się za częścią wudzialną
    '                                            If myDataGridView.Left + ThisColRight > RightVArea Then
    '                                                'ThisColWidth -= ThisColRight + myDataGridView.Left - RightVArea
    '                                                ThisColWidth -= LeftVArea + SumColWidth + ThisColWidth + offsetX - RightVArea
    '                                            End If

    '                                            tBox.Location = New System.Drawing.Point(LeftVArea + SumColWidth + offsetX, _
    '                                                                      myStatusBar.Location.Y + 1 + offsetY)
    '                                            tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
    '                                            tBox.Visible = True
    '                                            tBox.BringToFront()
    '                                            '-------------
    '                                            tBox.BackColor = System.Drawing.SystemColors.Window
    '                                            tBox.Font = New System.Drawing.Font(Softart_FormFontName, _
    '                                                                                   Softart_FormFontSize, _
    '                                                                                    System.Drawing.FontStyle.Regular, _
    '                                                                                    System.Drawing.GraphicsUnit.Point, _
    '                                                                                    CType(238, Byte))
    '                                            tBox.ForeColor = System.Drawing.Color.Navy
    '                                            tBox.BorderStyle = BorderStyle.Fixed3D
    '                                            tBox.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or _
    '                                                                     System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    '                                            '-------------




    '                                            tCmd = SzukajButtonONazwie(myForm, "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
    '                                            If tCmd Is Nothing Then
    '                                            Else
    '                                                tCmd.Location = New System.Drawing.Point(LeftVArea + SumColWidth + ThisColWidth - 16 + offsetX, _
    '                                                                         myStatusBar.Location.Y + 1 + offsetY)
    '                                                tCmd.Size = New System.Drawing.Size(16, 22)
    '                                                If Par_WyswLiWartFiltra Then
    '                                                    'OK, wyswietlaj...
    '                                                    If tBox.Width < tCmd.Width Then
    '                                                        tCmd.Visible = False
    '                                                    Else
    '                                                        tCmd.Visible = True
    '                                                    End If
    '                                                Else
    '                                                    tCmd.Visible = False
    '                                                End If
    '                                                tCmd.BringToFront()
    '                                            End If

    '                                            SumColWidth += ThisColWidth + 1
    '                                        End If
    '                                    End If
    '                                Else
    '                                    'poza obszarem widzialnym
    '                                    Exit For
    '                                End If
    '                                'MsgBox("po obsludze kolumny widzialnej")
    '                            End If
    '                        End If
    '                    End If
    '                Next

    '                'System.Windows.Forms.Application.DoEvents()

    '                'MsgBox("po obsludze widzialnych")




    '                'analizuj kolumny za obszarem widzialnym
    '                'If myDataGridView.FirstDisplayedScrollingColumnIndex + myDataGridView.Columns.GetColumnCount(DataGridViewElementStates.Displayed) < IleCol - 1 Then
    '                If myDataGridView.FirstDisplayedScrollingColumnIndex + myDataGridView.DisplayedColumnCount(True) < myColumnNo - 1 Then
    '                    'For LiCol = myDataGridView.FirstDisplayedScrollingColumnIndex + myDataGridView.Columns.GetColumnCount(DataGridViewElementStates.Displayed) To IleCol - 1
    '                    'For LiCol = myDataGridView.FirstDisplayedScrollingColumnIndex + myDataGridView.Columns.GetColumnCount(True) To myColumnNo - 1
    '                    For LiCol = myDataGridView.FirstDisplayedScrollingColumnIndex + myDataGridView.Columns.GetColumnCount(DataGridViewElementStates.Displayed) To myColumnNo - 1
    '                        If myDataGridView.Columns(LiCol).Visible Then
    '                            ThisColWidth = myDataGridView.GetCellDisplayRectangle(LiCol, 0, True).Width
    '                            ThisColLeft = myDataGridView.GetCellDisplayRectangle(LiCol, 0, True).Left
    '                            ThisColRight = myDataGridView.GetCellDisplayRectangle(LiCol, 0, True).Right

    '                            If ThisColWidth > 0 Then
    '                                If ThisColLeft >= RightVArea Then
    '                                    tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
    '                                    If tBox Is Nothing Then
    '                                    Else
    '                                        'tBox.Location = New system.drawing.Point(LeftVArea + offsetX, _
    '                                        '                        myStatusBar.Location.Y + 1 + offsetY)
    '                                        'tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
    '                                        tBox.Visible = False
    '                                        'tBox.BringToFront()
    '                                    End If

    '                                    tCmd = SzukajButtonONazwie(myForm, "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
    '                                    If tCmd Is Nothing Then
    '                                    Else
    '                                        'tCmd.Location = New system.drawing.Point(LeftVArea + ThisColWidth - 16 + offsetX, _
    '                                        '                       myStatusBar.Location.Y + 1 + offsetY)
    '                                        'tCmd.Size = New System.Drawing.Size(16, 22)
    '                                        tCmd.Visible = False
    '                                        'tCmd.BringToFront()
    '                                    End If
    '                                End If
    '                            End If
    '                        End If
    '                    Next
    '                End If
    '                'MsgBox("po obsludze PO widzialnych")
    '            End If
    '        End If
    '    End If
    'End Sub



    'pokaż filtry na ekranie
    Public Sub __PokazFiltryDGV2(ByVal myForm As Control, _
                         ByRef myDataGridView As DataGridView, _
                         ByRef myColumnNo As Integer, _
                         ByRef myStatusBar As StatusBar, _
                         ByRef FormLoad As Boolean, _
                         ByRef offsetX As Integer, _
                         ByRef offsetY As Integer)
        If Not FormLoad Then
            'Dim IleCol As Integer = DataViewCennik.Table.Columns.Count
            Dim IleCol As Integer = myColumnNo
            Dim LiCol As Integer = 0

            Dim ThisColWidth As Integer = 0
            Dim ThisColLeft As Integer = 0
            Dim ThisColRight As Integer = 0

            Dim NextColWidth As Integer = 0
            Dim NextColLeft As Integer = 0
            Dim NextColRight As Integer = 0

            Dim SumColWidth As Integer = 0
            Dim tBox As System.Windows.Forms.TextBox = Nothing
            Dim tCmd As System.Windows.Forms.Button = Nothing

            Dim LeftVArea As Integer = myStatusBar.Location.X + myStatusBar.Panels(0).Width
            Dim RightVArea As Integer = myStatusBar.Location.X + myStatusBar.Width


            'Dim dv As DataView = myDataGridView.DataSource
            'If dv.Count = 0 Then
            '    'tabela jest pusta - nie pokazuj filtra...ma być niewidoczny
            '    For LiCol = 0 To myColumnNo - 1
            '        If myDataGridView.Columns(LiCol).Visible Then
            '            ThisColWidth = myDataGridView.Columns(LiCol).Width

            '            tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
            '            If Not (tBox Is Nothing) Then
            '                tBox.Visible = False
            '            End If

            '            tCmd = SzukajButtonONazwie(myForm, "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
            '            If Not (tCmd Is Nothing) Then
            '                tCmd.Visible = False
            '            End If
            '        End If
            '    Next
            'Else
            For LiCol = 0 To myColumnNo - 1
                If myDataGridView.ColumnHeadersVisible Then
                    If myDataGridView.Columns(LiCol).Visible Then
                        ThisColLeft = myDataGridView.GetColumnDisplayRectangle(LiCol, True).Location.X + myDataGridView.Location.X
                        ThisColWidth = myDataGridView.GetColumnDisplayRectangle(LiCol, True).Size.Width
                        ThisColRight = ThisColLeft + ThisColWidth

                        If (ThisColWidth > 0) Then
                            'tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                            'tCmd = SzukajButtonONazwie(myForm, "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                            tBox = myForm.Controls("txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                            tCmd = myForm.Controls("cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))

                            If (ThisColRight <= LeftVArea) Or (ThisColLeft >= RightVArea) Then
                                'poza obszarem widzialnym
                                If Not (tBox Is Nothing) Then
                                    tBox.Visible = False
                                End If

                                If Not (tCmd Is Nothing) Then
                                    tCmd.Visible = False
                                End If

                            Else
                                'częściowo lub całkowicie w obszarze widzialnym
                                If Not (tBox Is Nothing) Then
                                    tBox.Location = New System.Drawing.Point(ThisColLeft + offsetX,
                                                              myStatusBar.Location.Y + 1 + offsetY)
                                    tBox.Size = New System.Drawing.Size(ThisColWidth, 20)

                                    tBox.BringToFront()
                                    tBox.BackColor = System.Drawing.SystemColors.Window
                                    tBox.Font = New System.Drawing.Font(Softart_FormFontName,
                                                                           Softart_FormFontSize,
                                                                            System.Drawing.FontStyle.Regular,
                                                                            System.Drawing.GraphicsUnit.Point,
                                                                            CType(238, Byte))
                                    tBox.ForeColor = System.Drawing.Color.Navy
                                    tBox.BorderStyle = BorderStyle.Fixed3D
                                    tBox.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or
                                                             System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
                                    tBox.Visible = True
                                End If
                                '-------------
                                If Not (tCmd Is Nothing) Then
                                    tCmd.Location = New System.Drawing.Point(ThisColRight - 16 + offsetX,
                                                             myStatusBar.Location.Y + 1 + offsetY)
                                    tCmd.Size = New System.Drawing.Size(16, 22)

                                    tCmd.BringToFront()
                                    If tBox.Width < tCmd.Width Then
                                        tCmd.Visible = False
                                    Else
                                        tCmd.Visible = True
                                    End If
                                End If

                            End If
                        End If
                    End If
                End If
                'System.Windows.Forms.Application.DoEvents()
            Next
        End If
        'End If
        System.Windows.Forms.Application.DoEvents()
    End Sub






    '***********************************************************
    '** Ustaw sę na pierwszym wierszu w tej kolumnie
    '***********************************************************

    Public Sub PrzejdzDoDGV(ByRef myDataGridView As DataGridView, _
                            ByRef myColumnNo As Integer)

        If myDataGridView.RowCount > 0 Then
            myDataGridView.CurrentCell = myDataGridView(myColumnNo, 0)
            myDataGridView.CurrentCell.Selected = True
            myDataGridView.Focus()
        End If
    End Sub






    '***********************************************************
    '** FILTROWANIE Data View
    '***********************************************************

    Public Sub FiltrujDataViewDGV(ByVal pForm As Control, _
                               ByRef pDataGridView As DataGridView, _
                               ByRef pColumnNo As Integer, _
                               ByRef pDataView As DataView, _
                               ByRef pStatusBar As StatusBar)
        __FiltrujDVDGV(pForm, pDataGridView, pColumnNo, pDataView, pStatusBar, "")
    End Sub

    Public Sub FiltrujDataViewDGV(ByVal pForm As Control, _
                               ByRef pDataGridView As DataGridView, _
                               ByRef pColumnNo As Integer, _
                               ByRef pDataView As DataView, _
                               ByRef pStatusBar As StatusBar, _
                               ByRef pFilter As String)
        __FiltrujDVDGV(pForm, pDataGridView, pColumnNo, pDataView, pStatusBar, pFilter)
    End Sub


    Private Sub __FiltrujDVDGV(ByVal myForm As Control, _
                          ByRef myDataGridView As DataGridView, _
                          ByRef myColumnNo As Integer, _
                          ByRef myDataView As DataView, _
                          ByRef myStatusBar As StatusBar, _
                          ByRef myFilter As String)
        Dim tBox As System.Windows.Forms.TextBox = Nothing
        Dim LiCol As Integer = 0
        Dim ColName As String = ""
        Dim ColIndex As Integer = 0
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

        'myDataView.RowFilter = oLokalnyFiltr

        Dim srcDataSet As DataSet = myDataView.Table.DataSet
        Dim srcDataTable As DataTable = myDataView.Table

        If srcDataTable.Rows.Count > 0 Then
            For LiCol = 0 To srcDataTable.Columns.Count - 1
                'tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                tBox = myForm.Controls("txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))

                If tBox Is Nothing Then
                Else
                    If Len(Trim(tBox.Text)) > 0 Then
                        tBox.BackColor = SoftartFilterBackColor
                        'RysujGradientTextBox(tBox)

                        'ColName = srcDataTable.Columns(LiCol).ColumnName

                        'nazwe filtrowanej kolumny pobieramy z DATAGRIDVIEW
                        'gdyż kolumny w DGV nie muszą być w takiej samej kolejnosci co w DATAVIEW
                        'Index kolumny = TAG
                        ColName = myDataGridView.Columns(LiCol).DataPropertyName
                        ColIndex = LiCol    'myDataGridView.Columns(LiCol).Tag

                        If TypeOf srcDataTable.Rows.Item(0).Item(ColIndex) Is String Then
                            'oLokalnyFiltr &= Separ & _
                            '        "(" & ColName & _
                            '        " LIKE '" & tBox.Text & "*')"
                            'Separ = " AND "



                            'podziel linie na poszczególne części
                            'words = tBox.Text.Split(splitChars, StringSplitOptions.RemoveEmptyEntries)
                            'analiza filtrow alternatywnych
                            words = tBox.Text.Split(splitChars, StringSplitOptions.None)
                            If words.Length = 1 Then

                                'jeden warunek...
                                If Left(tBox.Text, 3) = "***" Then
                                    'pole puste
                                    oLokalnyFiltr &= Separ &
                                        "(" & myDataView.Table.Columns(LiCol).ColumnName & "='')"

                                ElseIf Left(tBox.Text, 2) = "**" Then
                                    'nie zawiera tekstu
                                    oLokalnyFiltr &= Separ &
                                        "(NOT " & myDataView.Table.Columns(LiCol).ColumnName & " LIKE '*" & Mid(tBox.Text, 3) & "*') "


                                Else
                                    'sprawdz czy to nie jest przedział danych OD..DO
                                    SepPoz = InStr(tBox.Text, "..")
                                    If SepPoz > 0 Then
                                        'drukuj przedział
                                        TxtOd = Mid(tBox.Text, 1, SepPoz - 1)
                                        TxtDo = Mid(tBox.Text, SepPoz + 2)
                                        If Len(TxtOd) > 0 And Len(TxtDo) > 0 Then
                                            'fuiltruj przedzial
                                            oLokalnyFiltr &= Separ &
                                                "(" &
                                                myDataView.Table.Columns(LiCol).ColumnName & ">='" & TxtOd & "' and " &
                                                myDataView.Table.Columns(LiCol).ColumnName & "<='" & TxtDo & "'" &
                                                ") "
                                        Else
                                            If Len(TxtOd) > 0 Then
                                                'filtruj od
                                                oLokalnyFiltr &= Separ &
                                                    "(" & myDataView.Table.Columns(LiCol).ColumnName & ">='" & TxtOd & "')"
                                            Else
                                                'filtruj Do
                                                oLokalnyFiltr &= Separ &
                                                    "(" & myDataView.Table.Columns(LiCol).ColumnName & "<='" & TxtDo & "')"
                                            End If
                                        End If
                                    Else
                                        'filtruj pojedynczy rekord
                                        If Left(tBox.Text, 1) = "*" Then
                                            oLokalnyFiltr &= Separ &
                                            "(" & myDataView.Table.Columns(LiCol).ColumnName &
                                            " LIKE '" & tBox.Text & "*')"
                                        Else
                                            'nie ma gwiazdki - uzupełniamy....
                                            oLokalnyFiltr &= Separ &
                                            "(" & myDataView.Table.Columns(LiCol).ColumnName &
                                            " LIKE '" & tBox.Text & "*')"

                                        End If
                                    End If
                                End If
                            Else
                                'więcej niż jeden warunek - musi być spełniony TEN i TEN i TEN...
                                oLokalnyFiltr &= Separ & "("
                                For iwords = 0 To words.Length - 1
                                    'sprawdzamy kolejne warunki (i ten) 
                                    If Len(words(iwords)) > 0 Then
                                        If Left(words(iwords), 3) = "***" Then
                                            'pole puste
                                            oLokalnyFiltr &= "(" & myDataView.Table.Columns(LiCol).ColumnName & "='') OR "

                                        ElseIf Left(words(iwords), 2) = "**" Then
                                            'nie zawiera tekstu
                                            oLokalnyFiltr &= " (NOT " & myDataView.Table.Columns(LiCol).ColumnName &
                                                        " LIKE '*" & Mid(words(iwords), 3) & "*') OR "

                                        Else
                                            'sprawdz czy to nie jest przedział danych OD..DO
                                            SepPoz = InStr(words(iwords), "..")
                                            If SepPoz > 0 Then
                                                'drukuj przedział
                                                TxtOd = Mid(words(iwords), 1, SepPoz - 1)
                                                TxtDo = Mid(words(iwords), SepPoz + 2)
                                                If Len(TxtOd) > 0 And Len(TxtDo) > 0 Then
                                                    'fuiltruj przedzial
                                                    oLokalnyFiltr &= "(" &
                                                        myDataView.Table.Columns(LiCol).ColumnName & ">='" & TxtOd & "' AND " &
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
                                                'filtruj pojedynczy rekord
                                                If Left(tBox.Text, 1) = "*" Then
                                                    oLokalnyFiltr &= "(" & myDataView.Table.Columns(LiCol).ColumnName &
                                                                " LIKE '" & words(iwords) & "*') OR "
                                                Else
                                                    'nie ma gwiazdki - uzupełniamy....
                                                    oLokalnyFiltr &= "(" & myDataView.Table.Columns(LiCol).ColumnName &
                                                                " LIKE '" & words(iwords) & "*') OR "
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                                'usun ostatnie ' AND '
                                oLokalnyFiltr = Mid(oLokalnyFiltr, 1, Len(oLokalnyFiltr) - 4)
                                oLokalnyFiltr &= ")"
                            End If
                            Separ = " AND "



















                        ElseIf TypeOf srcDataTable.Rows.Item(0).Item(ColIndex) Is Integer Or _
                               TypeOf srcDataTable.Rows.Item(0).Item(ColIndex) Is Int16 Or _
                               TypeOf srcDataTable.Rows.Item(0).Item(ColIndex) Is Int32 Or _
                               TypeOf srcDataTable.Rows.Item(0).Item(ColIndex) Is Int64 Then
                            'Select Case IsNumeric(tBox.Text)
                            '    Case True
                            '        oLokalnyFiltr &= Separ & _
                            '                "(" & ColName & _
                            '                "=" & Trim(Str(tBox.Text)) & ")"
                            '        Separ = " AND "
                            '    Case Else
                            'End Select



                            'podziel linie na poszczególne części
                            'words = tBox.Text.Split(splitChars, StringSplitOptions.RemoveEmptyEntries)
                            'analiza filtrow alternatywnych
                            words = tBox.Text.Split(splitChars, StringSplitOptions.None)
                            If words.Length = 1 Then
                                If IsNumeric(tBox.Text) Then
                                    oLokalnyFiltr &= Separ &
                                        "(" & myDataView.Table.Columns(LiCol).ColumnName &
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
                                            oLokalnyFiltr &= Separ &
                                                "(" & myDataView.Table.Columns(LiCol).ColumnName &
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





                            'ElseIf TypeOf myDataView.Item(0).Item(LiCol) Is DBNull Then
                            '    oLokalnyFiltr &= Separ &
                            '        "(" & myDataView.Table.Columns(LiCol).ColumnName &
                            '        "=" & Trim(Str(tBox.Text)) & ")"
                            '    Separ = " AND "




                        ElseIf TypeOf srcDataTable.Rows.Item(0).Item(ColIndex) Is Double Then
                            Select Case IsNumeric(tBox.Text)
                                Case True
                                    oLokalnyFiltr &= Separ & _
                                            "(" & ColName & _
                                            "=" & Trim(Str(tBox.Text)) & ")"
                                    Separ = " AND "
                                Case Else
                            End Select

                        ElseIf TypeOf srcDataTable.Rows.Item(0).Item(ColIndex) Is Boolean Then
                            Select Case UCase(tBox.Text)
                                Case "T", "Y", "TAK", "YES", "TRUE"
                                    oLokalnyFiltr &= Separ & _
                                                    "(" & ColName & _
                                                    "=1)"
                                    Separ = " AND "
                                Case "N", "NIE", "NO", "FALSE"
                                    oLokalnyFiltr &= Separ & _
                                                    "(" & ColName & _
                                                    "=0)"
                                    Separ = " AND "
                                Case Else
                            End Select

                        ElseIf TypeOf srcDataTable.Rows.Item(0).Item(ColIndex) Is Guid Then
                            'oLokalnyFiltr &= Separ & _
                            '        "(" & ColName & _
                            '        " LIKE '" & tBox.Text.ToString & "*')"
                            'Separ = " AND "

                            'oLokalnyFiltr &= Separ & _
                            '        "(SUBSTRING(" & ColName & ",1," & Len(tBox.Text).ToString & ")" & _
                            '        "='" & tBox.Text.ToString & "')"
                            'Separ = " AND "

                            oLokalnyFiltr &= Separ & _
                                    "(" & ColName & _
                                    "='" & tBox.Text.ToString & "')"
                            Separ = " AND "
                        End If

                    Else
                        If tBox.ContainsFocus Then
                            tBox.BackColor = SoftartCursorBackColor
                        Else
                            tBox.BackColor = SoftartEditableBackColor
                        End If
                    End If
                End If
            Next
        End If

        Try
            myDataView.RowFilter = oLokalnyFiltr
            'myDataGridView.DataSource = myDataView

            myStatusBar.Panels(0).Text = "Il.rec.: " & myDataView.Count.ToString
        Catch Ex As System.Exception
            MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & Ex.Message, "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Finally
        End Try
    End Sub





    '******************************************
    '** procedura zamiany jednego znaku na inny
    '******************************************

    Public Function Char2Char(ByVal SrcString As String,
                              ByVal SrcChar As String,
                              ByVal DstChar As String) As String
        Dim pos As Integer = InStr(SrcString, SrcChar)
        Do While pos > 0
            If pos = 1 Then
                SrcString = DstChar & Mid(SrcString, 2)
            Else
                SrcString = Mid(SrcString, 1, pos - 1) & DstChar & Mid(SrcString, pos + 1)
            End If
            pos = InStr(SrcString, SrcChar)
        Loop
        Return SrcString
    End Function



    '***********************************************************
    '** Buduj filtr zlozony dla wszystkich kolumn
    '***********************************************************
    'oLokalnyFiltrTowary = ZbudujFiltrDGV(Me.SplitContainer2.Panel2, DataViewCennik, "%", "Cennik")

    Public Function ZbudujFiltrDGV(ByVal myForm As Control, _
                                   ByVal myDataView As DataView, _
                                   ByVal myAsterix As String, _
                                   ByVal myAlias As String) As String

        Dim tBox As System.Windows.Forms.TextBox = Nothing
        Dim LiCol As Integer = 0
        Dim ColumnNo As Integer = myDataView.Table.Columns.Count
        Dim FiltrDlaTabeli As String = ""
        Dim Separ As String = ""

        If ColumnNo > 0 Then
            For LiCol = 0 To ColumnNo - 1
                'tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                tBox = myForm.Controls("txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))

                If tBox Is Nothing Then
                Else
                    If Len(Trim(tBox.Text)) > 0 Then

                        If TypeOf myDataView.Item(0).Item(LiCol) Is String Then
                            FiltrDlaTabeli &= Separ & _
                                    "(" & IIf(myAlias.Length > 0, myAlias & ".", "") & myDataView.Table.Columns(LiCol).ColumnName & _
                                    " LIKE '" & Char2Char(tBox.Text, "*", myAsterix) & myAsterix & "')"
                            Separ = " AND "


                        ElseIf TypeOf myDataView.Item(0).Item(LiCol) Is Integer Or _
                               TypeOf myDataView.Item(0).Item(LiCol) Is Int16 Or _
                               TypeOf myDataView.Item(0).Item(LiCol) Is Int32 Or _
                               TypeOf myDataView.Item(0).Item(LiCol) Is Int64 Then
                            Select Case IsNumeric(tBox.Text)
                                Case True
                                    FiltrDlaTabeli &= Separ & _
                                            "(" & IIf(myAlias.Length > 0, myAlias & ".", "") & myDataView.Table.Columns(LiCol).ColumnName & _
                                            "=" & Trim(Str(tBox.Text)) & ")"
                                    Separ = " AND "
                                Case Else
                            End Select


                        ElseIf TypeOf myDataView.Item(0).Item(LiCol) Is Double Then
                            Select Case IsNumeric(tBox.Text)
                                Case True
                                    FiltrDlaTabeli &= Separ & _
                                            "(" & IIf(myAlias.Length > 0, myAlias & ".", "") & myDataView.Table.Columns(LiCol).ColumnName & _
                                            "=" & Trim(Str(tBox.Text)) & ")"
                                    Separ = " AND "
                                Case Else
                            End Select


                        ElseIf TypeOf myDataView.Item(0).Item(LiCol) Is Boolean Then
                            Select Case UCase(tBox.Text)
                                Case "T", "Y", "TAK", "YES", "TRUE"
                                    FiltrDlaTabeli &= Separ & _
                                                    "(" & IIf(myAlias.Length > 0, myAlias & ".", "") & myDataView.Table.Columns(LiCol).ColumnName & _
                                                    "=1)"
                                    Separ = " AND "
                                Case "N", "NIE", "NO", "FALSE"
                                    FiltrDlaTabeli &= Separ & _
                                                    "(" & IIf(myAlias.Length > 0, myAlias & ".", "") & myDataView.Table.Columns(LiCol).ColumnName & _
                                                    "=0)"
                                    Separ = " AND "
                                Case Else
                            End Select


                        ElseIf TypeOf myDataView.Item(0).Item(LiCol) Is Guid Then
                            FiltrDlaTabeli &= Separ & _
                                    "(" & IIf(myAlias.Length > 0, myAlias & ".", "") & myDataView.Table.Columns(LiCol).ColumnName & _
                                    "='" & Char2Char(tBox.Text.ToString, "*", myAsterix) & "')"
                            Separ = " AND "


                        End If
                    Else
                    End If
                End If
            Next
        End If
        Return (FiltrDlaTabeli)
    End Function














    '***********************************************************
    '** Czyszczenie filtra
    '***********************************************************

    Public Sub CzyscFiltryDGV(ByVal myForm As Control, _
                               ByRef myDataGridView As DataGridView, _
                               ByRef myColumnNo As Integer, _
                               ByRef myDataView As DataView, _
                               ByRef myStatusBar As StatusBar)

        Dim tBox As System.Windows.Forms.TextBox = Nothing
        Dim LiCol As Integer = 0

        oLokalnyFiltr = ""
        myDataView.RowFilter = oLokalnyFiltr

        'If myDataView.Count > 0 Then
        For LiCol = 0 To myColumnNo - 1
            'tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
            tBox = myForm.Controls("txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))

            If tBox Is Nothing Then
            Else
                If Len(tBox.Text) > 0 Then
                    tBox.Text = ""
                    tBox.BackColor = SoftartEditableBackColor
                    tBox.ForeColor = SoftartEditableForeColor
                End If
            End If
        Next
        'End If
    End Sub


    Public Sub CzyscFiltryDGV(ByVal myForm As Control, _
                           ByRef myDataGridView As DataGridView, _
                           ByRef myColumnNo As Integer, _
                           ByRef myDataView As DataView, _
                           ByRef myStatusBar As StatusBar, _
                           ByRef myFilterBar As StatusBar, _
                           ByVal myFormStart As Boolean)

        Dim tBox As System.Windows.Forms.TextBox = Nothing
        Dim LiCol As Integer = 0

        oLokalnyFiltr = ""
        myDataView.RowFilter = oLokalnyFiltr

        'If myDataView.Count > 0 Then
        For LiCol = 0 To myColumnNo - 1
            'tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
            tBox = myForm.Controls("txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))

            If tBox Is Nothing Then
            Else
                If Len(tBox.Text) > 0 Then
                    tBox.Text = ""
                    tBox.BackColor = SoftartEditableBackColor
                    tBox.ForeColor = SoftartEditableForeColor
                End If
            End If
        Next
        'End If

        'pokaż filtry po wyczyszceniu
        PokazFiltryColumnDGV(myForm, myDataGridView, myColumnNo, myFilterBar, myFormStart)
    End Sub



    '***********************************************************
    '** CZYSC filtry
    '***********************************************************

    Public Sub CzyscFiltryDataTableDGV(ByVal myForm As Control, _
                           ByRef myDataGridView As DataGridView, _
                           ByRef myColumnNo As Integer, _
                           ByRef myDataView As DataView, _
                           ByRef myDataSet As DataSet, _
                           ByRef myStatusBar As StatusBar)

        Dim tBox As System.Windows.Forms.TextBox = Nothing
        Dim LiCol As Integer = 0
        Dim Separ As String = ""
        Dim Sortowanie As String = ""

        oLokalnyFiltr = ""
        If myDataView.Count > 0 Then
            Sortowanie = myDataView.Sort

            For LiCol = 0 To myColumnNo - 1
                'tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                tBox = myForm.Controls("txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))

                If tBox Is Nothing Then
                Else
                    tBox.Text = ""
                    tBox.BackColor = SoftartEditableBackColor
                    tBox.ForeColor = SoftartEditableForeColor
                End If
            Next

            'Filtr sklada sie z filtru lokalnego i filtru zlozonego
            Dim SumFiltr As String = ""
            If Len(oLokalnyFiltr) > 0 Then
                SumFiltr = "(" & oLokalnyFiltr & ")"
                Separ = " AND "
            Else
                SumFiltr = ""
                Separ = ""
            End If
            If Len(oFiltrZlozony) > 0 Then
                SumFiltr = SumFiltr & Separ & "(" & oFiltrZlozony & ")"
            End If

            Try
                'myDataView = New DataView(myDataSet.Tables(0))
                myDataView = New DataView(myDataSet.Tables(0), SumFiltr, Sortowanie, DataViewRowState.OriginalRows)

                myDataView.AllowDelete = False
                myDataView.AllowNew = False
                myDataGridView.DataSource = myDataView
                myDataGridView.RowHeadersWidth = 50       'use if tablestyle
                myDataGridView.Refresh()

                myStatusBar.Panels(0).Text = "Il.rec.: " & myDataView.Count.ToString
            Catch Ex As System.Exception
                MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & Ex.Message, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            Finally
            End Try
        End If
    End Sub


    Public Sub CzyscFiltryDataTableDGVOproczMnie(ByVal myForm As Control, _
                           ByRef myDataGridView As DataGridView, _
                           ByRef myColumnNo As Integer, _
                           ByRef myDataView As DataView, _
                           ByRef myDataSet As DataSet, _
                           ByRef myStatusBar As StatusBar, _
                           ByRef Ja As System.Windows.Forms.TextBox)

        Dim tBox As System.Windows.Forms.TextBox = Nothing
        Dim LiCol As Integer = 0
        Dim Separ As String = ""
        Dim Sortowanie As String = ""

        oLokalnyFiltr = ""
        If myDataView.Count > 0 Then
            Sortowanie = myDataView.Sort

            For LiCol = 0 To myColumnNo - 1
                'tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                tBox = myForm.Controls("txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))

                If tBox Is Ja Then
                Else
                    tBox.Text = ""
                    tBox.BackColor = SoftartEditableBackColor
                    tBox.ForeColor = SoftartEditableForeColor
                End If
            Next

            __FiltrujDTDGV(myForm, myDataGridView, myColumnNo, myDataView, myDataSet, myStatusBar, Sortowanie)

        End If
    End Sub











    '***********************************************************
    '** FILTROWANIE DATA TABLE
    '***********************************************************

    Public Sub FiltrujDataTableDGV(ByVal pForm As Control, _
                                ByRef pDataGridView As DataGridView, _
                                ByRef pColumnNo As Integer, _
                                ByRef pDataView As DataView, _
                                ByRef pDataSet As DataSet, _
                                ByRef pStatusBar As StatusBar)
        __FiltrujDTDGV(pForm, pDataGridView, pColumnNo, pDataView, pDataSet, pStatusBar, "")
    End Sub

    Public Sub FiltrujDataTableDGV(ByVal pForm As Control, _
                            ByRef pDataGridView As DataGridView, _
                            ByRef pColumnNo As Integer, _
                            ByRef pDataView As DataView, _
                            ByRef pDataSet As DataSet, _
                            ByRef pStatusBar As StatusBar, _
                            ByRef pSort As String)
        __FiltrujDTDGV(pForm, pDataGridView, pColumnNo, pDataView, pDataSet, pStatusBar, pSort)
    End Sub

    'filtruje DataTable.Select
    Private Sub __FiltrujDTDGV(ByVal myForm As Control, _
                      ByRef myDataGridView As DataGridView, _
                      ByRef myColumnNo As Integer, _
                      ByRef myDataView As DataView, _
                      ByRef myDataSet As DataSet, _
                      ByRef myStatusBar As StatusBar, _
                      ByRef Sort As String)

        Dim tBox As System.Windows.Forms.TextBox = Nothing
        Dim LiCol As Integer = 0
        Dim Separ As String = ""
        Dim Sortowanie As String = Sort
        Dim ColName As String = ""
        Dim ColIndex As Integer = 0

        oLokalnyFiltr = ""
        'analizuj wszystkie pola i buduj filtr - może filtrować TYLKO JEDNĄ KOLUMNĘ
        myDataView.RowFilter = oLokalnyFiltr
        If myDataView.Count > 0 Then
            For LiCol = 0 To myColumnNo - 1
                'tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                tBox = myForm.Controls("txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))

                If tBox Is Nothing Then
                Else
                    'sprawdz czy wpisano jakiś filtr....
                    If Len(Trim(tBox.Text)) > 0 Then
                        tBox.BackColor = SoftartFilterBackColor
                        'RysujGradientTextBox(tBox)

                        'ColName = srcDataTable.Columns(LiCol).ColumnName

                        'nazwe filtrowanej kolumny pobieramy z DATAGRIDVIEW
                        'gdyż kolumny w DGV nie muszą być w takiej samej kolejnosci co w DATAVIEW
                        'Index kolumny = TAG
                        ColName = myDataGridView.Columns(LiCol).DataPropertyName
                        ColIndex = myDataGridView.Columns(LiCol).Tag

                        If TypeOf myDataView.Item(0).Item(ColIndex) Is String Then
                            oLokalnyFiltr &= Separ & _
                                    "(" & ColName & _
                                    " LIKE '" & tBox.Text & "%')"
                            Separ = " AND "

                        ElseIf TypeOf myDataView.Item(0).Item(ColIndex) Is Integer Or _
                               TypeOf myDataView.Item(0).Item(ColIndex) Is Int16 Or _
                               TypeOf myDataView.Item(0).Item(ColIndex) Is Int32 Or _
                               TypeOf myDataView.Item(0).Item(ColIndex) Is Int64 Then
                            Select Case IsNumeric(tBox.Text)
                                Case True
                                    oLokalnyFiltr &= Separ & _
                                            "(" & ColName & _
                                            "=" & Trim(Str(tBox.Text)) & ")"
                                    Separ = " AND "
                                Case Else
                            End Select

                        ElseIf TypeOf myDataView.Item(0).Item(ColIndex) Is Double Then
                            Select Case IsNumeric(tBox.Text)
                                Case True
                                    oLokalnyFiltr &= Separ & _
                                            "(" & ColName & _
                                            "=" & Trim(Str(tBox.Text)) & ")"
                                    Separ = " AND "
                                Case Else
                            End Select

                        ElseIf TypeOf myDataView.Item(0).Item(ColIndex) Is Boolean Then
                            Select Case UCase(tBox.Text)
                                Case "T", "Y", "TAK", "YES", "TRUE"
                                    oLokalnyFiltr &= Separ & _
                                                    "(" & ColName & _
                                                    "=1)"
                                    Separ = " AND "
                                Case "N", "NIE", "NO", "FALSE"
                                    oLokalnyFiltr &= Separ & _
                                                    "(" & ColName & _
                                                    "=0)"
                                    Separ = " AND "
                                Case Else
                            End Select

                        ElseIf TypeOf myDataView.Item(0).Item(ColIndex) Is Guid Then
                            'oLokalnyFiltr &= Separ & _
                            '        "(" & ColName & _
                            '        " LIKE '" & tBox.Text.ToString & "*')"
                            'Separ = " AND "

                            'oLokalnyFiltr &= Separ & _
                            '        "(SUBSTRING(" & ColName & ",1," & Len(tBox.Text).ToString & ")" & _
                            '        "='" & tBox.Text.ToString & "')"
                            'Separ = " AND "

                            oLokalnyFiltr &= Separ & _
                                    "(" & ColName & _
                                    "='" & tBox.Text.ToString & "')"
                            Separ = " AND "
                        Else        'nie zidentyfikowano typu pola....
                        End If



                    Else
                        'nie ma wartosci filtra dla tej kolumny
                        If tBox.ContainsFocus Then
                            tBox.BackColor = SoftartCursorBackColor
                        Else
                            tBox.BackColor = SoftartEditableBackColor
                        End If
                    End If
                End If
            Next

            'Filtr sklada sie z filtru lokalnego i filtru zlozonego
            Dim SumFiltr As String = ""
            If Len(oLokalnyFiltr) > 0 Then
                SumFiltr = "(" & oLokalnyFiltr & ")"
                Separ = " AND "
            Else
                SumFiltr = ""
                Separ = ""
            End If
            If Len(oFiltrZlozony) > 0 Then
                SumFiltr = SumFiltr & Separ & "(" & oFiltrZlozony & ")"
            End If

            Try
                'myDataView = New DataView(myDataSet.Tables(0), oLokalnyFiltr, Sortowanie, DataViewRowState.OriginalRows)
                myDataView = New DataView(myDataSet.Tables(0), SumFiltr, Sortowanie, DataViewRowState.OriginalRows)
                myDataView.AllowDelete = False
                myDataView.AllowNew = False
                myDataGridView.DataSource = myDataView
                myDataGridView.RowHeadersWidth = 50       'use if tablestyle
                myDataGridView.Refresh()

                myStatusBar.Panels(0).Text = "Il.rec.: " & myDataView.Count.ToString
            Catch Ex As System.Exception
                MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & Ex.Message, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            Finally
            End Try
        End If
    End Sub













    '***********************************************************
    '** FILTROWANIE ZLOZONE DATA TABLE
    '***********************************************************

    Public Sub FiltrExDataTableDGV(ByVal pForm As Control, _
                            ByRef pDataGridView As DataGridView, _
                            ByRef pColumnNo As Integer, _
                            ByRef pDataView As DataView, _
                            ByRef pDataSet As DataSet, _
                            ByRef pStatusBar As StatusBar, _
                            ByRef pFilTxt As String, _
                            ByRef pSort As String)
        __FiltrExDTDGV(pForm, pDataGridView, pColumnNo, pDataView, pDataSet, pStatusBar, pFilTxt, pSort)
    End Sub

    'filtruje DataTable.Select
    Private Sub __FiltrExDTDGV(ByVal myForm As Control, _
                      ByRef myDataGridView As DataGridView, _
                      ByRef myColumnNo As Integer, _
                      ByRef myDataView As DataView, _
                      ByRef myDataSet As DataSet, _
                      ByRef myStatusBar As StatusBar, _
                      ByRef myFilTxt As String, _
                      ByRef Sort As String)

        Dim tBox As System.Windows.Forms.TextBox = Nothing
        Dim LiCol As Integer = 0
        Dim Separ As String = ""
        Dim Sortowanie As String = Sort
        Dim ColName As String = ""
        Dim ColIndex As Integer = 0

        oFiltrZlozony = ""
        If Len(myFilTxt) > 0 Then
            'tBox = SzukajTextBoxONazwie(myForm, "txtFiltrZlozony")
            tBox = myForm.Controls("txtFiltrZlozony")

            If tBox Is Nothing Then
            Else
                tBox.BackColor = SoftartFilterBackColor

                If myDataView.Count > 0 Then
                    For LiCol = 0 To myColumnNo - 1

                        'ColName = srcDataTable.Columns(LiCol).ColumnName

                        'nazwe filtrowanej kolumny pobieramy z DATAGRIDVIEW
                        'gdyż kolumny w DGV nie muszą być w takiej samej kolejnosci co w DATAVIEW
                        'Index kolumny = TAG
                        ColName = myDataGridView.Columns(LiCol).DataPropertyName
                        ColIndex = myDataGridView.Columns(LiCol).Tag

                        If TypeOf myDataView.Item(0).Item(ColIndex) Is String Then
                            oFiltrZlozony &= Separ & _
                                    "(" & ColName & _
                                    " LIKE '*" & myFilTxt & "*')"
                            Separ = " OR "

                        ElseIf TypeOf myDataView.Item(0).Item(ColIndex) Is Integer Or _
                               TypeOf myDataView.Item(0).Item(ColIndex) Is Int16 Or _
                               TypeOf myDataView.Item(0).Item(ColIndex) Is Int32 Or _
                               TypeOf myDataView.Item(0).Item(ColIndex) Is Int64 Then
                            If IsNumeric(tBox.Text) Then
                                'oFiltrZlozony &= Separ & _
                                '        "(" & ColName & _
                                '        "=" & Trim(Str(tBox.Text)) & ")"
                                'Separ = " OR "
                            End If

                        ElseIf TypeOf myDataView.Item(0).Item(ColIndex) Is Double Then
                            If IsNumeric(tBox.Text) Then
                                'oFiltrZlozony &= Separ & _
                                '        "(" & ColName & _
                                '        "=" & Trim(Str(tBox.Text)) & ")"
                                'Separ = " OR "
                            End If

                        ElseIf TypeOf myDataView.Item(0).Item(ColIndex) Is Boolean Then
                            'Select Case UCase(tBox.Text)
                            '    Case "T", "Y", "TAK", "YES", "TRUE"
                            '        oFiltrZlozony &= Separ & _
                            '                        "(" & ColName & _
                            '                        "=1)"
                            '        Separ = " OR "
                            '    Case "N", "NIE", "NO", "FALSE"
                            '        oFiltrZlozony &= Separ & _
                            '                        "(" & ColName & _
                            '                        "=0)"
                            '        Separ = " OR "
                            '    Case Else
                            'End Select

                        ElseIf TypeOf myDataView.Item(0).Item(ColIndex) Is Guid Then
                            'oFiltrZlozony &= Separ & _
                            '        "(" & ColName & _
                            '        " LIKE '*" & tBox.Text.ToString & "*')"
                            'Separ = " OR "

                        Else        'nie zidentyfikowano typu pola....
                        End If
                    Next
                End If
            End If
        Else
            'tBox = SzukajTextBoxONazwie(myForm, "txtFiltrZlozony")
            tBox = myForm.Controls("txtFiltrZlozony")

            If tBox Is Nothing Then
            Else
                tBox.BackColor = SoftartEditableBackColor
            End If
        End If

        'Filtr sklada sie z filtru lokalnego i filtru zlozonego
        Dim SumFiltr As String = ""
        If Len(oLokalnyFiltr) > 0 Then
            SumFiltr = "(" & oLokalnyFiltr & ")"
            Separ = " AND "
        Else
            SumFiltr = ""
            Separ = ""
        End If
        If Len(oFiltrZlozony) > 0 Then
            SumFiltr = SumFiltr & Separ & "(" & oFiltrZlozony & ")"
        End If

        Try
            myDataView = New DataView(myDataSet.Tables(0), SumFiltr, Sortowanie, DataViewRowState.OriginalRows)

            myDataView.AllowDelete = False
            myDataView.AllowNew = False
            myDataGridView.DataSource = myDataView
            myDataGridView.RowHeadersWidth = 50       'use if tablestyle
            myDataGridView.Refresh()

            myStatusBar.Panels(0).Text = "Il.rec.: " & myDataView.Count.ToString
        Catch Ex As System.Exception
            MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & Ex.Message, "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Finally
        End Try
    End Sub




End Module
