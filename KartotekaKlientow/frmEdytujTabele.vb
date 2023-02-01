Public Class frmEdytujTabele

    Dim RowHeightSetter As Softart.myDataGridRowHeightSetter
    Dim pointInCell00 As Point
    Dim StartFormy As Boolean = True

    Dim dbSelectKlienci As String = "SELECT * FROM Klienci"

    Dim dbConnection As OleDb.OleDbConnection
    Dim dbCommandSelect As OleDb.OleDbCommand
    Dim dbCommandDelete As OleDb.OleDbCommand
    Dim dbCommandUpdate As OleDb.OleDbCommand
    Dim dbCommandInsert As OleDb.OleDbCommand
    Dim dbDataAdapter As OleDb.OleDbDataAdapter

    Dim sqlConnection As SqlClient.SqlConnection
    Dim sqlCommandSelect As SqlClient.SqlCommand
    Dim sqlCommandDelete As SqlClient.SqlCommand
    Dim sqlCommandUpdate As SqlClient.SqlCommand
    Dim sqlCommandInsert As SqlClient.SqlCommand
    Dim sqlDataAdapter As SqlClient.SqlDataAdapter

    Dim DataSetKlienci As New DataSet
    Dim DataViewKlienci As New DataView
    Dim ConnectionState As System.Data.ConnectionState

    Dim PoprawnaTabela As Boolean = False




    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        oGeneratorWydruku = ZdefiniujSzablonWydruku()
        Me.Close()
    End Sub

    Private Function ZdefiniujSzablonWydruku() As String
        Dim ij As Integer = 0
        Dim szablon As String = Trim(txtNazwa.Text) & vbCrLf
        Dim N As String = ""
        Dim T As String = ""
        Dim R As String = ""
        Dim H As String = ""
        Dim W As String = ""
        Dim I As String = ""

        Dim tBox As TextBox = Nothing

        'Dim TCol As New Softart.myDataGridTextBoxColumn
        'TCol.MappingName = MapName
        'TCol.HeaderText = HeadText
        'TCol.Width = ColWidth
        'TCol.Alignment = ColAllignment
        'TCol.Format = ColFormat
        'TabStyle.GridColumnStyles.Add(TCol)

        If PoprawnaTabela Then
            For ij = 0 To dagKlienci.TableStyles(0).GridColumnStyles.Count - 1
                N = dagKlienci.TableStyles(0).GridColumnStyles(ij).MappingName

                Select Case dagKlienci.TableStyles(0).GridColumnStyles(ij).GetType.ToString
                    Case "KartotekaKlientow.Softart.myDataGridTextBoxColumn"
                        T = "System.String"
                    Case "KartotekaKlientow.Softart.myDataGridNumColumn"
                        T = "System.Int32"
                    Case "KartotekaKlientow.Softart.myDataGridBoolColumn"
                        T = "System.Boolean"
                End Select

                R = pts2mm(dagKlienci.TableStyles(0).GridColumnStyles(ij).Width).ToString("N0")

                'H = dagKlienci.TableStyles(0).GridColumnStyles(ij).HeaderText
                tBox = SzukajTextBoxONazwie(Me, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ij)), 3))
                H = tBox.Text

                Select Case dagKlienci.TableStyles(0).GridColumnStyles(ij).Alignment
                    Case HorizontalAlignment.Center
                        W = "Centralnie"
                    Case HorizontalAlignment.Left
                        W = "Do lewej"
                    Case HorizontalAlignment.Right
                        W = "Do prawej"
                End Select

                I = "1"

                'Opis wydruku=NAZWAPOLA|TYPPOLA|SZEROKOSC|NAGLOWEK|WYROWNANIE|MAXILWIERSZY<crlf>
                szablon += NowePolaTabeliKlienci(N) & "|" & T & "|" & R & "|" & H & "|" & W & "|" & I & vbCrLf
            Next
        End If
        Return szablon
    End Function


    Private Function SzerokoscWydruku() As String
        Dim ij As Integer = 0
        Dim Szerokosc As Integer = 0
        If PoprawnaTabela Then
            For ij = 0 To dagKlienci.TableStyles(0).GridColumnStyles.Count - 1
                Szerokosc += pts2mm(dagKlienci.TableStyles(0).GridColumnStyles(ij).Width)
            Next
        End If
        Return Szerokosc
    End Function


    Private Sub frmEdytujTabele_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        'zbuduj kwerende SELECT
        Dim Gener As String = ""
        Dim pos As Integer = 0
        Dim kol As String = ""

        Dim N As String = ""
        Dim T As String = ""
        Dim R As String = ""
        Dim H As String = ""
        Dim W As String = ""
        Dim I As String = ""

        Dim Przec As String = " "

        Gener = _Generator
        dbSelectKlienci = "SELECT TOP 10 "

        PoprawnaTabela = False
        pos = InStr(Gener, vbCrLf)
        If pos > 0 Then
            txtNazwa.Text = Mid(Gener, 1, pos - 1)
            Gener = Mid(Gener, pos + 2)
            pos = InStr(Gener, vbCrLf)
            Do While pos > 0
                PoprawnaTabela = True

                kol = Mid(Gener, 1, pos - 1)
                Gener = Mid(Gener, pos + 2)

                pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK
                N = Mid(kol, 1, pos - 1)
                kol = Mid(kol, pos + 1)

                pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK
                T = Mid(kol, 1, pos - 1)
                kol = Mid(kol, pos + 1)

                pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
                R = Mid(kol, 1, pos - 1)
                kol = Mid(kol, pos + 1)

                pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
                H = Mid(kol, 1, pos - 1)
                kol = Mid(kol, pos + 1)

                pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
                W = Mid(kol, 1, pos - 1)
                I = Mid(kol, pos + 1)
                '---------------------------
                Select Case N
                    Case "KONTAKTYHANDLOWE"
                        ''---------------------------------------------------------------------
                        ''Kontakty HANDLOWE - nowe 05.09.2015
                        ''-----------------------------------------
                        'Public oUniqIdKon As String           'UNIQID            varchar(40)
                        'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
                        'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
                        'Public oPracownikKon As String        'PRACOWNIK        Text, 10     uzytkownik
                        'Public oTematKon As String            'TEMAT            Text, 50
                        'Public oUwagiKon As String            'UWAGI            Memo
                        'Public oWersjaKon As Integer          'WERSJA           Integer
                        dbSelectKlienci &= Przec & "(SELECT TOP 1 UWAGI FROM KontaktyHandlowe WHERE KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA ORDER BY DATAKONTAKTU DESC) AS KONTAKTYHANDLOWE"

                    Case "CODALEJ"
                        ''---------------------------------------------------------------------
                        ''CoDalejPlan
                        ''-----------------------------------------
                        'Public cdUNIQID As String      'UNIQID Text, 40
                        'Public cdIDENTKLIENTA As String      'IDENTKLIENTA Text, 40
                        'Public cdIDENT As String             'IDENT        Text, 40
                        'Public cdOPIS As String              'OPIS         memo
                        'Public cdWersja As Integer           'WERSJA       Integer
                        dbSelectKlienci &= Przec & "(SELECT TOP 1 OPIS FROM CoDalejPlan WHERE CoDalejPlan.IDENTKLIENTA=Klienci.IDENTKLIENTA) AS CODALEJ"

                    Case Else
                        'kolumna z tabeli klienci
                        dbSelectKlienci &= Przec & OryginalnePolaTabeliKlienci(N)

                End Select
                Przec = ","

                pos = InStr(Gener, vbCrLf)
            Loop
        End If
        dbSelectKlienci &= " FROM Klienci"

        '---------------------------------------
        If PoprawnaTabela Then
            DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

            If ParL_DataBaseType = "MS ACCESS" Then
                dbConnection = New OleDb.OleDbConnection(sConnectionKlienci)
                dbCommandSelect = New OleDb.OleDbCommand(dbSelectKlienci, dbConnection)
                'dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLKlienci, dbConnection)
                'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLKlienci, dbConnection)
                'dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLKlienci, dbConnection)
                dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

                'mapujemy domyslna nazwe tabeli
                Dim MapowanieTabeli As System.Data.Common.DataTableMapping
                MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Klienci")


                ' Add the RowUpdated event handler.
                AddHandler dbDataAdapter.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

                '------------------------------------------
                'wypelnij DATASET
                Try
                    dbConnection.Open()
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                Finally
                    ConnectionState = dbConnection.State
                End Try


            Else
                sqlConnection = New SqlClient.SqlConnection(sConnectionKlienci)
                sqlCommandSelect = New SqlClient.SqlCommand(dbSelectKlienci, sqlConnection)
                'sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLKlienci, sqlConnection)
                'sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLKlienci, sqlConnection)
                'sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLKlienci, sqlConnection)
                sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

                'mapujemy domyslna nazwe tabeli
                Dim MapowanieTabeli As System.Data.Common.DataTableMapping
                MapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_Klienci")

                ' Add the RowUpdated event handler.
                AddHandler sqlDataAdapter.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

                '------------------------------------------
                'wypelnij DATASET
                Try
                    sqlConnection.Open()
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                Finally
                    ConnectionState = sqlConnection.State
                End Try
            End If

            If ConnectionState = ConnectionState.Open Then
                Try
                    If ParL_DataBaseType = "MS ACCESS" Then
                        dbDataAdapter.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                        dbDataAdapter.Fill(DataSetKlienci)
                        dbConnection.Close()
                    Else
                        sqlDataAdapter.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                        sqlDataAdapter.Fill(DataSetKlienci)
                        sqlConnection.Close()
                    End If

                    'parametry DataGrid
                    dagKlienci.CaptionVisible = False
                    dagKlienci.CaptionText = DataSetKlienci.Tables(0).TableName
                    dagKlienci.ColumnHeadersVisible = True
                    dagKlienci.RowHeadersVisible = True
                    dagKlienci.BackgroundColor = System.Drawing.SystemColors.Control
                    dagKlienci.BorderStyle = BorderStyle.Fixed3D
                    dagKlienci.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                            Or System.Windows.Forms.AnchorStyles.Bottom) _
                                            Or System.Windows.Forms.AnchorStyles.Left) _
                                            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                    dagKlienci.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
                    dagKlienci.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))

                    'definiuj DataView
                    DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
                    DataViewKlienci.AllowDelete = False
                    DataViewKlienci.AllowEdit = False
                    DataViewKlienci.AllowNew = False
                    'wypelnienie DataGrid
                    'dagKlienci.SetDataBinding(DataSetKlienci, "TABELA_Klienci")
                    dagKlienci.DataSource = DataViewKlienci

                    ' Add a GridTableStyle and set the MappingName to the name of the DataTable.
                    Dim TSKlienci As New DataGridTableStyle
                    TSKlienci.MappingName = DataSetKlienci.Tables(0).TableName
                    TSKlienci.AlternatingBackColor = System.Drawing.SystemColors.Control
                    TSKlienci.BackColor = System.Drawing.SystemColors.ControlLight
                    TSKlienci.GridLineColor = System.Drawing.SystemColors.ControlDark
                    TSKlienci.ForeColor = System.Drawing.Color.Purple
                    TSKlienci.HeaderBackColor = System.Drawing.SystemColors.ScrollBar
                    TSKlienci.HeaderForeColor = System.Drawing.Color.Navy
                    TSKlienci.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                    TSKlienci.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow

                    ' Add a GridColumnStyle and set the MappingName 
                    ' to the name of a DataColumn in the DataTable. 
                    ' Set the HeaderText and Width properties. 


                    'generuj DataGrid
                    Dim Alignment As Windows.Forms.HorizontalAlignment
                    Dim szerokosc As Integer = 0
                    Dim ii As Integer = 0

                    Gener = _Generator
                    pos = InStr(Gener, vbCrLf)
                    If pos > 0 Then
                        'txtNazwa.Text = Mid(Gener, 1, pos - 1)
                        Gener = Mid(Gener, pos + 2)
                        pos = InStr(Gener, vbCrLf)
                        ii = 0
                        Do While pos > 0
                            kol = Mid(Gener, 1, pos - 1)
                            Gener = Mid(Gener, pos + 2)

                            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK
                            N = Mid(kol, 1, pos - 1)
                            kol = Mid(kol, pos + 1)

                            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK
                            T = Mid(kol, 1, pos - 1)
                            kol = Mid(kol, pos + 1)

                            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
                            R = Mid(kol, 1, pos - 1)
                            kol = Mid(kol, pos + 1)

                            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
                            H = Mid(kol, 1, pos - 1)
                            kol = Mid(kol, pos + 1)

                            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
                            W = Mid(kol, 1, pos - 1)
                            I = Mid(kol, pos + 1)
                            '---------------------------
                            If IsNumeric(R) Then
                                szerokosc = mm2pts(CInt(R))
                            Else
                                szerokosc = 100
                            End If

                            Select Case W
                                Case "Do lewej"
                                    Alignment = HorizontalAlignment.Left
                                Case "Do prawej"
                                    Alignment = HorizontalAlignment.Right
                                Case "Centralnie"
                                    Alignment = HorizontalAlignment.Center
                            End Select

                            Select Case T
                                Case "System.Boolean"
                                    LogColumn(TSKlienci, OryginalnePolaTabeliKlienci(N), N, szerokosc)
                                Case "System.String"
                                    TxtColumn(TSKlienci, OryginalnePolaTabeliKlienci(N), N, szerokosc, Alignment)
                                Case "System.Int32"
                                    NumColumn(TSKlienci, OryginalnePolaTabeliKlienci(N), N, szerokosc, Alignment, "N0")
                            End Select

                            'stworz zmienn¹ TXTBOX
                            Dim CtrlT As New TextBox
                            CtrlT.Name = "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
                            CtrlT.Visible = False
                            Me.Controls.Add(CtrlT)

                            CtrlT.BackColor = System.Drawing.SystemColors.Window
                            CtrlT.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
                            CtrlT.ForeColor = System.Drawing.Color.Navy
                            CtrlT.BorderStyle = BorderStyle.Fixed3D
                            CtrlT.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or _
                                                 System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
                            CtrlT.Text = H

                            pos = InStr(Gener, vbCrLf)
                            ii = ii + 1
                        Loop
                    End If

                    ' Add the DataGridTableStyle instance to the GridTableStylesCollection. 
                    dagKlienci.TableStyles.Add(TSKlienci)

                    'ustaw sie na pierwszym zapisie
                    If DataSetKlienci.Tables(0).Rows.Count > 0 Then
                        dagKlienci.Select(0)
                        dagKlienci.CurrentCell = New DataGridCell(0, 0)
                    End If

                Catch Ex As System.Exception
                    'ViewInLog(ex, Me.Name(), Nothing)
                    MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :", _
                        System.Windows.Forms.MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, _
                        MessageBoxDefaultButton.Button1)
                End Try
            End If
            '---------------------------------
            'dodaj StatusBar na koncu
            stbFiltry.BackColor = System.Drawing.SystemColors.Control
            stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
            stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)


            stbFiltry.Panels.Add("stbFiltryRowHead")
            stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbFiltry.Panels(0).Width = 50  'dagKlienci.TableStyles(0).RowHeaderWidth
            stbFiltry.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbFiltry.Panels(0).Text = "Nag³ówki: "

            stbFiltry.Panels.Add("stbFiltryReszta")
            stbFiltry.Panels(1).AutoSize = StatusBarPanelAutoSize.Spring
            'stbFiltry.Panels(1).Width = 150
            stbFiltry.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbFiltry.Panels(1).Text = ""

            stbFiltry.ShowPanels = True
            '-----------------------------------------

            '--------------------------------------------------------------------
            'definiuj przerwanie przy malowaniu DataGrid
            AddHandler Me.dagKlienci.Paint, New System.Windows.Forms.PaintEventHandler(AddressOf dagKlienci_Paint)

            'set the header width to something apporpriate
            dagKlienci.TableStyles(0).RowHeaderWidth = 50       'use if tablestyle
            'Me.dagKlienci.RowHeaderWidth = 80 'use if no tablestyle

            'set topleft corner point so we can locate the toprow
            If DataSetKlienci.Tables(0).Rows.Count > 0 Then
                pointInCell00 = New Point((dagKlienci.GetCellBounds(0, 0).X + 4), (dagKlienci.GetCellBounds(0, 0).Y + 4))
            Else
                pointInCell00 = New Point((dagKlienci.Left + 4), (dagKlienci.Top + 4))
            End If
            '--------------------------------------------------------------------
            'inicjowanie... 0=kolumna dla trybu Dingle/ 1=kolumna dla trybu multiline
            RowHeightSetter = New Softart.myDataGridRowHeightSetter(Me.dagKlienci, 0, 0)
            '--------------------------------------------------
            Me.WindowState = FormWindowState.Maximized
            RowHeightSetter.MultiHightsView()

            StartFormy = False
            PokazNaglowkiColumn(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
            lblSzerokosc.Text = "Szerokoœæ wydruku : " & Trim(Str(SzerokoscWydruku()))
        End If
    End Sub

    Private Sub dagKlienci_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKlienci.Resize
        If Not StartFormy Then
            PokazNaglowkiColumn(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
            lblSzerokosc.Text = "Szerokoœæ wydruku : " & Trim(Str(SzerokoscWydruku()))
        End If
    End Sub

    Private Sub dagKlienci_Scroll(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKlienci.Scroll
        If Not StartFormy Then
            PokazNaglowkiColumn(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
            lblSzerokosc.Text = "Szerokoœæ wydruku : " & Trim(Str(SzerokoscWydruku()))
        End If
    End Sub

    Private Sub dagKlienci_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKlienci.SizeChanged
        If Not StartFormy Then
            PokazNaglowkiColumn(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
            lblSzerokosc.Text = "Szerokoœæ wydruku : " & Trim(Str(SzerokoscWydruku()))
        End If
    End Sub

    Private Sub dagKlienci_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagKlienci.MouseMove
        'przypisz biezacy DataGrid do myGrid
        Dim myGrid As DataGrid = CType(sender, DataGrid)
        Dim hti As System.Windows.Forms.DataGrid.HitTestInfo
        hti = myGrid.HitTest(e.X, e.Y)
        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None     '"the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                If Not StartFormy Then
                    PokazNaglowkiColumn(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazNaglowkiColumn(Me, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    '**********************************************

    'wyswietla numer rekordu w RowHead
    Private Sub dagKlienci_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles dagKlienci.Paint
        Dim hti As Softart.MyDataGrid.HitTestInfo
        Dim row As Integer
        Dim yDelta As Integer
        Dim y As Integer

        Try
            If DataSetKlienci.Tables(0).Rows.Count > 0 Then
                hti = dagKlienci.HitTest(Me.pointInCell00)
                row = hti.Row
                yDelta = (dagKlienci.GetCellBounds(row, 0).Height + 1)
                y = (dagKlienci.GetCellBounds(row, 0).Top + 2)

                Dim cm As CurrencyManager
                Dim RowText As String
                cm = CType(Me.BindingContext(dagKlienci.DataSource, dagKlienci.DataMember), CurrencyManager)
                Do While ((y < (dagKlienci.Height - yDelta)) AndAlso (row < cm.Count))
                    RowText = (row + 1).ToString("N0")
                    e.Graphics.DrawString(RowText, dagKlienci.Font, New SolidBrush(Color.Navy), 12, y)
                    yDelta = (dagKlienci.GetCellBounds(row, 0).Height + 1)
                    y += yDelta
                    row += 1
                Loop
            End If
        Catch Ex As System.Exception
            'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
            '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    MessageBoxIcon.Information, _
            '    MessageBoxDefaultButton.Button1)
        End Try
    End Sub













    Public Sub PokazNaglowkiColumn(ByVal myForm As Control, _
                         ByVal myDataGrid As DataGrid, _
                         ByVal myColumnNo As Integer, _
                         ByVal myStatusBar As StatusBar, _
                         ByVal FormLoad As Boolean)
        If Not FormLoad Then
            'Dim IleCol As Integer = DataViewCennik.Table.Columns.Count
            Dim IleCol As Integer = myColumnNo
            Dim LiCol As Integer = 0
            Dim ThisColWidth As Integer = 0
            Dim ThisColLeft As Integer = 0
            Dim ThisColRight As Integer = 0
            Dim SumColWidth As Integer = 0
            Dim tBox As TextBox = Nothing

            Dim LeftVArea As Integer = myStatusBar.Location.X + myStatusBar.Panels(0).Width
            Dim RightVArea As Integer = myStatusBar.Location.X + myStatusBar.Width - 20


            Dim dv As DataView = myDataGrid.DataSource
            If dv.Count = 0 Then
                'tabela jest pusta...
                For LiCol = 0 To myColumnNo - 1
                    ThisColWidth = myDataGrid.TableStyles(0).GridColumnStyles(LiCol).Width
                    tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                    tBox.Location = New Point(LeftVArea, myStatusBar.Location.Y + 22)
                    tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
                    tBox.Visible = False
                    tBox.BringToFront()
                Next
            Else
                'analizuj kolumny przed obszarem widzialnym
                If myDataGrid.FirstVisibleColumn > 0 Then
                    For LiCol = 0 To myDataGrid.FirstVisibleColumn - 1
                        ThisColWidth = myDataGrid.TableStyles(0).GridColumnStyles(LiCol).Width

                        'ThisColWidth = myDataGrid.GetCellBounds(0, LiCol).Width
                        tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                        tBox.Location = New Point(LeftVArea, myStatusBar.Location.Y + 22)
                        tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
                        tBox.Visible = False
                        tBox.BringToFront()
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
                        tBox.Location = New Point(LeftVArea, _
                                                  myStatusBar.Location.Y + 22)
                        tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
                        tBox.Visible = False
                        tBox.BringToFront()
                    Else
                        'czy kolumna zaczyna sie przed czescia widzialn¹
                        If myDataGrid.Left + ThisColLeft < LeftVArea Then
                            ThisColWidth -= LeftVArea - ThisColLeft - myDataGrid.Left
                        End If
                        'czy kolumna koñczy siê za czêœci¹ wudzialn¹
                        If myDataGrid.Left + ThisColRight > RightVArea Then
                            ThisColWidth -= ThisColRight + myDataGrid.Left - RightVArea
                        End If

                        tBox.Location = New Point(LeftVArea + SumColWidth, _
                                                  myStatusBar.Location.Y + 2)
                        tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
                        tBox.Visible = True
                        tBox.BringToFront()

                        SumColWidth += ThisColWidth + 1
                    End If
                Next

                'analizuj kolumny za obszarem widzialnym
                If myDataGrid.FirstVisibleColumn + myDataGrid.VisibleColumnCount < IleCol - 1 Then
                    For LiCol = myDataGrid.FirstVisibleColumn + myDataGrid.VisibleColumnCount To IleCol - 1
                        ThisColWidth = myDataGrid.GetCellBounds(0, LiCol).Width
                        tBox = SzukajTextBoxONazwie(myForm, "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(LiCol)), 3))
                        tBox.Location = New Point(LeftVArea, myStatusBar.Location.Y + 22)
                        tBox.Size = New System.Drawing.Size(ThisColWidth, 20)
                        tBox.Visible = False
                        tBox.BringToFront()
                    Next
                End If
            End If
        End If
    End Sub
End Class