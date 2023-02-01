Public Class AkcjaMarketingowaWybierzKilka

    'Akcje Handlowe - Opis
    'Public aoIdentAkcji As String            'IDENT     Text 20
    'Public aoDataAkcji As String             'DATA      Text,10
    'Public aoOpisAkcji As String             'OPIS      Text,100
    'Public aoUwagiAkcji As String            'UWAGI     Text,10
    'Public aoMarkerAkcji As String           'MARKER    logical
    'Public aoWersjaAkcji As Integer           'WERSJA    Integer

    Dim dbSelect As String = sSelectSQLAkcjeOpis

    Dim dbConnectionAkcjeOpis As OleDb.OleDbConnection
    Dim dbCommandSelectAkcjeOpis As OleDb.OleDbCommand
    Dim dbCommandDeleteAkcjeOpis As OleDb.OleDbCommand
    Dim dbCommandUpdateAkcjeOpis As OleDb.OleDbCommand
    Dim dbCommandInsertAkcjeOpis As OleDb.OleDbCommand
    Dim dbDataAdapterAkcjeOpis As OleDb.OleDbDataAdapter

    Dim sqlConnectionAkcjeOpis As SqlClient.SqlConnection
    Dim sqlCommandSelectAkcjeOpis As SqlClient.SqlCommand
    Dim sqlCommandDeleteAkcjeOpis As SqlClient.SqlCommand
    Dim sqlCommandUpdateAkcjeOpis As SqlClient.SqlCommand
    Dim sqlCommandInsertAkcjeOpis As SqlClient.SqlCommand
    Dim sqlDataAdapterAkcjeOpis As SqlClient.SqlDataAdapter

    Dim DataSetAkcjeOpis As New DataSet
    Dim DataViewAkcjeOpis As New DataView
    Dim ConnectionState As System.Data.ConnectionState

    Dim OdDaty As String = ""
    Dim DoDaty As String = ""
    Dim TakenDate As Date = Nothing

    Private Sub WybierzAkcjeMarketingowa_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        oWybranoAkcjeMarketingowa = False
        '-------------------------------
        dtpOdDaty.Value = Mid(SysData(), 1, 7) & "-01"
        dtpDoDaty.Value = SysData()
        rbtWszyscy.Checked = True
        rbtWybierzWszystkich.Checked = True
        '---------------------------------------
        '-------------------------------------
        DataSetAkcjeOpis.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionAkcjeOpis = New OleDb.OleDbConnection(sConnectionAkcjeOpis)
            dbCommandSelectAkcjeOpis = New OleDb.OleDbCommand(sSelectSQLAkcjeOpis, dbConnectionAkcjeOpis)
            dbCommandDeleteAkcjeOpis = New OleDb.OleDbCommand(sDeleteSQLAkcjeOpis, dbConnectionAkcjeOpis)
            dbCommandUpdateAkcjeOpis = New OleDb.OleDbCommand(sUpdateSQLAkcjeOpis, dbConnectionAkcjeOpis)
            dbCommandInsertAkcjeOpis = New OleDb.OleDbCommand(sInsertSQLAkcjeOpis, dbConnectionAkcjeOpis)
            dbDataAdapterAkcjeOpis = New OleDb.OleDbDataAdapter(dbCommandSelectAkcjeOpis)

            Dim dbMapowanieTabeli As System.Data.Common.DataTableMapping
            dbMapowanieTabeli = dbDataAdapterAkcjeOpis.TableMappings.Add("Table", "TABELA_AkcjeOpis")

            Dim objParam As OleDb.OleDbParameter
            '------------------------------------------
            'komenda DELETE
            'parametry filtrowania
            objParam = dbCommandDeleteAkcjeOpis.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 20, "IDENT")
            objParam.SourceVersion = DataRowVersion.Original
            objParam = dbCommandDeleteAkcjeOpis.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Original
            dbDataAdapterAkcjeOpis.DeleteCommand = dbCommandDeleteAkcjeOpis

            '------------------------------------------
            'komenda UPDATE
            'parametry aktualizacji
            'dbCommandUpdateAkcjeOpis.Parameters.Add("@oIdentAkcji", OleDb.OleDbType.Char, 20, "IDENT")
            dbCommandUpdateAkcjeOpis.Parameters.Add("@oDataAkcji", OleDb.OleDbType.Char, 10, "DATA")
            dbCommandUpdateAkcjeOpis.Parameters.Add("@oOpisAkcji", OleDb.OleDbType.Char, 100, "OPIS")
            dbCommandUpdateAkcjeOpis.Parameters.Add("@oUwagiAkcji", OleDb.OleDbType.WChar, Nothing, "UWAGI")
            dbCommandUpdateAkcjeOpis.Parameters.Add("@oMarkerAkcji", OleDb.OleDbType.Boolean, Nothing, "MARKER")
            dbCommandUpdateAkcjeOpis.Parameters.Add("@oWersjaAkcji", OleDb.OleDbType.Integer, Nothing, "WERSJA")

            'parametry filtrowania
            objParam = dbCommandUpdateAkcjeOpis.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 20, "IDENT")
            objParam.SourceVersion = DataRowVersion.Original
            objParam = dbCommandUpdateAkcjeOpis.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Original
            dbDataAdapterAkcjeOpis.UpdateCommand = dbCommandUpdateAkcjeOpis
            '------------------------------------------
            'komenda INSERT
            objParam = dbCommandInsertAkcjeOpis.Parameters.Add("@oIdentAkcji", OleDb.OleDbType.Char, 20, "IDENT")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAkcjeOpis.Parameters.Add("@oDataAkcji", OleDb.OleDbType.Char, 10, "DATA")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAkcjeOpis.Parameters.Add("@oOpisAkcji", OleDb.OleDbType.Char, 100, "OPIS")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAkcjeOpis.Parameters.Add("@oUwagiAkcji", OleDb.OleDbType.WChar, Nothing, "UWAGI")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAkcjeOpis.Parameters.Add("@oMarkerAkcji", OleDb.OleDbType.Boolean, Nothing, "MARKER")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertAkcjeOpis.Parameters.Add("@oWersjaAkcji", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Current

            dbDataAdapterAkcjeOpis.InsertCommand = dbCommandInsertAkcjeOpis

            ' Add the RowUpdated event handler.
            AddHandler dbDataAdapterAkcjeOpis.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionAkcjeOpis.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = dbConnectionAkcjeOpis.State
            End Try
        Else
            sqlConnectionAkcjeOpis = New SqlClient.SqlConnection(sConnectionAkcjeOpis)
            sqlCommandSelectAkcjeOpis = New SqlClient.SqlCommand(sSelectSQLAkcjeOpis, sqlConnectionAkcjeOpis)
            sqlCommandDeleteAkcjeOpis = New SqlClient.SqlCommand(sDeleteSQLAkcjeOpis, sqlConnectionAkcjeOpis)
            sqlCommandUpdateAkcjeOpis = New SqlClient.SqlCommand(sUpdateSQLAkcjeOpis, sqlConnectionAkcjeOpis)
            sqlCommandInsertAkcjeOpis = New SqlClient.SqlCommand(sInsertSQLAkcjeOpis, sqlConnectionAkcjeOpis)
            sqlDataAdapterAkcjeOpis = New SqlClient.SqlDataAdapter(sqlCommandSelectAkcjeOpis)

            Dim sqlMapowanieTabeli As System.Data.Common.DataTableMapping
            sqlMapowanieTabeli = sqlDataAdapterAkcjeOpis.TableMappings.Add("Table", "TABELA_AkcjeOpis")

            Dim objParam As SqlClient.SqlParameter
            '------------------------------------------
            'komenda DELETE
            'parametry filtrowania
            objParam = sqlCommandDeleteAkcjeOpis.Parameters.Add("@orygSymbol", SqlDbType.Char, 20, "IDENT")
            objParam.SourceVersion = DataRowVersion.Original
            objParam = sqlCommandDeleteAkcjeOpis.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Original
            sqlDataAdapterAkcjeOpis.DeleteCommand = sqlCommandDeleteAkcjeOpis

            '------------------------------------------
            'komenda UPDATE
            'parametry aktualizacji
            'SQLCommandUpdateAkcjeOpis.Parameters.Add("@oIdentAkcji", SqlDbType.Char, 20, "IDENT")
            sqlCommandUpdateAkcjeOpis.Parameters.Add("@oDataAkcji", SqlDbType.Char, 10, "DATA")
            sqlCommandUpdateAkcjeOpis.Parameters.Add("@oOpisAkcji", SqlDbType.VarChar, 100, "OPIS")
            sqlCommandUpdateAkcjeOpis.Parameters.Add("@oUwagiAkcji", SqlDbType.Text, Nothing, "UWAGI")
            sqlCommandUpdateAkcjeOpis.Parameters.Add("@oMarkerAkcji", SqlDbType.Bit, Nothing, "MARKER")
            sqlCommandUpdateAkcjeOpis.Parameters.Add("@oWersjaAkcji", SqlDbType.Int, Nothing, "WERSJA")

            'parametry filtrowania
            objParam = sqlCommandUpdateAkcjeOpis.Parameters.Add("@orygSymbol", SqlDbType.Char, 20, "IDENT")
            objParam.SourceVersion = DataRowVersion.Original
            objParam = sqlCommandUpdateAkcjeOpis.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Original
            sqlDataAdapterAkcjeOpis.UpdateCommand = sqlCommandUpdateAkcjeOpis
            '------------------------------------------
            'komenda INSERT
            objParam = sqlCommandInsertAkcjeOpis.Parameters.Add("@oIdentAkcji", SqlDbType.Char, 20, "IDENT")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAkcjeOpis.Parameters.Add("@oDataAkcji", SqlDbType.Char, 10, "DATA")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAkcjeOpis.Parameters.Add("@oOpisAkcji", SqlDbType.VarChar, 100, "OPIS")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAkcjeOpis.Parameters.Add("@oUwagiAkcji", SqlDbType.Text, Nothing, "UWAGI")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAkcjeOpis.Parameters.Add("@oMarkerAkcji", SqlDbType.Bit, Nothing, "MARKER")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = sqlCommandInsertAkcjeOpis.Parameters.Add("@oWersjaAkcji", SqlDbType.Char, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Current

            sqlDataAdapterAkcjeOpis.InsertCommand = sqlCommandInsertAkcjeOpis

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterAkcjeOpis.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionAkcjeOpis.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionAkcjeOpis.State
            End Try
        End If


        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    dbDataAdapterAkcjeOpis.FillSchema(DataSetAkcjeOpis, SchemaType.Mapped, "TABELA_AkcjeOpis")
                    dbDataAdapterAkcjeOpis.Fill(DataSetAkcjeOpis)
                    dbConnectionAkcjeOpis.Close()
                Else
                    sqlDataAdapterAkcjeOpis.FillSchema(DataSetAkcjeOpis, SchemaType.Mapped, "TABELA_AkcjeOpis")
                    sqlDataAdapterAkcjeOpis.Fill(DataSetAkcjeOpis)
                    sqlConnectionAkcjeOpis.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetAkcjeOpis.Tables(0).PrimaryKey = New DataColumn() {DataSetAkcjeOpis.Tables(0).Columns("IDENT")}

                'parametry DataGrid
                dagAkcjeOpis.CaptionVisible = False
                dagAkcjeOpis.CaptionText = DataSetAkcjeOpis.Tables(0).TableName
                dagAkcjeOpis.ColumnHeadersVisible = True
                dagAkcjeOpis.RowHeadersVisible = True
                dagAkcjeOpis.BackgroundColor = System.Drawing.SystemColors.Control
                dagAkcjeOpis.BorderStyle = BorderStyle.Fixed3D
                dagAkcjeOpis.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                dagAkcjeOpis.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
                dagAkcjeOpis.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))

                'definiuj DataView
                DataViewAkcjeOpis = New DataView(DataSetAkcjeOpis.Tables(0))
                'DataViewAkcjeOpis.AllowDelete = False
                'DataViewAkcjeOpis.AllowEdit = False
                'DataViewAkcjeOpis.AllowNew = False
                'wypelnienie DataGrid



                'inicjuj markery
                If DataViewAkcjeOpis.Count > 0 Then
                    Dim drv As DataRowView = Nothing
                    For Each drv In DataViewAkcjeOpis
                        drv("MARKER") = False
                    Next
                    AktualizujBazeAkcjeOpis()
                End If






                'dagAkcjeOpis.SetDataBinding(DataSetAkcjeOpis, "TABELA_AkcjeOpis")
                dagAkcjeOpis.DataSource = DataViewAkcjeOpis

                ' Add a GridTableStyle and set the MappingName to the name of the DataTable.
                Dim TSAkcjeOpis As New DataGridTableStyle
                TSAkcjeOpis.MappingName = DataSetAkcjeOpis.Tables(0).TableName
                TSAkcjeOpis.AlternatingBackColor = System.Drawing.SystemColors.Control
                TSAkcjeOpis.BackColor = System.Drawing.SystemColors.ControlLight
                TSAkcjeOpis.GridLineColor = System.Drawing.SystemColors.ControlDark
                TSAkcjeOpis.ForeColor = System.Drawing.Color.Purple
                TSAkcjeOpis.HeaderBackColor = System.Drawing.SystemColors.ScrollBar
                TSAkcjeOpis.HeaderForeColor = System.Drawing.Color.Navy
                'TSAkcjeOpis.SelectionBackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(192, Byte), CType(0, Byte))
                TSAkcjeOpis.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                TSAkcjeOpis.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow

                'definiuj delegat
                Dim deleg As delegateDefColor = AddressOf Koloruj

                TxtColoredColumnF(deleg, TSAkcjeOpis, DataSetAkcjeOpis.Tables(0).Columns(0).ColumnName, "Ident", 100, HorizontalAlignment.Left)
                TxtColoredColumnF(deleg, TSAkcjeOpis, DataSetAkcjeOpis.Tables(0).Columns(1).ColumnName, "Data", 80, HorizontalAlignment.Left)
                TxtColoredColumnF(deleg, TSAkcjeOpis, DataSetAkcjeOpis.Tables(0).Columns(2).ColumnName, "Opis", 300, HorizontalAlignment.Left)
                TxtColoredColumnF(deleg, TSAkcjeOpis, DataSetAkcjeOpis.Tables(0).Columns(3).ColumnName, "", 0, HorizontalAlignment.Left)
                LogColumn(TSAkcjeOpis, DataSetAkcjeOpis.Tables(0).Columns(4).ColumnName, "Marker", 40)
                NumColoredColumnF(deleg, TSAkcjeOpis, DataSetAkcjeOpis.Tables(0).Columns(5).ColumnName, "", 0, HorizontalAlignment.Left, "F0")

                ' Add the DataGridTableStyle instance to the GridTableStylesCollection. 
                dagAkcjeOpis.TableStyles.Add(TSAkcjeOpis)
                'ustaw sie na pierwszym zapisie
                If DataSetAkcjeOpis.Tables(0).Rows.Count > 0 Then dagAkcjeOpis.Select(0)

            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            'dodaj StatusBar na koncu
            stbAkcjeOpis.BackColor = System.Drawing.SystemColors.Control
            stbAkcjeOpis.ForeColor = System.Drawing.Color.DarkGreen
            stbAkcjeOpis.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            stbAkcjeOpis.Panels.Add("stbPanelIleRec")
            stbAkcjeOpis.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbAkcjeOpis.Panels(0).Width = 80
            stbAkcjeOpis.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbAkcjeOpis.Panels(0).Text = "Iloœæ rec.: " & DataSetAkcjeOpis.Tables(0).Rows.Count.ToString

            stbAkcjeOpis.Panels.Add("stbPanelWiersz")
            stbAkcjeOpis.Panels(1).AutoSize = StatusBarPanelAutoSize.None
            stbAkcjeOpis.Panels(1).Width = 80
            stbAkcjeOpis.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbAkcjeOpis.Panels(1).Text = "Wiersz : " & dagAkcjeOpis.CurrentCell.RowNumber.ToString

            stbAkcjeOpis.Panels.Add("stbPanelKolumna")
            stbAkcjeOpis.Panels(2).AutoSize = StatusBarPanelAutoSize.None
            stbAkcjeOpis.Panels(2).Width = 80
            stbAkcjeOpis.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbAkcjeOpis.Panels(2).Text = "Kolumna : " & dagAkcjeOpis.CurrentCell.ColumnNumber.ToString

            stbAkcjeOpis.Panels.Add("stbPanelSort")
            stbAkcjeOpis.Panels(3).AutoSize = StatusBarPanelAutoSize.None
            stbAkcjeOpis.Panels(3).Width = 150
            stbAkcjeOpis.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbAkcjeOpis.Panels(3).Text = "Sort : "

            stbAkcjeOpis.Panels.Add("stbPanelFiltr")
            stbAkcjeOpis.Panels(4).AutoSize = StatusBarPanelAutoSize.None
            stbAkcjeOpis.Panels(4).Width = 150
            stbAkcjeOpis.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbAkcjeOpis.Panels(4).Text = "Filtr : "

            stbAkcjeOpis.Panels.Add("stbPanelReszta")
            stbAkcjeOpis.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
            'stbAkcjeOpis.Panels(5).Width = 150
            stbAkcjeOpis.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbAkcjeOpis.Panels(5).Text = ""

            stbAkcjeOpis.ShowPanels = True
            '---------------------------------
        End If
        InitAkcjeOpis() 'inicjuj zmienne
        '-----------------------------------------------------------------
    End Sub

    Private Sub WybierzAkcjeMarketingowa_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        RysujGradient(Me)
    End Sub


    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub





    Private Sub dagAkcjeOpis_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAkcjeOpis.CurrentCellChanged
        stbAkcjeOpis.Panels(1).Text = "Wiersz : " & dagAkcjeOpis.CurrentCell.RowNumber.ToString
        stbAkcjeOpis.Panels(2).Text = "Kolumna : " & dagAkcjeOpis.CurrentCell.ColumnNumber.ToString
    End Sub

    Private Sub dagAkcjeOpis_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagAkcjeOpis.MouseUp
        'przypisz biezacy DataGrid do myGrid
        Dim myGrid As DataGrid = CType(sender, DataGrid)
        Dim hti As System.Windows.Forms.DataGrid.HitTestInfo
        hti = myGrid.HitTest(e.X, e.Y)

        Dim message As String = "You clicked "

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None
                message &= "the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                message &= "cell at row " & hti.Row & ", col " & hti.Column
                stbAkcjeOpis.Panels(1).Text = "Wiersz : " & hti.Row.ToString
                stbAkcjeOpis.Panels(2).Text = "Kolumna : " & hti.Column.ToString
                myGrid.Select(hti.Row)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.Column
                stbAkcjeOpis.Panels(2).Text = "Kolumna : " & hti.Column.ToString
                stbAkcjeOpis.Panels(3).Text = "Sort : " & DataSetAkcjeOpis.Tables(0).Columns(hti.Column).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.Row
                stbAkcjeOpis.Panels(1).Text = "Wiersz : " & hti.Row.ToString
                myGrid.Select(hti.Row)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.Column
                stbAkcjeOpis.Panels(2).Text = "Kolumna : " & hti.Column.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.Row
                stbAkcjeOpis.Panels(1).Text = "Wiersz : " & hti.Row.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagAkcjeOpis_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAkcjeOpis.DoubleClick
        'cmdEdytuj_Click(sender, e)
        OznaczMarker()
    End Sub

    Private Sub dagAkcjeOpis_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagAkcjeOpis.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                OznaczMarker()
                'cmdEdytuj_Click(sender, e)
            Case Keys.Insert
                'cmdDodaj_Click(sender, e)
            Case Keys.Delete
                'cmdUsuñ_Click(sender, e)
            Case Else
        End Select
    End Sub

    '*************************************************
    '*** przenoszenie danych miêdzy rekordem i zmiennymi
    '*************************************************

    Private Sub InitAkcjeOpis()
        aoIdentAkcji = ""        'IDENT       Text, 10
        aoDataAkcji = SysData()
        aoOpisAkcji = ""
        aoUwagiAkcji = ""
        aoMarkerAkcji = False
        aoWersjaAkcji = 1        'WERSJA      Integer, 2
    End Sub

    Private Sub PobierzAkcjeOpis()
        'pobierz wartosci przed aktualizacja
        aoIdentAkcji = IIf(IsDBNull(dagAkcjeOpis.Item(dagAkcjeOpis.CurrentCell.RowNumber, 0)), "", dagAkcjeOpis.Item(dagAkcjeOpis.CurrentCell.RowNumber, 0))
        aoDataAkcji = IIf(IsDBNull(dagAkcjeOpis.Item(dagAkcjeOpis.CurrentCell.RowNumber, 1)), "", dagAkcjeOpis.Item(dagAkcjeOpis.CurrentCell.RowNumber, 1))
        aoOpisAkcji = IIf(IsDBNull(dagAkcjeOpis.Item(dagAkcjeOpis.CurrentCell.RowNumber, 2)), "", dagAkcjeOpis.Item(dagAkcjeOpis.CurrentCell.RowNumber, 2))
        aoUwagiAkcji = IIf(IsDBNull(dagAkcjeOpis.Item(dagAkcjeOpis.CurrentCell.RowNumber, 3)), "", dagAkcjeOpis.Item(dagAkcjeOpis.CurrentCell.RowNumber, 3))
        aoMarkerAkcji = IIf(IsDBNull(dagAkcjeOpis.Item(dagAkcjeOpis.CurrentCell.RowNumber, 4)), False, dagAkcjeOpis.Item(dagAkcjeOpis.CurrentCell.RowNumber, 4))
        aoWersjaAkcji = IIf(IsDBNull(dagAkcjeOpis.Item(dagAkcjeOpis.CurrentCell.RowNumber, 5)), 0, dagAkcjeOpis.Item(dagAkcjeOpis.CurrentCell.RowNumber, 5))
    End Sub

    Private Sub OznaczMarker()
        If DataSetAkcjeOpis.Tables(0).Rows.Count <= 0 Then
            'MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf & _
            '    "Proszê najpierw dopisaæ rekord do tabeli...", _
            '    "Uwaga :", _
            '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    MessageBoxIcon.Information, _
            '    MessageBoxDefaultButton.Button1)
        Else
            PobierzAkcjeOpis()
            Dim foundRow As DataRow
            ' Create an array for the key values to find.
            Dim findTheseVals(0) As Object
            findTheseVals(0) = aoIdentAkcji
            foundRow = DataSetAkcjeOpis.Tables(0).Rows.Find(findTheseVals)
            ' sprawdz czy jest...
            If Not (foundRow Is Nothing) Then
                'foundRow("IDENT") = aoidentAkcji
                'foundRow("DATA") = aoDataAkcji           'OPIS      Text 50
                'foundRow("OPIS") = aoOpisAkcji           'OPIS      Text 50
                'foundRow("UWAGI") = aoUwagiAkcji           'OPIS      Text 50
                foundRow("MARKER") = Not aoMarkerAkcji           'OPIS      Text 50
                foundRow("WERSJA") = ZmienWersje(aoWersjaAkcji)    'Wersja         Integer, 2

                'wyswietl ilosc rekordow
                stbAkcjeOpis.Panels(0).Text = "Iloœæ rec.: " & DataSetAkcjeOpis.Tables(0).Rows.Count.ToString

                AktualizujBazeAkcjeOpis()

                'aktualizuj DataGrid
                dagAkcjeOpis.Update()
            End If
        End If
    End Sub

    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujBazeAkcjeOpis()
        If ParL_DataBaseType = "MS ACCESS" Then
            Try
                dbConnectionAkcjeOpis.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If dbConnectionAkcjeOpis.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    dbDataAdapterAkcjeOpis.Update(DataSetAkcjeOpis, "TABELA_AkcjeOpis")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    dbDataAdapterAkcjeOpis.Fill(DataSetAkcjeOpis)
                End If
                dbConnectionAkcjeOpis.Close()
            End If
        Else
            Try
                sqlConnectionAkcjeOpis.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionAkcjeOpis.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterAkcjeOpis.Update(DataSetAkcjeOpis, "TABELA_AkcjeOpis")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterAkcjeOpis.Fill(DataSetAkcjeOpis)
                End If
                sqlConnectionAkcjeOpis.Close()
            End If
        End If
    End Sub

    '**************************************************************
    '** procedura okreslania koloru w DataGrid
    '**************************************************************

    'Akcje Handlowe - Opis
    'Public aoIdentAkcji As String            'IDENT     Text 20
    'Public aoDataAkcji As String             'DATA      Text,10
    'Public aoOpisAkcji As String             'OPIS      Text,100
    'Public aoUwagiAkcji As String            'UWAGI     Text,10
    'Public aoMarkerAkcji As String           'MARKER    logical
    'Public aoWersjaAkcji As Integer           'WERSJA    Integer

    Public Function Koloruj(ByVal iRow As Long) As Integer
        Dim Kolor As Integer = System.Drawing.Color.Purple.ToArgb
        Dim bMarker As Boolean = IIf(IsDBNull(DataViewAkcjeOpis.Item(iRow).Item("MARKER")), False, DataViewAkcjeOpis.Item(iRow).Item("MARKER"))
        Try
            If bMarker Then
                Kolor = System.Drawing.Color.Red.ToArgb
            Else
                Kolor = System.Drawing.Color.Navy.ToArgb
            End If

        Catch ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.Message & vbCrLf & ex.StackTrace,
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
        End Try

        Return (Kolor)
    End Function




    '**************************************************
    '** zapamietaj informacje o Akcjach...
    '**************************************************


    '------------------------------------------------
    'AkcJe handlowe - spec
    Dim dbSelectSQLAkcjeSpec As String = sSelectSQLAkcjeSpec

    Dim dbConnectionAkcjeSpec As OleDb.OleDbConnection
    Dim dbCommandSelectAkcjeSpec As OleDb.OleDbCommand
    Dim dbCommandDeleteAkcjeSpec As OleDb.OleDbCommand
    Dim dbCommandUpdateAkcjeSpec As OleDb.OleDbCommand
    Dim dbCommandInsertAkcjeSpec As OleDb.OleDbCommand
    Dim dbDataAdapterAkcjeSpec As OleDb.OleDbDataAdapter

    Dim sqlConnectionAkcjeSpec As SqlClient.SqlConnection
    Dim sqlCommandSelectAkcjeSpec As SqlClient.SqlCommand
    Dim sqlCommandDeleteAkcjeSpec As SqlClient.SqlCommand
    Dim sqlCommandUpdateAkcjeSpec As SqlClient.SqlCommand
    Dim sqlCommandInsertAkcjeSpec As SqlClient.SqlCommand
    Dim sqlDataAdapterAkcjeSpec As SqlClient.SqlDataAdapter

    Dim DataSetAkcjeSpec As DataSet
    Dim DataViewAkcjeSpec As DataView



    Private Sub cmdZapamietaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdZapamietaj.Click
        PobierzKlientow(True)
    End Sub


    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        PobierzKlientow(False)
    End Sub



    Private Sub PobierzKlientow(ByVal parInicjuj As Boolean)
        Dim pIdentAkcji As String = ""

        If rbtWszyscy.Checked Then
            OdDaty = ""
            DoDaty = ""
        ElseIf rbtObsluzeni.Checked Or rbtNieObsluzeni.Checked Then
            TakenDate = dtpOdDaty.Value
            OdDaty = TakenDate.ToString("yyyy-MM-dd")
            TakenDate = dtpDoDaty.Value
            DoDaty = TakenDate.ToString("yyyy-MM-dd")
        Else        'def branze
            OdDaty = ""
            DoDaty = ""
        End If
        '-------------------------------
        If _dsKlenci.Tables(0).Rows.Count = 0 Then
            MessageBox.Show("Specyfikacja klientów jest pusta.",
                "Uwaga:",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            DataSetAkcjeSpec = New DataSet
            DataViewAkcjeSpec = New DataView
            DataSetAkcjeSpec.Locale = New System.Globalization.CultureInfo("pl-PL")

            '---------------------------------------------------------------------
            'Akcje Handlowe - Opis
            'Public aoIdentAkcji As String            'IDENT     Text 20
            'Public aoDataAkcji As String             'DATA      Text,10
            'Public aoOpisAkcji As String             'OPIS      Text,100
            'Public aoUwagiAkcji As String            'UWAGI     Text,10
            'Public aoMarkerAkcji As String           'MARKER    logical
            'Public aoWersjaAkcji As Integer           'WERSJA    Integer
            '---------------------------------------------------------------------
            'Akcje Handlowe - Spec
            'Public asIdentAkcji As String            'IDENTAKCJI     Text 20
            'Public asIdentKlienta As String          'IDENTKLI       Text 6
            'Public asWersjaAkcji As Integer           'WERSJA    Integer
            '---------------------------------------------------------------------
            'Kontakty
            'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
            'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
            'Public oPracownikKon As String        'PRACOWNIK        Text, 2     uzytkownik
            'Public oTematKon As String            'TEMAT            Text, 50
            'Public oUwagiKon As String            'UWAGI            Memo
            'Public oWersjaKon As Integer          'WERSJA           Integer
            '---------------------------------------------------------------------


            'If rbtWszyscy.Checked Then
            '    dbSelectSQLAkcjeSpec = "SELECT IDENTAKCJI, " & _
            '                                  "IDENTKLI " & _
            '                           "FROM AkcjeSpec AS ASp INNER JOIN (SELECT * FROM AkcjeOpis WHERE MARKER='TRUE') AS AOp " & _
            '                           "ON ASp.IDENTAKCJI=AOp.IDENT " & _
            '                           " "
            'ElseIf rbtObsluzeni.Checked Then
            '    dbSelectSQLAkcjeSpec = "SELECT IDENTAKCJI, " & _
            '                                  "IDENTKLI " & _
            '                           "FROM AkcjeSpec AS ASp INNER JOIN (SELECT * FROM AkcjeOpis WHERE MARKER='TRUE') AS AOp " & _
            '                           "ON ASp.IDENTAKCJI=AOp.IDENT " & _
            '        " WHERE ((SELECT count(*) FROM KontaktyHandlowe WHERE (DATAKONTAKTU>='" & OdDaty & "') AND " & _
            '                                                    "(DATAKONTAKTU<='" & DoDaty & "') AND " & _
            '                                                    "(KontaktyHandlowe.IDENTKLIENTA=AkcjeSpec.IDENTKLI)) > 0) "
            'Else
            '    dbSelectSQLAkcjeSpec = "SELECT IDENTAKCJI, " & _
            '                                  "IDENTKLI " & _
            '                           "FROM AkcjeSpec AS ASp INNER JOIN (SELECT * FROM AkcjeOpis WHERE MARKER='TRUE') AS AOp " & _
            '                           "ON ASp.IDENTAKCJI=AOp.IDENT " & _
            '        " WHERE ((SELECT count(*) FROM KontaktyHandlowe WHERE (DATAKONTAKTU>='" & OdDaty & "') AND " & _
            '                                                    "(DATAKONTAKTU<='" & DoDaty & "') AND " & _
            '                                                    "(KontaktyHandlowe.IDENTKLIENTA=AkcjeSpec.IDENTKLI)) = 0) "
            'End If


            If rbtWszyscy.Checked Then
                dbSelectSQLAkcjeSpec = "SELECT DISTINCT " &
                                               "IDENTKLI " &
                                       "FROM AkcjeSpec AS ASp INNER JOIN (SELECT * FROM AkcjeOpis WHERE MARKER='TRUE') AS AOp " &
                                       "ON ASp.IDENTAKCJI=AOp.IDENT " &
                                       " "
            ElseIf rbtObsluzeni.Checked Then
                dbSelectSQLAkcjeSpec = "SELECT DISTINCT " &
                                              "IDENTKLI " &
                                       "FROM AkcjeSpec AS ASp INNER JOIN (SELECT * FROM AkcjeOpis WHERE MARKER='TRUE') AS AOp " &
                                       "ON ASp.IDENTAKCJI=AOp.IDENT " &
                    " WHERE ((SELECT count(*) FROM KontaktyHandlowe WHERE (DATAKONTAKTU>='" & OdDaty & "') AND " &
                                                                "(DATAKONTAKTU<='" & DoDaty & "') AND " &
                                                                "(KontaktyHandlowe.IDENTKLIENTA=Asp.IDENTKLI)) > 0) "
            ElseIf rbtnieObsluzeni.Checked Then
                dbSelectSQLAkcjeSpec = "SELECT DISTINCT " &
                                              "IDENTKLI " &
                                       "FROM AkcjeSpec AS ASp INNER JOIN (SELECT * FROM AkcjeOpis WHERE MARKER='TRUE') AS AOp " &
                                       "ON ASp.IDENTAKCJI=AOp.IDENT " &
                    " WHERE ((SELECT count(*) FROM KontaktyHandlowe WHERE (DATAKONTAKTU>='" & OdDaty & "') AND " &
                                                                "(DATAKONTAKTU<='" & DoDaty & "') AND " &
                                                                "(KontaktyHandlowe.IDENTKLIENTA=Asp.IDENTKLI)) = 0) "
            Else
                'def branze
                dbSelectSQLAkcjeSpec = "SELECT DISTINCT " &
                                               "IDENTKLI " &
                                       "FROM AkcjeSpec AS ASp INNER JOIN (SELECT * FROM AkcjeOpis WHERE MARKER='TRUE') AS AOp " &
                                       "ON ASp.IDENTAKCJI=AOp.IDENT " &
                                       " "
            End If








            If rbtWybierzWszystkich.Checked Then
                'kwerenda OK

            ElseIf rbtWybierzWspolnych.Checked Then
                'analizujemy oznaczenie akcji
                Dim IleOznaczono As Integer = 0
                Dim KluczAnalizy As String = ""
                Dim drv As DataRowView = Nothing
                If DataViewAkcjeOpis.Count > 0 Then
                    For Each drv In DataViewAkcjeOpis
                        If drv("MARKER") = True Then
                            IleOznaczono += 1
                            If Len(KluczAnalizy) = 0 Then
                                KluczAnalizy = " (AkcjeSpec.IDENTAKCJI='" & Trim(drv("IDENT")) & "') "
                            Else
                                KluczAnalizy &= " OR (AkcjeSpec.IDENTAKCJI='" & Trim(drv("IDENT")) & "') "
                            End If
                        End If
                    Next
                End If

                If Len(KluczAnalizy) = 0 Then
                    'cos dziwnego - nie ma co wybierac
                Else
                    If rbtWszyscy.Checked Then
                        dbSelectSQLAkcjeSpec &= " WHERE "
                    Else
                        dbSelectSQLAkcjeSpec &= " AND "
                    End If
                    dbSelectSQLAkcjeSpec &=
                    " ((SELECT count(*) FROM AkcjeSpec WHERE (AkcjeSpec.IDENTKLI=ASP.IDENTKLI) AND (" & KluczAnalizy & " )) = " & Trim(Str(IleOznaczono)) & ") "
                End If

            Else
            End If


            If ParL_DataBaseType = "MS ACCESS" Then
                'OK - mo¿na zapisaæ....
                dbConnectionAkcjeSpec = New OleDb.OleDbConnection(sConnectionAkcjeSpec)
                dbCommandSelectAkcjeSpec = New OleDb.OleDbCommand(dbSelectSQLAkcjeSpec, dbConnectionAkcjeSpec)
                'dbCommandDeleteAkcjeSpec = New OleDb.OleDbCommand(sDeleteSQLAkcjeSpec, dbConnectionAkcjeSpec)
                'dbCommandUpdateAkcjeSpec = New OleDb.OleDbCommand(sUpdateSQLAkcjeSpec, dbConnectionAkcjeSpec)
                'dbCommandInsertAkcjeSpec = New OleDb.OleDbCommand(sInsertSQLAkcjeSpec, dbConnectionAkcjeSpec)
                dbDataAdapterAkcjeSpec = New OleDb.OleDbDataAdapter(dbCommandSelectAkcjeSpec)

                Dim dbMapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
                dbMapowanieTabeliAkcjeSpec = dbDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_dbAkcjeSpec")

                Try
                    dbConnectionAkcjeSpec.Open()
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                Finally
                    ConnectionState = dbConnectionAkcjeSpec.State
                End Try
            Else
                'OK - mo¿na zapisaæ....
                sqlConnectionAkcjeSpec = New SqlClient.SqlConnection(sConnectionAkcjeSpec)
                sqlCommandSelectAkcjeSpec = New SqlClient.SqlCommand(dbSelectSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                'sqlCommandDeleteAkcjeSpec = New SqlClient.SqlCommand(sDeleteSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                'sqlCommandUpdateAkcjeSpec = New SqlClient.SqlCommand(sUpdateSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                'sqlCommandInsertAkcjeSpec = New SqlClient.SqlCommand(sInsertSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                sqlDataAdapterAkcjeSpec = New SqlClient.SqlDataAdapter(sqlCommandSelectAkcjeSpec)

                Dim sqlMapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
                sqlMapowanieTabeliAkcjeSpec = sqlDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_sqlAkcjeSpec")

                Try
                    sqlConnectionAkcjeSpec.Open()
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                Finally
                    ConnectionState = sqlConnectionAkcjeSpec.State
                End Try
            End If

            If ConnectionState = ConnectionState.Open Then
                Try
                    If ParL_DataBaseType = "MS ACCESS" Then
                        dbDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_dbAkcjeSpec")
                        dbDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                        dbConnectionAkcjeSpec.Close()
                    Else
                        sqlDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_sqlAkcjeSpec")
                        sqlDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                        sqlConnectionAkcjeSpec.Close()
                    End If

                    'zdefiniuj unikalny klucz wyszukiwania w tabeli
                    DataSetAkcjeSpec.Tables(0).PrimaryKey = New DataColumn() {DataSetAkcjeSpec.Tables(0).Columns("IDENTKLI")}

                    'definiuj DataView
                    DataViewAkcjeSpec = New DataView(DataSetAkcjeSpec.Tables(0))
                    DataViewAkcjeSpec.AllowDelete = False
                    DataViewAkcjeSpec.AllowEdit = False
                    DataViewAkcjeSpec.AllowNew = False

                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                End Try
            End If

            If DataViewAkcjeSpec.Count > 0 Then
                If rbtWszyscy.Checked Or rbtObsluzeni.Checked Or rbtNieObsluzeni.Checked Then
                    '...........................
                    'definiuj delegat
                    Dim deleg As delegateAktualizuj = _AktualizujKartKlientow
                    Dim FormRaport As New RaportWybierzAkcjeMarketingowa("",
                                                                         _dsKlenci,
                                                                         DataViewAkcjeSpec,
                                                                         deleg,
                                                                         parInicjuj)
                    FormRaport.ShowDialog()
                    oWybranoAkcjeMarketingowa = True
                    '...........................
                    MessageBox.Show("Pobra³em informacje z wybranych akcji marketingowych" & vbCrLf &
                                    "Wybra³em " & Trim(Str(DataViewAkcjeSpec.Count)) & " Klientów.",
                         "OK - skoñczy³em",
                              System.Windows.Forms.MessageBoxButtons.OK,
                             MessageBoxIcon.Information,
                             MessageBoxDefaultButton.Button1)


                Else
                    'def branze
                    'definiuj delegat
                    Dim deleg As delegateAktualizuj = _AktualizujKartKlientow
                    Dim FormRaport As New RaportWybierzAkcjeMarketingowaIDefBranze("",
                                                                             _dsKlenci,
                                                                             DataViewAkcjeSpec,
                                                                             deleg,
                                                                             txtBranza.Text,
                                                                             txtPodbranza.Text)
                    FormRaport.ShowDialog()
                    oWybranoAkcjeMarketingowa = False
                    '...........................
                    MessageBox.Show("Zdefiniowa³em Bran¿e na podstawie akcji marketingowych" & vbCrLf &
                                    "dla " & Trim(Str(DataViewAkcjeSpec.Count)) & " Klientów.",
                                            "OK - skoñczy³em",
                                            System.Windows.Forms.MessageBoxButtons.OK,
                                            MessageBoxIcon.Information,
                                            MessageBoxDefaultButton.Button1)
                End If
            Else
                MessageBox.Show("Nie znalaz³em klientów spe³niaj¹cych zadane warunki.",
                     "OK - skoñczy³em",
                          System.Windows.Forms.MessageBoxButtons.OK,
                         MessageBoxIcon.Information,
                         MessageBoxDefaultButton.Button1)
            End If
        End If
    End Sub










    Private Sub rbtWszyscy_CheckedChanged(sender As Object, e As EventArgs) Handles rbtWszyscy.CheckedChanged, rbtObsluzeni.CheckedChanged, rbtNieObsluzeni.CheckedChanged, rbtDefiniujBranze.CheckedChanged
        If rbtWszyscy.Checked Then
            lblOdDaty.Visible = False
            dtpOdDaty.Visible = False
            lblDoDaty.Visible = False
            dtpDoDaty.Visible = False

            LabelBranza.Visible = False
            txtBranza.Visible = False
            cmdWybierzBranze.Visible = False
            LabelPodBranza.Visible = False
            txtPodbranza.Visible = False
            cmdWybierzPodbranze.Visible = False

            cmdZapamietaj.Text = "Pobierz akcje"
            cmdDodaj.Visible = True

        ElseIf rbtObsluzeni.Checked Or rbtNieObsluzeni.Checked Then
            lblOdDaty.Visible = True
            dtpOdDaty.Visible = True
            lblDoDaty.Visible = True
            dtpDoDaty.Visible = True

            LabelBranza.Visible = False
            txtBranza.Visible = False
            cmdWybierzBranze.Visible = False
            LabelPodBranza.Visible = False
            txtPodbranza.Visible = False
            cmdWybierzPodbranze.Visible = False

            cmdZapamietaj.Text = "Pobierz akcje"
            cmdDodaj.Visible = True

        Else        'def branze
            lblOdDaty.Visible = False
            dtpOdDaty.Visible = False
            lblDoDaty.Visible = False
            dtpDoDaty.Visible = False

            LabelBranza.Visible = True
            txtBranza.Visible = True
            cmdWybierzBranze.Visible = True
            LabelPodBranza.Visible = True
            txtPodbranza.Visible = True
            cmdWybierzPodbranze.Visible = True

            cmdZapamietaj.Text = "Pobierz bran¿e"
            cmdDodaj.Visible = False
        End If
    End Sub

    'Private Sub rbtObsluzeni_CheckedChanged(sender As Object, e As EventArgs) Handles rbtObsluzeni.CheckedChanged
    '    If rbtWszyscy.Checked Then
    '        lblOdDaty.Visible = False
    '        dtpOdDaty.Visible = False
    '        lblDoDaty.Visible = False
    '        dtpDoDaty.Visible = False
    '    Else
    '        lblOdDaty.Visible = True
    '        dtpOdDaty.Visible = True
    '        lblDoDaty.Visible = True
    '        dtpDoDaty.Visible = True
    '    End If
    'End Sub

    'Private Sub rbtNieObsluzeni_CheckedChanged(sender As Object, e As EventArgs) Handles rbtNieObsluzeni.CheckedChanged
    '    If rbtWszyscy.Checked Then
    '        lblOdDaty.Visible = False
    '        dtpOdDaty.Visible = False
    '        lblDoDaty.Visible = False
    '        dtpDoDaty.Visible = False
    '    Else
    '        lblOdDaty.Visible = True
    '        dtpOdDaty.Visible = True
    '        lblDoDaty.Visible = True
    '        dtpDoDaty.Visible = True
    '    End If
    'End Sub
















    Private Sub txtBranza_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBranza.TextChanged
        TestLen(txtBranza, "BRAN¯A", 100)
    End Sub
    Private Sub txtBranza_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBranza.GotFocus
        StartEdycjiTxt(txtBranza)
    End Sub
    Private Sub txtBranza_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBranza.LostFocus
        KoniecEdycjiTxt(txtBranza)
    End Sub


    Private Sub txtPodBranza_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPodbranza.TextChanged
        TestLen(txtPodbranza, "PODBRAN¯A", 100)
    End Sub
    Private Sub txtPodBranza_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPodbranza.GotFocus
        StartEdycjiTxt(txtPodbranza)
    End Sub
    Private Sub txtPodBranza_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPodbranza.LostFocus
        KoniecEdycjiTxt(txtPodbranza)
    End Sub

    Private Sub cmdWybierzBranze_Click(sender As Object, e As EventArgs) Handles cmdWybierzBranze.Click
        oCoMamRobic = "WYBIERAJ"
        oIdBranzy = Trim(txtBranza.Text)
        Dim form As New KatalogBranze
        form.ShowDialog()
        If Len(Trim(oIdBranzy)) > 0 Then
            txtBranza.Text = oIdBranzy
        End If
    End Sub

    Private Sub cmdWybierzPodBranze_Click(sender As Object, e As EventArgs) Handles cmdWybierzPodbranze.Click
        oCoMamRobic = "WYBIERAJ"
        oIdBranzy = Trim(txtBranza.Text)
        oIdPodBranzy = Trim(txtPodbranza.Text)
        Dim form As New KatalogPodBranze
        form.ShowDialog()
        If Len(Trim(oIdBranzy)) > 0 Then
            txtBranza.Text = oIdBranzy
        End If
        If Len(Trim(oIdPodBranzy)) > 0 Then
            txtPodbranza.Text = oIdPodBranzy
        End If
    End Sub

End Class

