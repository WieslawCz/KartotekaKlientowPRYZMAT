Imports System.Drawing

Public Class _frmFiltrowanieWybierzWartoscPola

    Dim pointInCell00 As System.Drawing.Point
    Dim StartFormy As Boolean = True

    Dim dbSelectWartosc As String = ""

    Dim dbfConnection As Odbc.OdbcConnection
    Dim dbfCommandSelect As Odbc.OdbcCommand
    Dim dbfDataAdapter As Odbc.OdbcDataAdapter

    'Dim dbConnection As OleDb.OleDbConnection
    'Dim dbCommandSelect As OleDb.OleDbCommand
    'Dim dbDataAdapter As OleDb.OleDbDataAdapter

    Dim sqlConnection As SqlClient.SqlConnection
    Dim sqlCommandSelect As SqlClient.SqlCommand
    Dim sqlDataAdapter As SqlClient.SqlDataAdapter

    Dim DataSetWartosc As New DataSet
    Dim DataViewWartosc As New DataView
    Dim ConnectionState As System.Data.ConnectionState


    Private Sub WybierzWartoscPola_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = "Wartoœci w polu " & _NazwaPola
        'Me.Icon = My.Resources.MojaFirma_N
        Me.picLogo.Image = My.Resources.logomini
        'Me.cmdWybierz.Image = My.Resources._EXIT.ToBitmap
        '----------------------------------------
        If _DataView Is Nothing Then
            DataSetWartosc.Locale = New System.Globalization.CultureInfo("pl-PL")
            If Len(Trim(_Filtr)) = 0 Then
                dbSelectWartosc = "SELECT DISTINCT " & _NazwaPola & _
                                  " FROM " & _NazwaTabeli & _
                                  " ORDER BY " & _NazwaPola
            Else
                If ParL_DataBaseType = "MS ACCESS" Then
                    ''jesli  to ACCESS - zamien w filtrze * na %
                    'Dim pos As Integer
                    'pos = InStr(_Filtr, "*")
                    'Do While pos > 0
                    '    If pos = 1 Then
                    '        _Filtr = "%" & Mid(_Filtr, pos + 1)
                    '    Else
                    '        _Filtr = Mid(_Filtr, 1, pos - 1) & "%" & Mid(_Filtr, pos + 1)
                    '    End If
                    '    pos = InStr(_Filtr, "*")
                    'Loop
                Else
                End If
                '----------------------------------
                'dbSelectWartosc = "SELECT DISTINCT " & _NazwaPola & _
                '                  " FROM " & _NazwaTabeli & _
                '                  " WHERE " & _Filtr & _
                '                  " ORDER BY " & _NazwaPola

                dbSelectWartosc = "SELECT DISTINCT " & _NazwaPola & _
                                  " FROM " & _NazwaTabeli & _
                                  " ORDER BY " & _NazwaPola
            End If



            If Len(_ConnString) > 0 Then
                'baza DBF
                dbfConnection = New Odbc.OdbcConnection(_ConnString)
                dbfCommandSelect = New Odbc.OdbcCommand(dbSelectWartosc, dbfConnection)
                dbfDataAdapter = New Odbc.OdbcDataAdapter(dbfCommandSelect)

                'mapujemy domyslna nazwe tabeli
                Dim DBfMapowanieTabeli As System.Data.Common.DataTableMapping
                DBfMapowanieTabeli = dbfDataAdapter.TableMappings.Add("Table", "TABELA_dbfWartosc")
                '------------------------------------------
                'wypelnij DATASET
                Try
                    dbfConnection.Open()
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                Finally
                    ConnectionState = dbfConnection.State
                End Try
            Else
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbConnection = New OleDb.OleDbConnection(DataBaseConnection())
                    'dbCommandSelect = New OleDb.OleDbCommand(dbSelectWartosc, dbConnection)
                    'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

                    ''mapujemy domyslna nazwe tabeli
                    'Dim DBMapowanieTabeli As System.Data.Common.DataTableMapping
                    'DBMapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_dbWartosc")
                    ''------------------------------------------
                    ''wypelnij DATASET
                    'Try
                    '    dbConnection.Open()
                    'Catch Ex As System.Exception
                    '    ViewInLog(Ex, Me.Name(), Nothing)
                    '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    '    MessageBoxIcon.Information, _
                    '    '    MessageBoxDefaultButton.Button1)
                    'Finally
                    '    ConnectionState = dbConnection.State
                    'End Try
                Else
                    sqlConnection = New SqlClient.SqlConnection(DataBaseConnection())
                    sqlCommandSelect = New SqlClient.SqlCommand(dbSelectWartosc, sqlConnection)
                    sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

                    'mapujemy domyslna nazwe tabeli
                    Dim sqlMapowanieTabeli As System.Data.Common.DataTableMapping
                    sqlMapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_sqlWartosc")
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
            End If

            If ConnectionState = ConnectionState.Open Then
                Try
                    If Len(_ConnString) > 0 Then
                        dbfDataAdapter.FillSchema(DataSetWartosc, SchemaType.Mapped, "TABELA_dbfWartosc")
                        dbfDataAdapter.Fill(DataSetWartosc)
                        dbfConnection.Close()
                    Else
                        If ParL_DataBaseType = "MS ACCESS" Then
                            'dbDataAdapter.FillSchema(DataSetWartosc, SchemaType.Mapped, "TABELA_dbWartosc")
                            'dbDataAdapter.Fill(DataSetWartosc)
                            'dbConnection.Close()
                        Else
                            sqlDataAdapter.FillSchema(DataSetWartosc, SchemaType.Mapped, "TABELA_sqlWartosc")
                            sqlDataAdapter.Fill(DataSetWartosc)
                            sqlConnection.Close()
                        End If
                    End If

                    'zdefiniuj unikalny klucz wyszukiwania w tabeli
                    'DataSetWartosc.Tables(0).PrimaryKey = New DataColumn() {DataSetWartosc.Tables(0).Columns("NRKARTY")}

                    'definiuj DataView
                    DataViewWartosc = New DataView(DataSetWartosc.Tables(0))
                    DataViewWartosc.AllowDelete = False
                    DataViewWartosc.AllowEdit = False
                    DataViewWartosc.AllowNew = False

                    '-------------------------------------------------
                    'parametry DataGridView
                    '-------------------------------------------------
                    ' Initialize basic DataGridView properties.
                    dagWartosc.Dock = DockStyle.Fill
                    dagWartosc.BackgroundColor = System.Drawing.SystemColors.Control
                    dagWartosc.GridColor = System.Drawing.SystemColors.ControlDark
                    dagWartosc.ForeColor = System.Drawing.Color.Purple
                    dagWartosc.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize, _
                                                                 System.Drawing.FontStyle.Regular, _
                                                                 System.Drawing.GraphicsUnit.Point, _
                                                                 CType(238, Byte))
                    dagWartosc.BorderStyle = BorderStyle.Fixed3D
                    dagWartosc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                            Or System.Windows.Forms.AnchorStyles.Bottom) _
                                            Or System.Windows.Forms.AnchorStyles.Left) _
                                            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                    ' Set property values appropriate for read-only display and limited interactivity. 
                    dagWartosc.AllowUserToAddRows = False
                    dagWartosc.AllowUserToDeleteRows = False
                    dagWartosc.AllowUserToOrderColumns = True
                    dagWartosc.AllowUserToResizeColumns = True
                    dagWartosc.AllowUserToResizeRows = True
                    dagWartosc.ReadOnly = True
                    dagWartosc.SelectionMode = DataGridViewSelectionMode.CellSelect
                    dagWartosc.MultiSelect = False

                    'UWAGA: Ustawienie AutoSizeMode na AllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                    'dagWartosc.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                    dagWartosc.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None     'DisplayedCells
                    dagWartosc.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.None

                    dagWartosc.ColumnHeadersVisible = True
                    dagWartosc.ColumnHeadersHeight = 40
                    dagWartosc.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                    dagWartosc.RowHeadersVisible = True
                    dagWartosc.RowHeadersWidth = 50
                    dagWartosc.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing


                    ' Set the selection background color for all the cells.
                    dagWartosc.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                    dagWartosc.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                    dagWartosc.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                    dagWartosc.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                    dagWartosc.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize, _
                                                                 System.Drawing.FontStyle.Regular, _
                                                                 System.Drawing.GraphicsUnit.Point, _
                                                                 CType(238, Byte))
                    dagWartosc.DefaultCellStyle.NullValue = ""
                    dagWartosc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

                    ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                    ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                    dagWartosc.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                    dagWartosc.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                    ' Set the background color for all rows and for alternating rows.  
                    ' The value for alternating rows overrides the value for all rows. 
                    dagWartosc.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                    dagWartosc.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                    ' Set the row and column header styles.
                    dagWartosc.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                    dagWartosc.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                    dagWartosc.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                    dagWartosc.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                    '-----------------------------------
                    'wypelnienie DataGridView
                    dagWartosc.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                    dagWartosc.DataMember = DataSetWartosc.Tables(0).TableName
                    dagWartosc.DataSource = DataViewWartosc
                    '-----------------------------------

                    ' Add a GridColumnStyle and set the MappingName 
                    ' to the name of a DataColumn in the DataTable. 
                    ' Set the HeaderText and Width properties. 
                    TxtColumnView(dagWartosc, DataSetWartosc.Tables(0).Columns(0).ColumnName, _NazwaPola, dagWartosc.Width - 50, DataGridViewContentAlignment.MiddleLeft, True)

                    'ustaw sie na pierwszym zapisie
                    If DataSetWartosc.Tables(0).Rows.Count > 0 Then dagWartosc.Select()

                Catch Ex As System.Exception
                    'ViewInLog(ex, Me.Name(), Nothing)
                    MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :", _
                        System.Windows.Forms.MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, _
                        MessageBoxDefaultButton.Button1)
                End Try
            End If




        Else
            ' mamy dany DATAVIEW z danymi
            'Dim distinctTable As DataTable = _DataView.Table.DefaultView.ToTable(True, _NazwaPola)
            'DataTable distinctTable = originalTable.DefaultView.ToTable( /*distinct*/ true);

            DataViewWartosc = New DataView(_DataView.Table.DefaultView.ToTable(True, _NazwaPola))

            DataViewWartosc.AllowDelete = False
            DataViewWartosc.AllowEdit = False
            DataViewWartosc.AllowNew = False

            '-------------------------------------------------
            'parametry DataGridView
            '-------------------------------------------------
            ' Initialize basic DataGridView properties.
            dagWartosc.Dock = DockStyle.Fill
            dagWartosc.BackgroundColor = System.Drawing.SystemColors.Control
            dagWartosc.GridColor = System.Drawing.SystemColors.ControlDark
            dagWartosc.ForeColor = System.Drawing.Color.Purple
            dagWartosc.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize, _
                                                         System.Drawing.FontStyle.Regular, _
                                                         System.Drawing.GraphicsUnit.Point, _
                                                         CType(238, Byte))
            dagWartosc.BorderStyle = BorderStyle.Fixed3D
            dagWartosc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                    Or System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            ' Set property values appropriate for read-only display and limited interactivity. 
            dagWartosc.AllowUserToAddRows = False
            dagWartosc.AllowUserToDeleteRows = False
            dagWartosc.AllowUserToOrderColumns = True
            dagWartosc.AllowUserToResizeColumns = True
            dagWartosc.AllowUserToResizeRows = True
            dagWartosc.ReadOnly = True
            dagWartosc.SelectionMode = DataGridViewSelectionMode.CellSelect
            dagWartosc.MultiSelect = False

            'UWAGA: Ustawienie AutoSizeMode na AllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
            'dagWartosc.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            dagWartosc.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None     'DisplayedCells
            dagWartosc.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.None

            dagWartosc.ColumnHeadersVisible = True
            dagWartosc.ColumnHeadersHeight = 40
            dagWartosc.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
            dagWartosc.RowHeadersVisible = True
            dagWartosc.RowHeadersWidth = 50
            dagWartosc.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing


            ' Set the selection background color for all the cells.
            dagWartosc.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
            dagWartosc.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
            dagWartosc.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
            dagWartosc.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
            dagWartosc.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize, _
                                                         System.Drawing.FontStyle.Regular, _
                                                         System.Drawing.GraphicsUnit.Point, _
                                                         CType(238, Byte))
            dagWartosc.DefaultCellStyle.NullValue = ""
            dagWartosc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
            ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
            dagWartosc.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
            dagWartosc.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

            ' Set the background color for all rows and for alternating rows.  
            ' The value for alternating rows overrides the value for all rows. 
            dagWartosc.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
            dagWartosc.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

            ' Set the row and column header styles.
            dagWartosc.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
            dagWartosc.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
            dagWartosc.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
            dagWartosc.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
            '-----------------------------------
            'wypelnienie DataGridView
            dagWartosc.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
            dagWartosc.DataMember = DataViewWartosc.Table.TableName
            'dagWartosc.DataSource = DataViewWartosc
            '-----------------------------------


            ' Add a GridColumnStyle and set the MappingName 
            ' to the name of a DataColumn in the DataTable. 
            ' Set the HeaderText and Width properties. 
            TxtColumnView(dagWartosc, DataViewWartosc.Table.Columns(0).ColumnName, _NazwaPola, dagWartosc.Width - 50, DataGridViewContentAlignment.MiddleLeft, True)

            Me.Cursor = Cursors.WaitCursor
            dagWartosc.DataSource = DataViewWartosc
            Me.Cursor = Cursors.Default

            'ustaw sie na pierwszym zapisie
            If DataViewWartosc.Table.Rows.Count > 0 Then dagWartosc.Select()
        End If

        ''--------------------------------------------------------------------
        ' ''definiuj przerwanie przy malowaniau DataGrid
        ''AddHandler Me.dagWartosc.Paint, New System.Windows.Forms.PaintEventHandler(AddressOf dagWartosc_Paint)

        ''set the header width to something apporpriate
        'dagWartosc.RowHeadersWidth = 50       'use if tablestyle

        ''set topleft corner point so we can locate the toprow
        'If DataSetWartosc.Tables(0).Rows.Count > 0 Then
        '    pointInCell00 = New system.drawing.Point((dagWartosc.GetCellDisplayRectangle(0, 0, True).X + 4), _
        '                              (dagWartosc.GetCellDisplayRectangle(0, 0, True).Y + 4))
        'Else
        '    pointInCell00 = New system.drawing.Point((dagWartosc.Left + 4), (dagWartosc.Top + 4))
        'End If
        ''--------------------------------------------------------------------
    End Sub

    Private Sub WybierzWartoscPola_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        RysujGradient(Me)
    End Sub

    Private Sub WybierzWartoscPola_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
        dagWartosc.Columns(0).Width = dagWartosc.Width - 50
    End Sub




    Private Sub dagWartosc_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagWartosc.DoubleClick
        cmdWybierz_Click(sender, e)
    End Sub

    Private Sub dagWartosc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagWartosc.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                cmdWybierz_Click(sender, e)
            Case Keys.Insert, Keys.Add
            Case Keys.Delete, Keys.Subtract
            Case Else
        End Select
    End Sub





    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagWartosc_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagWartosc.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagWartosc.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture), _
                      dagWartosc.DefaultCellStyle.Font, _
                      b, _
                      e.RowBounds.Location.X + 15, _
                      e.RowBounds.Location.Y + 4)
        End Using
    End Sub






    Private Sub cmdWybierz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWybierz.Click
        If DataViewWartosc.Table.Rows.Count > 0 Then
            oWybralem = GetTxtField(dagWartosc, 0)
        Else
            oWybralem = ""        'IDENT       Text, 10
        End If
        Me.Close()
    End Sub

End Class