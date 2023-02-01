Public Class FiltrowanieWybierzWartoscPola

    Dim RowHeightSetter As Softart.myDataGridRowHeightSetter
    Dim pointInCell00 As Point
    Dim StartFormy As Boolean = True

    Dim dbSelectWartosc As String = ""

    Dim dbConnection As OleDb.OleDbConnection
    Dim dbCommandSelect As OleDb.OleDbCommand
    Dim dbDataAdapter As OleDb.OleDbDataAdapter

    Dim sqlConnection As SqlClient.SqlConnection
    Dim sqlCommandSelect As SqlClient.SqlCommand
    Dim sqlDataAdapter As SqlClient.SqlDataAdapter

    Dim DataSetWartosc As New DataSet
    Dim DataViewWartosc As New DataView
    Dim ConnectionState As System.Data.ConnectionState


    Private Sub WybierzWartoscPola_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        Me.Text = "Wartoœci w polu " & _NazwaPola
        '----------------------------------------
        DataSetWartosc.Locale = New System.Globalization.CultureInfo("pl-PL")
        If Len(Trim(_Filtr)) = 0 Then
            dbSelectWartosc = "SELECT DISTINCT " & _NazwaPola & _
                              " FROM " & _NazwaTabeli & _
                              " ORDER BY " & _NazwaPola
        Else
            'jesli  to ACCESS - zamien w filtrze * na %
            Dim pos As Integer
            pos = InStr(_Filtr, "*")
            Do While pos > 0
                If pos = 1 Then
                    _Filtr = "%" & Mid(_Filtr, pos + 1)
                Else
                    _Filtr = Mid(_Filtr, 1, pos - 1) & "%" & Mid(_Filtr, pos + 1)
                End If
                pos = InStr(_Filtr, "*")
            Loop
            '----------------------------------
            dbSelectWartosc = "SELECT DISTINCT " & _NazwaPola & _
                              " FROM " & _NazwaTabeli & _
                              " WHERE " & _Filtr & _
                              " ORDER BY " & _NazwaPola
        End If

        If ParL_DataBaseType = "MS ACCESS" Then

            dbConnection = New OleDb.OleDbConnection(sConnectionKlienci)
            dbCommandSelect = New OleDb.OleDbCommand(dbSelectWartosc, dbConnection)
            dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim DBMapowanieTabeli As System.Data.Common.DataTableMapping
            DBMapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_dbWartosc")
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

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    dbDataAdapter.FillSchema(DataSetWartosc, SchemaType.Mapped, "TABELA_dbWartosc")
                    dbDataAdapter.Fill(DataSetWartosc)
                    dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(DataSetWartosc, SchemaType.Mapped, "TABELA_sqlWartosc")
                    sqlDataAdapter.Fill(DataSetWartosc)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                'DataSetWartosc.Tables(0).PrimaryKey = New DataColumn() {DataSetWartosc.Tables(0).Columns("NRKARTY")}

                'parametry DataGrid
                dagWartosc.CaptionVisible = False
                dagWartosc.CaptionText = DataSetWartosc.Tables(0).TableName
                dagWartosc.ColumnHeadersVisible = True
                dagWartosc.RowHeadersVisible = True
                dagWartosc.BackgroundColor = System.Drawing.SystemColors.Control
                dagWartosc.BorderStyle = BorderStyle.Fixed3D
                dagWartosc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                dagWartosc.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
                dagWartosc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))

                'definiuj DataView
                DataViewWartosc = New DataView(DataSetWartosc.Tables(0))
                DataViewWartosc.AllowDelete = False
                DataViewWartosc.AllowEdit = False
                DataViewWartosc.AllowNew = False
                'wypelnienie DataGrid
                'dagwartosc.SetDataBinding(DataSetWartosc, "TABELA_Wartosc")
                dagWartosc.DataSource = DataViewWartosc

                ' Add a GridTableStyle and set the MappingName to the name of the DataTable.
                Dim TSWartosc As New DataGridTableStyle
                TSWartosc.MappingName = DataSetWartosc.Tables(0).TableName
                TSWartosc.AlternatingBackColor = System.Drawing.SystemColors.Control
                TSWartosc.BackColor = System.Drawing.SystemColors.ControlLight
                TSWartosc.GridLineColor = System.Drawing.SystemColors.ControlDark
                TSWartosc.ForeColor = System.Drawing.Color.Purple
                TSWartosc.HeaderBackColor = System.Drawing.SystemColors.ScrollBar
                TSWartosc.HeaderForeColor = System.Drawing.Color.Navy
                TSWartosc.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                TSWartosc.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow

                ' Add a GridColumnStyle and set the MappingName 
                ' to the name of a DataColumn in the DataTable. 
                ' Set the HeaderText and Width properties. 
                TxtColumn(TSWartosc, DataSetWartosc.Tables(0).Columns(0).ColumnName, _NazwaPola, dagWartosc.Width - 50, HorizontalAlignment.Left)

                ' Add the DataGridTableStyle instance to the GridTableStylesCollection. 
                dagWartosc.TableStyles.Add(TSWartosc)
                'ustaw sie na pierwszym zapisie
                If DataSetWartosc.Tables(0).Rows.Count > 0 Then dagWartosc.Select(0)

            Catch Ex As System.Exception
                'ViewInLog(ex, Me.Name(), Nothing)
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            End Try
        End If
    End Sub

    Private Sub WybierzWartoscPola_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        RysujGradient(Me)
    End Sub

    Private Sub WybierzWartoscPola_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
        dagWartosc.TableStyles(0).GridColumnStyles(0).Width = dagWartosc.Width - 50
    End Sub




    Private Sub dagWartosc_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagWartosc.DoubleClick
        cmdWybierz_Click(sender, e)
    End Sub

    Private Sub dagWartosc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagWartosc.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                cmdWybierz_Click(sender, e)
            Case Keys.Insert
            Case Keys.Delete
            Case Else
        End Select
    End Sub

    'wyswietla numer rekordu w RowHead
    Private Sub dagWartosc_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles dagWartosc.Paint
        Dim hti As Softart.MyDataGrid.HitTestInfo
        Dim row As Integer
        Dim yDelta As Integer
        Dim y As Integer

        Try
            If DataSetWartosc.Tables(0).Rows.Count > 0 Then
                hti = dagWartosc.HitTest(Me.pointInCell00)
                row = hti.Row
                yDelta = (dagWartosc.GetCellBounds(row, 0).Height + 1)
                y = (dagWartosc.GetCellBounds(row, 0).Top + 2)

                Dim cm As CurrencyManager
                Dim RowText As String
                cm = CType(Me.BindingContext(dagWartosc.DataSource, dagWartosc.DataMember), CurrencyManager)
                Do While ((y < (dagWartosc.Height - yDelta)) AndAlso (row < cm.Count))
                    RowText = (row + 1).ToString("N0")
                    e.Graphics.DrawString(RowText, dagWartosc.Font, New SolidBrush(Color.Navy), 12, y)
                    yDelta = (dagWartosc.GetCellBounds(row, 0).Height + 1)
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


    Private Sub cmdWybierz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWybierz.Click
        If DataSetWartosc.Tables(0).Rows.Count > 0 Then
            oWybralem = IIf(IsDBNull(dagWartosc.Item(dagWartosc.CurrentCell.RowNumber, 0)), "", dagWartosc.Item(dagWartosc.CurrentCell.RowNumber, 0))        'IDENT       Text    10
        Else
            oWybralem = ""        'IDENT       Text, 10
        End If
        Me.Close()
    End Sub

End Class