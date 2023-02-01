Public Class RaportWybierzAkcjeMarketingowa

    Private Sub RaportWykonania_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        lblFunkcja.Text = ""
        lblOperacja.Text = ""
        Progres.Maximum = 0
        Progres.Value = 0

        If _Ident Is Nothing Then
            lblFunkcja.Text = "Pobieram klient�w na podstawie wielko�ci sprzeda�y."
        ElseIf Len(_Ident) = 0 Then
            lblFunkcja.Text = "Pobieram informacje z wybranych akcji marketingowych."
        Else
            lblFunkcja.Text = "Pobieram informacje o akcji marketingowej " & _Ident & "."
        End If

        Me.Opacity = 0.8

        'Timer1.Interval = 200
        'Timer1.Enabled = True

        Timer2.Interval = 200
        Timer2.Enabled = True

    End Sub

    Private Sub RaportWykonania_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        RysujGradient(Me)
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False

        Dim dvKlienci As DataView = New DataView(_dsKlienci.Tables(0))
        Dim dsRow As DataRow = Nothing
        Dim dvRow As DataRowView = Nothing
        Dim SzukajKlienta(0) As Object
        Dim Klient As String = ""
        Dim WartSP As Double = 0
        Dim j As Long = 0
        Dim i As Long = 0

        If _InicjujKartKlientow Then
            'usun wszystkie markery
            lblOperacja.Text = "Inicjuje pole MARKER"
            Progres.Maximum = _dsKlienci.Tables(0).Rows.Count - 1
            For i = 0 To _dsKlienci.Tables(0).Rows.Count - 1
                Progres.Value = i

                dsRow = _dsKlienci.Tables(0).Rows.Item(i)
                If IsDBNull(dsRow("MARKER")) Then
                    dsRow("MARKER") = False
                    'dsRow("MARKERP") = False
                    'dvRow("WARTSPRZED") = 0

                    If _AktualizujKartKlientow Is Nothing Then
                    Else
                        _AktualizujKartKlientow()
                    End If
                Else
                    If dsRow("MARKER") Then
                        dsRow("MARKER") = False
                        dsRow("MARKERP") = False
                        'dvRow("WARTSPRZED") = 0

                        If _AktualizujKartKlientow Is Nothing Then
                        Else
                            _AktualizujKartKlientow()
                        End If
                    End If
                End If
                'System.Windows.Forms.Application.DoEvents()
            Next
        End If

        'USTAW MARKERY KLIENTOW Z TEJ AKCJI
        lblOperacja.Text = "Ustawiam pole MARKER dla wybranych klient�w."
        Progres.Maximum = _dvAkcjeSpec.Count - 1
        For i = 0 To _dvAkcjeSpec.Count - 1
            Progres.Value = i

            dvRow = _dvAkcjeSpec.Item(i)
            Klient = dvRow("IDENTKLI")
            'If _Ident Is Nothing Then
            '    WartSP = dvRow("WARTSPRZED")
            'Else
            '    WartSP = 0
            'End If
            '-----------------------------------------
            If _AktualizujKartKlientow Is Nothing Then
                dvKlienci.RowFilter = "IDENTKLIENTA='" & Klient & "'"
                If dvKlienci.Count() > 0 Then
                    For j = 1 To dvKlienci.Count
                        If dvKlienci.Item(j - 1).Item("AKTYWNY") = True Then
                            dvKlienci.Item(j - 1).Item("MARKER") = True
                            dvKlienci.Item(j - 1).Item("MARKERP") = True
                            'dvKlienci.Item(j - 1).Item("WARTSPRZED") = WartSP
                        End If
                    Next
                End If
            Else
                SzukajKlienta(0) = Klient
                dsRow = _dsKlienci.Tables(0).Rows.Find(SzukajKlienta)
                ' sprawdz czy jest...
                If Not (dsRow Is Nothing) Then
                    dsRow("MARKER") = True
                    dsRow("MARKERP") = True
                    'dsRow("WARTSPRZED") = WartSP

                    _AktualizujKartKlientow()
                Else
                    'brak klienta
                End If
            End If
            'System.Windows.Forms.Application.DoEvents()
        Next

        Me.Close()
    End Sub






















    Dim dbSelectKlienci As String = "SELECT * FROM Klienci"

    Dim dbConnection As OleDb.OleDbConnection
    Dim dbCommandSelect As OleDb.OleDbCommand
    Dim dbCommandDelete As OleDb.OleDbCommand
    Dim dbCommandUpdate As OleDb.OleDbCommand
    Dim dbCommandInsert As OleDb.OleDbCommand
    Dim dbDataAdapter As OleDb.OleDbDataAdapter

    Dim sqlConnectionKlienci As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienci As SqlClient.SqlCommand
    Dim sqlCommandDeleteKlienci As SqlClient.SqlCommand
    Dim sqlCommandUpdateKlienci As SqlClient.SqlCommand
    Dim sqlCommandInsertKlienci As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienci As SqlClient.SqlDataAdapter

    Dim DataSetKlienci As New DataSet
    Dim DataViewKlienci As New DataView
    Dim ConnectionState As System.Data.ConnectionState


    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Timer2.Enabled = False

        Dim ir As Integer = 0
        Dim Li As Integer = 0
        Dim dsRow As DataRow = Nothing
        Dim dvRow As DataRowView = Nothing
        Dim Klient As String = ""




        'dbSelectKlienci = sSelectSQLKlienci

        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionKlienci)
            'dbCommandSelect = New OleDb.OleDbCommand(dbSelectKlienci & " ORDER BY NRKARTY", dbConnection)
            ''dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLKlienci, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLKlienci, dbConnection)
            ''dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLKlienci, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Klienci")

            ''DBDeleteKlienci(dbCommandDelete, dbDataAdapter)
            'DBUpdateKlienci(dbCommandUpdate, dbDataAdapter)
            ''DBInsertKlienci(dbCommandInsert, dbDataAdapter)

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapter.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnection.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst�pi� b��d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnection.State
            'End Try


        Else
            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
            sqlCommandSelectKlienci = New SqlClient.SqlCommand(sSelectSQLKlienciAktywni & " ORDER BY NRKARTY", sqlConnectionKlienci)
            'sqlCommandDeleteKlienci = New SqlClient.SqlCommand(sDeleteSQLKlienci, sqlConnection)
            sqlCommandUpdateKlienci = New SqlClient.SqlCommand(sUpdateSQLKlienci, sqlConnectionKlienci)
            'sqlCommandInsertKlienci = New SqlClient.SqlCommand(sInsertSQLKlienci, sqlConnectionKlienci)
            sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            'SQLDeleteKlienci(sqlCommandDeleteKlienci, sqlDataAdapterKlienci)
            SQLUpdateKlienci(sqlCommandUpdateKlienci, sqlDataAdapterKlienci)
            'SQLInsertKlienci(sqlCommandInsertKlienci, sqlDataAdapterKlienci)

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterKlienci.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienci.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst�pi� b��d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionKlienci.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    dbDataAdapter.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    dbDataAdapter.Fill(DataSetKlienci)
                    dbConnection.Close()
                Else
                    sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("NRKARTY")}

                'definiuj DataView
                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
                DataViewKlienci.AllowDelete = False
                DataViewKlienci.AllowNew = False

                DataViewKlienci.Sort = "NRKARTY"
                '-------------------------------
                Dim i As Long
                Li = 0
                If _InicjujKartKlientow Then
                    'usun wszystkie markery
                    lblOperacja.Text = "Inicjuje pole MARKER"
                    Progres.Maximum = DataViewKlienci.Count - 1
                    System.Windows.Forms.Application.DoEvents()
                    For i = 0 To DataViewKlienci.Count - 1
                        Progres.Value = i
                        dvRow = DataViewKlienci.Item(i)
                        If (dvRow("MARKER") = True) Or (dvRow("MARKERP") = True) Then
                            If dvRow("MARKER") Then dvRow("MARKER") = False
                            If dvRow("MARKERP") Then dvRow("MARKERP") = False
                            '_AktualizujKartKlientow()
                            Li += 1
                        End If
                        'If Int(i / 100) = i / 100 Then System.Windows.Forms.Application.DoEvents()
                    Next
                End If
                'MsgBox("Wyzerowa�em = " & Str(Li))



                Li = 0
                If _dvAkcjeSpec.Count > 0 Then
                    'USTAW MARKERY KLIENTOW Z TEJ ANALIZY
                    lblOperacja.Text = "Ustawiam pole MARKER dla klient�w z wynik�w analizy."
                    Progres.Maximum = _dvAkcjeSpec.Count - 1
                    System.Windows.Forms.Application.DoEvents()
                    For i = 0 To _dvAkcjeSpec.Count - 1
                        Progres.Value = i

                        dvRow = _dvAkcjeSpec.Item(i)
                        Klient = dvRow("IDENTKLI")

                        ir = DataViewKlienci.Find(Klient)
                        ' sprawdz czy jest...
                        If ir >= 0 Then
                            If DataViewKlienci.Item(ir).Item("MARKER") = True Then
                            Else
                                DataViewKlienci.Item(ir).Item("MARKER") = True
                                DataViewKlienci.Item(ir).Item("MARKERP") = True
                                '_AktualizujKartKlientow()
                                Li += 1
                            End If

                        End If

                        'If Int(i / 100) = i / 100 Then System.Windows.Forms.Application.DoEvents()
                    Next
                    'MsgBox("Ustawi�em = " & Str(Li))

                    lblOperacja.Text = "Aktualizuj� tabel� klienci w Bazie Danych."
                    System.Windows.Forms.Application.DoEvents()
                    AktualizujBaze()
                    System.Windows.Forms.Application.DoEvents()
                    '-------------------------------
                End If

                'MsgBox("Ilo�� wybranych klient�w " & Trim(Str(Li)))

            Catch Ex As System.Exception
                'ViewInLog(ex, Me.Name(), Nothing)
                MessageBox.Show("W programie wyst�pi� b��d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End Try
        End If
        Me.Close()

    End Sub

    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujBaze()
        If ParL_DataBaseType = "MS ACCESS" Then
            Try
                dbConnection.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If dbConnection.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    dbDataAdapter.Update(DataSetKlienci, "TABELA_Klienci")
                Catch Ex As System.Exception
                    'ViewInLog(ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    dbDataAdapter.Fill(DataSetKlienci)
                End If
                dbConnection.Close()
            End If
        Else
            Try
                sqlConnectionKlienci.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionKlienci.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterKlienci.Update(DataSetKlienci, "TABELA_Klienci")
                Catch Ex As System.Exception
                    'ViewInLog(ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                End If
                sqlConnectionKlienci.Close()
            End If
        End If

    End Sub










End Class
