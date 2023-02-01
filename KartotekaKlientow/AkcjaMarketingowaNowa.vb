Public Class AkcjaMarketingowaNowa

    Private Sub NowaAkcjaMarketingowa_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        Me.PictureBox1.Image = My.Resources.logomini
        '-------------------------------
        txtIleZap.Text = _dvKlenci.Count.ToString("N0")
        txtIdent.Text = ""
        txtOpis.Text = ""
        dtpData.Value = SysData()
        txtUwagi.Text = ""
    End Sub

    Private Sub NowaAkcjaMarketingowa_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        RysujGradient(Me)
    End Sub

    '*******************************************
    '** procedury obslugi edycji
    '*******************************************

    Private Sub txtOpis_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOpis.TextChanged
        TestLen(txtOpis, "OPIS", 100)
    End Sub
    Private Sub txtOpis_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOpis.GotFocus
        StartEdycjiTxt(txtOpis)
    End Sub
    Private Sub txtOpis_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOpis.LostFocus
        KoniecEdycjiTxt(txtOpis)
    End Sub

    'Private Sub txtData_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtData.TextChanged
    '    TestLen(txtData, "DATA", 10)
    'End Sub
    'Private Sub txtData_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtData.GotFocus
    '    StartEdycjiTxt(txtData)
    'End Sub
    'Private Sub txtData_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtData.LostFocus
    '    KoniecEdycjiTxt(txtData)
    'End Sub

    Private Sub txtIdent_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIdent.TextChanged
        TestLen(txtIdent, "IDENTYFIKATOR", 20)
    End Sub
    Private Sub txtIdent_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIdent.GotFocus
        StartEdycjiTxt(txtIdent)
    End Sub
    Private Sub txtIdent_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIdent.LostFocus
        KoniecEdycjiTxt(txtIdent)
    End Sub

    Private Sub txtUwagi_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUwagi.GotFocus
        StartEdycjiTxt(txtUwagi)
    End Sub
    Private Sub txtUwagi_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUwagi.LostFocus
        KoniecEdycjiTxt(txtUwagi)
    End Sub

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub

    '**************************************************
    '** zapamietaj informacje o Akcjac...
    '**************************************************

    'Akcje Handlowe - Opis
    'Public aoIdentAkcji As String            'IDENT     Text 20
    'Public aoDataAkcji As String             'DATA      Text,10
    'Public aoOpisAkcji As String             'OPIS      Text,100
    'Public aoUwagiAkcji As String            'UWAGI     Memo
    'Public aoMarkerAkcji As boolean          'MARKER logical
    'Public aoWersjaAkcji As Integer          'WERSJA    Integer

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


    Dim DataSetAkcjeOpis As DataSet
    Dim DataViewAkcjeOpis As DataView

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

    Dim ConnectionState As System.Data.ConnectionState



    Private Sub cmdZapamietaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdZapamietaj.Click
        If _dvKlenci.Count = 0 Then
            MessageBox.Show("Specyfikacja klientów jest pusta.", _
                "Uwaga:", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            If Len(txtIdent.Text) = 0 Then
                MessageBox.Show("Proszê wpisaæ identyfikator akcji marketingowej", _
                    "Uwaga:", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            Else
                If CzyJestTakaAkcja(txtIdent.Text) Then
                    MessageBox.Show("Istnieje ju¿ akcja marketingowa o takim identyfikatorze", _
                        "Uwaga:", _
                        System.Windows.Forms.MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, _
                        MessageBoxDefaultButton.Button1)
                Else
                    DataSetAkcjeOpis = New DataSet
                    DataViewAkcjeOpis = New DataView
                    DataSetAkcjeOpis.Locale = New System.Globalization.CultureInfo("pl-PL")

                    If ParL_DataBaseType = "MS ACCESS" Then
                        ''OK - mo¿na zapisaæ....
                        'dbConnectionAkcjeOpis = New OleDb.OleDbConnection(sConnectionAkcjeOpis)
                        'dbCommandSelectAkcjeOpis = New OleDb.OleDbCommand(sSelectSQLAkcjeOpis, dbConnectionAkcjeOpis)
                        'dbCommandDeleteAkcjeOpis = New OleDb.OleDbCommand(sDeleteSQLAkcjeOpis, dbConnectionAkcjeOpis)
                        'dbCommandUpdateAkcjeOpis = New OleDb.OleDbCommand(sUpdateSQLAkcjeOpis, dbConnectionAkcjeOpis)
                        'dbCommandInsertAkcjeOpis = New OleDb.OleDbCommand(sInsertSQLAkcjeOpis, dbConnectionAkcjeOpis)
                        'dbDataAdapterAkcjeOpis = New OleDb.OleDbDataAdapter(dbCommandSelectAkcjeOpis)

                        'Dim MapowanieTabeliAkcjeOpis As System.Data.Common.DataTableMapping
                        'MapowanieTabeliAkcjeOpis = dbDataAdapterAkcjeOpis.TableMappings.Add("Table", "TABELA_AkcjeOpis")

                        'Dim objParamAkcjeOpis As OleDb.OleDbParameter
                        ''------------------------------------------
                        ''komenda DELETE
                        ''parametry filtrowania
                        'objParamAkcjeOpis = dbCommandDeleteAkcjeOpis.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 20, "IDENT")
                        'objParamAkcjeOpis.SourceVersion = DataRowVersion.Original
                        'objParamAkcjeOpis = dbCommandDeleteAkcjeOpis.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
                        'objParamAkcjeOpis.SourceVersion = DataRowVersion.Original
                        'dbDataAdapterAkcjeOpis.DeleteCommand = dbCommandDeleteAkcjeOpis

                        ''------------------------------------------
                        ''komenda UPDATE
                        ''parametry aktualizacji
                        ''dbCommandUpdateAkcjeOpis.Parameters.Add("@oIdentAkcji", OleDb.OleDbType.Char, 20, "IDENT")
                        'dbCommandUpdateAkcjeOpis.Parameters.Add("@oDataAkcji", OleDb.OleDbType.Char, 10, "DATA")
                        'dbCommandUpdateAkcjeOpis.Parameters.Add("@oOpisAkcji", OleDb.OleDbType.Char, 100, "OPIS")
                        'dbCommandUpdateAkcjeOpis.Parameters.Add("@oUwagiAkcji", OleDb.OleDbType.WChar, Nothing, "UWAGI")
                        'dbCommandUpdateAkcjeOpis.Parameters.Add("@oMarkerAkcji", OleDb.OleDbType.Boolean, Nothing, "MARKER")
                        'dbCommandUpdateAkcjeOpis.Parameters.Add("@oWersjaAkcji", OleDb.OleDbType.Integer, Nothing, "WERSJA")

                        ''parametry filtrowania
                        'objParamAkcjeOpis = dbCommandUpdateAkcjeOpis.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 20, "IDENT")
                        'objParamAkcjeOpis.SourceVersion = DataRowVersion.Original
                        'objParamAkcjeOpis = dbCommandUpdateAkcjeOpis.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
                        'objParamAkcjeOpis.SourceVersion = DataRowVersion.Original
                        'dbDataAdapterAkcjeOpis.UpdateCommand = dbCommandUpdateAkcjeOpis
                        ''------------------------------------------
                        ''komenda INSERT
                        'objParamAkcjeOpis = dbCommandInsertAkcjeOpis.Parameters.Add("@oIdentAkcji", OleDb.OleDbType.Char, 20, "IDENT")
                        'objParamAkcjeOpis.SourceVersion = DataRowVersion.Current
                        'objParamAkcjeOpis = dbCommandInsertAkcjeOpis.Parameters.Add("@oDataAkcji", OleDb.OleDbType.Char, 10, "DATA")
                        'objParamAkcjeOpis.SourceVersion = DataRowVersion.Current
                        'objParamAkcjeOpis = dbCommandInsertAkcjeOpis.Parameters.Add("@oOpisAkcji", OleDb.OleDbType.Char, 100, "OPIS")
                        'objParamAkcjeOpis.SourceVersion = DataRowVersion.Current
                        'objParamAkcjeOpis = dbCommandInsertAkcjeOpis.Parameters.Add("@oUwagiAkcji", OleDb.OleDbType.WChar, Nothing, "UWAGI")
                        'objParamAkcjeOpis.SourceVersion = DataRowVersion.Current
                        'objParamAkcjeOpis = dbCommandInsertAkcjeOpis.Parameters.Add("@oMarkerAkcji", OleDb.OleDbType.Boolean, Nothing, "MARKER")
                        'objParamAkcjeOpis.SourceVersion = DataRowVersion.Current
                        'objParamAkcjeOpis = dbCommandInsertAkcjeOpis.Parameters.Add("@oWersjaAkcji", OleDb.OleDbType.Integer, Nothing, "WERSJA")
                        'objParamAkcjeOpis.SourceVersion = DataRowVersion.Current

                        'dbDataAdapterAkcjeOpis.InsertCommand = dbCommandInsertAkcjeOpis

                        '' Add the RowUpdated event handler.
                        'AddHandler dbDataAdapterAkcjeOpis.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

                        ''------------------------------------------
                        ''wypelnij DATASET
                        'Try
                        '    dbConnectionAkcjeOpis.Open()
                        'Catch Ex As System.Exception
                        '    ViewInLog(Ex, Me.Name(), Nothing)
                        '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    '    MessageBoxIcon.Information, _
                        '    '    MessageBoxDefaultButton.Button1)
                        'Finally
                        '    ConnectionState = dbConnectionAkcjeOpis.State
                        'End Try
                    Else
                        'OK - mo¿na zapisaæ....
                        sqlConnectionAkcjeOpis = New SqlClient.SqlConnection(sConnectionAkcjeOpis)
                        sqlCommandSelectAkcjeOpis = New SqlClient.SqlCommand(sSelectSQLAkcjeOpis, sqlConnectionAkcjeOpis)
                        sqlCommandDeleteAkcjeOpis = New SqlClient.SqlCommand(sDeleteSQLAkcjeOpis, sqlConnectionAkcjeOpis)
                        sqlCommandUpdateAkcjeOpis = New SqlClient.SqlCommand(sUpdateSQLAkcjeOpis, sqlConnectionAkcjeOpis)
                        sqlCommandInsertAkcjeOpis = New SqlClient.SqlCommand(sInsertSQLAkcjeOpis, sqlConnectionAkcjeOpis)
                        sqlDataAdapterAkcjeOpis = New SqlClient.SqlDataAdapter(sqlCommandSelectAkcjeOpis)

                        Dim MapowanieTabeliAkcjeOpis As System.Data.Common.DataTableMapping
                        MapowanieTabeliAkcjeOpis = sqlDataAdapterAkcjeOpis.TableMappings.Add("Table", "TABELA_AkcjeOpis")

                        Dim objParamAkcjeOpis As SqlClient.SqlParameter
                        '------------------------------------------
                        'komenda DELETE
                        'parametry filtrowania
                        objParamAkcjeOpis = sqlCommandDeleteAkcjeOpis.Parameters.Add("@orygSymbol", SqlDbType.Char, 20, "IDENT")
                        objParamAkcjeOpis.SourceVersion = DataRowVersion.Original
                        objParamAkcjeOpis = sqlCommandDeleteAkcjeOpis.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
                        objParamAkcjeOpis.SourceVersion = DataRowVersion.Original
                        sqlDataAdapterAkcjeOpis.DeleteCommand = sqlCommandDeleteAkcjeOpis

                        '------------------------------------------
                        'komenda UPDATE
                        'parametry aktualizacji
                        'sqlCommandUpdateAkcjeOpis.Parameters.Add("@oIdentAkcji", sqlDbType.Char, 20, "IDENT")
                        sqlCommandUpdateAkcjeOpis.Parameters.Add("@oDataAkcji", SqlDbType.Char, 10, "DATA")
                        sqlCommandUpdateAkcjeOpis.Parameters.Add("@oOpisAkcji", SqlDbType.Char, 100, "OPIS")
                        sqlCommandUpdateAkcjeOpis.Parameters.Add("@oUwagiAkcji", SqlDbType.Text, Nothing, "UWAGI")
                        sqlCommandUpdateAkcjeOpis.Parameters.Add("@oMARKERAkcji", SqlDbType.Bit, Nothing, "MARKER")
                        sqlCommandUpdateAkcjeOpis.Parameters.Add("@oWersjaAkcji", SqlDbType.Int, Nothing, "WERSJA")

                        'parametry filtrowania
                        objParamAkcjeOpis = sqlCommandUpdateAkcjeOpis.Parameters.Add("@orygSymbol", SqlDbType.Char, 20, "IDENT")
                        objParamAkcjeOpis.SourceVersion = DataRowVersion.Original
                        objParamAkcjeOpis = sqlCommandUpdateAkcjeOpis.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
                        objParamAkcjeOpis.SourceVersion = DataRowVersion.Original
                        sqlDataAdapterAkcjeOpis.UpdateCommand = sqlCommandUpdateAkcjeOpis
                        '------------------------------------------
                        'komenda INSERT
                        objParamAkcjeOpis = sqlCommandInsertAkcjeOpis.Parameters.Add("@oIdentAkcji", SqlDbType.Char, 20, "IDENT")
                        objParamAkcjeOpis.SourceVersion = DataRowVersion.Current
                        objParamAkcjeOpis = sqlCommandInsertAkcjeOpis.Parameters.Add("@oDataAkcji", SqlDbType.Char, 10, "DATA")
                        objParamAkcjeOpis.SourceVersion = DataRowVersion.Current
                        objParamAkcjeOpis = sqlCommandInsertAkcjeOpis.Parameters.Add("@oOpisAkcji", SqlDbType.Char, 100, "OPIS")
                        objParamAkcjeOpis.SourceVersion = DataRowVersion.Current
                        objParamAkcjeOpis = sqlCommandInsertAkcjeOpis.Parameters.Add("@oUwagiAkcji", SqlDbType.Text, Nothing, "UWAGI")
                        objParamAkcjeOpis.SourceVersion = DataRowVersion.Current
                        objParamAkcjeOpis = sqlCommandInsertAkcjeOpis.Parameters.Add("@oMarkerAkcji", SqlDbType.Bit, Nothing, "MARKER")
                        objParamAkcjeOpis.SourceVersion = DataRowVersion.Current
                        objParamAkcjeOpis = sqlCommandInsertAkcjeOpis.Parameters.Add("@oWersjaAkcji", SqlDbType.Int, Nothing, "WERSJA")
                        objParamAkcjeOpis.SourceVersion = DataRowVersion.Current

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
                                'dbDataAdapterAkcjeOpis.FillSchema(DataSetAkcjeOpis, SchemaType.Mapped, "TABELA_AkcjeOpis")
                                'dbDataAdapterAkcjeOpis.Fill(DataSetAkcjeOpis)
                                'dbConnectionAkcjeOpis.Close()
                            Else
                                sqlDataAdapterAkcjeOpis.FillSchema(DataSetAkcjeOpis, SchemaType.Mapped, "TABELA_AkcjeOpis")
                                sqlDataAdapterAkcjeOpis.Fill(DataSetAkcjeOpis)
                                sqlConnectionAkcjeOpis.Close()
                            End If

                            'zdefiniuj unikalny klucz wyszukiwania w tabeli
                            DataSetAkcjeOpis.Tables(0).PrimaryKey = New DataColumn() {DataSetAkcjeOpis.Tables(0).Columns("IDENT")}

                            'definiuj DataView
                            DataViewAkcjeOpis = New DataView(DataSetAkcjeOpis.Tables(0))
                            DataViewAkcjeOpis.AllowDelete = False
                            DataViewAkcjeOpis.AllowEdit = False
                            DataViewAkcjeOpis.AllowNew = False

                        Catch Ex As System.Exception
                            ViewInLog(Ex, Me.Name(), Nothing)
                            'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                            '    System.Windows.Forms.MessageBoxButtons.OK, _
                            '    MessageBoxIcon.Information, _
                            '    MessageBoxDefaultButton.Button1)
                        End Try
                    End If






                    dbSelectSQLAkcjeSpec = "SELECT * FROM AkcjeSpec"
                    DataSetAkcjeSpec = New DataSet
                    DataViewAkcjeSpec = New DataView
                    DataSetAkcjeSpec.Locale = New System.Globalization.CultureInfo("pl-PL")

                    If ParL_DataBaseType = "MS ACCESS" Then
                        'dbConnectionAkcjeSpec = New OleDb.OleDbConnection(sConnectionAkcjeSpec)
                        'dbCommandSelectAkcjeSpec = New OleDb.OleDbCommand(dbSelectSQLAkcjeSpec, dbConnectionAkcjeSpec)
                        'dbCommandDeleteAkcjeSpec = New OleDb.OleDbCommand(sDeleteSQLAkcjeSpec, dbConnectionAkcjeSpec)
                        'dbCommandUpdateAkcjeSpec = New OleDb.OleDbCommand(sUpdateSQLAkcjeSpec, dbConnectionAkcjeSpec)
                        'dbCommandInsertAkcjeSpec = New OleDb.OleDbCommand(sInsertSQLAkcjeSpec, dbConnectionAkcjeSpec)
                        'dbDataAdapterAkcjeSpec = New OleDb.OleDbDataAdapter(dbCommandSelectAkcjeSpec)

                        'Dim MapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
                        'MapowanieTabeliAkcjeSpec = dbDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_AkcjeSpec")

                        'Dim objParamAkcjeSpec As OleDb.OleDbParameter
                        ''------------------------------------------
                        ''komenda DELETE
                        ''parametry filtrowania
                        'objParamAkcjeSpec = dbCommandDeleteAkcjeSpec.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 20, "IDENTAKCJI")
                        'objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
                        'objParamAkcjeSpec = dbCommandDeleteAkcjeSpec.Parameters.Add("@orygKlient", OleDb.OleDbType.Char, 10, "IDENTKLI")
                        'objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
                        'objParamAkcjeSpec = dbCommandDeleteAkcjeSpec.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
                        'objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
                        'dbDataAdapterAkcjeSpec.DeleteCommand = dbCommandDeleteAkcjeSpec

                        ''------------------------------------------
                        ''komenda UPDATE
                        ''parametry aktualizacji
                        ''dbCommandUpdateAkcjeSpec.Parameters.Add("@oIdentAkcji", OleDb.OleDbType.Char, 20, "IDENTAKCJI")
                        ''dbCommandUpdateAkcjeSpec.Parameters.Add("@oIdentKli", OleDb.OleDbType.Char, 10, "IDENTKLI")
                        'dbCommandUpdateAkcjeSpec.Parameters.Add("@oWersjaAkcji", OleDb.OleDbType.Integer, Nothing, "WERSJA")

                        ''parametry filtrowania
                        'objParamAkcjeSpec = dbCommandUpdateAkcjeSpec.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 20, "IDENTAKCJI")
                        'objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
                        'objParamAkcjeSpec = dbCommandUpdateAkcjeSpec.Parameters.Add("@orygKlient", OleDb.OleDbType.Char, 10, "IDENTKLI")
                        'objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
                        'objParamAkcjeSpec = dbCommandUpdateAkcjeSpec.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
                        'objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
                        'dbDataAdapterAkcjeSpec.UpdateCommand = dbCommandUpdateAkcjeSpec

                        ''------------------------------------------
                        ''komenda INSERT
                        'objParamAkcjeSpec = dbCommandInsertAkcjeSpec.Parameters.Add("@oIdentAkcji", OleDb.OleDbType.Char, 20, "IDENTAKCJI")
                        'objParamAkcjeSpec.SourceVersion = DataRowVersion.Current
                        'objParamAkcjeSpec = dbCommandInsertAkcjeSpec.Parameters.Add("@oIdentKli", OleDb.OleDbType.Char, 10, "IDENTKLI")
                        'objParamAkcjeSpec.SourceVersion = DataRowVersion.Current
                        'objParamAkcjeSpec = dbCommandInsertAkcjeSpec.Parameters.Add("@oWersjaAkcji", OleDb.OleDbType.Integer, Nothing, "WERSJA")
                        'objParamAkcjeSpec.SourceVersion = DataRowVersion.Current

                        'dbDataAdapterAkcjeSpec.InsertCommand = dbCommandInsertAkcjeSpec

                        '' Add the RowUpdated event handler.
                        'AddHandler dbDataAdapterAkcjeSpec.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

                        ''------------------------------------------
                        ''wypelnij DATASET
                        'Try
                        '    dbConnectionAkcjeSpec.Open()
                        'Catch Ex As System.Exception
                        '    ViewInLog(Ex, Me.Name(), Nothing)
                        '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    '    MessageBoxIcon.Information, _
                        '    '    MessageBoxDefaultButton.Button1)
                        'Finally
                        '    ConnectionState = dbConnectionAkcjeSpec.State
                        'End Try
                    Else
                        sqlConnectionAkcjeSpec = New SqlClient.SqlConnection(sConnectionAkcjeSpec)
                        sqlCommandSelectAkcjeSpec = New SqlClient.SqlCommand(dbSelectSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                        sqlCommandDeleteAkcjeSpec = New SqlClient.SqlCommand(sDeleteSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                        sqlCommandUpdateAkcjeSpec = New SqlClient.SqlCommand(sUpdateSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                        sqlCommandInsertAkcjeSpec = New SqlClient.SqlCommand(sInsertSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                        sqlDataAdapterAkcjeSpec = New SqlClient.SqlDataAdapter(sqlCommandSelectAkcjeSpec)

                        Dim MapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
                        MapowanieTabeliAkcjeSpec = sqlDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_AkcjeSpec")

                        Dim objParamAkcjeSpec As SqlClient.SqlParameter
                        '------------------------------------------
                        'komenda DELETE
                        'parametry filtrowania
                        objParamAkcjeSpec = sqlCommandDeleteAkcjeSpec.Parameters.Add("@orygSymbol", SqlDbType.Char, 20, "IDENTAKCJI")
                        objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
                        objParamAkcjeSpec = sqlCommandDeleteAkcjeSpec.Parameters.Add("@orygKlient", SqlDbType.Char, 10, "IDENTKLI")
                        objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
                        objParamAkcjeSpec = sqlCommandDeleteAkcjeSpec.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
                        objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
                        sqlDataAdapterAkcjeSpec.DeleteCommand = sqlCommandDeleteAkcjeSpec

                        '------------------------------------------
                        'komenda UPDATE
                        'parametry aktualizacji
                        'sqlCommandUpdateAkcjeSpec.Parameters.Add("@oIdentAkcji", sqlDbType.Char, 20, "IDENTAKCJI")
                        'sqlCommandUpdateAkcjeSpec.Parameters.Add("@oIdentKli", sqlDbType.Char, 10, "IDENTKLI")
                        sqlCommandUpdateAkcjeSpec.Parameters.Add("@oWersjaAkcji", SqlDbType.Int, Nothing, "WERSJA")

                        'parametry filtrowania
                        objParamAkcjeSpec = sqlCommandUpdateAkcjeSpec.Parameters.Add("@orygSymbol", SqlDbType.Char, 20, "IDENTAKCJI")
                        objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
                        objParamAkcjeSpec = sqlCommandUpdateAkcjeSpec.Parameters.Add("@orygKlient", SqlDbType.Char, 10, "IDENTKLI")
                        objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
                        objParamAkcjeSpec = sqlCommandUpdateAkcjeSpec.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
                        objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
                        sqlDataAdapterAkcjeSpec.UpdateCommand = sqlCommandUpdateAkcjeSpec

                        '------------------------------------------
                        'komenda INSERT
                        objParamAkcjeSpec = sqlCommandInsertAkcjeSpec.Parameters.Add("@oIdentAkcji", SqlDbType.Char, 20, "IDENTAKCJI")
                        objParamAkcjeSpec.SourceVersion = DataRowVersion.Current
                        objParamAkcjeSpec = sqlCommandInsertAkcjeSpec.Parameters.Add("@oIdentKli", SqlDbType.Char, 10, "IDENTKLI")
                        objParamAkcjeSpec.SourceVersion = DataRowVersion.Current
                        objParamAkcjeSpec = sqlCommandInsertAkcjeSpec.Parameters.Add("@oWersjaAkcji", SqlDbType.Int, Nothing, "WERSJA")
                        objParamAkcjeSpec.SourceVersion = DataRowVersion.Current

                        sqlDataAdapterAkcjeSpec.InsertCommand = sqlCommandInsertAkcjeSpec

                        ' Add the RowUpdated event handler.
                        AddHandler sqlDataAdapterAkcjeSpec.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

                        '------------------------------------------
                        'wypelnij DATASET
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
                                'dbDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_AkcjeSpec")
                                'dbDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                                'dbConnectionAkcjeSpec.Close()
                            Else
                                sqlDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_AkcjeSpec")
                                sqlDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                                sqlConnectionAkcjeSpec.Close()
                            End If

                            'zdefiniuj unikalny klucz wyszukiwania w tabeli
                            DataSetAkcjeSpec.Tables(0).PrimaryKey = New DataColumn() {DataSetAkcjeSpec.Tables(0).Columns("IDENTAKCJI"), DataSetAkcjeSpec.Tables(0).Columns("IDENTKLI")}

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

                    Dim NewRow As DataRow

                    'DOPISZ opis akcji
                    'stworz nowy rekord
                    NewRow = DataSetAkcjeOpis.Tables(0).NewRow()
                    NewRow("IDENT") = txtIdent.Text
                    NewRow("DATA") = dtpData.Value.ToString("yyyy-MM-dd")
                    NewRow("OPIS") = txtOpis.Text
                    NewRow("UWAGI") = txtUwagi.Text
                    NewRow("MARKER") = False
                    NewRow("WERSJA") = 1
                    'dodaj rekord do DataSet
                    DataSetAkcjeOpis.Tables(0).Rows.Add(NewRow)
                    AktualizujBazeAkcjeOpis()
                    '...........................
                    'definiuj delegat
                    Dim deleg As delegateAktualizuj = AddressOf AktualizujBazeAkcjeSpec

                    Dim FormRaport As New RaportNowaAkcjaMarketingowa(txtIdent.Text, _
                                                                         _dvKlenci, _
                                                                         DataSetAkcjeSpec, _
                                                                         deleg)
                    FormRaport.ShowDialog()
                    '...........................
                    MessageBox.Show("Zapamiêta³em informacje o nowej akcji marketingowej" & vbCrLf & txtIdent.Text, _
                        "OK - skoñczy³em", _
                        System.Windows.Forms.MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, _
                        MessageBoxDefaultButton.Button1)


                End If
            End If
        End If
    End Sub


    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujBazeAkcjeOpis()
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionAkcjeOpis.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnectionAkcjeOpis.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapterAkcjeOpis.Update(DataSetAkcjeOpis, "TABELA_AkcjeOpis")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapterAkcjeOpis.Fill(DataSetAkcjeOpis)
            '    End If
            '    dbConnectionAkcjeOpis.Close()
            'End If
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


    Private Sub AktualizujBazeAkcjeSpec()
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionAkcjeSpec.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnectionAkcjeSpec.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapterAkcjeSpec.Update(DataSetAkcjeSpec, "TABELA_AkcjeSpec")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
            '    End If
            '    dbConnectionAkcjeSpec.Close()
            'End If
        Else
            Try
                sqlConnectionAkcjeSpec.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionAkcjeSpec.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterAkcjeSpec.Update(DataSetAkcjeSpec, "TABELA_AkcjeSpec")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                End If
                sqlConnectionAkcjeSpec.Close()
            End If
        End If
    End Sub


End Class