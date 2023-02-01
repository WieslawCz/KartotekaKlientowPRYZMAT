Module modStrukturyLogAktywnosci





    '----------------------------------------------------------
    'Katalog LogAktywnosci
    '----------------------------------------------------------
    Public LG_UniqID As String       'UNIQID     STRING 40
    Public LG_Temat As String        'TEMAT      STRING 100
    Public LG_Katalog As String      'KATALOG    STRING 100
    Public LG_DataZapisu As String   'DATAZAPISU STRING 30
    Public LG_Pracownik As String    'PRACOWNIK  STRING 10
    Public LG_Operacja As String     'OPERACJA   STRING 20
    Public LG_Uwagi As String        'UWAGI      Text
    Public LG_Wersja As Integer      'WERSJA     Integer


    Public sConnectionLogAktywnosci As String = ConnectionStrings()
    Public HQConnectionLogAktywnosci As String = HQConnectionStrings()

    Public sSelectSQLLogAktywnosci As String = "SELECT " &
                                                    "UNIQID, " &
                                                    "TEMAT, " &
                                                    "KATALOG, " &
                                                    "DATAZAPISU, " &
                                                    "PRACOWNIK, " &
                                                    "OPERACJA, " &
                                                    "UWAGI, " &
                                                    "WERSJA " &
                                              "FROM LogAktywnosci"

    Public sDeleteSQLLogAktywnosci As String = "DELETE FROM LogAktywnosci " &
                                        "WHERE (UNIQID=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sUpdateSQLLogAktywnosci As String = "UPDATE LogAktywnosci SET " &
                                                    "TEMAT=@lgTEMAT, " &
                                                    "KATALOG=@lgKATALOG, " &
                                                    "DATAZAPISU=@lgDATAZAPISU, " &
                                                    "PRACOWNIK=@lgPRACOWNIK, " &
                                                    "OPERACJA=@lgOPERACJA, " &
                                                    "UWAGI=@lgUWAGI, " &
                                                    "WERSJA=@lgWersja " &
                                        "WHERE (UNIQID=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sInsertSQLLogAktywnosci As String = "INSERT INTO LogAktywnosci " &
                                           "( " &
                                                    "UNIQID, " &
                                                    "TEMAT, " &
                                                    "KATALOG, " &
                                                    "DATAZAPISU, " &
                                                    "PRACOWNIK, " &
                                                    "OPERACJA, " &
                                                    "UWAGI, " &
                                                    "WERSJA " &
                                         ") " &
                                "VALUES  ( " &
                                         "@lgUNIQID, " &
                                         "@lgTEMAT, " &
                                         "@lgKATALOG, " &
                                         "@lgDATAZAPISU, " &
                                         "@lgPRACOWNIK, " &
                                         "@lgOPERACJA, " &
                                         "@lgUWAGI, " &
                                         "@lgWERSJA " &
                                         ") "

    'SQLDeleteLogAktywnosci(sqlCommandDelete, sqlDataAdapter)
    'SQLUpdateLogAktywnosci(sqlCommandUpdate, sqlDataAdapter)
    'SQLInsertLogAktywnosci(sqlCommandInsert, sqlDataAdapter)

    Public Sub SQLDeleteLogAktywnosci(ByRef sqlCommandDelete As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter

        'parametry filtrowania
        sqlParam = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 40, "UNIQID")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original

        sqlDataAdapter.DeleteCommand = sqlCommandDelete
    End Sub

    Public Sub SQLUpdateLogAktywnosci(ByRef sqlCommandUpdate As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter

        'Public LG_UniqID As String       'UNIQID     STRING 40
        'Public LG_Temat As String        'TEMAT      STRING 100
        'Public LG_Katalog As String      'KATALOG    STRING 100
        'Public LG_DataZapisu As String   'DATAZAPISU STRING 30
        'Public LG_Pracownik As String    'PRACOWNIK  STRING 10
        'Public LG_Operacja As String     'OPERACJA   STRING 20
        'Public LG_Uwagi As String        'UWAGI      Text
        'Public LG_Wersja As Integer      'WERSJA     Integer

        'sqlCommandUpdate.Parameters.Add("@lgUNIQID", sqldbtype.varChar, 40, "UNIQID")
        sqlCommandUpdate.Parameters.Add("@lgTEMAT", SqlDbType.VarChar, 100, "TEMAT")
        sqlCommandUpdate.Parameters.Add("@lgKATALOG", SqlDbType.VarChar, 100, "KATALOG")
        sqlCommandUpdate.Parameters.Add("@lgDATAZAPISU", SqlDbType.VarChar, 30, "DATAZAPISU")
        sqlCommandUpdate.Parameters.Add("@lgPRACOWNIK", SqlDbType.VarChar, 10, "PRACOWNIK")
        sqlCommandUpdate.Parameters.Add("@lgOPERACJA", SqlDbType.VarChar, 20, "OPERACJA")
        sqlCommandUpdate.Parameters.Add("@lgUWAGI", SqlDbType.Text, Nothing, "UWAGI")
        sqlCommandUpdate.Parameters.Add("@lgWERSJA", SqlDbType.Int, Nothing, "WERSJA")

        'parametry filtrowania
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 40, "UNIQID")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.UpdateCommand = sqlCommandUpdate
    End Sub

    Public Sub SQLInsertLogAktywnosci(ByRef sqlCommandInsert As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter

        sqlParam = sqlCommandInsert.Parameters.Add("@lgUNIQID", SqlDbType.VarChar, 40, "UNIQID")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@lgTEMAT", SqlDbType.VarChar, 100, "TEMAT")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@lgKATALOG", SqlDbType.VarChar, 100, "KATALOG")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@lgDATAZAPISU", SqlDbType.VarChar, 30, "DATAZAPISU")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@lgPRACOWNIK", SqlDbType.VarChar, 10, "PRACOWNIK")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@lgOPERACJA", SqlDbType.VarChar, 20, "OPERACJA")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@lgUWAGI", SqlDbType.Text, Nothing, "UWAGI")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@lgWERSJA", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlDataAdapter.InsertCommand = sqlCommandInsert
    End Sub



    'Public Sub DBDeleteLogAktywnosci(ByRef dbCommandDelete As OleDb.OleDbCommand,
    '                       ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As OleDb.OleDbParameter

    '    dbParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 10, "IDENT")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original

    '    dbDataAdapter.DeleteCommand = dbCommandDelete
    'End Sub

    'Public Sub DBUpdateLogAktywnosci(ByRef dbCommandUpdate As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter

    '    'dbCommandUpdate.Parameters.Add("@oIdentUzytkownika", OleDb.OleDbType.varChar, 10, "IDENT")
    '    dbCommandUpdate.Parameters.Add("@oNazwaUzytkownika", OleDb.OleDbType.VarChar, 100, "NAZWA")
    '    dbCommandUpdate.Parameters.Add("@oFUNKCJAUzytkownika", OleDb.OleDbType.VarChar, 100, "FUNKCJA")
    '    dbCommandUpdate.Parameters.Add("@oHASLOUzytkownika", OleDb.OleDbType.VarChar, 100, "HASLO")
    '    dbCommandUpdate.Parameters.Add("@oUprawnieniaUzytkownika", OleDb.OleDbType.VarChar, 100, "UPRAWNIENIA")
    '    dbCommandUpdate.Parameters.Add("@oUWAGIUzytkownika", OleDb.OleDbType.WChar, Nothing, "UWAGI")
    '    dbCommandUpdate.Parameters.Add("@oWersjaUzytkownika", OleDb.OleDbType.Integer, Nothing, "WERSJA")

    '    'parametry filtrowania
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 10, "IDENT")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.UpdateCommand = dbCommandUpdate
    'End Sub

    'Public Sub DBInsertLogAktywnosci(ByRef dbCommandInsert As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter

    '    dbParam = dbCommandInsert.Parameters.Add("@oIdentUzytkownika", OleDb.OleDbType.VarChar, 10, "IDENT")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oNazwaUzytkownika", OleDb.OleDbType.VarChar, 100, "NAZWA")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oFunkcjaUzytkownika", OleDb.OleDbType.VarChar, 100, "FUNKCJA")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oHASLOUzytkownika", OleDb.OleDbType.VarChar, 100, "HASLO")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oUPRAWNIENIAUzytkownika", OleDb.OleDbType.VarChar, 100, "UPRAWNIENIA")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oUWAGIUzytkownika", OleDb.OleDbType.WChar, Nothing, "UWAGI")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oWersjaUzytkownika", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Current

    '    dbDataAdapter.InsertCommand = dbCommandInsert
    'End Sub






    '**************************************
    '** pobierz informacje o LogAktywnosci
    '**************************************
    'Public LG_UniqID As String       'UNIQID     STRING 40
    'Public LG_Temat As String        'TEMAT    STRING 100
    'Public LG_Katalog As String      'KATALOG    STRING 100
    'Public LG_DataZapisu As String   'DATAZAPISU STRING 30
    'Public LG_Pracownik As String    'PRACOWNIK  STRING 10
    'Public LG_Operacja As String     'OPERACJA   STRING 20
    'Public LG_Uwagi As String        'UWAGI      Text
    'Public LG_Wersja As Integer      'WERSJA     Integer

    Public Sub InitLogAktywnosci()
        LG_UniqID = Guid.NewGuid.ToString       'UNIQID     STRING 40
        LG_Temat = ""                           'TEMAT      STRING 100
        LG_Katalog = ""                         'KATALOG    STRING 100
        LG_DataZapisu = SysTime()               'DATAZAPISU STRING 30
        LG_Pracownik = Program_IdUzytkownika    'PRACOWNIK  STRING 10
        LG_Operacja = ""                        'OPERACJA   STRING 20
        LG_Uwagi = ""                           'UWAGI      Text
        LG_Wersja = 1                           'WERSJA     Integer
    End Sub


    Public Function IdentLogAktywnosci(ByVal IdentLG As String) As Boolean
        Dim dbSelectSQLLogAktywnosci As String = sSelectSQLLogAktywnosci & " WHERE IDENT='" & IdentLG & "'"

        Dim dbConnectionLogAktywnosci As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectLogAktywnosci As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterLogAktywnosci As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionLogAktywnosci As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectLogAktywnosci As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterLogAktywnosci As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetLogAktywnosci As New DataSet
        Dim DataViewLogAktywnosci As New DataView
        Dim ConnectionStateLogAktywnosci As System.Data.ConnectionState

        Dim Wynik As Boolean = False

        InitLogAktywnosci()

        DataSetLogAktywnosci.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionLogAktywnosci = New OleDb.OleDbConnection(sConnectionLogAktywnosci)
            'dbCommandSelectLogAktywnosci = New OleDb.OleDbCommand(dbSelectSQLLogAktywnosci, dbConnectionLogAktywnosci)
            'dbDataAdapterLogAktywnosci = New OleDb.OleDbDataAdapter(dbCommandSelectLogAktywnosci)

            ''mapujemy domyslna nazwe tabeli
            'Dim dbMapowanieTabeliLogAktywnosci As System.Data.Common.DataTableMapping
            'dbMapowanieTabeliLogAktywnosci = dbDataAdapterLogAktywnosci.TableMappings.Add("Table", "TABELA_LogAktywnosci")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionLogAktywnosci.Open()
            'Catch Ex As System.Exception
            'Finally
            '    ConnectionStateLogAktywnosci = dbConnectionLogAktywnosci.State
            'End Try
        Else
            sqlConnectionLogAktywnosci = New SqlClient.SqlConnection(sConnectionLogAktywnosci)
            sqlCommandSelectLogAktywnosci = New SqlClient.SqlCommand(dbSelectSQLLogAktywnosci, sqlConnectionLogAktywnosci)
            sqlDataAdapterLogAktywnosci = New SqlClient.SqlDataAdapter(sqlCommandSelectLogAktywnosci)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliLogAktywnosci As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliLogAktywnosci = sqlDataAdapterLogAktywnosci.TableMappings.Add("Table", "TABELA_LogAktywnosci")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionLogAktywnosci.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateLogAktywnosci = sqlConnectionLogAktywnosci.State
            End Try
        End If

        If ConnectionStateLogAktywnosci = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterLogAktywnosci.FillSchema(DataSetLogAktywnosci, SchemaType.Mapped, "TABELA_LogAktywnosci")
                    'dbDataAdapterLogAktywnosci.Fill(DataSetLogAktywnosci)
                    'dbConnectionLogAktywnosci.Close()
                Else
                    sqlDataAdapterLogAktywnosci.FillSchema(DataSetLogAktywnosci, SchemaType.Mapped, "TABELA_LogAktywnosci")
                    sqlDataAdapterLogAktywnosci.Fill(DataSetLogAktywnosci)
                    sqlConnectionLogAktywnosci.Close()
                End If

                'definiuj DataView
                DataViewLogAktywnosci = New DataView(DataSetLogAktywnosci.Tables(0))
                If DataViewLogAktywnosci.Count > 0 Then
                    'Public LG_UniqID As String       'UNIQID     STRING 40
                    'Public LG_Temat As String        'TEMAT    STRING 100
                    'Public LG_Katalog As String      'KATALOG    STRING 100
                    'Public LG_DataZapisu As String   'DATAZAPISU STRING 30
                    'Public LG_Pracownik As String    'PRACOWNIK  STRING 10
                    'Public LG_Operacja As String     'OPERACJA   STRING 20
                    'Public LG_Uwagi As String        'UWAGI      Text
                    'Public LG_Wersja As Integer      'WERSJA     Integer

                    LG_UniqID = IdentLG
                    LG_Temat = GetTxtDRVField(DataViewLogAktywnosci.Item(0), "TEMAT")
                    LG_Katalog = GetTxtDRVField(DataViewLogAktywnosci.Item(0), "KATALOG")
                    LG_DataZapisu = GetTxtDRVField(DataViewLogAktywnosci.Item(0), "DATAZAPISU")
                    LG_Pracownik = GetTxtDRVField(DataViewLogAktywnosci.Item(0), "PRACOWNIK")
                    LG_Operacja = GetTxtDRVField(DataViewLogAktywnosci.Item(0), "OPERACJA")
                    LG_Uwagi = GetTxtDRVField(DataViewLogAktywnosci.Item(0), "UWAGI")
                    LG_Wersja = GetIntDRVField(DataViewLogAktywnosci.Item(0), "WERSJA")

                    Wynik = True
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wynik)
    End Function









    '--------------------------------------
    'Wykonanie kweredny w BazieDanych
    '--------------------------------------
    'Public LG_UniqID As String       'UNIQID     STRING 40
    'Public LG_Temat As String        'TEMAT    STRING 100
    'Public LG_Katalog As String      'KATALOG    STRING 100
    'Public LG_DataZapisu As String   'DATAZAPISU STRING 30
    'Public LG_Pracownik As String    'PRACOWNIK  STRING 10
    'Public LG_Operacja As String     'OPERACJA   STRING 20
    'Public LG_Uwagi As String        'UWAGI      Text
    'Public LG_Wersja As Integer      'WERSJA     Integer

    Public LOG_Aktywny As Boolean = True

    Public LOG_OperacjaUsun As String = "Usunięcie zapisu"
    Public LOG_OperacjaDodaj As String = "Dodanie zapisu"
    Public LOG_OperacjaEdytuj As String = "Edycja zapisu"



    Public Sub DopiszDoLoguAktywnosci(ByVal parTemat As String,
                                       ByVal parKatalog As String,
                                       ByVal parDataZapisu As String,
                                       ByVal parPracownik As String,
                                       ByVal parOperacja As String,
                                       ByVal parUwagi As String)

        Dim sqlConnection As SqlClient.SqlConnection
        Dim sqlCommand As SqlClient.SqlCommand

        Dim Wynik As Integer = 0
        Dim Kwerenda As String = "INSERT INTO LogAktywnosci " &
                                   "(UNIQID " &
                                   ",TEMAT " &
                                   ",KATALOG " &
                                   ",DATAZAPISU " &
                                   ",PRACOWNIK " &
                                   ",OPERACJA " &
                                   ",UWAGI " &
                                   ",WERSJA) " &
                                "VALUES " &
                                   "('" & Guid.NewGuid().ToString & "' " &
                                   ",'" & parTemat & "' " &
                                   ",'" & parKatalog & "' " &
                                   ",'" & parDataZapisu & "' " &
                                   ",'" & parPracownik & "' " &
                                   ",'" & parOperacja & "' " &
                                   ",'" & parUwagi & "' " &
                                   ",1)"

        If LOG_Aktywny Then
            sqlConnection = New SqlClient.SqlConnection(DataBaseConnection())
            sqlCommand = New SqlClient.SqlCommand
            Try
                sqlCommand.Connection = sqlConnection
                sqlCommand.CommandType = CommandType.Text
                sqlCommand.CommandText = Kwerenda
                sqlConnection.Open()
                Wynik = sqlCommand.ExecuteNonQuery()
                If Wynik = -1 Then
                End If
                'txtRaport.Text &= "OK, aktualizacja wykonana poprawnie" & vbCrLf

            Catch ex As Exception
                'MessageBox.Show(ex.Message)
            Finally
                sqlConnection.Close()
            End Try
        End If
    End Sub





End Module
