Module modStrukturySlowniki





    ' ###### ###      #####  ###   ### ###  ### ### ### ### ### 
    '###     #####   ### ### ###   ### #### ### ### ### ### ###
    ' #####  ####    ### ### ### # ### ######## ### ######  ###
    '    ### ###     ### ### ######### ### #### ### ### ### ###
    '######  #######  #####  ##     ## ###  ### ### ### ### ###

    Public KatUzytkownicyMDB As String = oPlikKartoteki     'Katalog Towarów
    Public KatUzytkownicyBD As String = oBazaDanych
    Public KatUzytkowanicyTabela As String = "Uzytkownicy"









    '----------------------------------------------------------
    'Katalog Uzytkownicy
    '----------------------------------------------------------
    Public oIdentUzytkownika As String             'IDENT  Text    10
    Public oNazwaUzytkownika As String             'NAZWA             Text    100
    Public oFunkcjaUzytkownika As String           'FUNKCJA           Text    100
    Public oHasloUzytkownika As String             'HASLO             Text    100
    Public oUprawnieniaUzytkownika As String       'UPRAWNIENIA       Text    100
    Public oUwagiUzytkownika As String   'UWAGI          Text
    Public oWersjaUzytkownika As Integer 'WERSJA         Integer


    Public sConnectionUzytkownicy As String = ConnectionStrings()
    Public HQConnectionUzytkownicy As String = HQConnectionStrings()

    Public sSelectSQLUzytkownicy As String = "SELECT " &
                                                    "IDENT, " &
                                                    "NAZWA, " &
                                                    "FUNKCJA, " &
                                                    "HASLO, " &
                                                    "UPRAWNIENIA, " &
                                                    "UWAGI, " &
                                                    "WERSJA " &
                                                    "FROM Uzytkownicy"

    Public sDeleteSQLUzytkownicy As String = "DELETE FROM Uzytkownicy " &
                                        "WHERE (IDENT=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sUpdateSQLUzytkownicy As String = "UPDATE Uzytkownicy SET " &
                                                 "NAZWA=@oNazwaUzytkownika, " &
                                                 "FUNKCJA=@oFUNKCJAUzytkownika, " &
                                                 "HASLO=@oHASLOUzytkownika, " &
                                                 "UPRAWNIENIA=@oUPRAWNIENIAUzytkownika, " &
                                                 "UWAGI=@oUWAGIUzytkownika, " &
                                                 "WERSJA=@oWersjaUzytkownika " &
                                        "WHERE (IDENT=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sInsertSQLUzytkownicy As String = "INSERT INTO Uzytkownicy " &
                                         "(IDENT, " &
                                         "NAZWA, " &
                                         "FUNKCJA, " &
                                         "HASLO, " &
                                         "UPRAWNIENIA, " &
                                         "UWAGI, " &
                                         "WERSJA " &
                                         ") " &
                                "VALUES  (@oIdentUzytkownika, " &
                                         "@oNazwaUzytkownika, " &
                                         "@oFUNKCJAUzytkownika, " &
                                         "@oHASLOUzytkownika, " &
                                         "@oUPRAWNIENIAUzytkownika, " &
                                         "@oUwagiUzytkownika, " &
                                         "@oWersjaUzytkownika " &
                                         ")"

    'SQLDeleteUzytkownicy(sqlCommandDelete, sqlDataAdapter)
    'SQLUpdateUzytkownicy(sqlCommandUpdate, sqlDataAdapter)
    'SQLInsertUzytkownicy(sqlCommandInsert, sqlDataAdapter)

    Public Sub SQLDeleteUzytkownicy(ByRef sqlCommandDelete As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter

        'parametry filtrowania
        sqlParam = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 10, "IDENT")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original

        sqlDataAdapter.DeleteCommand = sqlCommandDelete
    End Sub

    Public Sub SQLUpdateUzytkownicy(ByRef sqlCommandUpdate As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter

        'sqlCommandUpdate.Parameters.Add("@oIdentUzytkownika", sqldbtype.varChar, 10, "IDENT")
        sqlCommandUpdate.Parameters.Add("@oNazwaUzytkownika", SqlDbType.VarChar, 100, "NAZWA")
        sqlCommandUpdate.Parameters.Add("@oFunkcjaUzytkownika", SqlDbType.VarChar, 100, "FUNKCJA")
        sqlCommandUpdate.Parameters.Add("@oHasloUzytkownika", SqlDbType.VarChar, 100, "HASLO")
        sqlCommandUpdate.Parameters.Add("@oUprawnieniaUzytkownika", SqlDbType.VarChar, 100, "UPRAWNIENIA")
        sqlCommandUpdate.Parameters.Add("@oUWAGIUzytkownika", SqlDbType.Text, Nothing, "UWAGI")
        sqlCommandUpdate.Parameters.Add("@oWersjaUzytkownika", SqlDbType.Int, Nothing, "WERSJA")

        'parametry filtrowania
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 10, "IDENT")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.UpdateCommand = sqlCommandUpdate
    End Sub

    Public Sub SQLInsertUzytkownicy(ByRef sqlCommandInsert As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter

        sqlParam = sqlCommandInsert.Parameters.Add("@oIdentUzytkownika", SqlDbType.VarChar, 10, "IDENT")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oNazwaUzytkownika", SqlDbType.VarChar, 100, "NAZWA")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oFunkcjaUzytkownika", SqlDbType.VarChar, 100, "FUNKCJA")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oHasloUzytkownika", SqlDbType.VarChar, 100, "HASLO")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oUprawnieniaUzytkownika", SqlDbType.VarChar, 100, "UPRAWNIENIA")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oUWAGIUzytkownika", SqlDbType.Text, Nothing, "UWAGI")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oWersjaUzytkownika", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlDataAdapter.InsertCommand = sqlCommandInsert
    End Sub



    'Public Sub DBDeleteUzytkownicy(ByRef dbCommandDelete As OleDb.OleDbCommand,
    '                       ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As OleDb.OleDbParameter

    '    dbParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 10, "IDENT")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original

    '    dbDataAdapter.DeleteCommand = dbCommandDelete
    'End Sub

    'Public Sub DBUpdateUzytkownicy(ByRef dbCommandUpdate As OleDb.OleDbCommand,
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

    'Public Sub DBInsertUzytkownicy(ByRef dbCommandInsert As OleDb.OleDbCommand,
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
    '** pobierz informacje o UZYTKOWNIKU
    '**************************************
    'Public oIdentUzytkownika As String           'IDENT  Text    10
    'Public oNazwaUzytkownika As String           'NAZWA             Text    100
    'Public oFunkcjaUzytkownika As String         'FUNKCJA           Text    100
    'Public oHasloUzytkownika As String           'HASLO             Text    100
    'Public oUprawnieniaUzytkownika As String     'UPRAWNIENIA          Text    100
    'Public oUwagiUzytkownika As String           'UWAGI          Text
    'Public oWersjaUzytkownika As Integer         'WERSJA         Integer

    Public Sub InitInfoUzytkownika()
        oIdentUzytkownika = ""
        oNazwaUzytkownika = ""
        oFunkcjaUzytkownika = ""
        oHasloUzytkownika = ""
        oUprawnieniaUzytkownika = ""
        oUwagiUzytkownika = ""
        oWersjaUzytkownika = 0
    End Sub


    Public Function IdentUzytkownika(ByVal IdentU As String) As Boolean
        Dim dbSelectSQLUzytkownicy As String = sSelectSQLUzytkownicy & " WHERE IDENT='" & IdentU & "'"

        Dim dbConnectionUzytkownicy As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectUzytkownicy As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterUzytkownicy As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionUzytkownicy As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectUzytkownicy As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterUzytkownicy As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetUzytkownicy As New DataSet
        Dim DataViewUzytkownicy As New DataView
        Dim ConnectionStateUzytkownicy As System.Data.ConnectionState

        Dim Wynik As Boolean = False

        InitInfoUzytkownika()

        DataSetUzytkownicy.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionUzytkownicy = New OleDb.OleDbConnection(sConnectionUzytkownicy)
            'dbCommandSelectUzytkownicy = New OleDb.OleDbCommand(dbSelectSQLUzytkownicy, dbConnectionUzytkownicy)
            'dbDataAdapterUzytkownicy = New OleDb.OleDbDataAdapter(dbCommandSelectUzytkownicy)

            ''mapujemy domyslna nazwe tabeli
            'Dim dbMapowanieTabeliUzytkownicy As System.Data.Common.DataTableMapping
            'dbMapowanieTabeliUzytkownicy = dbDataAdapterUzytkownicy.TableMappings.Add("Table", "TABELA_Uzytkownicy")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionUzytkownicy.Open()
            'Catch Ex As System.Exception
            'Finally
            '    ConnectionStateUzytkownicy = dbConnectionUzytkownicy.State
            'End Try
        Else
            sqlConnectionUzytkownicy = New SqlClient.SqlConnection(sConnectionUzytkownicy)
            sqlCommandSelectUzytkownicy = New SqlClient.SqlCommand(dbSelectSQLUzytkownicy, sqlConnectionUzytkownicy)
            sqlDataAdapterUzytkownicy = New SqlClient.SqlDataAdapter(sqlCommandSelectUzytkownicy)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliUzytkownicy As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliUzytkownicy = sqlDataAdapterUzytkownicy.TableMappings.Add("Table", "TABELA_Uzytkownicy")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionUzytkownicy.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateUzytkownicy = sqlConnectionUzytkownicy.State
            End Try
        End If

        If ConnectionStateUzytkownicy = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterUzytkownicy.FillSchema(DataSetUzytkownicy, SchemaType.Mapped, "TABELA_Uzytkownicy")
                    'dbDataAdapterUzytkownicy.Fill(DataSetUzytkownicy)
                    'dbConnectionUzytkownicy.Close()
                Else
                    sqlDataAdapterUzytkownicy.FillSchema(DataSetUzytkownicy, SchemaType.Mapped, "TABELA_Uzytkownicy")
                    sqlDataAdapterUzytkownicy.Fill(DataSetUzytkownicy)
                    sqlConnectionUzytkownicy.Close()
                End If

                'definiuj DataView
                DataViewUzytkownicy = New DataView(DataSetUzytkownicy.Tables(0))
                If DataViewUzytkownicy.Count > 0 Then
                    oIdentUzytkownika = IdentU
                    oNazwaUzytkownika = GetTxtDRVField(DataViewUzytkownicy.Item(0), "NAZWA")
                    oFunkcjaUzytkownika = GetTxtDRVField(DataViewUzytkownicy.Item(0), "FUNKCJA")
                    oHasloUzytkownika = GetTxtDRVField(DataViewUzytkownicy.Item(0), "HASLO")
                    oUprawnieniaUzytkownika = GetTxtDRVField(DataViewUzytkownicy.Item(0), "UPRAWNIENIA")
                    oUwagiUzytkownika = GetTxtDRVField(DataViewUzytkownicy.Item(0), "UWAGI")
                    oWersjaUzytkownika = GetIntDRVField(DataViewUzytkownicy.Item(0), "WERSJA")

                    Wynik = True
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wynik)
    End Function

















    '---------------------------------------------------------------------
    'Kontakty HANDLOWE - po staremu - do 05.09.2015
    '-----------------------------------------
    'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
    'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
    'Public oPracownikKon As String        'PRACOWNIK        Text, 10     uzytkownik
    'Public oTematKon As String            'TEMAT            Text, 50
    'Public oUwagiKon As String            'UWAGI            Memo
    'Public oWersjaKon As Integer          'WERSJA           Integer

    'Public sConnectionKontakty As String = ConnectionStrings()

    Public sSelectSQLKontaktyOld As String = "SELECT IDENTKLIENTA, " &
                                                 "DATAKONTAKTU, " &
                                                 "PRACOWNIK, " &
                                                 "TEMAT, " &
                                                 "UWAGI, " &
                                                 "WERSJA " &
                                          "FROM Kontakty"

    Public sUpdateSQLKontaktyOld As String = "UPDATE Kontakty SET " &
                                             "UNIQID=@oUniqIDKon," &
                                             "PRACOWNIK=@oPracownikKon," &
                                             "TEMAT=@oTematkon," &
                                             "UWAGI=@oUwagiKon," &
                                             "WERSJA=@oWersjaKon " &
                                    "WHERE (IDENTKLIENTA=@orygSymbol) AND " &
                                          "(DATAKONTAKTU=@orygData) AND " &
                                          "(WERSJA=@orygWersja)"

    Public sDeleteSQLKontaktyOld As String = "DELETE FROM Kontakty " &
                                    "WHERE (IDENTKLIENTA=@orygSymbol) AND " &
                                          "(DATAKONTAKTU=@orygData) AND " &
                                          "(WERSJA=@orygWersja)"


    Public Sub SQLDeleteKontaktyOld(ByRef sqlCommandDelete As SqlClient.SqlCommand,
                           ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda DELETE
        'parametry filtrowania
        sqlParam = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 6, "IDENTKLIENTA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygData", SqlDbType.VarChar, 10, "DATAKONTAKTU")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.DeleteCommand = sqlCommandDelete
    End Sub

    'Public Sub DBDeleteKontaktyOld(ByRef dbCommandDelete As OleDb.OleDbCommand,
    '                   ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda DELETE
    '    'parametry filtrowania
    '    dbParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 6, "IDENTKLIENTA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandDelete.Parameters.Add("@orygData", OleDb.OleDbType.VarChar, 10, "DATAKONTAKTU")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.DeleteCommand = dbCommandDelete
    'End Sub




    '---------------------------------------------------------------------
    'Kontakty HANDLOWE - nowe 05.09.2015
    '-----------------------------------------
    Public oUniqIdKon As String           'UNIQID            varchar(40)
    Public oIdentKon As String            'IDENTKLIENTA     Text, 6
    Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
    Public oPracownikKon As String        'PRACOWNIK        Text, 10     uzytkownik
    Public oTematKon As String            'TEMAT            Text, 50
    Public oUwagiKon As String            'UWAGI            Memo
    Public oWersjaKon As Integer          'WERSJA           Integer

    Public sConnectionKontakty As String = ConnectionStrings()
    Public HQConnectionKontakty As String = HQConnectionStrings()

    Public sSelectSQLKontakty As String = "SELECT UNIQID, " &
                                                 "IDENTKLIENTA, " &
                                                 "DATAKONTAKTU, " &
                                                 "PRACOWNIK, " &
                                                 "TEMAT, " &
                                                 "UWAGI, " &
                                                 "WERSJA " &
                                          "FROM KontaktyHandlowe"

    Public sDeleteSQLKontakty As String = "DELETE FROM KontaktyHandlowe " &
                                        "WHERE (UNIQID=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sUpdateSQLKontakty As String = "UPDATE KontaktyHandlowe SET " &
                                                 "IDENTKLIENTA=@oIdentKon," &
                                                 "DATAKONTAKTU=@oDataKon," &
                                                 "PRACOWNIK=@oPracownikKon," &
                                                 "TEMAT=@oTematkon," &
                                                 "UWAGI=@oUwagiKon," &
                                                 "WERSJA=@oWersjaKon " &
                                        "WHERE (UNIQID=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sInsertSQLKontakty As String = "INSERT INTO KontaktyHandlowe " &
                                         "(UNIQID, " &
                                           "IDENTKLIENTA, " &
                                           "DATAKONTAKTU," &
                                           "PRACOWNIK," &
                                           "TEMAT," &
                                           "UWAGI," &
                                           "WERSJA " &
                                           ") " &
                                "VALUES  (@oUniqIdKon," &
                                         "@oIdentKon," &
                                         "@oDataKon," &
                                         "@oPracownikKon," &
                                         "@oTematKon," &
                                         "@oUwagiKon," &
                                         "@oWersjaKon " &
                                         ")"


    'SQLDeleteKontaktyHandlowe(sqlCommandDelete, sqlDataAdapter)
    'SQLUpdateKontaktyHandlowe(sqlCommandUpdate, sqlDataAdapter)
    'SQLInsertKontaktyHandlowe(sqlCommandInsert, sqlDataAdapter)


    Public Sub SQLDeleteKontaktyHandlowe(ByRef sqlCommandDelete As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda DELETE
        'parametry filtrowania
        sqlParam = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 40, "UNIQID")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.DeleteCommand = sqlCommandDelete
    End Sub

    Public Sub SQLUpdateKontaktyHandlowe(ByRef sqlCommandUpdate As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda UPDATE
        'parametry aktualizacji
        'sqlCommandUpdate.Parameters.Add("@oUniqIdKon", SqlDbType.VarChar, 40, "UNIQID")
        sqlCommandUpdate.Parameters.Add("@oIdentKon", SqlDbType.VarChar, 6, "IDENTKLIENTA")
        sqlCommandUpdate.Parameters.Add("@oDataKon", SqlDbType.VarChar, 10, "DATAKONTAKTU")
        sqlCommandUpdate.Parameters.Add("@oPracownikKon", SqlDbType.VarChar, 10, "PRACOWNIK")
        sqlCommandUpdate.Parameters.Add("@oTematKon", SqlDbType.VarChar, 50, "TEMAT")
        sqlCommandUpdate.Parameters.Add("@oUwagiKon", SqlDbType.Text, Nothing, "UWAGI")
        sqlCommandUpdate.Parameters.Add("@oWersjaKon", SqlDbType.Int, Nothing, "WERSJA")

        'parametry filtrowania
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 40, "UNIQID")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.UpdateCommand = sqlCommandUpdate
    End Sub


    Public Sub SQLInsertKontaktyHandlowe(ByRef sqlCommandInsert As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda INSERT
        sqlParam = sqlCommandInsert.Parameters.Add("@oUniqIdKon", SqlDbType.VarChar, 40, "UNIQID")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oIdentKon", SqlDbType.VarChar, 6, "IDENTKLIENTA")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oDataKon", SqlDbType.VarChar, 10, "DATAKONTAKTU")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oPracownikKon", SqlDbType.VarChar, 10, "PRACOWNIK")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oTematKon", SqlDbType.VarChar, 50, "TEMAT")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oUwagiKon", SqlDbType.Text, Nothing, "UWAGI")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oWersjaKon", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlDataAdapter.InsertCommand = sqlCommandInsert
    End Sub



    'Public Sub DBDeleteKontaktyHandlowe(ByRef dbCommandDelete As OleDb.OleDbCommand,
    '                       ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda DELETE
    '    'parametry filtrowania
    '    dbParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.DeleteCommand = dbCommandDelete
    'End Sub

    'Public Sub DBUpdateKontaktyHandlowe(ByRef dbCommandUpdate As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter

    '    '------------------------------------------
    '    'komenda UPDATE
    '    'parametry aktualizacji
    '    'dbCommandUpdate.Parameters.Add("@oUniqIdKon", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbCommandUpdate.Parameters.Add("@oIdentKlientaKon", OleDb.OleDbType.VarChar, 6, "IDENTKLIENTA")
    '    dbCommandUpdate.Parameters.Add("@oDataKon", OleDb.OleDbType.VarChar, 10, "DATAKONTAKTU")
    '    dbCommandUpdate.Parameters.Add("@oPracownikKon", OleDb.OleDbType.VarChar, 10, "PRACOWNIK")
    '    dbCommandUpdate.Parameters.Add("@oTematKon", OleDb.OleDbType.VarChar, 50, "TEMAT")
    '    dbCommandUpdate.Parameters.Add("@oUwagiKon", OleDb.OleDbType.WChar, Nothing, "UWAGI")
    '    dbCommandUpdate.Parameters.Add("@oWersjaKon", OleDb.OleDbType.Integer, Nothing, "WERSJA")

    '    'parametry filtrowania
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.UpdateCommand = dbCommandUpdate
    'End Sub

    'Public Sub DBInsertKontaktyHandlowe(ByRef dbCommandInsert As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda INSERT
    '    dbParam = dbCommandInsert.Parameters.Add("@oUniqIdKon", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oIdentKlientaKon", OleDb.OleDbType.VarChar, 6, "IDENTKLIENTA")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oDataKon", OleDb.OleDbType.VarChar, 10, "DATAKONTAKTU")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oPracownikKon", OleDb.OleDbType.VarChar, 10, "PRACOWNIK")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oTematKon", OleDb.OleDbType.VarChar, 50, "TEMAT")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oUwagiKon", OleDb.OleDbType.WChar, Nothing, "UWAGI")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oWersjaKon", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Current

    '    dbDataAdapter.InsertCommand = dbCommandInsert
    'End Sub



    Public Function IdentOstatniKontakt(ByVal IdentK As String) As Boolean
        Dim dbSelectSQLKontaktyHandlowe As String = sSelectSQLKontakty & " WHERE IDENTKLIENTA='" & IdentK & "' ORDER BY DATAKONTAKTU DESC"

        Dim dbConnectionKontaktyHandlowe As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectKontaktyHandlowe As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterKontaktyHandlowe As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionKontaktyHandlowe As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectKontaktyHandlowe As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterKontaktyHandlowe As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetKontaktyHandlowe As New DataSet
        Dim DataViewKontaktyHandlowe As New DataView
        Dim ConnectionStateKontaktyHandlowe As System.Data.ConnectionState

        Dim Wynik As Boolean = False

        oUniqIdKon = ""
        oIdentKon = ""
        oDataKon = ""
        oPracownikKon = ""
        oTematKon = ""
        oUwagiKon = ""
        oWersjaKon = 1

        DataSetKontaktyHandlowe.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionKontaktyHandlowe = New OleDb.OleDbConnection(sConnectionKontakty)
            'dbCommandSelectKontaktyHandlowe = New OleDb.OleDbCommand(dbSelectSQLKontaktyHandlowe, dbConnectionKontaktyHandlowe)
            'dbDataAdapterKontaktyHandlowe = New OleDb.OleDbDataAdapter(dbCommandSelectKontaktyHandlowe)

            ''mapujemy domyslna nazwe tabeli
            'Dim dbMapowanieTabeliKontaktyHandlowe As System.Data.Common.DataTableMapping
            'dbMapowanieTabeliKontaktyHandlowe = dbDataAdapterKontaktyHandlowe.TableMappings.Add("Table", "TABELA_KontaktyHandlowe")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionKontaktyHandlowe.Open()
            'Catch Ex As System.Exception
            'Finally
            '    ConnectionStateKontaktyHandlowe = dbConnectionKontaktyHandlowe.State
            'End Try
        Else
            sqlConnectionKontaktyHandlowe = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectKontaktyHandlowe = New SqlClient.SqlCommand(dbSelectSQLKontaktyHandlowe, sqlConnectionKontaktyHandlowe)
            sqlDataAdapterKontaktyHandlowe = New SqlClient.SqlDataAdapter(sqlCommandSelectKontaktyHandlowe)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliKontaktyHandlowe As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliKontaktyHandlowe = sqlDataAdapterKontaktyHandlowe.TableMappings.Add("Table", "TABELA_KontaktyHandlowe")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKontaktyHandlowe.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateKontaktyHandlowe = sqlConnectionKontaktyHandlowe.State
            End Try
        End If

        If ConnectionStateKontaktyHandlowe = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterKontaktyHandlowe.FillSchema(DataSetKontaktyHandlowe, SchemaType.Mapped, "TABELA_KontaktyHandlowe")
                    'dbDataAdapterKontaktyHandlowe.Fill(DataSetKontaktyHandlowe)
                    'dbConnectionKontaktyHandlowe.Close()
                Else
                    sqlDataAdapterKontaktyHandlowe.FillSchema(DataSetKontaktyHandlowe, SchemaType.Mapped, "TABELA_KontaktyHandlowe")
                    sqlDataAdapterKontaktyHandlowe.Fill(DataSetKontaktyHandlowe)
                    sqlConnectionKontaktyHandlowe.Close()
                End If

                'definiuj DataView
                DataViewKontaktyHandlowe = New DataView(DataSetKontaktyHandlowe.Tables(0))
                If DataViewKontaktyHandlowe.Count > 0 Then
                    oUniqIdKon = GetTxtDRVField(DataViewKontaktyHandlowe.Item(0), "UNIQID")
                    oIdentKon = GetTxtDRVField(DataViewKontaktyHandlowe.Item(0), "IDENTKLIENTA")
                    oDataKon = GetTxtDRVField(DataViewKontaktyHandlowe.Item(0), "DATAKONTAKTU")
                    oPracownikKon = GetTxtDRVField(DataViewKontaktyHandlowe.Item(0), "PRACOWNIK")
                    oTematKon = GetTxtDRVField(DataViewKontaktyHandlowe.Item(0), "TEMAT")
                    oUwagiKon = GetTxtDRVField(DataViewKontaktyHandlowe.Item(0), "UWAGI")
                    oWersjaKon = GetIntDRVField(DataViewKontaktyHandlowe.Item(0), "WERSJA")

                    Wynik = True
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wynik)
    End Function
















    '---------------------------------------------------------------------
    'Akcje Handlowe - Opis
    Public aoIdentAkcji As String            'IDENT     Text 20
    Public aoDataAkcji As String             'DATA      Text,10
    Public aoOpisAkcji As String             'OPIS      Text,100
    Public aoUwagiAkcji As String            'UWAGI     Text,10
    Public aoMarkerAkcji As Boolean           'MARKER    logical
    Public aoWersjaAkcji As Integer           'WERSJA    Integer

    Public sConnectionAkcjeOpis As String = ConnectionStrings()
    Public HQConnectionAkcjeOpis As String = HQConnectionStrings()

    Public sSelectSQLAkcjeOpis As String = "SELECT IDENT, " &
                                                    "DATA, " &
                                                    "OPIS, " &
                                                    "UWAGI, " &
                                                    "MARKER, " &
                                                    "WERSJA " &
                                            "FROM AkcjeOpis"

    Public sDeleteSQLAkcjeOpis As String = "DELETE FROM AkcjeOpis " &
                                        "WHERE (IDENT=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sUpdateSQLAkcjeOpis As String = "UPDATE AkcjeOpis SET " &
                                                 "DATA=@oDataAkcji," &
                                                 "OPIS=@oOpisAkcji," &
                                                 "UWAGI=@oUwagiAkcji," &
                                                 "MARKER=@oMarkerAkcji," &
                                                 "WERSJA=@oWersjaAkcji " &
                                        "WHERE (IDENT=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sInsertSQLAkcjeOpis As String = "INSERT INTO AkcjeOpis (" &
                                                 "IDENT," &
                                                 "DATA," &
                                                 "OPIS," &
                                                 "UWAGI," &
                                                 "MARKER," &
                                                 "WERSJA " &
                                                 ") " &
                                        "VALUES  (" &
                                                 "@oIdentAkcji," &
                                                 "@oDataAkcji," &
                                                 "@oOpisAkcji," &
                                                 "@oUwagiAkcji," &
                                                 "@oMarkerAkcji," &
                                                 "@oWersjaAkcji " &
                                                 ")"


    'SQLDeleteAkcjeOpis(sqlCommandDelete, sqlDataAdapter)
    'SQLUpdateAkcjeOpis(sqlCommandUpdate, sqlDataAdapter)
    'SQLInsertAkcjeOpis(sqlCommandInsert, sqlDataAdapter)


    Public Sub SQLDeleteAkcjeOpis(ByRef sqlCommandDelete As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda DELETE
        'parametry filtrowania
        sqlParam = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.Char, 20, "IDENT")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.DeleteCommand = sqlCommandDelete
    End Sub

    Public Sub SQLUpdateAkcjeOpis(ByRef sqlCommandUpdate As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda UPDATE
        'parametry aktualizacji
        'SQLCommandUpdate.Parameters.Add("@oIdentAkcji", SqlDbType.Char, 20, "IDENT")
        sqlCommandUpdate.Parameters.Add("@oDataAkcji", SqlDbType.Char, 10, "DATA")
        sqlCommandUpdate.Parameters.Add("@oOpisAkcji", SqlDbType.VarChar, 100, "OPIS")
        sqlCommandUpdate.Parameters.Add("@oUwagiAkcji", SqlDbType.Text, Nothing, "UWAGI")
        sqlCommandUpdate.Parameters.Add("@oMarkerAkcji", SqlDbType.Bit, Nothing, "MARKER")
        sqlCommandUpdate.Parameters.Add("@oWersjaAkcji", SqlDbType.Int, Nothing, "WERSJA")

        'parametry filtrowania
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygSymbol", SqlDbType.Char, 20, "IDENT")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.UpdateCommand = sqlCommandUpdate
    End Sub


    Public Sub SQLInsertAkcjeOpis(ByRef sqlCommandInsert As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda INSERT
        sqlParam = sqlCommandInsert.Parameters.Add("@oIdentAkcji", SqlDbType.Char, 20, "IDENT")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oDataAkcji", SqlDbType.Char, 10, "DATA")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oOpisAkcji", SqlDbType.VarChar, 100, "OPIS")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oUwagiAkcji", SqlDbType.Text, Nothing, "UWAGI")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oMarkerAkcji", SqlDbType.Bit, Nothing, "MARKER")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oWersjaAkcji", SqlDbType.Char, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlDataAdapter.InsertCommand = sqlCommandInsert
    End Sub



    'Public Sub DBDeleteAkcjeOpis(ByRef dbCommandDelete As OleDb.OleDbCommand,
    '                       ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda DELETE
    '    'parametry filtrowania
    '    dbParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 20, "IDENT")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.DeleteCommand = dbCommandDelete
    'End Sub

    'Public Sub DBUpdateAkcjeOpis(ByRef dbCommandUpdate As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda UPDATE
    '    'parametry aktualizacji
    '    'dbCommandUpdate.Parameters.Add("@oIdentAkcji", OleDb.OleDbType.Char, 20, "IDENT")
    '    dbCommandUpdate.Parameters.Add("@oDataAkcji", OleDb.OleDbType.Char, 10, "DATA")
    '    dbCommandUpdate.Parameters.Add("@oOpisAkcji", OleDb.OleDbType.Char, 100, "OPIS")
    '    dbCommandUpdate.Parameters.Add("@oUwagiAkcji", OleDb.OleDbType.WChar, Nothing, "UWAGI")
    '    dbCommandUpdate.Parameters.Add("@oMarkerAkcji", OleDb.OleDbType.Boolean, Nothing, "MARKER")
    '    dbCommandUpdate.Parameters.Add("@oWersjaAkcji", OleDb.OleDbType.Integer, Nothing, "WERSJA")

    '    'parametry filtrowania
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 20, "IDENT")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.UpdateCommand = dbCommandUpdate
    'End Sub

    'Public Sub DBInsertAkcjeOpis(ByRef dbCommandInsert As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda INSERT
    '    dbParam = dbCommandInsert.Parameters.Add("@oIdentAkcji", OleDb.OleDbType.Char, 20, "IDENT")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oDataAkcji", OleDb.OleDbType.Char, 10, "DATA")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oOpisAkcji", OleDb.OleDbType.Char, 100, "OPIS")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oUwagiAkcji", OleDb.OleDbType.WChar, Nothing, "UWAGI")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oMarkerAkcji", OleDb.OleDbType.Boolean, Nothing, "MARKER")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oWersjaAkcji", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Current

    '    dbDataAdapter.InsertCommand = dbCommandInsert
    'End Sub



    Public Function IdentAkcjeOpis(ByVal IdentK As String) As Boolean
        Dim dbSelectSQLAkcjeOpis As String = sSelectSQLAkcjeOpis & " WHERE IDENT='" & IdentK & "' ORDER BY DATA DESC"

        Dim dbConnectionAkcjeOpis As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectAkcjeOpis As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterAkcjeOpis As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionAkcjeOpis As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectAkcjeOpis As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterAkcjeOpis As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetAkcjeOpis As New DataSet
        Dim DataViewAkcjeOpis As New DataView
        Dim ConnectionStateAkcjeOpis As System.Data.ConnectionState

        Dim Wynik As Boolean = False

        aoIdentAkcji = ""        'IDENT       Text, 10
        aoDataAkcji = SysData()
        aoOpisAkcji = ""
        aoUwagiAkcji = ""
        aoMarkerAkcji = False
        aoWersjaAkcji = 1        'WERSJA      Integer, 2

        DataSetAkcjeOpis.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionAkcjeOpis = New OleDb.OleDbConnection(sConnectionKontakty)
            'dbCommandSelectAkcjeOpis = New OleDb.OleDbCommand(dbSelectSQLAkcjeOpis, dbConnectionAkcjeOpis)
            'dbDataAdapterAkcjeOpis = New OleDb.OleDbDataAdapter(dbCommandSelectAkcjeOpis)

            ''mapujemy domyslna nazwe tabeli
            'Dim dbMapowanieTabeliAkcjeOpis As System.Data.Common.DataTableMapping
            'dbMapowanieTabeliAkcjeOpis = dbDataAdapterAkcjeOpis.TableMappings.Add("Table", "TABELA_AkcjeOpis")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionAkcjeOpis.Open()
            'Catch Ex As System.Exception
            'Finally
            '    ConnectionStateAkcjeOpis = dbConnectionAkcjeOpis.State
            'End Try
        Else
            sqlConnectionAkcjeOpis = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectAkcjeOpis = New SqlClient.SqlCommand(dbSelectSQLAkcjeOpis, sqlConnectionAkcjeOpis)
            sqlDataAdapterAkcjeOpis = New SqlClient.SqlDataAdapter(sqlCommandSelectAkcjeOpis)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliAkcjeOpis As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliAkcjeOpis = sqlDataAdapterAkcjeOpis.TableMappings.Add("Table", "TABELA_AkcjeOpis")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionAkcjeOpis.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateAkcjeOpis = sqlConnectionAkcjeOpis.State
            End Try
        End If

        If ConnectionStateAkcjeOpis = ConnectionState.Open Then
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

                'definiuj DataView
                DataViewAkcjeOpis = New DataView(DataSetAkcjeOpis.Tables(0))
                If DataViewAkcjeOpis.Count > 0 Then
                    aoIdentAkcji = GetTxtDRVField(DataViewAkcjeOpis.Item(0), "IDENT")
                    aoDataAkcji = GetTxtDRVField(DataViewAkcjeOpis.Item(0), "DATA")
                    aoOpisAkcji = GetTxtDRVField(DataViewAkcjeOpis.Item(0), "OPIS")
                    aoUwagiAkcji = GetTxtDRVField(DataViewAkcjeOpis.Item(0), "UWAGI")
                    aoMarkerAkcji = GetLogDRVField(DataViewAkcjeOpis.Item(0), "MARKER")
                    aoWersjaAkcji = GetIntDRVField(DataViewAkcjeOpis.Item(0), "WERSJA")

                    Wynik = True
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wynik)
    End Function











    '---------------------------------------------------------------------
    'Akcje Handlowe - Specyfikacja
    Public asIdentAkcji As String            'IDENTAKCJI     Text 20
    Public asIdentKlienta As String          'IDENTKLI       Text 6
    Public asWersjaAkcji As Integer           'WERSJA    Integer

    Public sConnectionAkcjeSpec As String = ConnectionStrings()
    Public HQConnectionAkcjeSpec As String = HQConnectionStrings()

    Public sSelectSQLAkcjeSpec As String = "SELECT IDENTAKCJI, " &
                                                   "IDENTKLI, " &
                                                   "WERSJA " &
                                                   "FROM AkcjeSpec"

    Public sDeleteSQLAkcjeSpec As String = "DELETE FROM AkcjeSpec " &
                                        "WHERE (IDENTAKCJI=@orygSymbol) AND " &
                                              "(IDENTKLI=@orygKlient) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sUpdateSQLAkcjeSpec As String = "UPDATE AkcjeSpec SET " &
                                                 "WERSJA=@oWersjaAkcji " &
                                        "WHERE (IDENTAKCJI=@orygSymbol) AND " &
                                              "(IDENTKLI=@orygKlient) AND " &
                                              "(WERSJA=@orygWersja)"


    Public sInsertSQLAkcjeSpec As String = "INSERT INTO AkcjeSpec (" &
                                                 "IDENTAKCJI," &
                                                 "IDENTKLI," &
                                                 "WERSJA " &
                                                 ") " &
                                        "VALUES  (" &
                                                 "@oIdentAkcji," &
                                                 "@oIdentkli," &
                                                 "@oWersjaAkcji " &
                                                 ")"


    'SQLDeleteAkcjeSpec(sqlCommandDelete, sqlDataAdapter)
    'SQLUpdateAkcjeSpec(sqlCommandUpdate, sqlDataAdapter)
    'SQLInsertAkcjeSpec(sqlCommandInsert, sqlDataAdapter)


    Public Sub SQLDeleteAkcjeSpec(ByRef sqlCommandDelete As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda DELETE
        'parametry filtrowania
        sqlParam = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.Char, 20, "IDENTAKCJI")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygKlient", SqlDbType.Char, 6, "IDENTKLI")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.DeleteCommand = sqlCommandDelete
    End Sub

    Public Sub SQLUpdateAkcjeSpec(ByRef sqlCommandUpdate As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda UPDATE
        'parametry aktualizacji
        'sqlCommandUpdate.Parameters.Add("@oIdentAkcji", SqlDbType.Char, 20, "IDENTAKCJI")
        'sqlCommandUpdate.Parameters.Add("@oIdentKli", SqlDbType.Char, 6, "IDENTKLI")
        sqlCommandUpdate.Parameters.Add("@oWersjaAkcji", SqlDbType.Int, Nothing, "WERSJA")

        'parametry filtrowania
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygSymbol", SqlDbType.Char, 20, "IDENTAKCJI")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygKlient", SqlDbType.Char, 6, "IDENTKLI")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.UpdateCommand = sqlCommandUpdate
    End Sub


    Public Sub SQLInsertAkcjeSpec(ByRef sqlCommandInsert As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda INSERT
        sqlParam = sqlCommandInsert.Parameters.Add("@oIdentAkcji", SqlDbType.Char, 20, "IDENTAKCJI")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oIdentKli", SqlDbType.Char, 6, "IDENTKLI")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oWersjaAkcji", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlDataAdapter.InsertCommand = sqlCommandInsert
    End Sub



    'Public Sub DBDeleteAkcjeSpec(ByRef dbCommandDelete As OleDb.OleDbCommand,
    '                       ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda DELETE
    '    'parametry filtrowania
    '    dbParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 20, "IDENTAKCJI")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandDelete.Parameters.Add("@orygKlient", OleDb.OleDbType.Char, 6, "IDENTKLI")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.DeleteCommand = dbCommandDelete
    'End Sub

    'Public Sub DBUpdateAkcjeSpec(ByRef dbCommandUpdate As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda UPDATE
    '    'parametry aktualizacji
    '    'dbCommandUpdate.Parameters.Add("@oIdentAkcji", OleDb.OleDbType.Char, 20, "IDENTAKCJI")
    '    'dbCommandUpdate.Parameters.Add("@oIdentKli", OleDb.OleDbType.Char, 6, "IDENTKLI")
    '    dbCommandUpdate.Parameters.Add("@oWersjaAkcji", OleDb.OleDbType.Integer, Nothing, "WERSJA")

    '    'parametry filtrowania
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 20, "IDENTAKCJI")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygKlient", OleDb.OleDbType.Char, 6, "IDENTKLI")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.UpdateCommand = dbCommandUpdate
    'End Sub

    'Public Sub DBInsertAkcjeSpec(ByRef dbCommandInsert As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda INSERT
    '    dbParam = dbCommandInsert.Parameters.Add("@oIdentAkcji", OleDb.OleDbType.Char, 20, "IDENTAKCJI")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oIdentKli", OleDb.OleDbType.Char, 6, "IDENTKLI")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oWersjaAkcji", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Current

    '    dbDataAdapter.InsertCommand = dbCommandInsert
    'End Sub



    Public Function IdentAkcjeSpec(ByVal IdentK As String, ByVal IdentA As String) As Boolean
        Dim dbSelectSQLAkcjeSpec As String = sSelectSQLAkcjeSpec & " WHERE IDENTKLI='" & IdentK & "' AND IDENTAKCJI='" & IdentA & "' ORDER BY DATA DESC"

        Dim dbConnectionAkcjeSpec As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectAkcjeSpec As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterAkcjeSpec As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionAkcjeSpec As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectAkcjeSpec As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterAkcjeSpec As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetAkcjeSpec As New DataSet
        Dim DataViewAkcjeSpec As New DataView
        Dim ConnectionStateAkcjeSpec As System.Data.ConnectionState

        Dim Wynik As Boolean = False

        asIdentAkcji = ""        'IDENT       Text, 10
        asIdentKlienta = ""        'IDENT       Text, 10
        asWersjaAkcji = 1        'WERSJA      Integer, 2

        DataSetAkcjeSpec.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionAkcjeSpec = New OleDb.OleDbConnection(sConnectionKontakty)
            'dbCommandSelectAkcjeSpec = New OleDb.OleDbCommand(dbSelectSQLAkcjeSpec, dbConnectionAkcjeSpec)
            'dbDataAdapterAkcjeSpec = New OleDb.OleDbDataAdapter(dbCommandSelectAkcjeSpec)

            ''mapujemy domyslna nazwe tabeli
            'Dim dbMapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
            'dbMapowanieTabeliAkcjeSpec = dbDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_AkcjeSpec")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionAkcjeSpec.Open()
            'Catch Ex As System.Exception
            'Finally
            '    ConnectionStateAkcjeSpec = dbConnectionAkcjeSpec.State
            'End Try
        Else
            sqlConnectionAkcjeSpec = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectAkcjeSpec = New SqlClient.SqlCommand(dbSelectSQLAkcjeSpec, sqlConnectionAkcjeSpec)
            sqlDataAdapterAkcjeSpec = New SqlClient.SqlDataAdapter(sqlCommandSelectAkcjeSpec)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliAkcjeSpec = sqlDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_AkcjeSpec")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionAkcjeSpec.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateAkcjeSpec = sqlConnectionAkcjeSpec.State
            End Try
        End If

        If ConnectionStateAkcjeSpec = ConnectionState.Open Then
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

                'definiuj DataView
                DataViewAkcjeSpec = New DataView(DataSetAkcjeSpec.Tables(0))
                If DataViewAkcjeSpec.Count > 0 Then
                    asIdentAkcji = GetTxtDRVField(DataViewAkcjeSpec.Item(0), "IDENTAKCJI")
                    asIdentKlienta = GetTxtDRVField(DataViewAkcjeSpec.Item(0), "IDENTKLI")
                    asWersjaAkcji = GetIntDRVField(DataViewAkcjeSpec.Item(0), "WERSJA")

                    Wynik = True
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wynik)
    End Function









    '---------------------------------------------------------------------
    'Osoby Kontaktowe
    Public oIdentOso As String            'IDENTKLIENTA     Text, 6
    Public oOsobaOso As String            'OSOBA            Text, 100
    Public oVIPOso As Boolean             'VIP              boolean
    Public oeMailOso As String            'EMAIL            Text, 100
    Public oTelefonOso As String          'TELEFON          Text, 100
    Public oTelefonKomOso As String       'TELEFONKOM       Text, 100
    Public oStanowiskoOso As String       'STANOWISKO       Text, 100
    Public oKompetencjeOso As String      'KOMPETENCJE      Text, 100
    Public oWersjaOso As Integer          'WERSJA           Integer

    Public sConnectionOsoby As String = ConnectionStrings()
    Public HQConnectionOsoby As String = HQConnectionStrings()

    'Public sSelectSQLOsoby As String = "SELECT * FROM Osoby WHERE IDENTKLIENTA='" & oIdentKli & "'"
    Public sSelectSQLOsoby As String = "SELECT " &
                                           "IDENTKLIENTA, " &
                                           "OSOBA," &
                                           "VIP," &
                                           "EMAIL," &
                                           "TELEFON," &
                                           "TELEFONKOM," &
                                           "STANOWISKO," &
                                           "KOMPETENCJE," &
                                           "WERSJA " &
                                        "FROM Osoby"


    Public sDeleteSQLOsoby As String = "DELETE FROM Osoby " &
                                        "WHERE (IDENTKLIENTA=@orygSymbol) AND " &
                                              "(OSOBA=@orygOsoba) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sUpdateSQLOsoby As String = "UPDATE Osoby SET " &
                                                 "VIP=@oVIPOso," &
                                                 "EMAIL=@oeMailOso," &
                                                 "TELEFON=@oTelefonOso," &
                                                 "TELEFONKOM=@oTelefonKomOso," &
                                                 "STANOWISKO=@oStanowiskoOso," &
                                                 "KOMPETENCJE=@oKompetencjeOso," &
                                                 "WERSJA=@oWersjaOso " &
                                        "WHERE (IDENTKLIENTA=@orygSymbol) AND " &
                                              "(OSOBA=@orygOsoba) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sInsertSQLOsoby As String = "INSERT INTO Osoby " &
                                         "(IDENTKLIENTA, " &
                                           "OSOBA," &
                                           "VIP," &
                                           "EMAIL," &
                                           "TELEFON," &
                                           "TELEFONKOM," &
                                           "STANOWISKO," &
                                           "KOMPETENCJE," &
                                           "WERSJA " &
                                           ") " &
                                "VALUES  (@oIdentOso," &
                                         "@oOsobaOso," &
                                         "@oVIPOso," &
                                         "@oemailOso," &
                                         "@oTelefonOso," &
                                         "@oTelefonKomOso," &
                                         "@oStanowiskoOso," &
                                         "@oKompetencjeOso," &
                                         "@oWersjaOso " &
                                         ")"


    'SQLDeleteOsoby(sqlCommandDelete, sqlDataAdapter)
    'SQLUpdateOsoby(sqlCommandUpdate, sqlDataAdapter)
    'SQLInsertOsoby(sqlCommandInsert, sqlDataAdapter)


    Public Sub SQLDeleteOsoby(ByRef sqlCommandDelete As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda DELETE
        'parametry filtrowania
        sqlParam = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.Char, 6, "IDENTKLIENTA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygOsoba", SqlDbType.VarChar, 100, "OSOBA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.DeleteCommand = sqlCommandDelete
    End Sub

    Public Sub SQLUpdateOsoby(ByRef sqlCommandUpdate As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda Update
        'parametry aktualizacji
        'sqlCommandUpdate.Parameters.Add("@oIdentOso", sqlDbType.Char, 6, "IDENTKLIENTA")
        'sqlCommandUpdate.Parameters.Add("@oOsobaOso", sqlDbType.varChar, 100, "OSOBA")
        sqlCommandUpdate.Parameters.Add("@oVIPOso", SqlDbType.Bit, Nothing, "VIP")
        sqlCommandUpdate.Parameters.Add("@oeMailOso", SqlDbType.VarChar, 100, "EMAIL")
        sqlCommandUpdate.Parameters.Add("@oTelefonOso", SqlDbType.VarChar, 100, "TELEFON")
        sqlCommandUpdate.Parameters.Add("@oTelefonKomOso", SqlDbType.VarChar, 100, "TELEFONKOM")
        sqlCommandUpdate.Parameters.Add("@oStanowiskoOso", SqlDbType.VarChar, 100, "STANOWISKO")
        sqlCommandUpdate.Parameters.Add("@oKompetencjeOso", SqlDbType.VarChar, 100, "KOMPETENCJE")
        sqlCommandUpdate.Parameters.Add("@oWersjaOso", SqlDbType.Int, Nothing, "WERSJA")

        'parametry filtrowania
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygSymbol", SqlDbType.Char, 6, "IDENTKLIENTA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygOsoba", SqlDbType.VarChar, 100, "OSOBA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.UpdateCommand = sqlCommandUpdate
    End Sub


    Public Sub SQLInsertOsoby(ByRef sqlCommandInsert As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda INSERT
        sqlParam = sqlCommandInsert.Parameters.Add("@oIdentOso", SqlDbType.Char, 6, "IDENTKLIENTA")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oOsobaOso", SqlDbType.VarChar, 100, "OSOBA")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oVIPOso", SqlDbType.Bit, Nothing, "VIP")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oeMailOso", SqlDbType.VarChar, 100, "EMAIL")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oTelefonOso", SqlDbType.VarChar, 100, "TELEFON")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oTelefonKomOso", SqlDbType.VarChar, 100, "TELEFONKOM")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oStanowiskoOso", SqlDbType.VarChar, 100, "STANOWISKO")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oKompetencjeOso", SqlDbType.VarChar, 100, "KOMPETENCJE")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oWersjaOso", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlDataAdapter.InsertCommand = sqlCommandInsert
    End Sub



    'Public Sub DBDeleteOsoby(ByRef dbCommandDelete As OleDb.OleDbCommand,
    '                       ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda DELETE
    '    'parametry filtrowania
    '    dbParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 6, "IDENTKLIENTA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandDelete.Parameters.Add("@orygOsoba", OleDb.OleDbType.VarChar, 100, "OSOBA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.DeleteCommand = dbCommandDelete
    'End Sub

    'Public Sub DBUpdateOsoby(ByRef dbCommandUpdate As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda Update
    '    'parametry aktualizacji
    '    'dbCommandUpdate.Parameters.Add("@oIdentOso", OleDb.OleDbType.varchar, 6, "IDENTKLIENTA")
    '    'dbCommandUpdate.Parameters.Add("@oOsobaOso", OleDb.OleDbType.varchar, 100, "OSOBA")
    '    dbCommandUpdate.Parameters.Add("@oVIPOso", OleDb.OleDbType.Boolean, Nothing, "VIP")
    '    dbCommandUpdate.Parameters.Add("@oStanowiskoOso", OleDb.OleDbType.VarChar, 100, "STANOWISKO")
    '    dbCommandUpdate.Parameters.Add("@oKompetencjeOso", OleDb.OleDbType.VarChar, 100, "KOMPETENCJE")
    '    dbCommandUpdate.Parameters.Add("@oTelefonOso", OleDb.OleDbType.VarChar, 100, "TELEFON")
    '    dbCommandUpdate.Parameters.Add("@oTelefonKomOso", OleDb.OleDbType.VarChar, 100, "TELEFONKOM")
    '    dbCommandUpdate.Parameters.Add("@oeMailOso", OleDb.OleDbType.VarChar, 100, "EMAIL")
    '    dbCommandUpdate.Parameters.Add("@oWersjaOso", OleDb.OleDbType.Integer, Nothing, "WERSJA")

    '    'parametry filtrowania
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 6, "IDENTKLIENTA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygOsoba", OleDb.OleDbType.VarChar, 100, "OSOBA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.UpdateCommand = dbCommandUpdate
    'End Sub

    'Public Sub DBInsertOsoby(ByRef dbCommandInsert As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda INSERT
    '    dbParam = dbCommandInsert.Parameters.Add("@oIdentOso", OleDb.OleDbType.VarChar, 6, "IDENTKLIENTA")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oOsobaOso", OleDb.OleDbType.VarChar, 100, "OSOBA")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oVIPOso", OleDb.OleDbType.Boolean, Nothing, "VIP")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oStanowiskoOso", OleDb.OleDbType.VarChar, 100, "STANOWISKO")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oKompetencjeOso", OleDb.OleDbType.VarChar, 100, "KOMPETENCJE")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oTelefonOso", OleDb.OleDbType.VarChar, 100, "TELEFON")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oTelefonKomOso", OleDb.OleDbType.VarChar, 100, "TELEFONKOM")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oeMailOso", OleDb.OleDbType.VarChar, 100, "EMAIL")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@oWersjaOso", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Current

    '    dbDataAdapter.InsertCommand = dbCommandInsert
    'End Sub



    Public Function IdentOsoby(ByVal IdentK As String, ByVal IdentO As String) As Boolean
        Dim dbSelectSQLOsoby As String = sSelectSQLOsoby & " WHERE IDENTKLIENTA='" & IdentK & "' AND OSOBA='" & IdentO & "' "

        Dim dbConnectionOsoby As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectOsoby As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterOsoby As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionOsoby As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectOsoby As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterOsoby As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetOsoby As New DataSet
        Dim DataViewOsoby As New DataView
        Dim ConnectionStateOsoby As System.Data.ConnectionState

        Dim Wynik As Boolean = False

        oIdentOso = ""
        oOsobaOso = ""
        oVIPOso = False
        oeMailOso = ""
        oTelefonOso = ""
        oTelefonKomOso = ""
        oStanowiskoOso = ""
        oKompetencjeOso = ""
        oWersjaOso = 1

        DataSetOsoby.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionOsoby = New OleDb.OleDbConnection(sConnectionKontakty)
            'dbCommandSelectOsoby = New OleDb.OleDbCommand(dbSelectSQLOsoby, dbConnectionOsoby)
            'dbDataAdapterOsoby = New OleDb.OleDbDataAdapter(dbCommandSelectOsoby)

            ''mapujemy domyslna nazwe tabeli
            'Dim dbMapowanieTabeliOsoby As System.Data.Common.DataTableMapping
            'dbMapowanieTabeliOsoby = dbDataAdapterOsoby.TableMappings.Add("Table", "TABELA_Osoby")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionOsoby.Open()
            'Catch Ex As System.Exception
            'Finally
            '    ConnectionStateOsoby = dbConnectionOsoby.State
            'End Try
        Else
            sqlConnectionOsoby = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectOsoby = New SqlClient.SqlCommand(dbSelectSQLOsoby, sqlConnectionOsoby)
            sqlDataAdapterOsoby = New SqlClient.SqlDataAdapter(sqlCommandSelectOsoby)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliOsoby As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliOsoby = sqlDataAdapterOsoby.TableMappings.Add("Table", "TABELA_Osoby")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionOsoby.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateOsoby = sqlConnectionOsoby.State
            End Try
        End If

        If ConnectionStateOsoby = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterOsoby.FillSchema(DataSetOsoby, SchemaType.Mapped, "TABELA_Osoby")
                    'dbDataAdapterOsoby.Fill(DataSetOsoby)
                    'dbConnectionOsoby.Close()
                Else
                    sqlDataAdapterOsoby.FillSchema(DataSetOsoby, SchemaType.Mapped, "TABELA_Osoby")
                    sqlDataAdapterOsoby.Fill(DataSetOsoby)
                    sqlConnectionOsoby.Close()
                End If

                'definiuj DataView
                DataViewOsoby = New DataView(DataSetOsoby.Tables(0))
                If DataViewOsoby.Count > 0 Then
                    oIdentOso = GetTxtDRVField(DataViewOsoby.Item(0), "IDENTKLIENTA")
                    oOsobaOso = GetTxtDRVField(DataViewOsoby.Item(0), "OSOBA")
                    oVIPOso = GetTxtDRVField(DataViewOsoby.Item(0), "VIP")
                    oeMailOso = GetTxtDRVField(DataViewOsoby.Item(0), "EMAIL")
                    oTelefonOso = GetTxtDRVField(DataViewOsoby.Item(0), "TELEFON")
                    oTelefonKomOso = GetTxtDRVField(DataViewOsoby.Item(0), "TELEFONKOM")
                    oStanowiskoOso = GetTxtDRVField(DataViewOsoby.Item(0), "STANOWISKO")
                    oKompetencjeOso = GetTxtDRVField(DataViewOsoby.Item(0), "KOMPETENCJE")
                    oWersjaOso = GetIntDRVField(DataViewOsoby.Item(0), "WERSJA")

                    Wynik = True
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wynik)
    End Function
























































    '---------------------------------------------------------------------
    'RodzajeKontaktow
    Public rkUniqId As String             'UNIQID           Text, 40
    Public rkRodzajKontaktu As String     'RODZAJKONTAKTU   Text, 50
    Public rkWersja As Integer            'WERSJA           Integer

    Public sConnectionRodzajeKontaktow As String = ConnectionStrings()
    Public HQConnectionRodzajeKontaktow As String = HQConnectionStrings()

    Public sSelectSQLRodzajeKontaktow As String = "SELECT UNIQID, " &
                                                         "RODZAJKONTAKTU," &
                                                         "WERSJA " &
                                                         "FROM RodzajeKontaktow"

    Public sDeleteSQLRodzajeKontaktow As String = "DELETE FROM RodzajeKontaktow " &
                                        "WHERE (UNIQID=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sUpdateSQLRodzajeKontaktow As String = "UPDATE RodzajeKontaktow SET " &
                                                 "RODZAJKONTAKTU=@rRodzajKontaktu," &
                                                 "WERSJA=@rWersja " &
                                        "WHERE (UNIQID=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sInsertSQLRodzajeKontaktow As String = "INSERT INTO RodzajeKontaktow " &
                                         "(UNIQID, " &
                                           "RODZAJKONTAKTU," &
                                           "WERSJA " &
                                           ") " &
                                "VALUES  (@rUNIQID," &
                                         "@rRODZAJKONTAKTU," &
                                         "@rWersja " &
                                         ")"


    'SQLDeleteRodzajeKontaktow(sqlCommandDelete, sqlDataAdapter)
    'SQLUpdateRodzajeKontaktow(sqlCommandUpdate, sqlDataAdapter)
    'SQLInsertRodzajeKontaktow(sqlCommandInsert, sqlDataAdapter)


    Public Sub SQLDeleteRodzajeKontaktow(ByRef sqlCommandDelete As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda DELETE
        'parametry filtrowania
        sqlParam = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 40, "UNIQID")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.DeleteCommand = sqlCommandDelete
    End Sub

    Public Sub SQLUpdateRodzajeKontaktow(ByRef sqlCommandUpdate As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        '------------------------------------------
        'komenda UPDATE
        'parametry aktualizacji
        'sqlCommandUpdate.Parameters.Add("@rUNIQID", SqlDbType.varChar, 40, "UNIQID")
        sqlCommandUpdate.Parameters.Add("@rRODZAJKONTAKTU", SqlDbType.VarChar, 50, "RODZAJKONTAKTU")
        sqlCommandUpdate.Parameters.Add("@rWersja", SqlDbType.Int, Nothing, "WERSJA")

        'parametry filtrowania
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 40, "UNIQID")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.UpdateCommand = sqlCommandUpdate

    End Sub


    Public Sub SQLInsertRodzajeKontaktow(ByRef sqlCommandInsert As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda INSERT
        sqlParam = sqlCommandInsert.Parameters.Add("@rUNIQID", SqlDbType.VarChar, 40, "UNIQID")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@rRODZAJKONTAKTU", SqlDbType.VarChar, 50, "RODZAJKONTAKTU")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@rWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlDataAdapter.InsertCommand = sqlCommandInsert
    End Sub



    'Public Sub DBDeleteRodzajeKontaktow(ByRef dbCommandDelete As OleDb.OleDbCommand,
    '                       ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda DELETE
    '    'parametry filtrowania
    '    dbParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.DeleteCommand = dbCommandDelete
    'End Sub

    'Public Sub DBUpdateRodzajeKontaktow(ByRef dbCommandUpdate As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda UPDATE
    '    'parametry aktualizacji
    '    'dbCommandUpdate.Parameters.Add("@rUNIQID", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbCommandUpdate.Parameters.Add("@rRODZAJKONTAKTU", OleDb.OleDbType.VarChar, 50, "RODZAJKONTAKTU")
    '    dbCommandUpdate.Parameters.Add("@rWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")

    '    'parametry filtrowania
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.UpdateCommand = dbCommandUpdate
    'End Sub

    'Public Sub DBInsertRodzajeKontaktow(ByRef dbCommandInsert As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda INSERT
    '    dbParam = dbCommandInsert.Parameters.Add("@rUNIQID", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@rRODZAJKONTAKTU", OleDb.OleDbType.VarChar, 50, "RODZAJKONTAKTU")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@rWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbDataAdapter.InsertCommand = dbCommandInsert
    'End Sub



    Public Function IdentRodzajeKontaktow(ByVal IdentK As String, ByVal IdentU As String) As Boolean
        Dim dbSelectSQLRodzajeKontaktow As String = sSelectSQLRodzajeKontaktow & " WHERE UNIQID='" & IdentU & "' "

        Dim dbConnectionRodzajeKontaktow As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectRodzajeKontaktow As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterRodzajeKontaktow As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionRodzajeKontaktow As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectRodzajeKontaktow As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterRodzajeKontaktow As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetRodzajeKontaktow As New DataSet
        Dim DataViewRodzajeKontaktow As New DataView
        Dim ConnectionStateRodzajeKontaktow As System.Data.ConnectionState

        Dim Wynik As Boolean = False

        rkUniqId = ""
        rkRodzajKontaktu = ""
        rkWersja = 0

        DataSetRodzajeKontaktow.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionRodzajeKontaktow = New OleDb.OleDbConnection(sConnectionKontakty)
            'dbCommandSelectRodzajeKontaktow = New OleDb.OleDbCommand(dbSelectSQLRodzajeKontaktow, dbConnectionRodzajeKontaktow)
            'dbDataAdapterRodzajeKontaktow = New OleDb.OleDbDataAdapter(dbCommandSelectRodzajeKontaktow)

            ''mapujemy domyslna nazwe tabeli
            'Dim dbMapowanieTabeliRodzajeKontaktow As System.Data.Common.DataTableMapping
            'dbMapowanieTabeliRodzajeKontaktow = dbDataAdapterRodzajeKontaktow.TableMappings.Add("Table", "TABELA_RodzajeKontaktow")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionRodzajeKontaktow.Open()
            'Catch Ex As System.Exception
            'Finally
            '    ConnectionStateRodzajeKontaktow = dbConnectionRodzajeKontaktow.State
            'End Try
        Else
            sqlConnectionRodzajeKontaktow = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectRodzajeKontaktow = New SqlClient.SqlCommand(dbSelectSQLRodzajeKontaktow, sqlConnectionRodzajeKontaktow)
            sqlDataAdapterRodzajeKontaktow = New SqlClient.SqlDataAdapter(sqlCommandSelectRodzajeKontaktow)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliRodzajeKontaktow As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliRodzajeKontaktow = sqlDataAdapterRodzajeKontaktow.TableMappings.Add("Table", "TABELA_RodzajeKontaktow")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionRodzajeKontaktow.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateRodzajeKontaktow = sqlConnectionRodzajeKontaktow.State
            End Try
        End If

        If ConnectionStateRodzajeKontaktow = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterRodzajeKontaktow.FillSchema(DataSetRodzajeKontaktow, SchemaType.Mapped, "TABELA_RodzajeKontaktow")
                    'dbDataAdapterRodzajeKontaktow.Fill(DataSetRodzajeKontaktow)
                    'dbConnectionRodzajeKontaktow.Close()
                Else
                    sqlDataAdapterRodzajeKontaktow.FillSchema(DataSetRodzajeKontaktow, SchemaType.Mapped, "TABELA_RodzajeKontaktow")
                    sqlDataAdapterRodzajeKontaktow.Fill(DataSetRodzajeKontaktow)
                    sqlConnectionRodzajeKontaktow.Close()
                End If

                'definiuj DataView
                DataViewRodzajeKontaktow = New DataView(DataSetRodzajeKontaktow.Tables(0))
                If DataViewRodzajeKontaktow.Count > 0 Then
                    rkUniqId = GetTxtDRVField(DataViewRodzajeKontaktow.Item(0), "UNIQID")
                    rkRodzajKontaktu = GetTxtDRVField(DataViewRodzajeKontaktow.Item(0), "RODZAJKONTAKTU")
                    rkWersja = GetIntDRVField(DataViewRodzajeKontaktow.Item(0), "WERSJA")

                    Wynik = True
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wynik)
    End Function























    '---------------------------------------------------------------------
    'SlownikCoDalej
    Public scdIDENT As String             'IDENT       Text, 40
    Public scdOPIS As String              'OPIS        memo
    Public scdWersja As Integer           'WERSJA      Integer

    Public sConnectionSlownikCoDalej As String = ConnectionStrings()
    Public HQConnectionSlownikCoDalej As String = HQConnectionStrings()

    Public sSelectSQLSlownikCoDalej As String = "SELECT IDENT, " &
                                                       "OPIS," &
                                                       "WERSJA " &
                                                       "FROM SlownikCoDalej"

    Public sDeleteSQLSlownikCoDalej As String = "DELETE FROM SlownikCoDalej " &
                                        "WHERE (IDENT=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sUpdateSQLSlownikCoDalej As String = "UPDATE SlownikCoDalej SET " &
                                                 "OPIS=@scdOPIS," &
                                                 "WERSJA=@scdWersja " &
                                        "WHERE (IDENT=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sInsertSQLSlownikCoDalej As String = "INSERT INTO SlownikCoDalej " &
                                         "(IDENT, " &
                                           "OPIS," &
                                           "WERSJA " &
                                           ") " &
                                "VALUES  (@scdIDENT," &
                                         "@scdOPIS," &
                                         "@scdWersja " &
                                         ")"


    'SQLDeleteSlownikCoDalej(sqlCommandDelete, sqlDataAdapter)
    'SQLUpdateSlownikCoDalej(sqlCommandUpdate, sqlDataAdapter)
    'SQLInsertSlownikCoDalej(sqlCommandInsert, sqlDataAdapter)

    Public Sub SQLDeleteSlownikCoDalej(ByRef sqlCommandDelete As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter

        sqlParam = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 20, "IDENT")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original

        sqlDataAdapter.DeleteCommand = sqlCommandDelete
    End Sub

    Public Sub SQLUpdateSlownikCoDalej(ByRef sqlCommandUpdate As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter

        'sqlCommandUpdate.Parameters.Add("@scdIDENT", SqlDbType.VarChar, 20, "IDENT")
        sqlCommandUpdate.Parameters.Add("@scdOPIS", SqlDbType.Text, Nothing, "OPIS")
        sqlCommandUpdate.Parameters.Add("@scdWERSJA", SqlDbType.Int, Nothing, "WERSJA")

        'parametry filtrowania
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 20, "IDENT")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original

        sqlDataAdapter.UpdateCommand = sqlCommandUpdate
    End Sub

    Public Sub SQLInsertSlownikCoDalej(ByRef sqlCommandInsert As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter

        sqlParam = sqlCommandInsert.Parameters.Add("@scdIDENT", SqlDbType.VarChar, 20, "IDENT")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@scdOPIS", SqlDbType.Text, Nothing, "OPIS")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@scdWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlDataAdapter.InsertCommand = sqlCommandInsert
    End Sub


    'Public Sub DBDeleteSlownikCoDalej(ByRef dbCommandDelete As OleDb.OleDbCommand,
    '                       ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As OleDb.OleDbParameter

    '    dbParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 20, "IDENT")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original

    '    dbDataAdapter.DeleteCommand = dbCommandDelete
    'End Sub

    'Public Sub DBUpdateSlownikCoDalej(ByRef dbCommandUpdate As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter

    '    'dbCommandUpdate.Parameters.Add("@scdIDENT", OleDb.OleDbType.VarChar, 20, "IDENT")
    '    dbCommandUpdate.Parameters.Add("@scdOPIS", OleDb.OleDbType.WChar, Nothing, "OPIS")
    '    dbCommandUpdate.Parameters.Add("@scdWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")

    '    'parametry filtrowania
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 20, "IDENT")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original

    '    dbDataAdapter.UpdateCommand = dbCommandUpdate
    'End Sub

    'Public Sub DBInsertSlownikCoDalej(ByRef dbCommandInsert As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter

    '    dbParam = dbCommandInsert.Parameters.Add("@scdIDENT", OleDb.OleDbType.VarChar, 20, "IDENT")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@scdOPIS", OleDb.OleDbType.WChar, Nothing, "OPIS")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@scdWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Current

    '    dbDataAdapter.InsertCommand = dbCommandInsert
    'End Sub













    '---------------------------------------------------------------------
    'CoDalejPlan
    Public cdUNIQID As String      'UNIQID Text, 40
    Public cdIDENTKLIENTA As String      'IDENTKLIENTA Text, 40
    Public cdIDENT As String             'IDENT        Text, 40
    Public cdOPIS As String              'OPIS         memo
    Public cdWersja As Integer           'WERSJA       Integer

    Public sConnectionCoDalej As String = ConnectionStrings()
    Public HQConnectionCoDalej As String = HQConnectionStrings()

    Public sSelectSQLCoDalej As String = "SELECT UNIQID, " &
                                                       "IDENTKLIENTA, " &
                                                       "IDENT," &
                                                       "OPIS," &
                                                       "WERSJA " &
                                                       "FROM CoDalejPlan"

    Public sDeleteSQLCoDalej As String = "DELETE FROM CoDalejPlan " &
                                        "WHERE (UNIQID=@orygSymbol) AND " &
                                              "(IDENTKLIENTA=@orygIdent) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sUpdateSQLCoDalej As String = "UPDATE CoDalejPlan SET " &
                                                 "IDENT=@cdIDENT," &
                                                 "OPIS=@cdOPIS," &
                                                 "WERSJA=@cdWersja " &
                                        "WHERE (UNIQID=@orygSymbol) AND " &
                                              "(IDENTKLIENTA=@orygIdent) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sInsertSQLCoDalej As String = "INSERT INTO CoDalejPlan " &
                                         "(UNIQID, " &
                                           "IDENTKLIENTA, " &
                                           "IDENT," &
                                           "OPIS," &
                                           "WERSJA " &
                                           ") " &
                                "VALUES  (@cdUNIQID," &
                                         "@cdIDENTKLIENTA," &
                                         "@cdIDENT," &
                                         "@cdOPIS," &
                                         "@cdWersja " &
                                         ")"


    'SQLDeleteCoDalej(sqlCommandDelete, sqlDataAdapter)
    'SQLUpdateCoDalej(sqlCommandUpdate, sqlDataAdapter)
    'SQLInsertCoDalej(sqlCommandInsert, sqlDataAdapter)

    Public Sub SQLDeleteCoDalej(ByRef sqlCommandDelete As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter

        sqlParam = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 40, "UNIQID")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygIDENT", SqlDbType.VarChar, 6, "IDENTKLIENTA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original

        sqlDataAdapter.DeleteCommand = sqlCommandDelete
    End Sub

    Public Sub SQLUpdateCoDalej(ByRef sqlCommandUpdate As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter

        'sqlCommandUpdate.Parameters.Add("@cdUNIQID", SqlDbType.VarChar, 40, "UNIQID")
        'sqlCommandUpdate.Parameters.Add("@cdIDENTKLIENTA", SqlDbType.VarChar, 6, "IDENTKLIENTA")
        sqlCommandUpdate.Parameters.Add("@cdIDENT", SqlDbType.VarChar, 20, "IDENT")
        sqlCommandUpdate.Parameters.Add("@cdOPIS", SqlDbType.Text, Nothing, "OPIS")
        sqlCommandUpdate.Parameters.Add("@cdWERSJA", SqlDbType.Int, Nothing, "WERSJA")

        'parametry filtrowania
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 40, "UNIQID")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygIdent", SqlDbType.VarChar, 6, "IDENTKLIENTA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original

        sqlDataAdapter.UpdateCommand = sqlCommandUpdate
    End Sub

    Public Sub SQLInsertCoDalej(ByRef sqlCommandInsert As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter

        sqlParam = sqlCommandInsert.Parameters.Add("@cdUNIQID", SqlDbType.VarChar, 40, "UNIQID")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@cdIDENTKLIENTA", SqlDbType.VarChar, 6, "IDENTKLIENTA")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@cdIDENT", SqlDbType.VarChar, 20, "IDENT")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@cdOPIS", SqlDbType.Text, Nothing, "OPIS")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@cdWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlDataAdapter.InsertCommand = sqlCommandInsert
    End Sub


    'Public Sub DBDeleteCoDalej(ByRef dbCommandDelete As OleDb.OleDbCommand,
    '                       ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As OleDb.OleDbParameter

    '    dbParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandDelete.Parameters.Add("@orygIDENT", OleDb.OleDbType.VarChar, 6, "IDENTKLIENTA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original

    '    dbDataAdapter.DeleteCommand = dbCommandDelete
    'End Sub

    'Public Sub DBUpdateCoDalej(ByRef dbCommandUpdate As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter

    '    'dbCommandUpdate.Parameters.Add("@cdUNIQID", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    'dbCommandUpdate.Parameters.Add("@cdIDENTKLIENTA", OleDb.OleDbType.VarChar, 6, "IDENTKLIENTA")
    '    dbCommandUpdate.Parameters.Add("@cdIDENT", OleDb.OleDbType.VarChar, 20, "IDENT")
    '    dbCommandUpdate.Parameters.Add("@cdOPIS", OleDb.OleDbType.WChar, Nothing, "OPIS")
    '    dbCommandUpdate.Parameters.Add("@cdWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")

    '    'parametry filtrowania
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygIDENT", OleDb.OleDbType.VarChar, 6, "IDENTKLIENTA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original

    '    dbDataAdapter.UpdateCommand = dbCommandUpdate
    'End Sub

    'Public Sub DBInsertCoDalej(ByRef dbCommandInsert As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter

    '    dbParam = dbCommandInsert.Parameters.Add("@cdUNIQID", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@cdIDENTKLIENTA", OleDb.OleDbType.VarChar, 6, "IDENTKLIENTA")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@cdIDENT", OleDb.OleDbType.VarChar, 20, "IDENT")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@cdOPIS", OleDb.OleDbType.WChar, Nothing, "OPIS")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@cdWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Current

    '    dbDataAdapter.InsertCommand = dbCommandInsert
    'End Sub




    Public Function IdentCoDalej(ByVal IdentK As String) As Boolean
        Dim dbSelectSQLCoDalej As String = sSelectSQLCoDalej & " WHERE IDENTKLIENTA='" & IdentK & "' "

        Dim dbConnectionCoDalej As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectCoDalej As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterCoDalej As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionCoDalej As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectCoDalej As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterCoDalej As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetCoDalej As New DataSet
        Dim DataViewCoDalej As New DataView
        Dim ConnectionStateCoDalej As System.Data.ConnectionState

        Dim Wynik As Boolean = False

        cdUNIQID = ""
        cdIDENTKLIENTA = ""
        cdIDENT = ""
        cdOPIS = ""
        cdWersja = 1

        DataSetCoDalej.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionCoDalej = New OleDb.OleDbConnection(sConnectionKontakty)
            'dbCommandSelectCoDalej = New OleDb.OleDbCommand(dbSelectSQLCoDalej, dbConnectionCoDalej)
            'dbDataAdapterCoDalej = New OleDb.OleDbDataAdapter(dbCommandSelectCoDalej)

            ''mapujemy domyslna nazwe tabeli
            'Dim dbMapowanieTabeliCoDalej As System.Data.Common.DataTableMapping
            'dbMapowanieTabeliCoDalej = dbDataAdapterCoDalej.TableMappings.Add("Table", "TABELA_CoDalej")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionCoDalej.Open()
            'Catch Ex As System.Exception
            'Finally
            '    ConnectionStateCoDalej = dbConnectionCoDalej.State
            'End Try
        Else
            sqlConnectionCoDalej = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectCoDalej = New SqlClient.SqlCommand(dbSelectSQLCoDalej, sqlConnectionCoDalej)
            sqlDataAdapterCoDalej = New SqlClient.SqlDataAdapter(sqlCommandSelectCoDalej)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliCoDalej As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliCoDalej = sqlDataAdapterCoDalej.TableMappings.Add("Table", "TABELA_CoDalej")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionCoDalej.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateCoDalej = sqlConnectionCoDalej.State
            End Try
        End If

        If ConnectionStateCoDalej = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterCoDalej.FillSchema(DataSetCoDalej, SchemaType.Mapped, "TABELA_CoDalej")
                    'dbDataAdapterCoDalej.Fill(DataSetCoDalej)
                    'dbConnectionCoDalej.Close()
                Else
                    sqlDataAdapterCoDalej.FillSchema(DataSetCoDalej, SchemaType.Mapped, "TABELA_CoDalej")
                    sqlDataAdapterCoDalej.Fill(DataSetCoDalej)
                    sqlConnectionCoDalej.Close()
                End If

                'definiuj DataView
                DataViewCoDalej = New DataView(DataSetCoDalej.Tables(0))
                If DataViewCoDalej.Count > 0 Then
                    cdUNIQID = GetTxtDRVField(DataViewCoDalej.Item(0), "UNIQID")
                    cdIDENTKLIENTA = GetTxtDRVField(DataViewCoDalej.Item(0), "IDENTKLIENTA")
                    cdIDENT = GetTxtDRVField(DataViewCoDalej.Item(0), "IDENT")
                    cdOPIS = GetTxtDRVField(DataViewCoDalej.Item(0), "OPIS")
                    cdWersja = GetIntDRVField(DataViewCoDalej.Item(0), "WERSJA")

                    Wynik = True
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wynik)
    End Function















    '---------------------------------------------------------------------
    'Branze
    Public brIdBranzy As String            'IDBRANZY         Text, 
    Public brWersja As Integer             'WERSJA           Integer

    Public sConnectionBranze As String = ConnectionStrings()
    Public HQConnectionBranze As String = HQConnectionStrings()

    Public sSelectSQLBranze As String = "SELECT IDBRANZY, " &
                                               "WERSJA " &
                                        "FROM Branze"

    Public sDeleteSQLBranze As String = "DELETE FROM Branze " &
                                        "WHERE (IDBRANZY=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sUpdateSQLBranze As String = "UPDATE Branze SET " &
                                                 "WERSJA=@brWersja " &
                                        "WHERE (IDBRANZY=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sInsertSQLBranze As String = "INSERT INTO Branze " &
                                         "(IDBRANZY, " &
                                          "WERSJA " &
                                          ") " &
                                "VALUES  (@brIDBRANZY," &
                                         "@brWersja " &
                                         ")"


    'SQLDeleteBranze(sqlCommandDelete, sqlDataAdapter)
    'SQLUpdateBranze(sqlCommandUpdate, sqlDataAdapter)
    'SQLInsertBranze(sqlCommandInsert, sqlDataAdapter)


    Public Sub SQLDeleteBranze(ByRef sqlCommandDelete As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda DELETE
        'parametry filtrowania
        sqlParam = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 100, "IDBRANZY")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.DeleteCommand = sqlCommandDelete
    End Sub

    Public Sub SQLUpdateBranze(ByRef sqlCommandUpdate As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        '------------------------------------------
        'komenda UPDATE
        'parametry aktualizacji
        'sqlCommandUpdate.Parameters.Add("@brIDBRANZY", SqlDbType.varChar, 100, "IDBRANZY")
        sqlCommandUpdate.Parameters.Add("@brWersja", SqlDbType.Int, Nothing, "WERSJA")

        'parametry filtrowania
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 100, "IDBRANZY")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.UpdateCommand = sqlCommandUpdate

    End Sub


    Public Sub SQLInsertBranze(ByRef sqlCommandInsert As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda INSERT
        sqlParam = sqlCommandInsert.Parameters.Add("@brIDBRANZY", SqlDbType.VarChar, 100, "IDBRANZY")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@brWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlDataAdapter.InsertCommand = sqlCommandInsert
    End Sub



    'Public Sub DBDeleteBranze(ByRef dbCommandDelete As OleDb.OleDbCommand,
    '                       ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda DELETE
    '    'parametry filtrowania
    '    dbParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.DeleteCommand = dbCommandDelete
    'End Sub

    'Public Sub DBUpdateBranze(ByRef dbCommandUpdate As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda UPDATE
    '    'parametry aktualizacji
    '    'dbCommandUpdate.Parameters.Add("@rUNIQID", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbCommandUpdate.Parameters.Add("@rRODZAJKONTAKTU", OleDb.OleDbType.VarChar, 50, "RODZAJKONTAKTU")
    '    dbCommandUpdate.Parameters.Add("@rWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")

    '    'parametry filtrowania
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.UpdateCommand = dbCommandUpdate
    'End Sub

    'Public Sub DBInsertBranze(ByRef dbCommandInsert As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda INSERT
    '    dbParam = dbCommandInsert.Parameters.Add("@rUNIQID", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@rRODZAJKONTAKTU", OleDb.OleDbType.VarChar, 50, "RODZAJKONTAKTU")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@rWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbDataAdapter.InsertCommand = dbCommandInsert
    'End Sub



    Public Function IdentBranze(ByVal IdentB As String) As Boolean
        Dim dbSelectSQLBranze As String = sSelectSQLBranze & " WHERE IDBRANZA='" & IdentB & "' "

        Dim dbConnectionBranze As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectBranze As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterBranze As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionBranze As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectBranze As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterBranze As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetBranze As New DataSet
        Dim DataViewBranze As New DataView
        Dim ConnectionStateBranze As System.Data.ConnectionState

        Dim Wynik As Boolean = False

        brIdBranzy = ""
        brWersja = 0

        DataSetBranze.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionBranze = New OleDb.OleDbConnection(sConnectionKontakty)
            'dbCommandSelectBranze = New OleDb.OleDbCommand(dbSelectSQLBranze, dbConnectionBranze)
            'dbDataAdapterBranze = New OleDb.OleDbDataAdapter(dbCommandSelectBranze)

            ''mapujemy domyslna nazwe tabeli
            'Dim dbMapowanieTabeliBranze As System.Data.Common.DataTableMapping
            'dbMapowanieTabeliBranze = dbDataAdapterBranze.TableMappings.Add("Table", "TABELA_Branze")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionBranze.Open()
            'Catch Ex As System.Exception
            'Finally
            '    ConnectionStateBranze = dbConnectionBranze.State
            'End Try
        Else
            sqlConnectionBranze = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectBranze = New SqlClient.SqlCommand(dbSelectSQLBranze, sqlConnectionBranze)
            sqlDataAdapterBranze = New SqlClient.SqlDataAdapter(sqlCommandSelectBranze)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliBranze As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliBranze = sqlDataAdapterBranze.TableMappings.Add("Table", "TABELA_Branze")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionBranze.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateBranze = sqlConnectionBranze.State
            End Try
        End If

        If ConnectionStateBranze = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterBranze.FillSchema(DataSetBranze, SchemaType.Mapped, "TABELA_Branze")
                    'dbDataAdapterBranze.Fill(DataSetBranze)
                    'dbConnectionBranze.Close()
                Else
                    sqlDataAdapterBranze.FillSchema(DataSetBranze, SchemaType.Mapped, "TABELA_Branze")
                    sqlDataAdapterBranze.Fill(DataSetBranze)
                    sqlConnectionBranze.Close()
                End If

                'definiuj DataView
                DataViewBranze = New DataView(DataSetBranze.Tables(0))
                If DataViewBranze.Count > 0 Then
                    brIdBranzy = GetTxtDRVField(DataViewBranze.Item(0), "IDBRANZY")
                    brWersja = GetIntDRVField(DataViewBranze.Item(0), "WERSJA")

                    Wynik = True
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wynik)
    End Function











    '---------------------------------------------------------------------
    'PodPodBranze
    Public pbrIdBranzy As String            'IDBRANZY         Text, 
    Public pbrIdPodBranzy As String         'IDPODBRANZY         Text, 
    Public pbrWersja As Integer             'WERSJA           Integer

    Public sConnectionPodBranze As String = ConnectionStrings()
    Public HQConnectionPodBranze As String = HQConnectionStrings()

    Public sSelectSQLPodBranze As String = "SELECT IDBRANZY, " &
                                               "IDPODBRANZY, " &
                                               "WERSJA " &
                                        "FROM PodBranze"

    Public sDeleteSQLPodBranze As String = "DELETE FROM PodBranze " &
                                        "WHERE (IDBRANZY=@orygSymbol) AND " &
                                              "(IDPODBRANZY=@orygSymbol2) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sUpdateSQLPodBranze As String = "UPDATE PodBranze SET " &
                                                 "WERSJA=@pbrWersja " &
                                        "WHERE (IDBRANZY=@orygSymbol) AND " &
                                              "(IDPODBRANZY=@orygSymbol2) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sInsertSQLPodBranze As String = "INSERT INTO PodBranze " &
                                         "(IDBRANZY, " &
                                          "IDPODBRANZY, " &
                                          "WERSJA " &
                                          ") " &
                                "VALUES  (@pbrIDBRANZY," &
                                         "@pbrIDPODBRANZY, " &
                                         "@pbrWersja " &
                                         ")"


    'SQLDeletePodBranze(sqlCommandDelete, sqlDataAdapter)
    'SQLUpdatePodBranze(sqlCommandUpdate, sqlDataAdapter)
    'SQLInsertPodBranze(sqlCommandInsert, sqlDataAdapter)


    Public Sub SQLDeletePodBranze(ByRef sqlCommandDelete As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda DELETE
        'parametry filtrowania
        sqlParam = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 100, "IDBRANZY")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygSymbol2", SqlDbType.VarChar, 100, "IDPODBRANZY")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.DeleteCommand = sqlCommandDelete
    End Sub

    Public Sub SQLUpdatePodBranze(ByRef sqlCommandUpdate As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        '------------------------------------------
        'komenda UPDATE
        'parametry aktualizacji
        'sqlCommandUpdate.Parameters.Add("@pbrIDBRANZY", SqlDbType.varChar, 100, "IDBRANZY")
        'sqlCommandUpdate.Parameters.Add("@pbrIDPODBRANZY", SqlDbType.varChar, 100, "IDPODBRANZY")
        sqlCommandUpdate.Parameters.Add("@pbrWersja", SqlDbType.Int, Nothing, "WERSJA")

        'parametry filtrowania
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 100, "IDBRANZY")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygSymbol2", SqlDbType.VarChar, 100, "IDPODBRANZY")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.UpdateCommand = sqlCommandUpdate

    End Sub


    Public Sub SQLInsertPodBranze(ByRef sqlCommandInsert As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda INSERT
        sqlParam = sqlCommandInsert.Parameters.Add("@pbrIDBRANZY", SqlDbType.VarChar, 100, "IDBRANZY")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@pbrIDPODBRANZY", SqlDbType.VarChar, 100, "IDPODBRANZY")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@pbrWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlDataAdapter.InsertCommand = sqlCommandInsert
    End Sub



    'Public Sub DBDeletePodBranze(ByRef dbCommandDelete As OleDb.OleDbCommand,
    '                       ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda DELETE
    '    'parametry filtrowania
    '    dbParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.DeleteCommand = dbCommandDelete
    'End Sub

    'Public Sub DBUpdatePodBranze(ByRef dbCommandUpdate As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda UPDATE
    '    'parametry aktualizacji
    '    'dbCommandUpdate.Parameters.Add("@rUNIQID", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbCommandUpdate.Parameters.Add("@rRODZAJKONTAKTU", OleDb.OleDbType.VarChar, 50, "RODZAJKONTAKTU")
    '    dbCommandUpdate.Parameters.Add("@rWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")

    '    'parametry filtrowania
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbParam = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.UpdateCommand = dbCommandUpdate
    'End Sub

    'Public Sub DBInsertPodBranze(ByRef dbCommandInsert As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParam As New OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda INSERT
    '    dbParam = dbCommandInsert.Parameters.Add("@rUNIQID", OleDb.OleDbType.VarChar, 40, "UNIQID")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@rRODZAJKONTAKTU", OleDb.OleDbType.VarChar, 50, "RODZAJKONTAKTU")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbParam = dbCommandInsert.Parameters.Add("@rWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParam.SourceVersion = DataRowVersion.Current
    '    dbDataAdapter.InsertCommand = dbCommandInsert
    'End Sub



    Public Function IdentPodBranze(ByVal IdentB As String, ByVal IdentPB As String) As Boolean
        Dim dbSelectSQLPodBranze As String = sSelectSQLPodBranze & " WHERE (IDBRANZA='" & IdentB & "') AND (IDPODBRANZA='" & IdentPB & "') "

        Dim dbConnectionPodBranze As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectPodBranze As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterPodBranze As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionPodBranze As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectPodBranze As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterPodBranze As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetPodBranze As New DataSet
        Dim DataViewPodBranze As New DataView
        Dim ConnectionStatePodBranze As System.Data.ConnectionState

        Dim Wynik As Boolean = False

        pbrIdBranzy = ""
        pbrIdPodBranzy = ""
        pbrWersja = 0

        DataSetPodBranze.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionPodBranze = New OleDb.OleDbConnection(sConnectionKontakty)
            'dbCommandSelectPodBranze = New OleDb.OleDbCommand(dbSelectSQLPodBranze, dbConnectionPodBranze)
            'dbDataAdapterPodBranze = New OleDb.OleDbDataAdapter(dbCommandSelectPodBranze)

            ''mapujemy domyslna nazwe tabeli
            'Dim dbMapowanieTabeliPodBranze As System.Data.Common.DataTableMapping
            'dbMapowanieTabeliPodBranze = dbDataAdapterPodBranze.TableMappings.Add("Table", "TABELA_PodBranze")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionPodBranze.Open()
            'Catch Ex As System.Exception
            'Finally
            '    ConnectionStatePodBranze = dbConnectionPodBranze.State
            'End Try
        Else
            sqlConnectionPodBranze = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectPodBranze = New SqlClient.SqlCommand(dbSelectSQLPodBranze, sqlConnectionPodBranze)
            sqlDataAdapterPodBranze = New SqlClient.SqlDataAdapter(sqlCommandSelectPodBranze)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliPodBranze As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliPodBranze = sqlDataAdapterPodBranze.TableMappings.Add("Table", "TABELA_PodBranze")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionPodBranze.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStatePodBranze = sqlConnectionPodBranze.State
            End Try
        End If

        If ConnectionStatePodBranze = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterPodBranze.FillSchema(DataSetPodBranze, SchemaType.Mapped, "TABELA_PodBranze")
                    'dbDataAdapterPodBranze.Fill(DataSetPodBranze)
                    'dbConnectionPodBranze.Close()
                Else
                    sqlDataAdapterPodBranze.FillSchema(DataSetPodBranze, SchemaType.Mapped, "TABELA_PodBranze")
                    sqlDataAdapterPodBranze.Fill(DataSetPodBranze)
                    sqlConnectionPodBranze.Close()
                End If

                'definiuj DataView
                DataViewPodBranze = New DataView(DataSetPodBranze.Tables(0))
                If DataViewPodBranze.Count > 0 Then
                    pbrIdBranzy = GetTxtDRVField(DataViewPodBranze.Item(0), "IDBRANZY")
                    pbrIdPodBranzy = GetTxtDRVField(DataViewPodBranze.Item(0), "IDPODBRANZY")
                    pbrWersja = GetIntDRVField(DataViewPodBranze.Item(0), "WERSJA")

                    Wynik = True
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wynik)
    End Function










End Module
