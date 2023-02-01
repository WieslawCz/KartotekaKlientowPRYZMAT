Module modStrukturyBazaNadrzedna

    '**********************************************
    ' struktury do obsługi BAZY NADRZĘDNDJ
    ' Przeznaczone dla Przedstawiciela Handlowego działającego OF-LINE
    '**********************************************





    '**************************************
    '** Connections Strings
    '**************************************

    Public Function HQConnectionStrings() As String
        oConnectionString = HQDataBaseConnection()

        HQConnectionUzytkownicy = oConnectionString
        HQConnectionKlienci = oConnectionString
        HQConnectionKontakty = oConnectionString
        HQConnectionRodzajeKontaktow = oConnectionString
        HQConnectionOsoby = oConnectionString
        HQConnectionObroty = oConnectionString
        HQConnectionObrotyMies = oConnectionString
        HQConnectionAkcjeOpis = oConnectionString
        HQConnectionAkcjeSpec = oConnectionString
        HQConnectionAnalizyZakupu = oConnectionString
        HQConnectionDaneDoRaportu = oConnectionString
        HQConnectionSlownikCoDalej = oConnectionString
        HQConnectionCoDalej = oConnectionString
        HQConnectionWersja = oConnectionString
        '---------------
        Return oConnectionString
    End Function





    '**************************************
    '** Connections Strings
    '**************************************

    Public Function HQDataBaseConnection() As String
        Dim ConnString As String = ""
        If ParL_HQDataBaseType = "MS ACCESS" Then
            'w ACCESS wersja przechowywana jest w Tabeli WERSJA
            'w polu IDENT, integer
            ConnString = HQACCESSDataBaseConnection()
        Else
            'w SQL wersja i nazwa programu przechowywana jest w procedurach wbudowanych
            ConnString = HQSQLDataBaseConnection()
        End If
        Return ConnString
    End Function




    Public Function HQACCESSDataBaseConnection() As String
        'Return ("Provider=""Microsoft.Jet.OLEDB.4.0"";" & _
        '    "Data Source=""" & ParL_HQPlikMDB & """;" & _
        '    "Persist Security Info=False")

        'exclusive connection
        'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\mydatabase.mdb;Mode=Share Exclusive;User Id=admin;Password=;

        If OSArchitectureIs64bit() Then
            ''ACE
            'If you are an application developer using OLEDB, set the Provider argument 
            'of the ConnectionString property to “Microsoft.ACE.OLEDB.12.0”. 

            Return ("Provider=""Microsoft.ACE.OLEDB.12.0"";" & _
                "Data Source=""" & ParL_HQPlikMDB & """;" & _
                "Persist Security Info=False;")

            'If you are an application developer using ODBC to connect to Microsoft Office Access data, 
            'set the Connection String to “Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=path to mdb/accdb file”
            'If you are an application developer using ODBC to connect to Microsoft Office Excel data, 
            'set the Connection String to “Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=path to xls/xlsx/xlsm/xlsb file”
            'Return ("Driver={Microsoft Access Driver (*.mdb, *.accdb)};" & _
            '        "DBQ=""" & ParL_HQKatZbiorow & "\" & oPlikKartoteki & """ ")
        Else
            'Jet 4,0
            Return ("Provider=""Microsoft.Jet.OLEDB.4.0"";" & _
                "Data Source=""" & ParL_HQPlikMDB & """;" & _
                "Persist Security Info=False;")
            'gdy hasło do bazy danych
            'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\mydatabase.mdb;
            'Jet OLEDB:Database Password=MyDbPassword;
        End If
    End Function




    Public Function HQSQLDataBaseConnection() As String
        Dim MyCS As String

        'baza SQL Serwer 2000
        'MyCS = "Data Source=" & oLocalHostName & ";" & _
        '       "Integrated Security=SSPI;" & _
        '       "Initial Catalog=" & oBazaDanychSlowniki & ""

        'exclusive connection
        'Data Source=MyData.sdf;File Mode=Exclusive;Persist Security Info=False;

        If ParL_HQAutentykacja = "Windows" Then
            MyCS = "Data Source=" & ParL_HQSerwer & ";" & _
                   "Integrated Security=SSPI;" & _
                   "Initial Catalog=" & ParL_HQDataBase & ""
        Else
            'MyCS = "Data Source=" & ParL_HQSerwer & ";" & _
            '       "User ID=" & PARL_HQLogin & ";" & _
            '       "Integrated Security=SSPI;" & _
            '       "Initial Catalog=" & oBazaDanychSlowniki & ""

            'MyCS = "Provider=SQLOLEDB.1;" & _
            '       "Persist Security Info=False;" & _
            '       "User ID=WCz;" & _
            '       "Initial Catalog=Softart;" & _
            '       "Data Source=SERWER1;" & _
            '       "Use Procedure for Prepare=1;" & _
            '       "Auto Translate=True;" & _
            '       "Packet Size=4096;" & _
            '       "Workstation ID=WCZNOTEBOOK;" & _
            '       "Use Encryption for Data=False;" & _
            '       "Tag with column collation when possible=False"

            MyCS = "Persist Security Info=False;" & _
                   "User ID=" & ParL_HQLogin & ";" & _
                   "Password=" & ParL_HQPassword & ";" & _
                   "Initial Catalog=" & ParL_HQDataBase & ";" & _
                   "Data Source=" & ParL_HQSerwer & ";" & _
                   "Packet Size=4096"
        End If

        Return (MyCS)
    End Function




    '**************************************
    '** sprawdzanie srodowiska
    '**************************************

    Public Function HQSrodowiskoOK() As Boolean
        If ParL_HQDataBaseType = "MS ACCESS" Then
            If IO.File.Exists(ParL_HQPlikMDB) Then
                Return (True)
            Else
                MessageBox.Show("Nie znalazłem pliku z Bazą Danych" & vbCrLf & _
                    ParL_HQPlikMDB, _
                    "Sprawdzanie środowiska", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
                Return (False)
            End If
        Else
            Return (True)
        End If
    End Function






    '**************************************
    '** sprawdzanie wersji ACCESS/MSSQL
    '**************************************

    Public Function HQWersjaBazyDanych() As Integer
        If ParL_HQDataBaseType = "MS ACCESS" Then
            'w ACCESS wersja przechowywana jest w Tabeli WERSJA
            'w polu IDENT, integer
            Return HQWersjaBazyDanychACCESS()
        Else
            'w SQL wersja i nazwa programu przechowywana jest w procedurach wbudowanych
            Return HQWersjaBazyDanychMSSQL()
        End If
    End Function


    Public Function HQWersjaBazyDanychACCESS() As Integer
        Dim sConnectionWersja As String = HQACCESSDataBaseConnection()

        Dim dbConnectionWersja As New OleDb.OleDbConnection(HQConnectionWersja)
        Dim dbCommandSelectWersja As New OleDb.OleDbCommand(sSelectSQLWersja, dbConnectionWersja)
        Dim dbDataAdapterWersja As New OleDb.OleDbDataAdapter(dbCommandSelectWersja)
        Dim DataSetWersja As New DataSet
        Dim DataViewWersja As New DataView

        Dim oWersjaDB As Integer = 0

        DataSetWersja.Locale = New System.Globalization.CultureInfo("pl-PL")

        Dim DBMapowanieTabeli As System.Data.Common.DataTableMapping
        dbDataAdapterWersja.TableMappings.Clear()
        DBMapowanieTabeli = dbDataAdapterWersja.TableMappings.Add("Table", "TABELA_Wersja")
        '------------------------------------------
        'wypelnij DATASET
        oWersjaBazyDanych = 0
        Try
            dbConnectionWersja.Open()
        Catch Ex As System.Exception
        End Try

        If dbConnectionWersja.State = ConnectionState.Open Then
            Try
                dbDataAdapterWersja.FillSchema(DataSetWersja, SchemaType.Mapped, "TABELA_Wersja")
                dbDataAdapterWersja.Fill(DataSetWersja)
                dbConnectionWersja.Close()
                'definiuj DataView
                DataViewWersja = New DataView(DataSetWersja.Tables(0))
                If DataViewWersja.Count > 0 Then
                    Dim ii As Integer
                    For ii = 0 To DataViewWersja.Count - 1
                        oWersjaDB = DataViewWersja.Item(ii).Item(0)
                    Next
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (oWersjaDB)
    End Function

    Public Function HQWersjaBazyDanychMSSQL() As Int16
        Dim oWersjaDB As Int16 = 0
        Dim oNazwaDB As String = ""
        '---------------------------------------
        'sprawdz czy to prawidłowa baza danych SoftArt -  ma procedure softart_nazwaprogramu
        Dim ir As Long = 0
        Dim OK As Boolean = False
        Dim Odpowiedz As String = ""

        Dim qConnectionString As String = HQSQLDataBaseConnection()
        Dim qSelectSQL As String = "SELECT * FROM sysobjects " & _
                                   "WHERE xtype='P'" & _
                                   "ORDER BY name"
        Dim qConnection As New SqlClient.SqlConnection(qConnectionString)
        Dim qCommandSelect As New SqlClient.SqlCommand(qSelectSQL, qConnection)
        Dim qDataAdapter As New SqlClient.SqlDataAdapter(qCommandSelect)
        Dim qDataSetTabela As New DataSet
        Dim qDataViewTabela As New DataView

        Dim sqlMapowanieTabeli As System.Data.Common.DataTableMapping
        sqlMapowanieTabeli = qDataAdapter.TableMappings.Add("Table", "TABELA")

        qDataSetTabela.Locale = New System.Globalization.CultureInfo("pl-PL")
        '------------------------------------------
        'wypelnij DATASET
        Try
            qConnection.Open()
        Catch Ex As System.Exception
            'MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & Ex.Message, "Uwaga :", _
            '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    MessageBoxIcon.Information, _
            '    MessageBoxDefaultButton.Button1)
        End Try

        If qConnection.State = ConnectionState.Open Then
            Try
                qDataAdapter.FillSchema(qDataSetTabela, SchemaType.Mapped, "TABELA")
                qDataAdapter.Fill(qDataSetTabela)
                qConnection.Close()

                'definiuj DataView
                qDataViewTabela = New DataView(qDataSetTabela.Tables(0))
                If qDataViewTabela.Count > 0 Then
                    'sprawdz czy to baza danych softart
                    For ir = 0 To qDataViewTabela.Count - 1
                        If qDataViewTabela.Item(ir).Item("NAME") = "softart_nazwaprogramu" Then
                            OK = True
                            Exit For
                        End If
                    Next
                End If
            Catch Ex As System.Exception
                'MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & Ex.Message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            'sprawdz czy to baza danych programu SOFTART-MAGAZYN
            If OK Then  'to jest baza danych softart
                Dim nConnectionString As String = HQSQLDataBaseConnection()
                Dim nStoredProc As String = "softart_nazwaprogramu"
                Dim nConnection As New SqlClient.SqlConnection(nConnectionString)
                Dim nCommand As New SqlClient.SqlCommand
                Dim nParam As SqlClient.SqlParameter
                nCommand.CommandText = nStoredProc  'xSelectSQL
                nCommand.CommandType = CommandType.StoredProcedure
                nParam = nCommand.Parameters.Add(New SqlClient.SqlParameter("@name", SqlDbType.VarChar, 50))
                nParam.Direction = ParameterDirection.Output
                nCommand.Connection = nConnection
                '---------------------------------------
                Dim wConnectionString As String = HQSQLDataBaseConnection()
                Dim wStoredProc As String = "softart_wersjaprogramu"
                Dim wConnection As New SqlClient.SqlConnection(wConnectionString)
                Dim wCommand As New SqlClient.SqlCommand
                Dim wParam As SqlClient.SqlParameter
                wCommand.CommandText = wStoredProc  'xSelectSQL
                wCommand.CommandType = CommandType.StoredProcedure
                wParam = wCommand.Parameters.Add(New SqlClient.SqlParameter("@wersja", SqlDbType.Int, Nothing))
                wParam.Direction = ParameterDirection.Output
                wCommand.Connection = wConnection
                '---------------------------------------
                Try
                    nConnection.Open()
                    nCommand.ExecuteNonQuery()
                    nConnection.Close()
                    'podstaw parametry wyjsciowe
                    oNazwaDB = CType(nCommand.Parameters("@name").Value, String)
                    If oNazwaDB = "KartotekaKlientowPryzmat" Then
                        wConnection.Open()
                        wCommand.ExecuteNonQuery()
                        wConnection.Close()
                        'podstaw parametry wyjsciowe
                        oWersjaDB = CType(wCommand.Parameters("@wersja").Value, Int16)
                    End If

                Catch ex As Exception
                    'MessageBox.Show(ex.Message)
                End Try
            Else
            End If
        End If
        Return (oWersjaDB)
    End Function

    ''********************************************************************
    ''*** testowanie wystąpienia błędu wspóbieżności
    ''*******************************************************************

    'Public Sub dbOnRowUpdated(ByVal sender As Object, ByVal args As OleDb.OleDbRowUpdatedEventArgs)
    '    Dim argsKomenda As String = ""
    '    'sprawdzamy liczbe aktualizowanych rekordow
    '    If args.RecordsAffected = 0 Then
    '        Select Case args.StatementType
    '            Case StatementType.Delete
    '                argsKomenda = "DELETE"
    '            Case StatementType.Insert
    '                argsKomenda = "INSERT"
    '            Case StatementType.Select
    '                argsKomenda = "SELECT"
    '            Case StatementType.Update
    '                argsKomenda = "UPDATE"
    '        End Select

    '        Dim aTXT As String
    '        aTXT = args.Command.ToString
    '        aTXT = args.Errors.ToString
    '        aTXT = args.Row.RowError.ToString
    '        aTXT = args.Status.ToString
    '        aTXT = args.TableMapping.ToString

    '        'opis bledu = args.Errors.ToString()
    '        'MessageBox.Show("Podczas wykonywania komendy " & argsKomenda & vbCrLf & _
    '        '                            "wystąpił błąd wspólnego dostępu do bazy danych..." & vbCrLf & _
    '        '                            "Aktualizacje rekordu " & args.Row(0) & " zostaną utracone" & vbCrLf & _
    '        '                            "i pobrany będzie bieżący rekord z bazy danych..." & vbCrLf & vbCrLf & _
    '        '                            "Proszę powtórzyć aktualizacje i spróbować ponownie zapisać rekord do bazy.", _
    '        '    "Błąd podczas aktualizacji bazy :", _
    '        '    System.Windows.Forms.MessageBoxButtons.OK, _
    '        '    MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1)

    '        MessageBox.Show("Podczas wykonywania komendy " & argsKomenda & vbCrLf & _
    '                                    "wystąpił błąd wspólnego dostępu do bazy danych..." & vbCrLf & _
    '                                    "Aktualizacje rekordu zostaną utracone" & vbCrLf & _
    '                                    "i pobrany będzie bieżący rekord z bazy danych..." & vbCrLf & vbCrLf & _
    '                                    "Proszę powtórzyć aktualizacje i spróbować ponownie zapisać rekord do bazy.", _
    '            "Błąd podczas aktualizacji bazy :", _
    '            System.Windows.Forms.MessageBoxButtons.OK, _
    '            MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1)

    '        MessageBox.Show("Podczas wykonywania komendy " & argsKomenda & vbCrLf & _
    '                                    "Command=" & args.Command.ToString & vbCrLf & _
    '                                    "Errors=" & args.Errors.ToString & vbCrLf & _
    '                                    "RowError=" & args.Row.RowError.ToString & vbCrLf & _
    '                                    "Status=" & args.Status.ToString & vbCrLf & _
    '                                    "StatementType=" & args.StatementType.ToString & vbCrLf & _
    '                                    "TableMapping=" & args.TableMapping.ToString, _
    '            "Błąd podczas aktualizacji bazy :", _
    '            System.Windows.Forms.MessageBoxButtons.OK, _
    '            MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1)

    '        'args.Row.RowError ="Optimistic Concurrency Violation Encountered"
    '        args.Status = UpdateStatus.SkipCurrentRow

    '        oWystapilConcurrencyException = True
    '    End If
    'End Sub


    'Public Sub sqlOnRowUpdated(ByVal sender As Object, ByVal args As SqlClient.SqlRowUpdatedEventArgs)
    '    Dim argsKomenda As String = ""
    '    'sprawdzamy liczbe aktualizowanych rekordow
    '    If args.RecordsAffected = 0 Then
    '        Select Case args.StatementType
    '            Case StatementType.Delete
    '                argsKomenda = "DELETE"
    '            Case StatementType.Insert
    '                argsKomenda = "INSERT"
    '            Case StatementType.Select
    '                argsKomenda = "SELECT"
    '            Case StatementType.Update
    '                argsKomenda = "UPDATE"
    '        End Select

    '        Dim aTXT As String
    '        aTXT = args.Command.ToString
    '        aTXT = args.Errors.ToString
    '        aTXT = args.Row.RowError.ToString
    '        aTXT = args.Status.ToString
    '        aTXT = args.TableMapping.ToString

    '        'opis bledu = args.Errors.ToString()
    '        MessageBox.Show("Podczas wykonywania komendy " & argsKomenda & vbCrLf & _
    '                                    "wystąpił błąd wspólnego dostępu do bazy danych..." & vbCrLf & _
    '                                    "Aktualizacje rekordu " & args.Row(0) & " zostaną utracone" & vbCrLf & _
    '                                    "i pobrany będzie bieżący rekord z bazy danych..." & vbCrLf & vbCrLf & _
    '                                    "Proszę powtórzyć aktualizacje i spróbować ponownie zapisać rekord do bazy.", _
    '            "Błąd podczas aktualizacji bazy :", _
    '            System.Windows.Forms.MessageBoxButtons.OK, _
    '            MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1)

    '        MessageBox.Show("Podczas wykonywania komendy " & argsKomenda & vbCrLf & _
    '                                    "Command=" & args.Command.ToString & vbCrLf & _
    '                                    "Errors=" & args.Errors.ToString & vbCrLf & _
    '                                    "RowError=" & args.Row.RowError.ToString & vbCrLf & _
    '                                    "Status=" & args.Status.ToString & vbCrLf & _
    '                                    "StatementType=" & args.StatementType.ToString & vbCrLf & _
    '                                    "TableMapping=" & args.TableMapping.ToString, _
    '            "Błąd podczas aktualizacji bazy :", _
    '            System.Windows.Forms.MessageBoxButtons.OK, _
    '            MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1)

    '        'args.Row.RowError ="Optimistic Concurrency Violation Encountered"
    '        args.Status = UpdateStatus.SkipCurrentRow

    '        oWystapilConcurrencyException = True
    '    End If
    'End Sub

    '********************************************************************
    '*** Ilosc rekordow w tabeli
    '*******************************************************************

    Public Function HQRecCount(ByVal TableName As String) As Long
        Dim SelectSQL As String = "SELECT count(*) AS COUNT FROM " & TableName
        Dim dbDataSet As New DataSet
        Dim dr As DataRow
        Dim Wynik As Long = 0

        dbDataSet.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_HQDataBaseType = "MS ACCESS" Then
        Else
            Dim sqlConnection As New SqlClient.SqlConnection(HQDataBaseConnection())
            Dim sqlCommandSelect As New SqlClient.SqlCommand(SelectSQL, sqlConnection)
            Dim sqlDataAdapter As New SqlClient.SqlDataAdapter(sqlCommandSelect)
            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeli As System.Data.Common.DataTableMapping
            sqlMapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_RecCount")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnection.Open()
            Catch Ex As System.Exception
            End Try

            If sqlConnection.State = ConnectionState.Open Then
                Try
                    sqlDataAdapter.FillSchema(dbDataSet, SchemaType.Mapped, "TABELA_RecCount")
                    sqlDataAdapter.Fill(dbDataSet)
                    sqlConnection.Close()
                    'definiuj dbDataView
                    If dbDataSet.Tables(0).Rows.Count > 0 Then
                        dr = dbDataSet.Tables(0).Rows.Item(0)
                        Wynik = IIf(IsDBNull(dr("COUNT")), 0, dr("COUNT"))
                    End If
                Catch Ex As System.Exception
                End Try
            End If
        End If
        Return (Wynik)
    End Function







End Module
