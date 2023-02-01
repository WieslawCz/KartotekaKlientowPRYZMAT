' Do referencji projektu trzeba dodać
' Microsoft.SqlServer.ConnectionInfo
' Microsoft.SqlServer.Management.Sdk.Sfc
Imports Microsoft.SqlServer
'Imports Microsoft.SqlServer.Management.Smo
'Imports Microsoft.SqlServer.Management.Smo.Extended
'Imports Microsoft.SqlServer.Management.Sdk.Sfc



Module _modSQLServer


    '--------------------------------------
    'Wykonanie kweredny w BazieDanych
    '--------------------------------------
    Private Sub SQLWExecuteNonQuery(ByVal Kwerenda As String)
        Dim sqlConnection As SqlClient.SqlConnection
        Dim sqlCommand As SqlClient.SqlCommand
        Dim Wynik As Integer = 0

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
    End Sub



    '--------------------------------------
    'Wyliczanie wartości w Bazie Danych
    '--------------------------------------

    Public Function SQLExecuteScalarInt(ByVal Kwerenda As String) As Integer
        Dim sqlConnection As SqlClient.SqlConnection
        Dim sqlCommand As SqlClient.SqlCommand
        Dim Wynik As Integer = 0

        sqlConnection = New SqlClient.SqlConnection(DataBaseConnection())
        sqlCommand = New SqlClient.SqlCommand
        Try
            sqlCommand.Connection = sqlConnection
            sqlCommand.CommandType = CommandType.Text
            sqlCommand.CommandText = Kwerenda
            sqlConnection.Open()
            Wynik = sqlCommand.ExecuteScalar()

        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        Finally
            sqlConnection.Close()
        End Try
        Return Wynik
    End Function


    Public Function SQLExecuteScalarDbl(ByVal Kwerenda As String) As Double
        Dim sqlConnection As SqlClient.SqlConnection
        Dim sqlCommand As SqlClient.SqlCommand
        Dim Wynik As Double? = 0

        sqlConnection = New SqlClient.SqlConnection(DataBaseConnection())
        sqlCommand = New SqlClient.SqlCommand
        Try
            sqlCommand.Connection = sqlConnection
            sqlCommand.CommandType = CommandType.Text
            sqlCommand.CommandText = Kwerenda
            sqlConnection.Open()
            Wynik = sqlCommand.ExecuteScalar()

        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        Finally
            sqlConnection.Close()
        End Try
        Return IIf(Wynik.HasValue, Wynik.Value, 0.0)
    End Function

    Public Function SQLExecuteScalarTxt(ByVal Kwerenda As String) As String
        Dim sqlConnection As SqlClient.SqlConnection
        Dim sqlCommand As SqlClient.SqlCommand
        Dim Wynik As String = ""

        sqlConnection = New SqlClient.SqlConnection(DataBaseConnection())
        sqlCommand = New SqlClient.SqlCommand
        Try
            sqlCommand.Connection = sqlConnection
            sqlCommand.CommandType = CommandType.Text
            sqlCommand.CommandText = Kwerenda
            sqlConnection.Open()
            Wynik = sqlCommand.ExecuteScalar()

        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        Finally
            sqlConnection.Close()
        End Try
        Return Wynik
    End Function

    Public Function SQLExecuteScalarLog(ByVal Kwerenda As String) As Boolean
        Dim sqlConnection As SqlClient.SqlConnection
        Dim sqlCommand As SqlClient.SqlCommand
        Dim Wynik As Boolean = False

        sqlConnection = New SqlClient.SqlConnection(DataBaseConnection())
        sqlCommand = New SqlClient.SqlCommand
        Try
            sqlCommand.Connection = sqlConnection
            sqlCommand.CommandType = CommandType.Text
            sqlCommand.CommandText = Kwerenda
            sqlConnection.Open()
            Wynik = sqlCommand.ExecuteScalar()

        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        Finally
            sqlConnection.Close()
        End Try
        Return Wynik
    End Function




















    ''parametry wspolpracy z baza danych SQL
    'Public ParL_Autentykacja As String = "" 'sposób autentykacji
    'Public ParL_Serwer As String = "" 'opis biezacej instalacji
    'Public ParL_Login As String = "" 'opis biezacej instalacji
    'Public ParL_Password As String = "" 'opis biezacej instalacji
    'Public ParL_DataBase As String = "" 'opis biezacej instalacji

    '--------------------------------------
    ' tworzy PUSTĄ strukturę Bazy Danych na podstawie struktury
    ' bieżąco obslugiwanej Bazy
    ' SERWER : parL_serwer
    ' UPRAWNIENIA : ParL_Autentykacja / Parl_Login / Parl_Password
    ' BAZA ŻRÓDŁOWA : Parl_DataBase
    ' BAZA DOCELOWA : pDestDataBase
    ' CZY KOPOIOWAC DANE : pCopyData
    '---------------------------------------

    'Public Function KlonujBazeDanychSQL(ByVal pCoKlonowac As String,
    '                               ByVal pDestDatabaseName As String,
    '                               ByVal pCopyData As Boolean) As Boolean
    '    Dim Result As Boolean = False
    '    Dim SrcDataBaseServer As String = ""
    '    Dim SrcDataBaseAuthentication As String = ""
    '    Dim SrcDataBaseLogin As String = ""
    '    Dim SrcDataBasePass As String = ""
    '    Dim SrcDataBaseName As String = ""

    '    If pCoKlonowac = "Aktualna Baza Danych" Then
    '        SrcDataBaseServer = ParL_Serwer
    '        SrcDataBaseAuthentication = ParL_Autentykacja
    '        SrcDataBaseLogin = ParL_Login
    '        SrcDataBasePass = ParL_Password
    '        SrcDataBaseName = ParL_DataBase
    '    Else
    '        SrcDataBaseServer = ParL_HQSerwer
    '        SrcDataBaseAuthentication = ParL_HQAutentykacja
    '        SrcDataBaseLogin = ParL_HQLogin
    '        SrcDataBasePass = ParL_HQPassword
    '        SrcDataBaseName = ParL_HQDataBase
    '    End If
    '    '-----------------------
    '    Dim SrcServer As Server
    '    Try
    '        ' Set Source SQL Server Instance Information
    '        SrcServer = New Server(SrcDataBaseServer)
    '        If SrcDataBaseAuthentication = "Windows" Then
    '            'autentykacja Windows
    '            SrcServer.ConnectionContext.LoginSecure = True
    '            'SrcServer.ConnectionContext.Login = ""
    '            'SrcServer.ConnectionContext.Password = ""
    '        Else
    '            'autentykacja sql
    '            SrcServer.ConnectionContext.LoginSecure = False
    '            SrcServer.ConnectionContext.Login = SrcDataBaseLogin
    '            SrcServer.ConnectionContext.Password = SrcDataBasePass
    '        End If
    '        Dim SqlServerVersion As Integer = SrcServer.Information.VersionMajor

    '        ' Set Source Database Name [Database to Copy]
    '        Dim SrcDataBase As Database = SrcServer.Databases(SrcDataBaseName)

    '        Dim transfer As New Transfer(SrcDataBase)          ' Set Transfer Class Source Database
    '        transfer.CopyAllObjects = True                     ' Yes I want to Copy All the Database Objects
    '        transfer.CopyAllDefaults = True                    ' kopiuj wszystkie defaults
    '        transfer.DropDestinationObjectsFirst = True        ' In case if the Destination Database / Objects Exists Drop them First
    '        'transfer.UseDestinationTransaction = True
    '        transfer.CopySchema = True                         ' Copy Database Schema
    '        transfer.CopyData = pCopyData                      ' Copy Database Data Get Value from bCopyData Parameter

    '        transfer.Options.AnsiFile = True
    '        transfer.Options.Indexes = True
    '        transfer.Options.DriAll = True
    '        transfer.Options.SchemaQualify = True
    '        transfer.Options.WithDependencies = False
    '        transfer.Options.IncludeIfNotExists = True        ' Include If Not Exists Clause in the Script

    '        transfer.DestinationServer = ParL_Serwer           ' Set Destination SQL Server Instance Name
    '        If ParL_Autentykacja = "Windows" Then
    '            'autentykacja Windows
    '            transfer.DestinationLoginSecure = True
    '            'transfer.DestinationLogin = ""
    '            'transfer.DestinationPassword = ""
    '        Else
    '            'autentykacja sql
    '            transfer.DestinationLoginSecure = False
    '            transfer.DestinationLogin = ParL_Login
    '            transfer.DestinationPassword = ParL_Password
    '        End If
    '        transfer.DestinationDatabase = pDestDatabaseName   ' Set Destination Database Name
    '        transfer.CreateTargetDatabase = True               ' Create The Database in Destination Server

    '        '' załóż pusta Bazę Danych
    '        'Dim DstDataBase As New Microsoft.SqlServer.Management.Smo.Database(SrcServer, pDestDatabaseName)        ' Set Destination Database Name
    '        'DstDataBase.Create()                                                 ' Create Empty Database at Destination

    '        transfer.TransferData()               ' Start Transfer
    '        Result = True






    '        '//Set Source SQL Server Instance Information
    '        'Server Server = New Server(DBHelper.SourceSQLServer); 

    '        '//Set Source Database Name [Database to Copy]
    '        'Database Database = Server.Databases[DBHelper.SourceDatabase]; 

    '        '//Set Transfer Class Source Database
    '        'transfer transfer = New Transfer(Database);

    '        '//Yes I want to Copy All the Database Objects
    '        'transfer.CopyAllObjects = True;

    '        '//In case if the Destination Database / Objects Exists Drop them First
    '        'transfer.DropDestinationObjectsFirst = True;

    '        '//Copy Database Schema
    '        'transfer.CopySchema = True;

    '        '//Copy Database Data Get Value from bCopyData Parameter
    '        'transfer.CopyData = bCopyData;

    '        '//Set Destination SQL Server Instance Name
    '        'transfer.DestinationServer = DBHelper.DestinationSQLServer;

    '        '//Create The Database in Destination Server
    '        'transfer.CreateTargetDatabase = True; 

    '        '//Set Destination Database Name
    '        'Database ddatabase = New Database(Server, DBHelper.DestinationDatabase);

    '        '//Create Empty Database at Destination
    '        'ddatabase.Create();

    '        '//Set Destination Database Name
    '        'transfer.DestinationDatabase = DBHelper.DestinationDatabase;

    '        '//Include If Not Exists Clause in the Script
    '        'transfer.Options.IncludeIfNotExists = True; 

    '        '//Start Transfer
    '        'transfer.TransferData();

    '        '//Release Server variable
    '        'Server = null;


    '    Catch Ex As System.Exception
    '        MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & Ex.Message, "Uwaga :",
    '            System.Windows.Forms.MessageBoxButtons.OK,
    '            MessageBoxIcon.Information,
    '            MessageBoxDefaultButton.Button1)
    '    Finally
    '        SrcServer = Nothing                   ' Release Server variable
    '    End Try
    '    Return Result
    'End Function


    'Public Sub KlonujBazeDanychSQL(ByVal pDestDatabaseName As String)

    '    Dim SrcDataBaseServer As String = ParL_Serwer
    '    Dim SrcDataBaseAuthentication As String = ParL_Autentykacja
    '    Dim SrcDataBaseLogin As String = ParL_Login
    '    Dim SrcDataBasePass As String = ParL_Password
    '    Dim SrcDataBaseName As String = ParL_DataBase
    '    '-----------------------
    '    Dim SrcServer As Server
    '    Try
    '        ' Set Source SQL Server Instance Information
    '        SrcServer = New Server(SrcDataBaseServer)
    '        If SrcDataBaseAuthentication = "Windows" Then
    '            'autentykacja Windows
    '            SrcServer.ConnectionContext.LoginSecure = True
    '            'SrcServer.ConnectionContext.Login = ""
    '            'SrcServer.ConnectionContext.Password = ""
    '        Else
    '            'autentykacja sql
    '            SrcServer.ConnectionContext.LoginSecure = False
    '            SrcServer.ConnectionContext.Login = SrcDataBaseLogin
    '            SrcServer.ConnectionContext.Password = SrcDataBasePass
    '        End If
    '        Dim SqlServerVersion As Integer = SrcServer.Information.VersionMajor

    '        ' Set Source Database Name [Database to Copy]
    '        Dim SrcDataBase As Database = SrcServer.Databases(SrcDataBaseName)

    '        Dim transfer As New Transfer(SrcDataBase)          ' Set Transfer Class Source Database
    '        transfer.CopyAllObjects = True                     ' Yes I want to Copy All the Database Objects
    '        transfer.CopyAllDefaults = True                    ' kopiuj wszystkie defaults
    '        transfer.DropDestinationObjectsFirst = True        ' In case if the Destination Database / Objects Exists Drop them First
    '        'transfer.UseDestinationTransaction = True
    '        transfer.CopySchema = True                         ' Copy Database Schema
    '        transfer.CopyData = True                           ' Copy Database Data Get Value from bCopyData Parameter

    '        transfer.Options.AnsiFile = True
    '        transfer.Options.Indexes = True
    '        transfer.Options.DriAll = True
    '        transfer.Options.SchemaQualify = True
    '        transfer.Options.WithDependencies = False
    '        transfer.Options.IncludeIfNotExists = True        ' Include If Not Exists Clause in the Script

    '        transfer.DestinationServer = ParL_Serwer           ' Set Destination SQL Server Instance Name
    '        If ParL_Autentykacja = "Windows" Then
    '            'autentykacja Windows
    '            transfer.DestinationLoginSecure = True
    '            'transfer.DestinationLogin = ""
    '            'transfer.DestinationPassword = ""
    '        Else
    '            'autentykacja sql
    '            transfer.DestinationLoginSecure = False
    '            transfer.DestinationLogin = ParL_Login
    '            transfer.DestinationPassword = ParL_Password
    '        End If
    '        transfer.DestinationDatabase = pDestDatabaseName   ' Set Destination Database Name
    '        'transfer.CreateTargetDatabase = True               ' Create The Database in Destination Server

    '        ' załóż pusta Bazę Danych
    '        Dim DstDataBase As New Microsoft.SqlServer.Management.Smo.Database(SrcServer, pDestDatabaseName)        ' Set Destination Database Name
    '        DstDataBase.Create()                                                 ' Create Empty Database at Destination

    '        transfer.TransferData()               ' Start Transfer


    '    Catch Ex As System.Exception
    '        MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & Ex.Message, "Uwaga :",
    '            System.Windows.Forms.MessageBoxButtons.OK,
    '            MessageBoxIcon.Information,
    '            MessageBoxDefaultButton.Button1)
    '    Finally
    '        SrcServer = Nothing                   ' Release Server variable
    '    End Try
    'End Sub








    '*******************************************
    ' obsługa bazy SQL
    '*******************************************

    Public Function CzyJestPolaczenieZSQL() As Boolean
        Dim sqlConnectionStr As String = DataBaseConnection()
        Dim sqlConnection As New SqlClient.SqlConnection(sqlConnectionStr)
        Dim sqlCommand As New SqlClient.SqlCommand
        Dim NazwaSerwera As String = ""

        Try
            sqlCommand.Connection = sqlConnection
            sqlCommand.CommandType = CommandType.Text
            sqlConnection.Open()
            '------------------------------------------------
            sqlCommand.CommandText = "SELECT @@SERVERNAME"
            NazwaSerwera = sqlCommand.ExecuteScalar()
            '------------------------------------------------

        Catch ex As Exception
            MessageBox.Show("Błąd połączenia z bazą danych SQL" & vbCrLf & "na serwerze " & ParL_Serwer & "...", "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
            'MessageBox.Show(ex.Message)
        Finally
            sqlConnection.Close()
        End Try

        Return (Len(NazwaSerwera) > 0)
    End Function




    '===========================================
    ' Czy jest już BAZA DANYCH na bieżącym serwerze...
    '===========================================

    ''Add reference To Microsoft.SqlServer.Management
    'Try
    'Dim srv As New Smo.Server("serverAddress")
    ''Server exists
    'If srv.Databases.Contains("dbName") Then
    ''dbExists
    'Else
    ''db not exists
    'End If
    'Catch ex As Exception
    ''server not exists
    'End Try

    Public Function CzyJestJuzTakaBazaDanychSQL(ByVal NazwaBazy As String) As Boolean
        Dim sqlConnectionStr As String = DataBaseConnection()
        Dim sqlConnection As New SqlClient.SqlConnection(sqlConnectionStr)
        Dim JestBaza As Boolean = False

        Dim sSelectSQL As String = "exec sp_catalogs_rowset;2 NULL"
        Dim sqlCommandSelect As New SqlClient.SqlCommand(sSelectSQL, sqlConnection)
        Dim sqlDataAdapter As New SqlClient.SqlDataAdapter(sqlCommandSelect)
        Dim sqlDataSet As New DataSet
        Dim sqlDataView As New DataView

        sqlDataSet.Locale = New System.Globalization.CultureInfo("pl-PL")

        Dim DBMapowanieTabeli As System.Data.Common.DataTableMapping
        DBMapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_Database")
        'wypelnij DATASET

        Try
            sqlConnection.Open()
        Catch Ex As System.Exception
        End Try

        If sqlConnection.State = ConnectionState.Open Then
            Try
                sqlDataAdapter.FillSchema(sqlDataSet, SchemaType.Mapped, "TABELA_Database")
                sqlDataAdapter.Fill(sqlDataSet)
                sqlConnection.Close()

                'definiuj sqlDataView
                sqlDataView = New DataView(sqlDataSet.Tables(0))
                If sqlDataView.Count > 0 Then
                    Dim dr As DataRow
                    Dim db As String
                    For Each dr In sqlDataView.Table.Rows
                        db = dr.Item("CATALOG_NAME")
                        If UCase(db) = UCase(NazwaBazy) Then
                            JestBaza = True
                            Exit For
                        End If
                    Next
                End If
            Catch Ex As System.Exception
                MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            Finally
                If sqlConnection.State = ConnectionState.Open Then
                    sqlConnection.Close()
                End If
            End Try
        End If
        '------------------------------------------------
        Return (JestBaza)
    End Function











    'pobranie katalogów plików bazy danych
    'Select Case name, physical_name As current_file_location
    'FROM sys.master_files

    Public Function PobierzKatalogZPlikamiSQL() As String
        Dim sqlConnectionStr As String = DataBaseConnection()
        Dim sqlConnection As New SqlClient.SqlConnection(sqlConnectionStr)
        Dim Katalog As String = ""

        'Dim sSelectSQL As String = "select name, physical_name from sys.database_files"
        Dim sSelectSQL As String = "select name, filename from master.dbo.sysdatabases where name like 'master'"
        Dim sqlCommandSelect As New SqlClient.SqlCommand(sSelectSQL, sqlConnection)
        Dim sqlDataAdapter As New SqlClient.SqlDataAdapter(sqlCommandSelect)
        Dim sqlDataSet As New DataSet
        Dim sqlDataView As New DataView

        sqlDataSet.Locale = New System.Globalization.CultureInfo("pl-PL")
        Dim sqlMapowanieTabeli As System.Data.Common.DataTableMapping
        sqlMapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_Database")
        Try
            sqlConnection.Open()
        Catch Ex As System.Exception
        End Try

        If sqlConnection.State = ConnectionState.Open Then
            Try
                sqlDataAdapter.FillSchema(sqlDataSet, SchemaType.Mapped, "TABELA_Database")
                sqlDataAdapter.Fill(sqlDataSet)
                sqlConnection.Close()

                'definiuj sqlDataView
                sqlDataView = New DataView(sqlDataSet.Tables(0))
                If sqlDataView.Count > 0 Then
                    Dim dr As DataRow
                    Dim dbn As String = ""
                    Dim pos As Integer = 0
                    For Each dr In sqlDataView.Table.Rows
                        dbn = GetTxtDRField(dr, "NAME")
                        If Trim(dbn) = "master" Then
                            Katalog = GetTxtDRField(dr, "FILENAME")
                            pos = InStr(Katalog, "master")
                            If pos > 0 Then
                                Katalog = Mid(Katalog, 1, pos - 1)
                            End If
                            Exit For
                        End If
                    Next
                End If
            Catch Ex As System.Exception
                MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            Finally
                If sqlConnection.State = ConnectionState.Open Then
                    sqlConnection.Close()
                End If
            End Try
        End If
        '------------------------------------------------
        Return (Katalog)
    End Function











    '================================
    ' TIMEOUT'y MS SQL
    '================================
    '1. Ustawienia po stronie MS SQL
    '''''''''''''''''''''''''''''''''
    'MS SQL Serwer stwarza możliwość ustalenia limitu czasu (ang timeout)
    'potrzebnego na wykonanie pojedynczego zapytania jezyka SQL. Domyślna
    'wartość limitu tego parametru wynosi 600 sekund, czyli 10 minut. Oznacza to,
    'że wykonanie każdego zapytania SQL, które nie zakończy się w czasie poniżej
    '600 sekund, zostanie przerwane i anulowane. Takie przerwanie może
    'uniemożliwić wykonanie niektórych operacji. Sytuacja ta
    'może zaistnieć, gdy program próbuje wykonać zapytanie SQL na dużej ilości
    'danych, na zakończenie którego potrzeba więcej czasu niż ustalony limit (np.
    'podczas konwersji bazy danych, archiwizacji lub odtwarzania dużej ilosci
    'dokumentów). Problem ten można rozwiązać zwiększajac limit czasu na
    'serwerze. W tym celu można użyc narzędzia osql.exe dostarczanego wraz z
    'motorem bazy danych MS SQL Server.
    '
    'Aby zmienić limit czasu na serwerze MS SQL, uruchom na komputerze z
    'zainstalowanym motorem bazy danych okno wiersza poleceń systemu Windows
    'i wprowadz polecenie wedug poniższego wzoru:
    '
    'C:\>osql -S "<serwer>" -U "<uzytkownik>" -P "<haslo>" -Q
    '    "sp_configure @configname = 'remote query timeout', @configvalue = '<wartosc_limitu>'"
    '
    'Opis poszczególnych parametrów:
    '• <serwer> - nazwa serwera MS SQL z bazą programu (Z230-HP)
    '• <uzytkownik> - nazwa użytkownika serwera MS SQL, który ma uprawnienia do wykonywania poleceń
    '       konfiguracyjnych na serwerze MSSQL (np. użytkownik „sa”)
    '• <haslo> - haslo użytkownika <uzytkownik>
    '• <wartosc_limitu> - liczba sekund, do jakiej chcemy zwiększyć limit
    '       czasu wykonywania zapytań (domyślnie 600, maksymalnie 2147483647).
    '
    'Po zmianie konfiguracji za pomocą procedury sp_configure, należy wykonać
    'polecenie RECONFIGURE, aby serwer zaczal dzialać z nowymi ustawieniami
    'parametru.
    '
    'C:\>osql -S "<serwer>" -U "<uytkownik>" -P "<haslo>" –Q "RECONFIGURE"
    '
    'AGORA Bytom :
    'C:\>sqlcmd -S "Z230-HP" -U "sa" -P "AGORABytom" -Q
    '    "sp_configure @configname = 'remote query timeout', @configvalue = '3600'"
    '
    'C:\>sqlcmd -S "Z230-HP" -U "sa" -P "AGORABytom" –Q "RECONFIGURE"
    '===================================================================
    '2. Ustawienie po stronie Visual Studio...
    ''''''''''''''''''''''''''''''''''''''''''
    'Increase the Query timeout and Connection timeout values in Visual Studio using the procedures documented below. 
    'Changing the Connection Timeout: - Limit czasu na otwarcie kwerendy
    '1.In Visual Studio IDE, enable Server Explorer by navigating to View ->Server Explorer
    '2.In the Server Explorer, right click on the connection to SQL Server where the CLR objects are being deployed and choose Modify Connection.
    '3.Click on Advanced button on the Modify Connection window.
    '4.In the Advanced Properties window change the Connect Timeout value under Initialization section to a higher value.
    '
    'Changing the Query Timeout: - limit czasu na wykonanie kwerendy
    '1.In Visual Studio IDE, navigate to Tools -> Options ->Database Tools ->Query and View Designers
    '2.You can either uncheck the option Cancel long running query or change the value of Cancel after option to a higher value.
    '===================================================================
    '3. Ustawienie w programie
    ''''''''''''''''''''''''''''''''''''''''''
    'Connection Timeout - czas na uzyskanie połączenia
    'Autentykacja WINDOWS
    'MyCS = "Data Source=" & ParL_Serwer & ";" & _
    '   "Integrated Security=SSPI;Connection Timeout=30;" & _
    '   "Initial Catalog=" & ParL_DataBase & ""
    'Autentykacja SQL
    'MyCS = "Persist Security Info=False;" & _
    '   "User ID=" & ParL_Login & ";" & _
    '   "Password=" & ParL_Password & ";" & _
    '   "Initial Catalog=" & ParL_DataBase & ";" & _
    '   "Data Source=" & ParL_Serwer & ";" & _
    '   "Connection Timeout=30;" & _
    '   "Packet Size=4096"
    '
    'Query Timeout - czas na uzyskanie wyników zapytania...
    '
    '===================================================================
    '4. Ustawienie w rejestrach Windows
    ''''''''''''''''''''''''''''''''''''
    ' You have to increase the query timeout allowed on the connection. 
    ' Unfortunately there is no user-interface-exposed way to do this. You have to search the Windows Registry 
    ' for the key "QueryTimeoutSeconds" and increase the value. I increased mine from 60 to 360 and that made 
    ' my schema compare timeouts disappear in Visual Studio 2012 and Visual Studio 2013.
    '===================================================================







    'ConnectionTimeout specifies how long, in seconds, should the code wait before timing out from trying to OPEN a connection. It relates directly to the line connection.Open();
    'CommandTimeout specified how long, in seconds, should the command wait before timing out. This relates to calls such as Fill(), ExecuteXXX (Reader, Scalar, NoQuery) and such. It's relelated to the actual SQL code you are trying to run (be it inline or an sproc).
    'CommandTimeout would seem to be your problem.You can't globally set it like the Connect Time. But you can create a utlity function, so that the code is centralized.
    'set the commandTimeout in AppSettings.... and use this utility function
    '
    'internal static SqlCommand GetSqlCommand()
    '{
    'SqlCommand command = new SqlCommand();
    'command.CommandTimeout =
    'ConfigurationSettings.AppSettings["commandTimeout"];
    'return command;
    '}




    'Parametry czasu obsługi kwerend
    '================MOJAFIRMA
    ' Instalacja            Server SQL                  Visual Studio           Program Comnn String    Komputer
    '                       Remote Query Timeout        SQL Connect Timeout     Connection Timeout      QueryTimeoutSeconds
    '                       domyślnie=600 sek                                   domyślnie=30 sek
    '--------------------------------------------------------------------------------------------------------------------
    'TRES Server Win2012    3600                                                600
    '--------------------------------------------------------------------------------------------------------------------


    '================OBSŁUGA PARKINGU
    ' Instalacja            Server SQL                  Visual Studio           Program Comnn String    Komputer
    '                       Remote Query Timeout        SQL Connect Timeout     Connection Timeout      QueryTimeoutSeconds
    '--------------------------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------------------------



    '==================================================
    'Zwiększenie CZASU WYKONANIA KWERENDY SELEC T - pobranie danych z bazy
    'Domyślnie jest 39 sek
    '
    '            If ParL_DataBaseType = "MS ACCESS" Then
    '            Else
    '               'sqlConnectionDokumenty = New SqlClient.SqlConnection(sConnectionDokumentyOpis)
    '                sqlConnectionDokumenty = New SqlClient.SqlConnection(SQLDataBaseConnection("120"))
    '                sqlCommandSelectDokumenty = New SqlClient.SqlCommand(dbSelectSQLDokumenty, sqlConnectionDokumenty)
    '               'sqlCommandDeleteDokumenty = New sqlclient.sqlCommand(sDeleteSQLDokumenty, sqlconnectionDokumenty)
    '               'sqlCommandUpdateDokumenty = New sqlclient.sqlCommand(sUpdateSQLDokumenty, sqlconnectionDokumenty)
    '               'sqlCommandInsertDokumenty = New sqlclient.sqlCommand(sInsertSQLDokumenty, sqlconnectionDokumenty)
    '                sqlDataAdapterDokumenty = New SqlClient.SqlDataAdapter(sqlCommandSelectDokumenty)

    '                DataSetDokumenty = New DataSet
    '                DataViewDokumenty = New DataView

    '                DataSetDokumenty.Locale = New System.Globalization.CultureInfo("pl-PL")

    '                Dim sqlMapowanieTabeliDokumenty As System.Data.Common.DataTableMapping
    '                sqlMapowanieTabeliDokumenty = sqlDataAdapterDokumenty.TableMappings.Add("Table", "TABELA_Dokumenty")

    '               '------------------------------------------
    '               'wypelnij DATASET
    '                Try
    '                    sqlCommandSelectDokumenty.CommandTimeout = 120      '!!!!!!!!!!!!!!!!!!!!!!!!!!!!Określenie czasu wykonania komendy SELECT....
    '                    sqlConnectionDokumenty.Open()
    '                Catch Ex As System.Exception
    '                    ViewInLog(Ex, Me.Name(), Nothing)
    '                    'MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & ex.message, "Uwaga :", _
    '                    '    System.Windows.Forms.MessageBoxButtons.OK, _
    '                    '    MessageBoxIcon.Information, _
    '                    '    MessageBoxDefaultButton.Button1)
    '                Finally
    '                    ConnectionState = sqlConnectionDokumenty.State
    '                End Try
    '            End If
    '==================================================













    ' select n Column z Tabeli...
    '---------------------------------------
    '   Create procedure TopNcolumns (@tableName varchar(100),@n int) 
    'as 
    'Declare @s varchar(2000) 
    'set @s='' 
    'If exists(Select * from information_Schema.tables where table_name=@tablename and table_type='Base Table')
    'Begin
    'If @n>=0 
    'Begin 
    'set rowcount @n 
    'Select @s=@s+','+ column_name from information_schema.columns 
    'where table_name=@tablename order by ordinal_position 
    'Set rowcount 0 
    'Set @s=substring(@s,2,len(@s)-1) 
    'Exec('Select '+@s+' from '+@tablename) 
    'End 
    'else 
    'Select 'Negative values are not allowed' as Error 
    'End
    'else
    'Select 'Table '+@tableName+' does not exist' as Error

    'exec topncolumns @tableName='DokumentySpec',@n=10 





    '------------------------
    ' Zapisz info w logu SQL Servera - SEVERITY=INFOTRMATIONAL
    '------------------------


    Public Sub MSSQL_LogInfo(ByVal parINFO As String)
        Dim Wynik As Integer = 0
        Dim sqlCommand As SqlClient.SqlCommand = Nothing
        Try
            'sqlConnection = New SqlClient.SqlConnection(SQLDataBaseConnection())
            sqlCommand = New SqlClient.SqlCommand
            sqlCommand.Connection = New SqlClient.SqlConnection(DataBaseConnection())
            sqlCommand.CommandType = CommandType.Text
            'sqlCommand.CommandText = sqlAlterParametry123_1
            sqlCommand.CommandText = "xp_logevent 123123 , '" & oNazwaProgramu & " " & parINFO & "' , informational"
            sqlCommand.Connection.Open()
            Wynik = sqlCommand.ExecuteNonQuery()
            If Wynik = -1 Then
            End If

        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        Finally
            sqlCommand.Connection.Close()
        End Try
    End Sub










    '=========================
    ' RESTORE DATABASE
    '================================
    'To restore a database to a New location, And optionally rename the database
    '1. Connect to the appropriate instance of the SQL Server Database Engine, And then in Object Explorer, click the server name to expandthe server tree.
    '2. Right-click Databases, And then click Restore Database. The Restore Database dialog box opens.
    '3. On the General page, use the Source section to specify the source And location of the backup sets to restore. Select one of thefollowing options:
    '•Database  •Select the database to restore from the drop-down list. The list contains only databases that have been backed up according to the msdb backup history.
    '           Note If the backup Is taken from a different server, the destination server will Not have the backup history informationfor the specified database. 
    '           In this case, select Device to manually specify the file Or device to restore.
    '•Device    •Click the browse (...) button to open the Select backup devices dialog box. 
    '           In the Backup media type box, select one of the listed device types. 
    '           To select one Or more devices for the Backup media box, click Add. 
    '           After you add the devices you want To the Backup media list box, click OK To Return To the General page. 
    '           In the Source Device: Database list box, Select the name Of the database which should be restored.
    '           Note This list Is only available When Device Is selected. Only databases that have backups On the selected device will be available.
    '4. In the Destination section, the Database box Is automatically populated with the name of the database to be restored.To change the name of the database, enter the New name in the Database box.
    '5. In the Restore to box, leave the default as To the last backup taken Or click on Timeline to access the Backup Timeline dialog box to manually select a point in time to stop the recovery action. 
    '6. In the Backup sets to restore grid, select the backups to restore. This grid displays the backups available for the specified location. By default, a recovery plan Is suggested. 
    '   To override the suggested recovery plan, you can change the selections in the grid.Backups that depend on the restoration of an earlier backup are automatically deselected when the earlier backup Is deselected.
    '7. To specify the New location of the database files, select the Files page, And then click Relocate all files to folder. 
    '   Provide a New location For the Data file folder And Log filefolder. Alternatively you can keep the same folders 
    '================================
    'Imports System.Data.SqlClient

    'Public Class Form1
    '    Dim con As SqlConnection
    '    Dim cmd As SqlCommand
    '    Dim dread As SqlDataReader

    '    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '        server(".")
    '        server(".\sqlexpress")
    '    End Sub

    '    Sub server(ByVal str As String)
    '        con = New SqlConnection("Data Source=" & str & ";Database=Master;integrated security=SSPI;")
    '        con.Open()
    '        cmd = New SqlCommand("select *  from sysservers  where srvproduct='SQL Server'", con)
    '        dread = cmd.ExecuteReader
    '        While dread.Read
    '            cmbserver.Items.Add(dread(2))
    '        End While
    '        dread.Close()
    '    End Sub

    '    Sub connection()
    '        con = New SqlConnection("Data Source=" & Trim(cmbserver.Text) & ";Database=Master;integrated security=SSPI;")
    '        con.Open()
    '        cmbdatabase.Items.Clear()
    '        cmd = New SqlCommand("select * from sysdatabases", con)
    '        dread = cmd.ExecuteReader
    '        While dread.Read
    '            cmbdatabase.Items.Add(dread(0))
    '        End While
    '        dread.Close()
    '    End Sub

    '    Private Sub cmbserver_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbserver.SelectedIndexChanged
    '        connection()
    '    End Sub

    '    Sub query(ByVal que As String)
    '        On Error Resume Next
    '        cmd = New SqlCommand(que, con)
    '        cmd.ExecuteNonQuery()
    '    End Sub

    '    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    '        If ProgressBar1.Value = 100 Then
    '            Timer1.Enabled = False
    '            ProgressBar1.Visible = False
    '            MsgBox("Successfully Done")
    '        Else
    '            ProgressBar1.Value = ProgressBar1.Value + 5
    '        End If
    '    End Sub

    '    Sub blank(ByVal str As String)
    '        If cmbserver.Text = "" Or cmbdatabase.Text = "" Then
    '            MsgBox("Server Name & Database Blank Field")
    '            Exit Sub
    '        Else
    '            If str = "backup" Then
    '                SaveFileDialog1.FileName = cmbdatabase.Text
    '                SaveFileDialog1.ShowDialog()
    '                Timer1.Enabled = True
    '                ProgressBar1.Visible = True
    '                Dim s As String
    '                s = SaveFileDialog1.FileName
    '                query("backup database " & cmbdatabase.Text & " to disk='" & s & "'")
    '            ElseIf str = "restore" Then
    '                OpenFileDialog1.ShowDialog()
    '                Timer1.Enabled = True
    '                ProgressBar1.Visible = True
    '                query("RESTORE DATABASE " & cmbdatabase.Text & " FROM disk='" & OpenFileDialog1.FileName & "'")
    '            End If
    '        End If
    '    End Sub

    '    Private Sub cmbbackup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbbackup.Click
    '        blank("backup")
    '    End Sub

    '    Private Sub cmdrestore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdrestore.Click
    '        blank("restore")
    '    End Sub
    'End Class


End Module
