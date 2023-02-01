Imports System.Drawing.Printing

Module modIdentyfikacja

    '**************************************
    '** inicjowanie list
    '**************************************

    Public Sub InicjujListeWaluty(ByVal Lista As ComboBox)
        Lista.Items.Add("AUD - dolar australijski")
        Lista.Items.Add("CZK - korona czeska")
        Lista.Items.Add("DKK - korona duñska")
        Lista.Items.Add("EEK - korona estoñska")
        Lista.Items.Add("EUR - euro")
        Lista.Items.Add("JPY - jeny japoñskie")
        Lista.Items.Add("CAD - dolar kanadyjski")
        Lista.Items.Add("NOK - korona norweska")
        Lista.Items.Add("CHF - frank szwajcarski")
        Lista.Items.Add("SEK - korona szwedzka")
        Lista.Items.Add("USD - dolar amerykañski")
        Lista.Items.Add("GPB - funt brytyjski")
        Lista.Items.Add("HUF - forint wêgierski")
        Lista.Items.Add("HKD - dolar hongkoñski")
        Lista.Items.Add("CYP - funt cypryjski")
        Lista.Items.Add("UAK - hrywna ukraiñska")
        Lista.Items.Add("SKK - korona s³owacka")
        Lista.Items.Add("MTL - lira maltañska")
        Lista.Items.Add("LIT - lit litewski")
        Lista.Items.Add("LVL - ³at ³otewski")
        Lista.Items.Add("ZAR - rad po³udniowoafrykañski")
        Lista.Items.Add("RUB - rubel rosyjski")
        Lista.Items.Add("SIT - tolar s³oweñski")
    End Sub

    Public Sub InicjujListeDniTygodnia(ByVal Lista As ComboBox)
        Lista.Items.Add("Poniedzia³ek")
        Lista.Items.Add("Wtorek")
        Lista.Items.Add("Œroda")
        Lista.Items.Add("Czwartek")
        Lista.Items.Add("Pi¹tek")
        Lista.Items.Add("Sobota")
        Lista.Items.Add("Niedziela")
    End Sub




    Public Sub InicjujListeLiczb(ByVal Lista As ComboBox)
        Dim ii As Integer
        For ii = 0 To 360
            Lista.Items.Add(Trim(Str(ii)))
        Next ii
    End Sub

    Public Sub InicjujListeProcent(ByVal Lista As ComboBox)
        Dim ii As Integer
        For ii = 0 To 100
            Lista.Items.Add(Trim(Str(ii)))
        Next ii
    End Sub

    Public Sub InicjujListeLat(ByVal Lista As ComboBox)
        Dim ii As Integer
        For ii = Now.Year - 10 To Now.Year + 90
            Lista.Items.Add(Trim(Str(ii)))
        Next ii
    End Sub

    Public Sub InicjujListeMiesiecy(ByVal Lista As ComboBox)
        Dim ii As Integer
        For ii = 1 To 12
            Lista.Items.Add(Right("00" & Trim(Str(ii)), 2))
        Next ii
    End Sub

    Public Sub InicjujListeWyrownaniaTekstu(ByVal Lista As ComboBox)
        Lista.Items.Add("Do lewej")
        Lista.Items.Add("Do prawej")
        Lista.Items.Add("Centralnie")
    End Sub


    Public Sub InicjujSposobWyboruDostawcy(ByVal Lista As ComboBox)
        Lista.Items.Add("Brak danych")
        Lista.Items.Add("Wolny zakup")
        Lista.Items.Add("Zapytanie ofertowe")
        Lista.Items.Add("Przetarg")
        Lista.Items.Add("Platforma zakupowa")
    End Sub





    ''Osoby Kontaktowe
    'Public oIdentOso As String            'IDENTKLIENTA     Text, 6
    'Public oOsobaOso As String            'OSOBA            Text, 100
    'Public oVIPOso As Boolean             'VIP              boolean
    'Public oStanowiskoOso As String       'STANOWISKO       Text, 100
    'Public oKompetencjeOso As String      'KOMPETENCJE      Text, 100
    'Public oTelefonOso As String          'TELEFON          Text, 100
    'Public oTelefonKomOso As String       'TELEFONKOM       Text, 100
    'Public oeMailOso As String            'EMAIL            Text, 100
    'Public oWersjaOso As Integer          'WERSJA           Integer

    Public Function OsobaKontaktowa(ByVal klient As String) As String
        Dim dbConnection As String = ConnectionStrings()
        Dim dbSelectSQL As String = sSelectSQLOsoby & " WHERE (IDENTKLIENTA='" & klient & "')  and (VIP='TRUE') ORDER BY OSOBA"

        Dim DataSetOsoby As New DataSet
        Dim DataViewOsoby As New DataView

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
            'Dim dbConnectionOsoby As New OleDb.OleDbConnection(dbConnection)
            'Dim dbCommandSelectOsoby As New OleDb.OleDbCommand(dbSelectSQL, dbConnectionOsoby)
            'Dim dbDataAdapterOsoby As New OleDb.OleDbDataAdapter(dbCommandSelectOsoby)
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapterOsoby.TableMappings.Add("Table", "TABELA_Osoby")
            ''------------------------------------------
            'Try
            '    dbConnectionOsoby.Open()
            'Catch Ex As System.Exception
            'End Try
            'If dbConnectionOsoby.State = ConnectionState.Open Then
            '    Try
            '        dbDataAdapterOsoby.FillSchema(DataSetOsoby, SchemaType.Mapped, "TABELA_Osoby")
            '        dbDataAdapterOsoby.Fill(DataSetOsoby)
            '        dbConnectionOsoby.Close()
            '        'definiuj DataView
            '        DataViewOsoby = New DataView(DataSetOsoby.Tables(0))

            '        If DataViewOsoby.Count > 0 Then
            '            oIdentOso = GetTxtDRVField(DataViewOsoby.Item(0), "IDENTKLIENTA")
            '            oOsobaOso = GetTxtDRVField(DataViewOsoby.Item(0), "OSOBA")
            '            oVIPOso = GetLogDRVField(DataViewOsoby.Item(0), "VIP")
            '            oStanowiskoOso = GetTxtDRVField(DataViewOsoby.Item(0), "STANOWISKO")
            '            oKompetencjeOso = GetTxtDRVField(DataViewOsoby.Item(0), "KOMPETENCJE")
            '            oTelefonOso = GetTxtDRVField(DataViewOsoby.Item(0), "TELEFON")
            '            oTelefonKomOso = GetTxtDRVField(DataViewOsoby.Item(0), "TELEFONKOM")
            '            oeMailOso = GetTxtDRVField(DataViewOsoby.Item(0), "EMAIL")
            '            oWersjaOso = GetIntDRVField(DataViewOsoby.Item(0), "WERSJA")
            '        End If

            '    Catch Ex As System.Exception
            '    End Try
            'End If
        Else
            Dim sqlConnectionOsoby As New SqlClient.SqlConnection(dbConnection)
            Dim sqlCommandSelectOsoby As New SqlClient.SqlCommand(dbSelectSQL, sqlConnectionOsoby)
            Dim sqlDataAdapterOsoby As New SqlClient.SqlDataAdapter(sqlCommandSelectOsoby)
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterOsoby.TableMappings.Add("Table", "TABELA_Osoby")
            '------------------------------------------
            Try
                sqlConnectionOsoby.Open()
            Catch Ex As System.Exception
            End Try
            If sqlConnectionOsoby.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterOsoby.FillSchema(DataSetOsoby, SchemaType.Mapped, "TABELA_Osoby")
                    sqlDataAdapterOsoby.Fill(DataSetOsoby)
                    sqlConnectionOsoby.Close()
                    'definiuj DataView
                    DataViewOsoby = New DataView(DataSetOsoby.Tables(0))

                    If DataViewOsoby.Count > 0 Then
                        oIdentOso = GetTxtDRVField(DataViewOsoby.Item(0), "IDENTKLIENTA")
                        oOsobaOso = GetTxtDRVField(DataViewOsoby.Item(0), "OSOBA")
                        oVIPOso = GetLogDRVField(DataViewOsoby.Item(0), "VIP")
                        oeMailOso = GetTxtDRVField(DataViewOsoby.Item(0), "EMAIL")
                        oTelefonOso = GetTxtDRVField(DataViewOsoby.Item(0), "TELEFON")
                        oTelefonKomOso = GetTxtDRVField(DataViewOsoby.Item(0), "TELEFONKOM")
                        oStanowiskoOso = GetTxtDRVField(DataViewOsoby.Item(0), "STANOWISKO")
                        oKompetencjeOso = GetTxtDRVField(DataViewOsoby.Item(0), "KOMPETENCJE")
                        oWersjaOso = GetIntDRVField(DataViewOsoby.Item(0), "WERSJA")
                    End If

                Catch Ex As System.Exception
                End Try
            End If
        End If

        Return (Trim(oOsobaOso))
    End Function





    Public Sub IdentKlienta(ByVal Ident As String)
        Dim sConnectionKlienci As String = ConnectionStrings()
        Dim sSelectSQLKlienci As String = "SELECT * FROM Klienci WHERE IDENTKLIENTA='" & Ident & "'"

        Dim DataSetKlienci As New DataSet
        Dim DataViewKlienci As New DataView
        Dim SymbolKlienta As String = ""

        oIdentKli = ""
        oNIPKli = ""
        'oNrTOFIKli = 0
        oNrTOFITxtKli = ""
        'oKartaPaybackKli = False
        'oNryKartPaybackKli = ""
        'oKategoriaAKli = False
        'oKategoriaBKli = False
        'oKategoriaCKli = False
        oNazwa1Kli = ""
        oMiejscKli = ""
        oKodPoczKli = ""
        oUlicaKli = ""
        oNumNrDomuKli = 0
        oNrDomuKli = ""
        oIDDomuKli = ""
        'oRejonKli = ""
        oTelefonKli = ""
        oFaxKli = ""
        oeMailKli = ""
        'oOsobaKontaktowaKli = ""
        oRodzSprzetuLKli = False
        oRodzSprzetuAKli = False
        oRodzSprzetuTKli = False
        'oIloSprzetuKli = ""
        'oOfertaKli = ""
        'oZuzycieKli = ""
        oPKontaktKli = ""
        'oSkupWartoscKli = ""
        'oSkupOpiekunKli = ""
        'oSkupOKontaktRKli = Mid(SysData, 1, 4)
        'oSkupOKontaktTKli = ""
        'oSkupOKontaktDKli = ""
        'oSkupNKontaktRKli = Mid(SysData, 1, 4)
        'oSkupNKontaktTKli = ""
        'oSkupNKontaktDKli = ""
        'oSkupPlanKli = ""
        'oSkupUwagikli = ""
        'oGdzieKupujeKli = ""
        'oSprzedOpiekunkli = ""
        'oSprzedWartoscKli = ""
        'oSprzedOKontaktRKli = Mid(SysData, 1, 4)
        'oSprzedOKontaktTKli = ""
        'oSprzedOKontaktDKli = ""
        'oSprzedNKontaktRKli = Mid(SysData, 1, 4)
        'oSprzedNKontaktTKli = ""
        'oSprzedNKontaktDKli = ""
        'oSprzedUwagiKli = ""
        'oSprzedWagaKli = ""
        'oSprzedTestKli = False
        'oWpisalKli = ""
        'oDzialaniaAkcyjneKli = ""
        oUwagikli = ""
        'oMarkerKli = False
        'oMarkerPKli = False
        'oWersjaKli = 1          'WERSJA         Integer, 2


        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'Dim dbConnectionKlienci As New OleDb.OleDbConnection(sConnectionKlienci)
            'Dim dbCommandSelectKlienci As New OleDb.OleDbCommand(sSelectSQLKlienci, dbConnectionKlienci)
            'Dim dbDataAdapterKlienci As New OleDb.OleDbDataAdapter(dbCommandSelectKlienci)
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionKlienci.Open()
            'Catch Ex As System.Exception
            'End Try

            'If dbConnectionKlienci.State = ConnectionState.Open Then
            '    Try
            '        dbDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
            '        dbDataAdapterKlienci.Fill(DataSetKlienci)
            '        dbConnectionKlienci.Close()
            '        'definiuj DataView
            '        DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
            '        If DataViewKlienci.Count > 0 Then
            '            oIdentKli = GetTxtDRVField(DataViewKlienci.Item(0), "IDENTKLIENTA")
            '            oNrTOFITxtKli = GetTxtDRVField(DataViewKlienci.Item(0), "NRTOFITXT")
            '            oNazwa1Kli = GetTxtDRVField(DataViewKlienci.Item(0), "NAZWA1")
            '            oKodPoczKli = GetTxtDRVField(DataViewKlienci.Item(0), "KODPOCZTOWY")
            '            oMiejscKli = GetTxtDRVField(DataViewKlienci.Item(0), "MIEJSCOWOSC")
            '            oUlicaKli = GetTxtDRVField(DataViewKlienci.Item(0), "ULICA")
            '            oNrDomuKli = GetTxtDRVField(DataViewKlienci.Item(0), "NRDOMU")
            '            oIDDomuKli = GetTxtDRVField(DataViewKlienci.Item(0), "IDDOMU")
            '            oNIPKli = GetTxtDRVField(DataViewKlienci.Item(0), "NIP")
            '            oTelefonKli = GetTxtDRVField(DataViewKlienci.Item(0), "TELEFON")
            '            oFaxKli = GetTxtDRVField(DataViewKlienci.Item(0), "FAX")
            '            oeMailKli = GetTxtDRVField(DataViewKlienci.Item(0), "EMAIL")

            '            oRodzSprzetuLKli = GetLogDRVField(DataViewKlienci.Item(0), "RODZSPRZETUL")
            '            oRodzSprzetuAKli = GetLogDRVField(DataViewKlienci.Item(0), "RODZSPRZETUA")
            '            oRodzSprzetuTKli = GetLogDRVField(DataViewKlienci.Item(0), "RODZSPRZETUT")
            '            oPKontaktKli = GetTxtDRVField(DataViewKlienci.Item(0), "PKONTAKT")

            '            oUwagikli = GetTxtDRVField(DataViewKlienci.Item(0), "UWAGI")
            '        End If
            '    Catch Ex As System.Exception
            '    End Try
            'End If
        Else
            Dim sqlConnectionKlienci As New SqlClient.SqlConnection(sConnectionKlienci)
            Dim sqlCommandSelectKlienci As New SqlClient.SqlCommand(sSelectSQLKlienci, sqlConnectionKlienci)
            Dim sqlDataAdapterKlienci As New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienci.Open()
            Catch Ex As System.Exception
            End Try

            If sqlConnectionKlienci.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()
                    'definiuj DataView
                    DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
                    If DataViewKlienci.Count > 0 Then
                        oIdentKli = GetTxtDRVField(DataViewKlienci.Item(0), "IDENTKLIENTA")
                        oNrTOFITxtKli = GetTxtDRVField(DataViewKlienci.Item(0), "NRTOFITXT")
                        oNazwa1Kli = GetTxtDRVField(DataViewKlienci.Item(0), "NAZWA1")
                        oKodPoczKli = GetTxtDRVField(DataViewKlienci.Item(0), "KODPOCZTOWY")
                        oMiejscKli = GetTxtDRVField(DataViewKlienci.Item(0), "MIEJSCOWOSC")
                        oUlicaKli = GetTxtDRVField(DataViewKlienci.Item(0), "ULICA")
                        oNrDomuKli = GetTxtDRVField(DataViewKlienci.Item(0), "NRDOMU")
                        oIDDomuKli = GetTxtDRVField(DataViewKlienci.Item(0), "IDDOMU")
                        oNIPKli = GetTxtDRVField(DataViewKlienci.Item(0), "NIP")
                        oTelefonKli = GetTxtDRVField(DataViewKlienci.Item(0), "TELEFON")
                        oFaxKli = GetTxtDRVField(DataViewKlienci.Item(0), "FAX")
                        oeMailKli = GetTxtDRVField(DataViewKlienci.Item(0), "EMAIL")
                        oUwagikli = GetTxtDRVField(DataViewKlienci.Item(0), "UWAGI")

                        oRodzSprzetuLKli = GetLogDRVField(DataViewKlienci.Item(0), "RODZSPRZETUL")
                        oRodzSprzetuAKli = GetLogDRVField(DataViewKlienci.Item(0), "RODZSPRZETUA")
                        oRodzSprzetuTKli = GetLogDRVField(DataViewKlienci.Item(0), "RODZSPRZETUT")
                        oPKontaktKli = GetTxtDRVField(DataViewKlienci.Item(0), "PKONTAKT")
                    End If
                Catch Ex As System.Exception
                End Try
            End If
        End If
    End Sub




    Public Function IdentTofiKlienta(ByVal TOFI As String) As String
        Dim dbConnection As String = ConnectionStrings()
        Dim dbSelectSQL As String = "SELECT * FROM Klienci WHERE NRTOFITXT LIKE '%" & TOFI & "%' "

        Dim DataSetKlienci As New DataSet
        Dim DataViewKlienci As New DataView
        Dim SymbolKlienta As String = ""

        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'Dim dbConnectionKlienci As New OleDb.OleDbConnection(dbConnection)
            'Dim dbCommandSelectKlienci As New OleDb.OleDbCommand(dbSelectSQL, dbConnectionKlienci)
            'Dim dbDataAdapterKlienci As New OleDb.OleDbDataAdapter(dbCommandSelectKlienci)
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")
            ''------------------------------------------
            'Try
            '    dbConnectionKlienci.Open()
            'Catch Ex As System.Exception
            'End Try

            'If dbConnectionKlienci.State = ConnectionState.Open Then
            '    Try
            '        dbDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
            '        dbDataAdapterKlienci.Fill(DataSetKlienci)
            '        dbConnectionKlienci.Close()
            '        'definiuj DataView
            '        DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
            '        If DataViewKlienci.Count > 0 Then
            '            'SymbolKlienta = DataViewKlienci.Item(0).Item(0)
            '            SymbolKlienta = GetTxtDRVField(DataViewKlienci.Item(0), "IDENTKLIENTA")
            '        End If

            '    Catch Ex As System.Exception
            '    End Try
            'End If
        Else
            Dim sqlConnectionKlienci As New SqlClient.SqlConnection(dbConnection)
            Dim sqlCommandSelectKlienci As New SqlClient.SqlCommand(dbSelectSQL, sqlConnectionKlienci)
            Dim sqlDataAdapterKlienci As New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")
            '------------------------------------------
            Try
                sqlConnectionKlienci.Open()
            Catch Ex As System.Exception
            End Try

            If sqlConnectionKlienci.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()
                    'definiuj DataView
                    DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
                    If DataViewKlienci.Count > 0 Then
                        'SymbolKlienta = DataViewKlienci.Item(0).Item(0)
                        SymbolKlienta = GetTxtDRVField(DataViewKlienci.Item(0), "IDENTKLIENTA")
                    End If

                Catch Ex As System.Exception
                End Try
            End If
        End If

        Return (Trim(SymbolKlienta))
    End Function



    Public Function CzyJestKlient(ByVal Ident As String) As Boolean
        Dim sConnectionKlienci As String = ConnectionStrings()
        Dim sSelectSQLKlienci As String = "SELECT * FROM Klienci WHERE IDENTKLIENTA='" & Ident & "'"
        Dim Ret As Boolean = False

        Dim DataSetKlienci As New DataSet
        Dim DataViewKlienci As New DataView

        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'Dim dbConnectionKlienci As New OleDb.OleDbConnection(sConnectionKlienci)
            'Dim dbCommandSelectKlienci As New OleDb.OleDbCommand(sSelectSQLKlienci, dbConnectionKlienci)
            'Dim dbDataAdapterKlienci As New OleDb.OleDbDataAdapter(dbCommandSelectKlienci)
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionKlienci.Open()
            'Catch Ex As System.Exception
            'End Try

            'If dbConnectionKlienci.State = ConnectionState.Open Then
            '    Try
            '        dbDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
            '        dbDataAdapterKlienci.Fill(DataSetKlienci)
            '        dbConnectionKlienci.Close()
            '        'definiuj DataView
            '        DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
            '        Ret = (DataViewKlienci.Count > 0)

            '    Catch Ex As System.Exception
            '    End Try
            'End If
        Else
            Dim sqlConnectionKlienci As New SqlClient.SqlConnection(sConnectionKlienci)
            Dim sqlCommandSelectKlienci As New SqlClient.SqlCommand(sSelectSQLKlienci, sqlConnectionKlienci)
            Dim sqlDataAdapterKlienci As New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienci.Open()
            Catch Ex As System.Exception
            End Try

            If sqlConnectionKlienci.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()
                    'definiuj DataView
                    DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
                    Ret = (DataViewKlienci.Count > 0)

                Catch Ex As System.Exception
                End Try
            End If
        End If
        Return (Ret)
    End Function


    Public Sub IdentIdKlienta(ByVal Ident As String)
        Dim sConnectionKlienci As String = ConnectionStrings()
        Dim sSelectSQLKlienci As String = "SELECT * FROM Klienci WHERE IDENTKLIENTA='" & Ident & "'"

        Dim DataSetKlienci As New DataSet
        Dim DataViewKlienci As New DataView
        Dim SymbolKlienta As String = ""

        o2NrTOFITxtKli = ""
        o2IdentKli = ""
        o2Nazwa1Kli = ""
        o2KodPoczKli = ""
        o2MiejscKli = ""
        o2UlicaKli = ""

        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'Dim dbConnectionKlienci As New OleDb.OleDbConnection(sConnectionKlienci)
            'Dim dbCommandSelectKlienci As New OleDb.OleDbCommand(sSelectSQLKlienci, dbConnectionKlienci)
            'Dim dbDataAdapterKlienci As New OleDb.OleDbDataAdapter(dbCommandSelectKlienci)
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionKlienci.Open()
            'Catch Ex As System.Exception
            'End Try

            'If dbConnectionKlienci.State = ConnectionState.Open Then
            '    Try
            '        dbDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
            '        dbDataAdapterKlienci.Fill(DataSetKlienci)
            '        dbConnectionKlienci.Close()
            '        'definiuj DataView
            '        DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
            '        If DataViewKlienci.Count > 0 Then
            '            'o2IdentKli = DataViewKlienci.Item(0).Item("IDENTKLIENTA")
            '            'o2NrTOFITxtKli = DataViewKlienci.Item(0).Item("NRTOFITXT")
            '            'o2Nazwa1Kli = DataViewKlienci.Item(0).Item("NAZWA1")
            '            'o2KodPoczKli = DataViewKlienci.Item(0).Item("KODPOCZTOWY")
            '            'o2MiejscKli = DataViewKlienci.Item(0).Item("MIEJSCOWOSC")
            '            'o2UlicaKli = DataViewKlienci.Item(0).Item("ULICA")
            '            o2IdentKli = GetTxtDRVField(DataViewKlienci.Item(0), "IDENTKLIENTA")
            '            o2NrTOFITxtKli = GetTxtDRVField(DataViewKlienci.Item(0), "NRTOFITXT")
            '            o2Nazwa1Kli = GetTxtDRVField(DataViewKlienci.Item(0), "NAZWA1")
            '            o2KodPoczKli = GetTxtDRVField(DataViewKlienci.Item(0), "KODPOCZTOWY")
            '            o2MiejscKli = GetTxtDRVField(DataViewKlienci.Item(0), "MIEJSCOWOSC")
            '            o2UlicaKli = GetTxtDRVField(DataViewKlienci.Item(0), "ULICA")
            '        End If
            '    Catch Ex As System.Exception
            '    End Try
            'End If
        Else
            Dim sqlConnectionKlienci As New SqlClient.SqlConnection(sConnectionKlienci)
            Dim sqlCommandSelectKlienci As New SqlClient.SqlCommand(sSelectSQLKlienci, sqlConnectionKlienci)
            Dim sqlDataAdapterKlienci As New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienci.Open()
            Catch Ex As System.Exception
            End Try

            If sqlConnectionKlienci.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()
                    'definiuj DataView
                    DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
                    If DataViewKlienci.Count > 0 Then
                        'o2IdentKli = DataViewKlienci.Item(0).Item("IDENTKLIENTA")
                        'o2NrTOFITxtKli = DataViewKlienci.Item(0).Item("NRTOFITXT")
                        'o2Nazwa1Kli = DataViewKlienci.Item(0).Item("NAZWA1")
                        'o2KodPoczKli = DataViewKlienci.Item(0).Item("KODPOCZTOWY")
                        'o2MiejscKli = DataViewKlienci.Item(0).Item("MIEJSCOWOSC")
                        'o2UlicaKli = DataViewKlienci.Item(0).Item("ULICA")
                        o2IdentKli = GetTxtDRVField(DataViewKlienci.Item(0), "IDENTKLIENTA")
                        o2NrTOFITxtKli = GetTxtDRVField(DataViewKlienci.Item(0), "NRTOFITXT")
                        o2Nazwa1Kli = GetTxtDRVField(DataViewKlienci.Item(0), "NAZWA1")
                        o2KodPoczKli = GetTxtDRVField(DataViewKlienci.Item(0), "KODPOCZTOWY")
                        o2MiejscKli = GetTxtDRVField(DataViewKlienci.Item(0), "MIEJSCOWOSC")
                        o2UlicaKli = GetTxtDRVField(DataViewKlienci.Item(0), "ULICA")
                    End If
                Catch Ex As System.Exception
                End Try
            End If
        End If
    End Sub


    Public Function OstatniIdKlienta() As String
        Dim sConnectionKlienci As String = ConnectionStrings()
        Dim sSelectSQLKlienci As String = "SELECT * FROM Klienci ORDER BY IDENTKLIENTA DESC"

        Dim DataSetKlienci As New DataSet
        Dim DataViewKlienci As New DataView
        Dim SymbolKlienta As String = ""

        o2IdentKli = ""

        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'Dim dbConnectionKlienci As New OleDb.OleDbConnection(sConnectionKlienci)
            'Dim dbCommandSelectKlienci As New OleDb.OleDbCommand(sSelectSQLKlienci, dbConnectionKlienci)
            'Dim dbDataAdapterKlienci As New OleDb.OleDbDataAdapter(dbCommandSelectKlienci)
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionKlienci.Open()
            'Catch Ex As System.Exception
            'End Try

            'If dbConnectionKlienci.State = ConnectionState.Open Then
            '    Try
            '        dbDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
            '        dbDataAdapterKlienci.Fill(DataSetKlienci)
            '        dbConnectionKlienci.Close()

            '        'definiuj DataView
            '        DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
            '        If DataViewKlienci.Count > 0 Then
            '            o2IdentKli = DataViewKlienci.Item(0).Item("IDENTKLIENTA")
            '        End If

            '    Catch Ex As System.Exception
            '    End Try
            'End If
        Else
            Dim sqlConnectionKlienci As New SqlClient.SqlConnection(sConnectionKlienci)
            Dim sqlCommandSelectKlienci As New SqlClient.SqlCommand(sSelectSQLKlienci, sqlConnectionKlienci)
            Dim sqlDataAdapterKlienci As New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienci.Open()
            Catch Ex As System.Exception
            End Try

            If sqlConnectionKlienci.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()

                    'definiuj DataView
                    DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
                    If DataViewKlienci.Count > 0 Then
                        o2IdentKli = DataViewKlienci.Item(0).Item("IDENTKLIENTA")
                    End If

                Catch Ex As System.Exception
                End Try
            End If
        End If
        Return (o2IdentKli)
    End Function



    Public Function OstTransakcje(ByVal Ident As String, ByVal PokazIlosc As Boolean) As String

        Dim dbSelectSQL As String = "SELECT " & _
                                            "IDENTKLIENTA, " & _
                                            "'A' AS IDENTKAT, " & _
                                            "DATA AS DATASP, " & _
                                            "LILOSPRZED, " & _
                                            "AILOSPRZED, " & _
                                            "LOILOSPRZED, " & _
                                            "AOILOSPRZED " & _
                                    "FROM OBROTY WHERE IDENTKLIENTA='" & Ident & "' " & _
                                    "UNION ALL " & _
                                    "SELECT " & _
                                            "IDENTKLIENTA, " & _
                                            "'B' AS IDENTKAT, " & _
                                            "MIESIAC+'-??' AS DATASP, " & _
                                            "LILOSPRZED, " & _
                                            "AILOSPRZED, " & _
                                            "LOILOSPRZED, " & _
                                            "AOILOSPRZED " & _
                                    "FROM OBROTYMIES WHERE IDENTKLIENTA='" & Ident & "' " & _
                                    "ORDER BY IDENTKAT, DATASP DESC"

        Dim DataSetObroty As New DataSet
        Dim DataViewObroty As New DataView
        Dim Transakcje As String = ""
        Dim ii As Integer = 0

        DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'Dim dbConnectionObroty As New OleDb.OleDbConnection(sConnectionObroty)
            'Dim dbCommandSelectObroty As New OleDb.OleDbCommand(dbSelectSQL, dbConnectionObroty)
            'Dim dbDataAdapterObroty As New OleDb.OleDbDataAdapter(dbCommandSelectObroty)
            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionObroty.Open()
            'Catch Ex As System.Exception
            'End Try

            'If dbConnectionObroty.State = ConnectionState.Open Then
            '    Try
            '        dbDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
            '        dbDataAdapterObroty.Fill(DataSetObroty)
            '        dbConnectionObroty.Close()
            '        'definiuj DataView
            '        DataViewObroty = New DataView(DataSetObroty.Tables(0))
            '        If DataViewObroty.Count > 0 Then
            '            For ii = 0 To 4
            '                If ii < DataViewObroty.Count Then
            '                    If PokazIlosc Then
            '                        Transakcje &= DataViewObroty.Item(ii).Item("DATASP") & " " & _
            '                                      IIf(DataViewObroty.Item(ii).Item("LILOSPRZED") > 0, DataViewObroty.Item(ii).Item("LILOSPRZED") & "L ", "") & _
            '                                      IIf(DataViewObroty.Item(ii).Item("AILOSPRZED") > 0, DataViewObroty.Item(ii).Item("AILOSPRZED") & "A ", "") & _
            '                                      IIf(DataViewObroty.Item(ii).Item("LOILOSPRZED") > 0, DataViewObroty.Item(ii).Item("LOILOSPRZED") & "LO ", "") & _
            '                                      IIf(DataViewObroty.Item(ii).Item("AOILOSPRZED") > 0, DataViewObroty.Item(ii).Item("AOILOSPRZED") & "AO ", "") & "| "

            '                        'Transakcje &= DataViewObroty.Item(ii).Item("DATASP") & " " & _
            '                        '              DataViewObroty.Item(ii).Item("LILOSPRZED") & "L " & _
            '                        '              DataViewObroty.Item(ii).Item("AILOSPRZED") & "A " & _
            '                        '              DataViewObroty.Item(ii).Item("LOILOSPRZED") & "LO " & _
            '                        '              DataViewObroty.Item(ii).Item("AOILOSPRZED") & "AO | "
            '                    Else
            '                        Transakcje &= DataViewObroty.Item(ii).Item("DATASP") & "  "
            '                    End If
            '                End If
            '            Next ii
            '        End If

            '    Catch Ex As System.Exception
            '    End Try
            'End If
        Else
            Dim sqlConnectionObroty As New SqlClient.SqlConnection(sConnectionObroty)
            Dim sqlCommandSelectObroty As New SqlClient.SqlCommand(dbSelectSQL, sqlConnectionObroty)
            Dim sqlDataAdapterObroty As New SqlClient.SqlDataAdapter(sqlCommandSelectObroty)
            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObroty.Open()
            Catch Ex As System.Exception
            End Try

            If sqlConnectionObroty.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    sqlDataAdapterObroty.Fill(DataSetObroty)
                    sqlConnectionObroty.Close()
                    'definiuj DataView
                    DataViewObroty = New DataView(DataSetObroty.Tables(0))
                    If DataViewObroty.Count > 0 Then
                        For ii = 0 To 4
                            If ii < DataViewObroty.Count Then
                                If PokazIlosc Then
                                    Transakcje &= DataViewObroty.Item(ii).Item("DATASP") & " " & _
                                                  IIf(DataViewObroty.Item(ii).Item("LILOSPRZED") > 0, DataViewObroty.Item(ii).Item("LILOSPRZED") & "L ", "") & _
                                                  IIf(DataViewObroty.Item(ii).Item("AILOSPRZED") > 0, DataViewObroty.Item(ii).Item("AILOSPRZED") & "A ", "") & _
                                                  IIf(DataViewObroty.Item(ii).Item("LOILOSPRZED") > 0, DataViewObroty.Item(ii).Item("LOILOSPRZED") & "LO ", "") & _
                                                  IIf(DataViewObroty.Item(ii).Item("AOILOSPRZED") > 0, DataViewObroty.Item(ii).Item("AOILOSPRZED") & "AO ", "") & "| "
                                Else
                                    Transakcje &= DataViewObroty.Item(ii).Item("DATASP") & "  "
                                End If
                            End If
                        Next ii
                    End If

                Catch Ex As System.Exception
                End Try
            End If
        End If
        Return (Trim(Transakcje))
    End Function


    '**********************************************
    ' sprawdz czy jest Akcja Marketingowa o takim identyfikatorze
    '**********************************************

    Public Function CzyJestTakaAkcja(ByVal Ident As String) As Boolean
        Dim dbConnection As String = ConnectionStrings()
        Dim dbSelectSQL As String = "SELECT * FROM AkcjeOpis WHERE IDENT='" & Ident & "'"
        Dim DataSetAkcjeOps As New DataSet
        Dim DataViewAkcjeOps As New DataView
        Dim RetJest As Boolean = False

        DataSetAkcjeOps.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'Dim dbConnectionAkcjeOps As New OleDb.OleDbConnection(dbConnection)
            'Dim dbCommandSelectAkcjeOps As New OleDb.OleDbCommand(dbSelectSQL, dbConnectionAkcjeOps)
            'Dim dbDataAdapterAkcjeOps As New OleDb.OleDbDataAdapter(dbCommandSelectAkcjeOps)

            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapterAkcjeOps.TableMappings.Add("Table", "TABELA_AkcjeOps")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionAkcjeOps.Open()
            'Catch Ex As System.Exception
            'End Try

            'If dbConnectionAkcjeOps.State = ConnectionState.Open Then
            '    Try
            '        dbDataAdapterAkcjeOps.FillSchema(DataSetAkcjeOps, SchemaType.Mapped, "TABELA_AkcjeOps")
            '        dbDataAdapterAkcjeOps.Fill(DataSetAkcjeOps)
            '        dbConnectionAkcjeOps.Close()
            '        'definiuj DataView
            '        DataViewAkcjeOps = New DataView(DataSetAkcjeOps.Tables(0))
            '        RetJest = (DataViewAkcjeOps.Count > 0)
            '    Catch Ex As System.Exception
            '    End Try
            'End If
        Else
            Dim sqlConnectionAkcjeOps As New SqlClient.SqlConnection(dbConnection)
            Dim sqlCommandSelectAkcjeOps As New SqlClient.SqlCommand(dbSelectSQL, sqlConnectionAkcjeOps)
            Dim sqlDataAdapterAkcjeOps As New SqlClient.SqlDataAdapter(sqlCommandSelectAkcjeOps)

            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterAkcjeOps.TableMappings.Add("Table", "TABELA_AkcjeOps")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionAkcjeOps.Open()
            Catch Ex As System.Exception
            End Try

            If sqlConnectionAkcjeOps.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterAkcjeOps.FillSchema(DataSetAkcjeOps, SchemaType.Mapped, "TABELA_AkcjeOps")
                    sqlDataAdapterAkcjeOps.Fill(DataSetAkcjeOps)
                    sqlConnectionAkcjeOps.Close()
                    'definiuj DataView
                    DataViewAkcjeOps = New DataView(DataSetAkcjeOps.Tables(0))
                    RetJest = (DataViewAkcjeOps.Count > 0)
                Catch Ex As System.Exception
                End Try
            End If
        End If
        Return (RetJest)
    End Function


    '**********************************************
    ' sprawdz ile klientow obejmuje akcja
    '**********************************************

    Public Function IleKlientowObejmujeAkcja(ByVal Ident As String) As Long
        Dim dbConnection As String = ConnectionStrings()
        Dim dbSelectSQL As String = "SELECT * FROM AkcjeSpec WHERE IDENTAKCJI='" & Ident & "'"
        Dim DataSetAkcjeSpec As New DataSet
        Dim DataViewAkcjeSpec As New DataView
        Dim IleJest As Long = 0

        DataSetAkcjeSpec.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'Dim dbConnectionAkcjeSpec As New OleDb.OleDbConnection(dbConnection)
            'Dim dbCommandSelectAkcjeSpec As New OleDb.OleDbCommand(dbSelectSQL, dbConnectionAkcjeSpec)
            'Dim dbDataAdapterAkcjeSpec As New OleDb.OleDbDataAdapter(dbCommandSelectAkcjeSpec)

            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_AkcjeSpec")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionAkcjeSpec.Open()
            'Catch Ex As System.Exception
            'End Try

            'If dbConnectionAkcjeSpec.State = ConnectionState.Open Then
            '    Try
            '        dbDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_AkcjeSpec")
            '        dbDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
            '        dbConnectionAkcjeSpec.Close()
            '        'definiuj DataView
            '        DataViewAkcjeSpec = New DataView(DataSetAkcjeSpec.Tables(0))
            '        IleJest = (DataViewAkcjeSpec.Count)
            '    Catch Ex As System.Exception
            '    End Try
            'End If
        Else
            Dim sqlConnectionAkcjeSpec As New SqlClient.SqlConnection(dbConnection)
            Dim sqlCommandSelectAkcjeSpec As New SqlClient.SqlCommand(dbSelectSQL, sqlConnectionAkcjeSpec)
            Dim sqlDataAdapterAkcjeSpec As New SqlClient.SqlDataAdapter(sqlCommandSelectAkcjeSpec)

            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_AkcjeSpec")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionAkcjeSpec.Open()
            Catch Ex As System.Exception
            End Try

            If sqlConnectionAkcjeSpec.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_AkcjeSpec")
                    sqlDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                    sqlConnectionAkcjeSpec.Close()
                    'definiuj DataView
                    DataViewAkcjeSpec = New DataView(DataSetAkcjeSpec.Tables(0))
                    IleJest = (DataViewAkcjeSpec.Count)
                Catch Ex As System.Exception
                End Try
            End If
        End If

        Return (IleJest)
    End Function




    '-----------------------------------
    ' pobierz kontakty do zmiennej text
    ' KLIENT | DATA | OPERATOR | RODZAJ KONTaktu | UWAGI
    '----------------------------------
    'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
    'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
    'Public oPracownikKon As String        'PRACOWNIK        Text, 2     uzytkownik
    'Public oTematKon As String            'TEMAT            Text, 50
    'Public oUwagiKon As String            'UWAGI            Memo
    'Public oWersjaKon As Integer          'WERSJA           Integer
    '----------------------------------

    Public Function PobierzKontaktyDoTXT(ByVal Ident As String, _
                                         ByVal MyFont As System.Drawing.Font, _
                                         ByVal e As PrintPageEventArgs) As String
        Dim TextRET As String = ""
        Dim dbConnectionkontakty As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectkontakty As OleDb.OleDbCommand = Nothing
        'Dim dbCommandDeletekontakty As OleDb.OleDbCommand = Nothing
        'Dim dbCommandUpdatekontakty As OleDb.OleDbCommand = Nothing
        'Dim dbCommandInsertkontakty As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterkontakty As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionkontakty As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectkontakty As SqlClient.SqlCommand = Nothing
        'Dim sqlCommandDeletekontakty As SqlClient.SqlCommand = Nothing
        'Dim sqlCommandUpdatekontakty As SqlClient.SqlCommand = Nothing
        'Dim sqlCommandInsertkontakty As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterkontakty As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetkontakty As New DataSet
        Dim DataViewkontakty As New DataView

        Dim ConnectionState As System.Data.ConnectionState
        '--------------------------------
        Dim sConnectionkontakty As String = ConnectionStrings()
        Dim dbSelectSQLkontakty As String = "SELECT * FROM kontaktyHandlowe WHERE IDENTKLIENTA='" & Ident & "' " &
                                            " ORDER BY DATAKONTAKTU DESC"
        Dim i As Integer = 0
        Dim iT As Integer = 0
        Dim IleLiTXT As Integer = 0
        Dim drv As DataRowView = Nothing

        oIdentKon = Ident
        oDataKon = SysData()
        oPracownikKon = ""
        oTematKon = ""
        oUwagiKon = ""
        oWersjaKon = 1          'WERSJA         Integer, 2

        DataSetkontakty.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionkontakty = New OleDb.OleDbConnection(sConnectionkontakty)
            'dbCommandSelectkontakty = New OleDb.OleDbCommand(dbSelectSQLkontakty, dbConnectionkontakty)
            'dbDataAdapterkontakty = New OleDb.OleDbDataAdapter(dbCommandSelectkontakty)

            'Dim dbMapowanieTabelikontakty As System.Data.Common.DataTableMapping
            'dbMapowanieTabelikontakty = dbDataAdapterkontakty.TableMappings.Add("Table", "TABELA_kontakty")
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionkontakty.Open()
            'Catch Ex As System.Exception
            '    'ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnectionkontakty.State
            'End Try
        Else
            sqlConnectionkontakty = New SqlClient.SqlConnection(sConnectionkontakty)
            sqlCommandSelectkontakty = New SqlClient.SqlCommand(dbSelectSQLkontakty, sqlConnectionkontakty)
            sqlDataAdapterkontakty = New SqlClient.SqlDataAdapter(sqlCommandSelectkontakty)
            Dim sqlMapowanieTabelikontakty As System.Data.Common.DataTableMapping
            sqlMapowanieTabelikontakty = sqlDataAdapterkontakty.TableMappings.Add("Table", "TABELA_kontakty")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionkontakty.Open()
            Catch Ex As System.Exception
                'ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionkontakty.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterkontakty.FillSchema(DataSetkontakty, SchemaType.Mapped, "TABELA_kontakty")
                    'dbDataAdapterkontakty.Fill(DataSetkontakty)
                    'dbConnectionkontakty.Close()
                Else
                    sqlDataAdapterkontakty.FillSchema(DataSetkontakty, SchemaType.Mapped, "TABELA_kontakty")
                    sqlDataAdapterkontakty.Fill(DataSetkontakty)
                    sqlConnectionkontakty.Close()
                End If

                'definiuj DataView
                DataViewkontakty = New DataView(DataSetkontakty.Tables(0))
                If DataViewkontakty.Count > 0 Then
                    For i = 0 To DataViewkontakty.Count() - 1
                        drv = DataViewkontakty.Item(i)

                        'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
                        'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
                        'Public oPracownikKon As String        'PRACOWNIK        Text, 2     uzytkownik
                        'Public oTematKon As String            'TEMAT            Text, 50
                        'Public oUwagiKon As String            'UWAGI            Memo
                        'Public oWersjaKon As Integer          'WERSJA           Integer

                        oIdentKon = GetTxtDRVField(drv, "IDENTKLIENTA")
                        oDataKon = GetTxtDRVField(drv, "DATAKONTAKTU")
                        oPracownikKon = GetTxtDRVField(drv, "PRACOWNIK")
                        oTematKon = GetTxtDRVField(drv, "TEMAT")
                        oUwagiKon = GetTxtDRVField(drv, "UWAGI")
                        oWersjaKon = GetTxtDRVField(drv, "WERSJA")

                        ' KLIENT | DATA | OPERATOR | RODZAJ KONTaktu | UWAGI
                        'dodaj pierwszy wiersz
                        TextRET &= oIdentKon & "|" & _
                                   oDataKon & "|" & _
                                   oPracownikKon & "|" & _
                                   oTematKon & "|" & _
                                   DajLinieTextuNr(oUwagiKon, 1, cm2pts(6.5), MyFont, e) & _
                                   vbCrLf

                        'sprawdz czy wiecej wierszy do wydrukowania
                        IleLiTXT = IleLiniiPotrzebaNaText(oUwagiKon, cm2pts(6.5), MyFont, e)
                        If IleLiTXT > 1 Then
                            For iT = 2 To IleLiTXT
                                TextRET &= "||||" & _
                                           DajLinieTextuNr(oUwagiKon, iT, cm2pts(6.5), MyFont, e) & _
                                           vbCrLf
                            Next
                        End If
                    Next
                End If


            Catch Ex As System.Exception
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try
        End If

        Return (TextRET)
    End Function

End Module
