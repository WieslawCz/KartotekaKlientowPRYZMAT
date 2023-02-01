Public Class frmAnalizaCzestKontaktowGrupowa

    '*********************************************************
    '** Analiza kontaktow handlowych
    '*********************************************************
    Dim StartFormy As Boolean = True
    Dim _OCoMamRobic As String = oCoMamRobic

    '---------------------------------------------
    'SCROLL-parametry obslugi scrollingu poziomego
    Dim StartPointInCell00 As System.Drawing.Point             'początek ektanu przed scrollingiem
    Dim ScrollXOffset As Integer = 0            'bieżący punkt ekranu wskazywany przez kursor podczas scrollingu
    Dim HorizScrollLastTime As String = ""      'Chwila rozpoczecia scrollingu - filtry wyswietlaner będą po 1 sekundzie...
    ' HorizScrollTimer - zegar odliczający czas od poczatku scrollingu
    '---------------------------------------------

    Dim dbSelectSQLKontakty As String = sSelectSQLKontakty

    Dim dbConnectionKontakty As OleDb.OleDbConnection
    Dim dbCommandSelectKontakty As OleDb.OleDbCommand
    Dim dbCommandDeleteKontakty As OleDb.OleDbCommand
    Dim dbCommandUpdateKontakty As OleDb.OleDbCommand
    Dim dbCommandInsertKontakty As OleDb.OleDbCommand
    Dim dbDataAdapterKontakty As OleDb.OleDbDataAdapter

    Dim sqlConnectionKontakty As SqlClient.SqlConnection
    Dim sqlCommandSelectKontakty As SqlClient.SqlCommand
    Dim sqlCommandDeleteKontakty As SqlClient.SqlCommand
    Dim sqlCommandUpdateKontakty As SqlClient.SqlCommand
    Dim sqlCommandInsertKontakty As SqlClient.SqlCommand
    Dim sqlDataAdapterKontakty As SqlClient.SqlDataAdapter

    Dim DataSetKontakty As New DataSet
    Dim DataViewKontakty As New DataView
    Dim ConnectionState As System.Data.ConnectionState
    '--------------------------------
    Dim dbSelectSQLKlienci As String = ""       'sSelectSQLKlienci

    Dim dbConnectionKlienci As OleDb.OleDbConnection
    Dim dbCommandSelectKlienci As OleDb.OleDbCommand
    Dim dbCommandDeleteKlienci As OleDb.OleDbCommand
    Dim dbCommandUpdateKlienci As OleDb.OleDbCommand
    Dim dbCommandInsertKlienci As OleDb.OleDbCommand
    Dim dbDataAdapterKlienci As OleDb.OleDbDataAdapter

    Dim sqlConnectionKlienci As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienci As SqlClient.SqlCommand
    Dim sqlCommandDeleteKlienci As SqlClient.SqlCommand
    Dim sqlCommandUpdateKlienci As SqlClient.SqlCommand
    Dim sqlCommandInsertKlienci As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienci As SqlClient.SqlDataAdapter

    Dim DataSetKlienci As New DataSet
    Dim DataViewKlienci As New DataView
    '-------------------------

    Dim DataDo As String = SysData()
    Dim Dataod As String = WyliczDate(SysData(), -7)
    Dim IleKontaktowOd As Integer = 1
    Dim IleKontaktowDo As Integer = 1
    Dim ListaRK As String = ""

    'Zaloz robocze struktury...
    Dim DataSetRaport As DataSet = Nothing
    Dim robTable As DataTable = Nothing
    Dim DataViewRaport As DataView = Nothing
    Dim robRow As DataRow = Nothing
    Dim robRowView As DataRowView = Nothing

    Private Sub frmAnalizaCzestKontaktowGrupowa_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        InicjujListeLiczb(cbxIleKontaktowOd)
        InicjujListeLiczb(cbxIleKontaktowDo)

        dtpOdDaty.Value = Dataod
        dtpDoDaty.Value = DataDo

        cbxIleKontaktowOd.Text = Trim(Str(IleKontaktowOd))
        cbxIleKontaktowDo.Text = Trim(Str(IleKontaktowDo))

        rabLiczbaMniejRowne.Checked = True

        txtRodzajeKontaktow.Text = ""
        rabRKWszystkie.Checked = True

        '-----------------
        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionKlienci = New OleDb.OleDbConnection(sConnectionKlienci)
            'dbCommandSelectKlienci = New OleDb.OleDbCommand(dbSelectSQLKlienci, dbConnectionKlienci)
            ''dbCommandDeleteKlienci = New OleDb.OleDbCommand(sDeleteSQLKlienci, dbConnectionKlienci)
            ''dbCommandUpdateKlienci = New OleDb.OleDbCommand(sUpdateSQLKlienci, dbConnectionKlienci)
            ''dbCommandInsertKlienci = New OleDb.OleDbCommand(sInsertSQLKlienci, dbConnectionKlienci)
            'dbDataAdapterKlienci = New OleDb.OleDbDataAdapter(dbCommandSelectKlienci)

            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeliKlienci As System.Data.Common.DataTableMapping
            'MapowanieTabeliKlienci = dbDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")


            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapterKlienci.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionKlienci.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnectionKlienci.State
            'End Try

        Else
            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
            sqlCommandSelectKlienci = New SqlClient.SqlCommand(sSelectSQLKlienciAktywni, sqlConnectionKlienci)
            'sqlCommandDeleteKlienci = New SqlClient.SqlCommand(sDeleteSQLKlienci, sqlConnectionKlienci)
            'sqlCommandUpdateKlienci = New SqlClient.SqlCommand(sUpdateSQLKlienci, sqlConnectionKlienci)
            'sqlCommandInsertKlienci = New SqlClient.SqlCommand(sInsertSQLKlienci, sqlConnectionKlienci)
            sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliKlienci As System.Data.Common.DataTableMapping
            MapowanieTabeliKlienci = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterKlienci.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienci.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & ex.message, "Uwaga :", _
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
                    dbDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    dbDataAdapterKlienci.Fill(DataSetKlienci)
                    dbConnectionKlienci.Close()
                Else
                    sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                'DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("IDENTKLIENTA"), _
                '                                                         DataSetKlienci.Tables(0).Columns("DATAKONTAKTU")}
                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))

            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try
            '---------------------------------
        End If







        '-----------------
        DataSetKontakty.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionKontakty = New OleDb.OleDbConnection(sConnectionKontakty)
            dbCommandSelectKontakty = New OleDb.OleDbCommand(dbSelectSQLKontakty, dbConnectionKontakty)
            'dbCommandDeleteKontakty = New OleDb.OleDbCommand(sDeleteSQLKontakty, dbConnectionKontakty)
            'dbCommandUpdateKontakty = New OleDb.OleDbCommand(sUpdateSQLKontakty, dbConnectionKontakty)
            'dbCommandInsertKontakty = New OleDb.OleDbCommand(sInsertSQLKontakty, dbConnectionKontakty)
            dbDataAdapterKontakty = New OleDb.OleDbDataAdapter(dbCommandSelectKontakty)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliKontakty As System.Data.Common.DataTableMapping
            MapowanieTabeliKontakty = dbDataAdapterKontakty.TableMappings.Add("Table", "TABELA_Kontakty")


            ' Add the RowUpdated event handler.
            AddHandler dbDataAdapterKontakty.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionKontakty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = dbConnectionKontakty.State
            End Try

        Else
            sqlConnectionKontakty = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectKontakty = New SqlClient.SqlCommand(dbSelectSQLKontakty, sqlConnectionKontakty)
            'sqlCommandDeleteKontakty = New SqlClient.SqlCommand(sDeleteSQLKontakty, sqlConnectionKontakty)
            'sqlCommandUpdateKontakty = New SqlClient.SqlCommand(sUpdateSQLKontakty, sqlConnectionKontakty)
            'sqlCommandInsertKontakty = New SqlClient.SqlCommand(sInsertSQLKontakty, sqlConnectionKontakty)
            sqlDataAdapterKontakty = New SqlClient.SqlDataAdapter(sqlCommandSelectKontakty)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliKontakty As System.Data.Common.DataTableMapping
            MapowanieTabeliKontakty = sqlDataAdapterKontakty.TableMappings.Add("Table", "TABELA_Kontakty")


            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterKontakty.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKontakty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionKontakty.State
            End Try
        End If
        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    dbDataAdapterKontakty.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_Kontakty")
                    dbDataAdapterKontakty.Fill(DataSetKontakty)
                    dbConnectionKontakty.Close()
                Else
                    sqlDataAdapterKontakty.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_Kontakty")
                    sqlDataAdapterKontakty.Fill(DataSetKontakty)
                    sqlConnectionKontakty.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetKontakty.Tables(0).PrimaryKey = New DataColumn() {DataSetKontakty.Tables(0).Columns("UNIQID")}
                DataViewKontakty = New DataView(DataSetKontakty.Tables(0))

            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try
            '---------------------------------
        End If






        'dodaj StatusBar na koncu
        stbRaport.BackColor = System.Drawing.SystemColors.Control
        stbRaport.ForeColor = System.Drawing.Color.DarkGreen
        stbRaport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                Or System.Windows.Forms.AnchorStyles.Left) _
                                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        stbRaport.Panels.Add("stbPanelIleRec")
        stbRaport.Panels(0).AutoSize = StatusBarPanelAutoSize.None
        stbRaport.Panels(0).Width = 80
        stbRaport.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
        stbRaport.Panels(0).Text = "Il.zap.: "

        'stbRAPORT.Panels.Add("stbPanelWiersz")
        'stbRAPORT.Panels(1).AutoSize = StatusBarPanelAutoSize.None
        'stbRAPORT.Panels(1).Width = 80
        'stbRAPORT.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
        'stbRAPORT.Panels(1).Text = "Wiersz : " & dagRAPORT.CurrentCell.RowNumber.ToString

        'stbRAPORT.Panels.Add("stbPanelKolumna")
        'stbRAPORT.Panels(2).AutoSize = StatusBarPanelAutoSize.None
        'stbRAPORT.Panels(2).Width = 80
        'stbRAPORT.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
        'stbRAPORT.Panels(2).Text = "Kolumna : " & dagRAPORT.CurrentCell.ColumnNumber.ToString

        'stbRAPORT.Panels.Add("stbPanelSort")
        'stbRAPORT.Panels(3).AutoSize = StatusBarPanelAutoSize.None
        'stbRAPORT.Panels(3).Width = 150
        'stbRAPORT.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
        'stbRAPORT.Panels(3).Text = "Sort : "

        'stbRAPORT.Panels.Add("stbPanelFiltr")
        'stbRAPORT.Panels(4).AutoSize = StatusBarPanelAutoSize.None
        'stbRAPORT.Panels(4).Width = 150
        'stbRAPORT.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
        'stbRAPORT.Panels(4).Text = "Filtr : "

        stbRaport.Panels.Add("stbPanelReszta")
        stbRaport.Panels(1).AutoSize = StatusBarPanelAutoSize.Spring
        'stbRAPORT.Panels(1).Width = 150
        stbRaport.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
        stbRaport.Panels(1).Text = ""

        stbRaport.ShowPanels = True
        '---------------------------------
        Timer1.Interval = 200
        Timer1.Enabled = True
    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        cmdAnalizuj_Click(sender, e)
    End Sub




    Private Sub cmdStop_Click(sender As Object, e As EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub

    Private Sub frmAnalizaCzestKontaktowGrupowa_Paint(sender As Object, e As PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub






    Private Sub cmdAnalizuj_Click(sender As Object, e As EventArgs) Handles cmdAnalizuj.Click
        Dim robPH As String = ""

        Me.Cursor = Cursors.WaitCursor
        Dataod = dtpOdDaty.Value.ToString("yyyy-MM-dd")
        DataDo = dtpDoDaty.Value.ToString("yyyy-MM-dd")
        IleKontaktowOd = CInt(cbxIleKontaktowOd.Text)
        IleKontaktowDo = CInt(cbxIleKontaktowDo.Text)
        ListaRK = txtRodzajeKontaktow.Text
        '----------------------
        'Zaloz robocze struktury...
        DataSetRaport = New DataSet()
        robTable = New DataTable()
        DataViewRaport = Nothing
        robRow = Nothing
        robRowView = Nothing

        'dbCommandUpdate.Parameters.Add("@oIdentKli", OleDb.OleDbType.Char, 6, "NRKARTY")
        'dbCommandUpdate.Parameters.Add("@oNrTofiTxtKli", OleDb.OleDbType.Char, 100, "NRTOFITXT")
        'dbCommandUpdate.Parameters.Add("@oNAZWA1Kli", OleDb.OleDbType.Char, 100, "NAZWAKLIENTA")
        'dbCommandUpdate.Parameters.Add("@oMiejscKli", OleDb.OleDbType.Char, 50, "MIEJSCOWOSC")
        'dbCommandUpdate.Parameters.Add("@oKodPoczKli", OleDb.OleDbType.Char, 10, "KODPOCZTOWY")
        'dbCommandUpdate.Parameters.Add("@oUlicaKli", OleDb.OleDbType.Char, 50, "ULICA")
        'dbCommandUpdate.Parameters.Add("@oNrDomuKli", OleDb.OleDbType.Char, 10, "NRDOMU")
        'dbCommandUpdate.Parameters.Add("@oIDDomuKli", OleDb.OleDbType.Char, 10, "IDDOMU")
        'dbCommandUpdate.Parameters.Add("@oRejonKli", OleDb.OleDbType.Char, 20, "REJONKLIENTA")


        Dim robCol1 As DataColumn = robTable.Columns.Add("KONTAKTY", GetType(System.Int32))
        Dim robCol2 As DataColumn = robTable.Columns.Add("PRZEDSTAWICIEL", GetType(System.String))
        Dim robCol3 As DataColumn = robTable.Columns.Add("IDENTKLIENTA", GetType(System.String))
        Dim robCol4 As DataColumn = robTable.Columns.Add("NRTOFI", GetType(System.String))
        Dim robCol5 As DataColumn = robTable.Columns.Add("NAZWA", GetType(System.String))
        Dim robCol6 As DataColumn = robTable.Columns.Add("MIEJSCOWOSC", GetType(System.String))
        Dim robCol7 As DataColumn = robTable.Columns.Add("KODPOCZTOWY", GetType(System.String))
        Dim robCol8 As DataColumn = robTable.Columns.Add("ULICA", GetType(System.String))
        Dim robCol9 As DataColumn = robTable.Columns.Add("NRDOMU", GetType(System.String))
        Dim robColA As DataColumn = robTable.Columns.Add("IDDOMU", GetType(System.String))
        Dim robColB As DataColumn = robTable.Columns.Add("REJON", GetType(System.String))
        Dim robColC As DataColumn = robTable.Columns.Add("MARKER", GetType(System.Boolean))
        Dim robColD As DataColumn = robTable.Columns.Add("MARKERP", GetType(System.Boolean))
        DataSetRaport.Tables.Add(robTable)
        'definiuj DataView
        DataViewRaport = New DataView(DataSetRaport.Tables(0))
        DataViewRaport.Sort = "IDENTKLIENTA"
        '-------------------------------------------
        Dim drv As DataRowView
        Dim rowNo As Integer = 0
        Dim Analizuj As Boolean = False
        Dim ileLi As Integer = 0
        Dim RodzajK As String = ""
        Dim pos As Integer = 0
        Dim i As Integer = 0


        'najpierw dopisujemy do listy wszystkich klientów z ilością kontaktów=0
        DataViewKontakty.Sort = "IDENTKLIENTA"
        If DataViewKlienci.Count > 0 Then
            For Each drv In DataViewKlienci
                oIdentKon = GetTxtDRVField(drv, "NRKARTY")

                rowNo = DataViewKontakty.Find(oIdentKon)
                'nie było kontaktu...
                If rowNo < 0 Then
                    ' DOPISUJEMY TEKO KLIENTA
                    robRow = DataSetRaport.Tables(0).NewRow()

                    robRow("PRZEDSTAWICIEL") = ""
                    robRow("KONTAKTY") = 0
                    robRow("IDENTKLIENTA") = GetTxtDRVField(drv, "NRKARTY")
                    robRow("NRTOFI") = GetTxtDRVField(drv, "NRTOFITXT")
                    robRow("NAZWA") = GetTxtDRVField(drv, "NAZWAKLIENTA")
                    robRow("MIEJSCOWOSC") = GetTxtDRVField(drv, "MIEJSCOWOSC")
                    robRow("KODPOCZTOWY") = GetTxtDRVField(drv, "KODPOCZTOWY")
                    robRow("ULICA") = GetTxtDRVField(drv, "ULICA")
                    robRow("NRDOMU") = GetTxtDRVField(drv, "NRDOMU")
                    robRow("IDDOMU") = GetTxtDRVField(drv, "IDDOMU")
                    robRow("REJON") = GetTxtDRVField(drv, "REJONKLIENTA")
                    robRow("MARKER") = False
                    robRow("MARKERP") = False

                    DataSetRaport.Tables(0).Rows.Add(robRow)
                End If
            Next
        End If
        '-----------------------------


        'Public oUniqIdKon As String           'UNIQID            varchar(40)
        'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
        'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
        'Public oPracownikKon As String        'PRACOWNIK        Text, 10     uzytkownik
        'Public oTematKon As String            'TEMAT            Text, 50
        'Public oUwagiKon As String            'UWAGI            Memo
        'Public oWersjaKon As Integer          'WERSJA           Integer

        If DataViewKontakty.Count > 0 Then
            For Each drv In DataViewKontakty
                oIdentKon = GetTxtDRVField(drv, "IDENTKLIENTA")
                oDataKon = GetTxtDRVField(drv, "DATAKONTAKTU")
                oPracownikKon = GetTxtDRVField(drv, "PRACOWNIK")
                oTematKon = GetTxtDRVField(drv, "TEMAT")
                oUwagiKon = GetTxtDRVField(drv, "UWAGI")

                'czy rodzaj kontaktu do analizy...
                If rabRKWszystkie.Checked Then
                    Analizuj = True
                ElseIf rabRKLista.Checked Then
                    'sprawdz czy Rodzaj Kontaktu z listy
                    Analizuj = False
                    ileLi = IleLiniiMaText(ListaRK)
                    If ileLi > 0 Then
                        For i = 1 To ileLi
                            RodzajK = UCase(LiniaTxtNr(ListaRK, i))
                            pos = InStr(RodzajK, "*")
                            If UCase(oTematKon) = RodzajK Then
                                Analizuj = True
                                Exit For
                            ElseIf pos > 1 Then
                                'jest * w szukanym Rodzaju Kontaktu - wyrzuć wszystko od gwiazdki - sprawdz czy zawiera się odo pocz tekstu
                                RodzajK = Mid(RodzajK, 1, pos - 1)
                                If InStr(UCase(oTematKon), RodzajK, ) = 1 Then
                                    Analizuj = True
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                Else
                End If


                'jesli miesci się między datami - dopisz ilościowo....
                If oDataKon <= DataDo And oDataKon >= Dataod And Analizuj Then
                    IdentKlienta(oIdentKon)

                    rowNo = DataViewRaport.Find(oIdentKon)
                    If rowNo < 0 Then
                        ' DOPISUJEMY TEKO KLIENTA
                        robRow = DataSetRaport.Tables(0).NewRow()
                        robRow("PRZEDSTAWICIEL") = oPracownikKon
                        robRow("KONTAKTY") = 1
                        robRow("IDENTKLIENTA") = oIdentKon
                        robRow("NRTOFI") = oNrTOFITxtKli
                        robRow("NAZWA") = oNazwa1Kli
                        robRow("MIEJSCOWOSC") = oMiejscKli
                        robRow("KODPOCZTOWY") = oKodPoczKli
                        robRow("ULICA") = oUlicaKli
                        robRow("NRDOMU") = oNrDomuKli
                        robRow("IDDOMU") = oIDDomuKli
                        robRow("REJON") = oRejonKli
                        robRow("MARKER") = False
                        robRow("MARKERP") = False
                        DataSetRaport.Tables(0).Rows.Add(robRow)
                    Else
                        'dopisz informacje o kontaktach...
                        robRowView = DataViewRaport.Item(rowNo)
                        robRowView("KONTAKTY") += 1

                        robPH = robRowView("PRZEDSTAWICIEL")
                        If InStr(robPH, oPracownikKon) < 0 Then
                            robRowView("PRZEDSTAWICIEL") &= " " & oPracownikKon
                        End If
                    End If
                End If
            Next
        End If
        '-----------------------------
        'mamy ilościowo w wybranym okresie - analizuj zależnie od relacji
        Dim ii As Integer
        If DataViewRaport.Count() > 0 Then
            If rabLiczbaRowne.Checked Then
                'usun zapisy rozne ilosci analizowanych kontaktow
                For ii = DataViewRaport.Count To 1 Step -1
                    drv = DataViewRaport.Item(ii - 1)
                    If Not (drv("KONTAKTY") = IleKontaktowOd) Then
                        drv.Delete()
                    End If
                Next

            ElseIf rabLiczbaRozne.Checked Then
                'usun zapisy rowne ilosci analizowanych kontaktow
                For ii = DataViewRaport.Count To 1 Step -1
                    drv = DataViewRaport.Item(ii - 1)
                    If Not (drv("KONTAKTY") <> IleKontaktowOd) Then
                        drv.Delete()
                    End If
                Next

            ElseIf rabLiczbaMniesze.Checked Then
                'usun zapisy wieksze i rowne ilosci analizowanych kontaktow
                For ii = DataViewRaport.Count To 1 Step -1
                    drv = DataViewRaport.Item(ii - 1)
                    If Not (drv("KONTAKTY") < IleKontaktowOd) Then
                        drv.Delete()
                    End If
                Next

            ElseIf rabLiczbaMniejRowne.Checked Then
                'usun zapisy wieksze od ilosci analizowanych kontaktow
                For ii = DataViewRaport.Count To 1 Step -1
                    drv = DataViewRaport.Item(ii - 1)
                    If Not (drv("KONTAKTY") <= IleKontaktowOd) Then
                        drv.Delete()
                    End If
                Next

            ElseIf rabLiczbaWieksze.Checked Then
                'usun zapisy mniejsze i rowne od ilosci analizowanych kontaktow
                For ii = DataViewRaport.Count To 1 Step -1
                    drv = DataViewRaport.Item(ii - 1)
                    If Not (drv("KONTAKTY") > IleKontaktowOd) Then
                        drv.Delete()
                    End If
                Next

            ElseIf rabLiczbaWieRowne.Checked Then
                'usun zapisy mniejsze od ilosci analizowanych kontaktow
                For ii = DataViewRaport.Count To 1 Step -1
                    drv = DataViewRaport.Item(ii - 1)
                    If Not (drv("KONTAKTY") >= IleKontaktowOd) Then
                        drv.Delete()
                    End If
                Next

            ElseIf rabLiczbaOdDo.Checked Then
                'usun zapisy od do analizowanych kontaktow
                For ii = DataViewRaport.Count To 1 Step -1
                    drv = DataViewRaport.Item(ii - 1)
                    If Not ((drv("KONTAKTY") >= IleKontaktowOd) And (drv("KONTAKTY") <= IleKontaktowDo)) Then
                        drv.Delete()
                    End If
                Next

            End If
        End If
        '-----------------------------
        '-------------------------------------------------
        'parametry DataGridView
        '-------------------------------------------------
        ' Initialize basic DataGridView properties.
        dagRaport.BackgroundColor = System.Drawing.SystemColors.Control
        dagRaport.GridColor = System.Drawing.SystemColors.ControlDark
        dagRaport.ForeColor = System.Drawing.Color.Purple
        dagRaport.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
        dagRaport.BorderStyle = BorderStyle.Fixed3D
        'dagraport.Dock = DockStyle.Fill
        dagRaport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

        ' Set property values appropriate for read-only display and limited interactivity. 
        dagRaport.AllowUserToAddRows = False
        dagRaport.AllowUserToDeleteRows = False
        dagRaport.AllowUserToOrderColumns = True
        dagRaport.AllowUserToResizeColumns = True
        dagRaport.AllowUserToResizeRows = True
        dagRaport.ReadOnly = True
        dagRaport.SelectionMode = DataGridViewSelectionMode.CellSelect
        dagRaport.MultiSelect = False

        'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
        'dagraport.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        dagRaport.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
        dagRaport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        dagRaport.ColumnHeadersVisible = True
        dagRaport.ColumnHeadersHeight = 40
        dagRaport.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
        dagRaport.RowHeadersVisible = True
        dagRaport.RowHeadersWidth = 50
        dagRaport.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

        ' Set the selection background color for all the cells.
        dagRaport.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
        dagRaport.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
        dagRaport.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
        dagRaport.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
        dagRaport.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
        dagRaport.DefaultCellStyle.NullValue = ""
        dagRaport.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        dagRaport.DefaultCellStyle.WrapMode = DataGridViewTriState.True


        ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
        ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
        dagRaport.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
        dagRaport.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

        ' Set the background color for all rows and for alternating rows.  
        ' The value for alternating rows overrides the value for all rows. 
        dagRaport.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
        dagRaport.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

        ' Set the row and column header styles.
        dagRaport.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
        dagRaport.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
        dagRaport.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
        dagRaport.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
        '-----------------------------------
        'wypelnienie DataGridView
        dagRaport.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
        'dagRaport.DataMember = DataSetRaport.Tables(0).TableName
        'dagraport.DataSource = DataViewraport
        '-----------------------------------

        dagRaport.Columns.Clear()

        TxtColumnView(dagRaport, DataSetRaport.Tables(0).Columns(0).ColumnName, "Kontakty", 60, HorizontalAlignment.Left, True)
        TxtColumnView(dagRaport, DataSetRaport.Tables(0).Columns(1).ColumnName, "Przedstaw.", 60, HorizontalAlignment.Left, True)
        TxtColumnView(dagRaport, DataSetRaport.Tables(0).Columns(2).ColumnName, "Id. klienta", 60, HorizontalAlignment.Left, True)
        TxtColumnView(dagRaport, DataSetRaport.Tables(0).Columns(3).ColumnName, "Nr TOFI", 200, HorizontalAlignment.Left, True)
        TxtColumnView(dagRaport, DataSetRaport.Tables(0).Columns(4).ColumnName, "Nazwa", 200, HorizontalAlignment.Left, True)
        TxtColumnView(dagRaport, DataSetRaport.Tables(0).Columns(5).ColumnName, "Miejscowoś", 200, HorizontalAlignment.Left, True)
        TxtColumnView(dagRaport, DataSetRaport.Tables(0).Columns(6).ColumnName, "Kod pocztowy", 60, HorizontalAlignment.Left, True)
        TxtColumnView(dagRaport, DataSetRaport.Tables(0).Columns(7).ColumnName, "Ulica", 200, HorizontalAlignment.Left, True)
        TxtColumnView(dagRaport, DataSetRaport.Tables(0).Columns(8).ColumnName, "Nr domu", 60, HorizontalAlignment.Left, True)
        TxtColumnView(dagRaport, DataSetRaport.Tables(0).Columns(9).ColumnName, "Id domu", 60, HorizontalAlignment.Left, True)
        TxtColumnView(dagRaport, DataSetRaport.Tables(0).Columns(10).ColumnName, "Rejon", 2000, HorizontalAlignment.Left, True)
        LogColumnView(dagRaport, DataSetRaport.Tables(0).Columns(11).ColumnName, "Marker", 40, True)
        LogColumnView(dagRaport, DataSetRaport.Tables(0).Columns(12).ColumnName, "Marker P", 40, True)

        Me.Cursor = Cursors.WaitCursor
        dagRaport.DataSource = DataViewRaport
        Me.Cursor = Cursors.Default

        'ustaw sie na pierwszym zapisie
        If DataSetRaport.Tables(0).Rows.Count > 0 Then
            dagRaport.CurrentCell = dagRaport(0, 0)
            dagRaport.CurrentCell.Selected = True
        End If




        '------------------------------------
        'dodaj StatusBar na koncu
        stbRaport.Panels.Clear()

        stbRaport.BackColor = System.Drawing.SystemColors.Control
        stbRaport.ForeColor = System.Drawing.Color.DarkGreen
        stbRaport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        stbRaport.Panels.Add("stbPanelIleRec")
        stbRaport.Panels(0).AutoSize = StatusBarPanelAutoSize.None
        stbRaport.Panels(0).Width = 80
        stbRaport.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
        stbRaport.Panels(0).Text = "Ilość rec.: " & DataSetRaport.Tables(0).Rows.Count.ToString

        stbRaport.Panels.Add("stbPanelWiersz")
        stbRaport.Panels(1).AutoSize = StatusBarPanelAutoSize.None
        stbRaport.Panels(1).Width = 80
        stbRaport.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
        stbRaport.Panels(1).Text = "Wiersz : " & dagRaport.CurrentCell.RowIndex.ToString

        stbRaport.Panels.Add("stbPanelKolumna")
        stbRaport.Panels(2).AutoSize = StatusBarPanelAutoSize.None
        stbRaport.Panels(2).Width = 80
        stbRaport.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
        stbRaport.Panels(2).Text = "Kolumna : " & dagRaport.CurrentCell.ColumnIndex.ToString

        stbRaport.Panels.Add("stbPanelSort")
        stbRaport.Panels(3).AutoSize = StatusBarPanelAutoSize.None
        stbRaport.Panels(3).Width = 150
        stbRaport.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
        stbRaport.Panels(3).Text = "Sort : "

        stbRaport.Panels.Add("stbPanelFiltr")
        stbRaport.Panels(4).AutoSize = StatusBarPanelAutoSize.None
        stbRaport.Panels(4).Width = 150
        stbRaport.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
        stbRaport.Panels(4).Text = "Filtr : "

        stbRaport.Panels.Add("stbPanelReszta")
        stbRaport.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
        'stbRaport.Panels(5).Width = 150
        stbRaport.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
        stbRaport.Panels(5).Text = ""

        stbRaport.ShowPanels = True





        '---------------------------------
        'dodaj StatusBar na koncu
        stbFiltry.Panels.Clear()

        stbFiltry.BackColor = System.Drawing.SystemColors.Control
        stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
        stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

        stbFiltry.Panels.Add("stbFiltryRowHead")
        stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
        stbFiltry.Panels(0).Width = 50  'dagRaport.TableStyles(0).RowHeaderWidth
        stbFiltry.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
        stbFiltry.Panels(0).Text = "Filtry"

        stbFiltry.Panels.Add("stbFiltryReszta")
        stbFiltry.Panels(1).AutoSize = StatusBarPanelAutoSize.Spring
        'stbFiltry.Panels(1).Width = 150
        stbFiltry.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
        stbFiltry.Panels(1).Text = ""

        stbFiltry.ShowPanels = True

        'dodaj ctrl do pobierania szablonu
        'cmdWszystko.Size = New Size(20, 20)
        'cmdWszystko.Location = New system.drawing.Point(stbFiltry.Location.X + _
        '                                 stbFiltry.Size.Width - _
        '                                 cmdWszystko.Size.Width, _
        '                                 stbFiltry.Location.Y + 2)
        'cmdWszystko.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or _
        '                            System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

        cmdWszystko.Size = New Size(20, 22)
        cmdWszystko.Location = New System.Drawing.Point(stbFiltry.Location.X + 50 - 20,
                                             stbFiltry.Location.Y)
        cmdWszystko.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or
                                        System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)








        '--------------------------------------------------------------------
        'set the header width to something apporpriate
        dagRaport.RowHeadersWidth = 50       'use if tablestyle
        'Me.dagraport.RowHeaderWidth = 80 'use if no tablestyle
        '--------------------------------------------------------------------
        'set topleft corner point so we can locate the toprow
        If DataSetRaport.Tables(0).Rows.Count > 0 Then
            StartPointInCell00 = New System.Drawing.Point((dagRaport.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (dagRaport.GetCellDisplayRectangle(0, 0, True).Y + 4))
            ScrollXOffset = 10000 * dagRaport.FirstDisplayedScrollingColumnIndex +
                        dagRaport.GetCellDisplayRectangle(dagRaport.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
        Else
            'brak zapisow - poczatek DataGrid
            StartPointInCell00 = New System.Drawing.Point((dagRaport.Left + 4), (dagRaport.Top + 4))
            ScrollXOffset = 0
        End If
        '--------------------------------------------------
        'inicjowanie zegara dla scrollingu poziomego
        HorizScrollLastTime = ""
        HorizScrollTimer.Interval = 2 * 1000
        HorizScrollTimer.Enabled = True
        '--------------------------------------------------
        'Stworz zmienne do filtrowania
        'Dim ii As Integer = 0
        For ii = 0 To DataViewRaport.Table.Columns.Count - 1
            'stworz zmienną TXTBOX
            Dim CtrlT As New System.Windows.Forms.TextBox
            CtrlT.Name = "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
            CtrlT.Visible = False
            Me.Controls.Add(CtrlT)
            AddHandler CtrlT.TextChanged, New EventHandler(AddressOf txtFiXX_TextChanged)
            AddHandler CtrlT.GotFocus, New EventHandler(AddressOf txtFiXX_GotFocus)
            AddHandler CtrlT.LostFocus, New EventHandler(AddressOf txtFiXX_LostFocus)
            AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXX_KeyDown)

            'stworz zmienną BUTTON
            Dim CtrlB As New System.Windows.Forms.Button
            CtrlB.Name = "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
            CtrlB.Image = cmdFi.Image
            CtrlB.ImageAlign = ContentAlignment.MiddleCenter
            CtrlB.Visible = False
            Me.Controls.Add(CtrlB)
            AddHandler CtrlB.Click, New EventHandler(AddressOf cmdFiXX_Click)
            AddHandler CtrlB.GotFocus, New EventHandler(AddressOf cmdFiXX_GotFocus)
            AddHandler CtrlB.LostFocus, New EventHandler(AddressOf cmdFiXX_LostFocus)
        Next
        Me.Refresh()
        '--------------------------------------------------
        StartFormy = False
        dagraport_CurrentCellChanged(sender, e)
        InicjujFiltryColumnDGV(Me, dagRaport, DataViewRaport.Table.Columns.Count, stbFiltry, StartFormy)


        stbRaport.Panels(0).Text = "Il.zap.: " & DataViewRaport.Count.ToString
        Me.Cursor = Cursors.Default
    End Sub


    Private Sub dagraport_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagRaport.CurrentCellChanged
        If Not StartFormy Then
            If dagRaport.CurrentCell Is Nothing Then
                stbRaport.Panels(1).Text = "Wi: "
                stbRaport.Panels(2).Text = "Ko: "
            Else
                stbRaport.Panels(1).Text = "Wi: " & dagRaport.CurrentCell.RowIndex.ToString
                stbRaport.Panels(2).Text = "Ko: " & dagRaport.CurrentCell.ColumnIndex.ToString
            End If
        End If
    End Sub

    '-=------------------------------------------
    'obsluga zegara scrollingu poziomego
    '-=------------------------------------------
    Private Sub HorizScrollTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HorizScrollTimer.Tick
        HorizScrollTimer.Enabled = False
        If Len(HorizScrollLastTime) = 0 Then
            'nie zainicjowano scrollingu
        Else
            'zainicjowano scrolling - sprawdz czy juz trwa 2 sec
            If IleSekundMinelo(HorizScrollLastTime, SysGodz()) >= 0 Then
                If dagRaport.RowCount > 0 Then
                    If ScrollXOffset <> 10000 * dagRaport.FirstDisplayedScrollingColumnIndex +
                                    dagRaport.GetCellDisplayRectangle(dagRaport.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X Then

                        PokazFiltryColumnDGV(Me, dagRaport, DataViewRaport.Table.Columns.Count, stbFiltry, StartFormy)
                        '-----------------------------
                        ScrollXOffset = 10000 * dagRaport.FirstDisplayedScrollingColumnIndex +
                                    dagRaport.GetCellDisplayRectangle(dagRaport.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X
                        HorizScrollLastTime = ""
                    End If
                Else
                    PokazFiltryColumnDGV(Me, dagRaport, DataViewRaport.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Else
            End If
        End If
        HorizScrollTimer.Enabled = True
    End Sub


    Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles MyBase.Scroll
        Dim rc As DataGridViewSelectedRowCollection = dagRaport.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                dagRaport.Rows(nextrow).Selected = True
            End If
        End If
    End Sub


    '*********************************************************
    '***** Filtrowanie wg poszczególnych pól
    '*********************************************************

    'AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXX_KeyDown)

    Private Sub txtFiXX_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Not StartFormy Then
            Select Case e.KeyCode
                'Case Keys.Enter
                'Case Keys.Insert, Keys.Add
                'Case Keys.Delete, Keys.Subtract

                Case Keys.Down
                    PrzejdzDoDGV(dagRaport, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            FiltrujDataViewDGV(Me, dagRaport, DataViewRaport.Table.Columns.Count, DataViewRaport, stbRaport)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me, dagRaport, Mid(sender.name, 6, 3), "raport")
    End Sub
    Private Sub cmdFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiFiltraCmd(Me, sender)
    End Sub
    Private Sub cmdFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraCmd(Me, sender)
    End Sub


    '-------------------------------------------------
    ' filtrowanie 
    '-------------------------------------------------

    Private Sub cmdWszystko_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWszystko.Click
        Me.Cursor = Cursors.WaitCursor
        oLokalnyFiltr = ""
        Try
            StartFormy = True
            CzyscFiltryDGV(Me, dagRaport, DataViewRaport.Table.Columns.Count, DataViewRaport, stbRaport)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbRaport.Panels(0).Text = "Il.zap.: " & DataViewRaport.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub



    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagRaport_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagRaport.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagRaport.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagRaport.DefaultCellStyle.Font,
                      b,
                      e.RowBounds.Location.X + 15,
                      e.RowBounds.Location.Y + 4)
        End Using
    End Sub



    '*************************************************
    '*** Odświezanie filtrow kolumnowych 
    '*************************************************

    Private Sub stbFiltry_PanelClick(sender As Object, e As StatusBarPanelClickEventArgs) Handles stbFiltry.PanelClick
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagRaport, DataViewRaport.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagRaport_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagRaport.CellFormatting
    '    If Not StartFormy Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagRaport, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagRaport, e.RowIndex, 4)

    '        Select Case UserName = Program_IdUzytkownika
    '            Case True
    '                e.CellStyle.ForeColor = Color.Red
    '            Case Else
    '                Select Case Mid(Upr, _Uprawnnienia_NrZnaku, 1)
    '                    Case "A"
    '                        e.CellStyle.ForeColor = Color.Green
    '                    Case "S"
    '                        e.CellStyle.ForeColor = Color.Purple
    '                    Case "U"
    '                        e.CellStyle.ForeColor = Color.Navy
    '                    Case Else
    '                        e.CellStyle.ForeColor = Color.Black
    '                End Select
    '        End Select
    '        '-----------------------
    '        ''zmieniamy BackColor
    '        'Dim Wal As String = GetTxtField(dagKursyWalut, e.RowIndex, 0)
    '        'Select Case Wal
    '        '    Case "EUR"
    '        '        e.CellStyle.BackColor = Color.White
    '        '    Case Else
    '        'End Select
    '        '-----------------------
    '    End If
    'End Sub






    '*********************************************************    '*********************************************************
    '***** obsluga dag
    '*********************************************************

    Private Sub dagRaport_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagRaport.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagRaport, DataViewRaport.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagRaport_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagRaport.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagRaport, DataViewRaport.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagRaport_Scroll(sender As Object, e As ScrollEventArgs) Handles dagRaport.Scroll
        'If (Not StartFormy) And (DataViewRaport.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me, dagRaport, DataViewRaport.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me, dagRaport, DataViewRaport.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If

        If (Not StartFormy) And (DataViewRaport.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagRaport_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagRaport.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagRaport, DataViewRaport.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagRaport_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagRaport.MouseMove
        'przypisz biezacy DataGrid do myGrid
        Dim myGridView As DataGridView = CType(sender, DataGridView)
        Dim hti As System.Windows.Forms.DataGridView.HitTestInfo
        hti = myGridView.HitTest(e.X, e.Y)

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None     '"the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                'If Not StartFormy Then
                '    PokazFiltryColumnDGV(Me, dagRaport, DataViewRaport.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me, dagRaport, DataViewRaport.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagRaport_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagRaport.MouseUp
        'przypisz biezacy DataGrid do myGrid
        Dim myGridView As DataGridView = CType(sender, DataGridView)
        Dim hti As System.Windows.Forms.DataGridView.HitTestInfo
        hti = myGridView.HitTest(e.X, e.Y)

        Dim message As String = "You clicked "

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None
                message &= "the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                message &= "cell at row " & hti.RowIndex & ", col " & hti.ColumnIndex
                stbRaport.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbRaport.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagRaport(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbRaport.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbRaport.Panels(3).Text = "Sort: " & DataSetRaport.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbRaport.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagRaport(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbRaport.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbRaport.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub



    Private Sub dagRaport_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagRaport.DoubleClick
        oIdentKon = GetTxtField(dagRaport, 2)
        IdentKlienta(oIdentKon)
        Dim Form As New KatalogKontaktow(oIdentKon)
        Form.ShowDialog()
    End Sub

    Private Sub dagRaport_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagRaport.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                oIdentKon = GetTxtField(dagRaport, 2)
                IdentKlienta(oIdentKon)
                Dim Form As New KatalogKontaktow(oIdentKon)
                Form.ShowDialog()
            Case Keys.Insert
            Case Keys.Delete
            Case Else
        End Select
    End Sub













































    '--------------------------------
    ' wybor relacji
    '--------------------------------


    Private Sub rabLiczbaRowne_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabLiczbaRowne.CheckedChanged
        lblOd.Visible = True
        lblOd.Text = "Ilość kontaktów . . . . . . . . . . . ."
        cbxIleKontaktowOd.Visible = True
        lblDo.Visible = False
        cbxIleKontaktowDo.Visible = False
        'cbxIleKontaktowOd_TextChanged(Me, System.EventArgs.Empty)
    End Sub

    Private Sub rabLiczbaRozne_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabLiczbaRozne.CheckedChanged
        lblOd.Visible = True
        lblOd.Text = "Ilość kontaktów . . . . . . . . . . . ."
        cbxIleKontaktowOd.Visible = True
        lblDo.Visible = False
        cbxIleKontaktowDo.Visible = False
    End Sub

    Private Sub rabLiczbaMniesze_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabLiczbaMniesze.CheckedChanged
        lblOd.Visible = True
        lblOd.Text = "Ilość kontaktów . . . . . . . . . . . ."
        cbxIleKontaktowOd.Visible = True
        lblDo.Visible = False
        cbxIleKontaktowDo.Visible = False
    End Sub

    Private Sub rabLiczbaMniejRowne_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabLiczbaMniejRowne.CheckedChanged
        lblOd.Visible = True
        lblOd.Text = "Ilość kontaktów . . . . . . . . . . . ."
        cbxIleKontaktowOd.Visible = True
        lblDo.Visible = False
        cbxIleKontaktowDo.Visible = False
    End Sub

    Private Sub rabLiczbaWieksze_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabLiczbaWieksze.CheckedChanged
        lblOd.Visible = True
        lblOd.Text = "Ilość kontaktów . . . . . . . . . . . ."
        cbxIleKontaktowOd.Visible = True
        lblDo.Visible = False
        cbxIleKontaktowDo.Visible = False
    End Sub

    Private Sub rabLiczbaWieRowne_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabLiczbaWieRowne.CheckedChanged
        lblOd.Visible = True
        lblOd.Text = "Ilość kontaktów . . . . . . . . . . . ."
        cbxIleKontaktowOd.Visible = True
        lblDo.Visible = False
        cbxIleKontaktowDo.Visible = False
    End Sub

    Private Sub rabLiczbaOdDo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rabLiczbaOdDo.CheckedChanged
        lblOd.Visible = True
        lblOd.Text = "Ilość kontaktów Od . . . . . . . . . . . ."
        cbxIleKontaktowOd.Visible = True
        lblDo.Visible = True
        lblOd.Text = "Ilość kontaktów Do . . . . . . . . . . . ."
        cbxIleKontaktowDo.Visible = True
    End Sub

    '--------------------------------
    ' automatyczne ustalanie zakresu dat
    '--------------------------------

    Private Sub cmdT1_Click(sender As Object, e As EventArgs) Handles cmdT1.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -1 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT2_Click(sender As Object, e As EventArgs) Handles cmdT2.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -2 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT3_Click(sender As Object, e As EventArgs) Handles cmdT3.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -3 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT4_Click(sender As Object, e As EventArgs) Handles cmdT4.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -4 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT5_Click(sender As Object, e As EventArgs) Handles cmdT5.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -5 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT6_Click(sender As Object, e As EventArgs) Handles cmdT6.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -6 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT7_Click(sender As Object, e As EventArgs) Handles cmdT7.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -7 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT8_Click(sender As Object, e As EventArgs) Handles cmdT8.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -8 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT9_Click(sender As Object, e As EventArgs) Handles cmdT9.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -9 * 7)
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdT10_Click(sender As Object, e As EventArgs) Handles cmdT10.Click
        dtpOdDaty.Value = WyliczDate(SysData(), -10 * 7)
        dtpDoDaty.Value = SysData()
    End Sub




    Private Sub cmdM1_Click(sender As Object, e As EventArgs) Handles cmdM1.Click
        dtpOdDaty.Value = PopMiesiac(SysData())
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM2_Click(sender As Object, e As EventArgs) Handles cmdM2.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(SysData()))
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM3_Click(sender As Object, e As EventArgs) Handles cmdM3.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(PopMiesiac(SysData())))
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM4_Click(sender As Object, e As EventArgs) Handles cmdM4.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(SysData()))))
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM5_Click(sender As Object, e As EventArgs) Handles cmdM5.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(SysData())))))
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM6_Click(sender As Object, e As EventArgs) Handles cmdM6.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(SysData()))))))
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM7_Click(sender As Object, e As EventArgs) Handles cmdM7.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(SysData())))))))
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM8_Click(sender As Object, e As EventArgs) Handles cmdM8.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(SysData()))))))))
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM9_Click(sender As Object, e As EventArgs) Handles cmdM9.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(SysData())))))))))
        dtpDoDaty.Value = SysData()
    End Sub

    Private Sub cmdM10_Click(sender As Object, e As EventArgs) Handles cmdM10.Click
        dtpOdDaty.Value = PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(PopMiesiac(SysData()))))))))))
        dtpDoDaty.Value = SysData()
    End Sub







    '-------------------------------
    'wybranie klientow z analizy w kartotece klientow
    '----------------------------------------------------------------

    Private Sub cmdZapamietaj_Click(sender As Object, e As EventArgs) Handles cmdZapamietaj.Click
        If DataSetRaport Is Nothing Then
            MessageBox.Show("Brak wyników analizy.",
                "Uwaga:",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            If DataSetRaport.Tables(0).Rows.Count = 0 Then
                MessageBox.Show("Analiza nie wybrała żadnych klientów.",
                    "Uwaga:",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            Else

                '...........................
                'definiuj delegat
                Dim deleg As delegateAktualizuj = _AktualizujKartKlientow
                Dim FormRaport As New RaportWybierzAnalizeGrupowaKontaktow(_dsKlenci,
                                                                     DataViewRaport,
                                                                     deleg,
                                                                     True)
                FormRaport.ShowDialog()
                '...........................
                MessageBox.Show("Wybrałem klientów na podstawie wyników analizy.",
                    "OK - skończyłem",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)


            End If
        End If
    End Sub


    Private Sub cmdRodzajeKontaktow_Click(sender As Object, e As EventArgs) Handles cmdRodzajeKontaktow.Click
        oCoMamRobic = "WYBIERAJ"
        oRodzajKontaktu = ""
        Dim form As New KatalogRodzajowKontaktow
        form.ShowDialog()
        If Len(Trim(oRodzajKontaktu)) > 0 Then
            txtRodzajeKontaktow.Text &= oRodzajKontaktu & vbCrLf
        End If
    End Sub

    Private Sub mnuWybierzAM_Click(sender As Object, e As EventArgs) Handles mnuWybierzAM.Click
        oWybranoAkcjeMarketingowa = False
        'definiuj delegat
        'Dim deleg As delegateAktualizuj = AddressOf AktualizujBaze
        'Dim Form As New AkcjaMarketingowaWybierz(DataSetKlienci, deleg)
        Dim Form As New AkcjaMarketingowaWybierz(DataSetRaport, Nothing)
        Form.ShowDialog()

        If oWybranoAkcjeMarketingowa Then
            DataViewRaport.RowFilter = "MARKER='True'"
            dagRaport.Update()
            stbRaport.Panels(0).Text = "Il.zap.: " & DataViewRaport.Count.ToString
        End If

    End Sub


End Class