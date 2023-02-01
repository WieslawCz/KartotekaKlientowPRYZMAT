Public Class AkcjaMarketingowaWybierzWgObrotow

    Dim KwotaMin As Double = 0
    Dim OdDaty1 As String = ""
    Dim DoDaty1 As String = ""
    Dim OdDaty2 As String = ""
    Dim DoDaty2 As String = ""
    Dim TakenDate As Date = Nothing

    Private Sub WybierzAkcjeMarketingowa_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        oWybranoAkcjeMarketingowa = False
        dtpOdDaty1.Value = Mid(SysData(), 1, 4) & "-01-01"
        dtpDoDaty1.Value = SysData()

        dtpOdDaty2.Value = Mid(SysData(), 1, 4) & "-01-01"
        dtpDoDaty2.Value = SysData()

        cbxJakSieZmienilo.Items.Add("Zwiêkszyli sprzeda¿")
        cbxJakSieZmienilo.Items.Add("Sprzeda¿ taka sama")
        cbxJakSieZmienilo.Items.Add("Zmniejszyli sprzeda¿")
        cbxJakSieZmienilo.Text = "Zwiêkszyli sprzeda¿"

        txtKwotaMin.Text = 0.ToString("F2")
    End Sub

    Private Sub WybierzAkcjeMarketingowa_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        RysujGradient(Me)
    End Sub


    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub


    '**************************************************
    '** zapamietaj informacje o klientach
    '**************************************************


    '---------------------------
    Dim dbSelectSQLObrotyMies As String = sSelectSQLObrotyMies

    Dim dbConnectionObrotyMies As OleDb.OleDbConnection
    Dim dbCommandSelectObrotyMies As OleDb.OleDbCommand
    Dim dbCommandDeleteObrotyMies As OleDb.OleDbCommand
    Dim dbCommandUpdateObrotyMies As OleDb.OleDbCommand
    Dim dbCommandInsertObrotyMies As OleDb.OleDbCommand
    Dim dbDataAdapterObrotyMies As OleDb.OleDbDataAdapter

    Dim sqlConnectionObrotyMies As SqlClient.SqlConnection
    Dim sqlCommandSelectObrotyMies As SqlClient.SqlCommand
    Dim sqlCommandDeleteObrotyMies As SqlClient.SqlCommand
    Dim sqlCommandUpdateObrotyMies As SqlClient.SqlCommand
    Dim sqlCommandInsertObrotyMies As SqlClient.SqlCommand
    Dim sqlDataAdapterObrotyMies As SqlClient.SqlDataAdapter

    Dim DataSetObrotyMies As New DataSet
    Dim DataViewObrotyMies As New DataView
    '---------------------------
    Dim dbSelectSQLObroty As String = sSelectSQLObroty

    Dim dbConnectionObroty As OleDb.OleDbConnection
    Dim dbCommandSelectObroty As OleDb.OleDbCommand
    Dim dbCommandDeleteObroty As OleDb.OleDbCommand
    Dim dbCommandUpdateObroty As OleDb.OleDbCommand
    Dim dbCommandInsertObroty As OleDb.OleDbCommand
    Dim dbDataAdapterObroty As OleDb.OleDbDataAdapter

    Dim sqlConnectionObroty As SqlClient.SqlConnection
    Dim sqlCommandSelectObroty As SqlClient.SqlCommand
    Dim sqlCommandDeleteObroty As SqlClient.SqlCommand
    Dim sqlCommandUpdateObroty As SqlClient.SqlCommand
    Dim sqlCommandInsertObroty As SqlClient.SqlCommand
    Dim sqlDataAdapterObroty As SqlClient.SqlDataAdapter

    Dim DataSetObroty As New DataSet
    Dim DataViewObroty As New DataView
    '---------------------------
    Dim dbSelectSQLKlienci As String = ""       'sSelectSQLKlienci & " ORDER BY IDENTKLIENTA"

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
    '---------------------------
    Dim ConnectionState As System.Data.ConnectionState


    Dim DataTableRobo As New System.Data.DataTable
    Dim DataSetRobo As New System.Data.DataSet
    Dim DataViewRobo As New System.Data.DataView


    Private Sub cmdZapamietaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdZapamietaj.Click
        TakenDate = dtpOdDaty1.Value
        OdDaty1 = TakenDate.ToString("yyyy-MM-dd")
        TakenDate = dtpDoDaty1.Value
        DoDaty1 = TakenDate.ToString("yyyy-MM-dd")

        TakenDate = dtpOdDaty2.Value
        OdDaty2 = TakenDate.ToString("yyyy-MM-dd")
        TakenDate = dtpDoDaty2.Value()
        DoDaty2 = TakenDate.ToString("yyyy-MM-dd")

        KwotaMin = GetDblFromText(txtKwotaMin.Text)
        '-------------------------------
        If _dsKlenci.Tables(0).Rows.Count = 0 Then
            MessageBox.Show("Specyfikacja klientów jest pusta.",
                "Uwaga:",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            'zakladamy strukture robocz¹ do zapisywania wynikow testowania
            '-------------------------------
            'Roboczy Dataset
            DataSetRobo = New System.Data.DataSet
            DataTableRobo = New System.Data.DataTable
            DataViewRobo = New System.Data.DataView

            Dim robCol1 As DataColumn = DataTableRobo.Columns.Add("IDENTKLI", GetType(System.String))
            Dim robCol2 As DataColumn = DataTableRobo.Columns.Add("WARTSPRZED1", GetType(System.Double))
            Dim robCol3 As DataColumn = DataTableRobo.Columns.Add("WARTSPRZED2", GetType(System.Double))
            DataSetRobo.Tables.Add(DataTableRobo)
            'zdefiniuj unikalny klucz wyszukiwania w tabeli
            DataSetRobo.Tables(0).PrimaryKey = New DataColumn() {DataSetRobo.Tables(0).Columns("IDENTKLI")}
            'definiuj DataView
            DataViewRobo = New DataView(DataSetRobo.Tables(0))

            '-----------------------------
            'analizuj klientow
            '-----------------------------

            'DataSet
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
                '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    '    MessageBoxIcon.Information, _
                '    '    MessageBoxDefaultButton.Button1)
                'Finally
                '    ConnectionState = dbConnectionKlienci.State
                'End Try

            Else
                sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
                sqlCommandSelectKlienci = New SqlClient.SqlCommand(sSelectSQLKlienciAktywni & " ORDER BY IDENTKLIENTA", sqlConnectionKlienci)
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
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
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
                    DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("NRKARTY")}
                    DataViewKlienci = New DataView(DataSetKlienci.Tables(0))

                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                End Try
                '---------------------------------
            End If






            '-----------------------------
            'analizuj Obroty Mies
            '-----------------------------

            'DataSet
            DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")
            If ParL_DataBaseType = "MS ACCESS" Then
                'dbConnectionObrotyMies = New OleDb.OleDbConnection(sConnectionObrotyMies)
                'dbCommandSelectObrotyMies = New OleDb.OleDbCommand(dbSelectSQLObrotyMies, dbConnectionObrotyMies)
                ''dbCommandDeleteObrotyMies = New OleDb.OleDbCommand(sDeleteSQLObrotyMies, dbConnectionObrotyMies)
                ''dbCommandUpdateObrotyMies = New OleDb.OleDbCommand(sUpdateSQLObrotyMies, dbConnectionObrotyMies)
                ''dbCommandInsertObrotyMies = New OleDb.OleDbCommand(sInsertSQLObrotyMies, dbConnectionObrotyMies)
                'dbDataAdapterObrotyMies = New OleDb.OleDbDataAdapter(dbCommandSelectObrotyMies)

                ''DataSet
                'DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")

                ''mapujemy domyslna nazwe tabeli
                'Dim MapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
                'MapowanieTabeliObrotyMies = dbDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")

                '' Add the RowUpdated event handler.
                'AddHandler dbDataAdapterObrotyMies.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

                ''------------------------------------------
                ''wypelnij DATASET
                'Try
                '    dbConnectionObrotyMies.Open()
                'Catch Ex As System.Exception
                '    ViewInLog(Ex, Me.Name(), Nothing)
                '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    '    MessageBoxIcon.Information, _
                '    '    MessageBoxDefaultButton.Button1)
                'Finally
                '    ConnectionState = dbConnectionObrotyMies.State
                'End Try

            Else
                sqlConnectionObrotyMies = New SqlClient.SqlConnection(sConnectionObrotyMies)
                sqlCommandSelectObrotyMies = New SqlClient.SqlCommand(dbSelectSQLObrotyMies, sqlConnectionObrotyMies)
                'sqlCommandDeleteObrotyMies = New SqlClient.SqlCommand(sDeleteSQLObrotyMies, sqlConnectionObrotyMies)
                'sqlCommandUpdateObrotyMies = New SqlClient.SqlCommand(sUpdateSQLObrotyMies, sqlConnectionObrotyMies)
                'sqlCommandInsertObrotyMies = New SqlClient.SqlCommand(sInsertSQLObrotyMies, sqlConnectionObrotyMies)
                sqlDataAdapterObrotyMies = New SqlClient.SqlDataAdapter(sqlCommandSelectObrotyMies)


                'mapujemy domyslna nazwe tabeli
                Dim MapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
                MapowanieTabeliObrotyMies = sqlDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")

                ' Add the RowUpdated event handler.
                AddHandler sqlDataAdapterObrotyMies.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

                '------------------------------------------
                'wypelnij DATASET
                Try
                    sqlConnectionObrotyMies.Open()
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                Finally
                    ConnectionState = sqlConnectionObrotyMies.State
                End Try
            End If

            If ConnectionState = ConnectionState.Open Then
                Try
                    If ParL_DataBaseType = "MS ACCESS" Then
                        dbDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                        dbDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                        dbConnectionObrotyMies.Close()
                    Else
                        sqlDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                        sqlDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                        sqlConnectionObrotyMies.Close()
                    End If

                    'zdefiniuj unikalny klucz wyszukiwania w tabeli
                    DataSetObrotyMies.Tables(0).PrimaryKey = New DataColumn() {DataSetObrotyMies.Tables(0).Columns("IDENTKLIENTA"),
                                                                               DataSetObrotyMies.Tables(0).Columns("MIESIAC")}
                    DataViewObrotyMies = New DataView(DataSetObrotyMies.Tables(0))

                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                End Try
                '---------------------------------
            End If






            '-----------------------------
            'analizuj Obroty w ostatnim miesiacu
            '-----------------------------

            'DataSet
            DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")
            If ParL_DataBaseType = "MS ACCESS" Then
                'dbConnectionObroty = New OleDb.OleDbConnection(sConnectionObroty)
                'dbCommandSelectObroty = New OleDb.OleDbCommand(dbSelectSQLObroty, dbConnectionObroty)
                ''dbCommandDeleteObroty = New OleDb.OleDbCommand(sDeleteSQLObroty, dbConnectionObroty)
                ''dbCommandUpdateObroty = New OleDb.OleDbCommand(sUpdateSQLObroty, dbConnectionObroty)
                ''dbCommandInsertObroty = New OleDb.OleDbCommand(sInsertSQLObroty, dbConnectionObroty)
                'dbDataAdapterObroty = New OleDb.OleDbDataAdapter(dbCommandSelectObroty)

                ''DataSet
                'DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")

                ''mapujemy domyslna nazwe tabeli
                'Dim MapowanieTabeliObroty As System.Data.Common.DataTableMapping
                'MapowanieTabeliObroty = dbDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")

                '' Add the RowUpdated event handler.
                'AddHandler dbDataAdapterObroty.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

                ''------------------------------------------
                ''wypelnij DATASET
                'Try
                '    dbConnectionObroty.Open()
                'Catch Ex As System.Exception
                '    ViewInLog(Ex, Me.Name(), Nothing)
                '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    '    MessageBoxIcon.Information, _
                '    '    MessageBoxDefaultButton.Button1)
                'Finally
                '    ConnectionState = dbConnectionObroty.State
                'End Try

            Else
                sqlConnectionObroty = New SqlClient.SqlConnection(sConnectionObroty)
                sqlCommandSelectObroty = New SqlClient.SqlCommand(dbSelectSQLObroty, sqlConnectionObroty)
                'sqlCommandDeleteObroty = New SqlClient.SqlCommand(sDeleteSQLObroty, sqlConnectionObroty)
                'sqlCommandUpdateObroty = New SqlClient.SqlCommand(sUpdateSQLObroty, sqlConnectionObroty)
                'sqlCommandInsertObroty = New SqlClient.SqlCommand(sInsertSQLObroty, sqlConnectionObroty)
                sqlDataAdapterObroty = New SqlClient.SqlDataAdapter(sqlCommandSelectObroty)

                'mapujemy domyslna nazwe tabeli
                Dim MapowanieTabeliObroty As System.Data.Common.DataTableMapping
                MapowanieTabeliObroty = sqlDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")

                ' Add the RowUpdated event handler.
                AddHandler sqlDataAdapterObroty.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

                '------------------------------------------
                'wypelnij DATASET
                Try
                    sqlConnectionObroty.Open()
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                Finally
                    ConnectionState = sqlConnectionObroty.State
                End Try
            End If

            If ConnectionState = ConnectionState.Open Then
                Try
                    If ParL_DataBaseType = "MS ACCESS" Then
                        dbDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                        dbDataAdapterObroty.Fill(DataSetObroty)
                        dbConnectionObroty.Close()
                    Else
                        sqlDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                        sqlDataAdapterObroty.Fill(DataSetObroty)
                        sqlConnectionObroty.Close()
                    End If

                    'zdefiniuj unikalny klucz wyszukiwania w tabeli
                    DataSetObroty.Tables(0).PrimaryKey = New DataColumn() {DataSetObroty.Tables(0).Columns("IDENTKLIENTA"),
                                                                           DataSetObroty.Tables(0).Columns("DATA")}
                    DataViewObroty = New DataView(DataSetObroty.Tables(0))

                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                End Try
                '---------------------------------
            End If










            ' analizuj poszczegolnych klientow
            ' a dla klientow analizuj obroty miesieczne
            ' dopisuj do bazy AnalizyZakupu
            Dim drKlienci As DataRow
            Dim drRoboczy As DataRow
            Dim drvObroty As DataRowView
            Dim drvObrotyMies As DataRowView

            Dim ik As Integer = 0
            Dim io As Integer = 0
            Dim ir As Integer = 0
            Dim SymbKlienta As String = ""
            Dim IleRecKlienci As Integer = 0

            Dim WSprzed1 As Double = 0
            Dim WSprzed2 As Double = 0

            Dim findObrotyMies(1) As Object
            Dim foundRow As DataRow = Nothing
            Dim biezMiesiac As String = ""

            Dim iRok As Integer = 0
            Dim iMies As Integer = 0
            Dim i As Integer = 0

            IleRecKlienci = DataSetKlienci.Tables(0).Rows.Count
            ProgressBar1.Maximum = IleRecKlienci
            For ik = 0 To IleRecKlienci - 1
                ProgressBar1.Value = ik
                drKlienci = DataSetKlienci.Tables(0).Rows.Item(ik)
                SymbKlienta = DataSetKlienci.Tables(0).Rows.Item(ik).Item("NRKARTY")
                WSprzed1 = 0
                WSprzed2 = 0

                ''---------------------------------------------------------------------
                ''Obroty
                'Public oIdentObr As String            'IDENTKLIENTA     Text, 6
                'Public oDataObr As String             'DATA             Text,10

                'Public oLWartSprzedObr As Double      'LWARTSPRZED      Double
                'Public oLIloSprzedObr As Double       'LILOPRZED        Double
                'Public oLMarSprzedObr As Double       'LMARPRZED        Double

                'Public oAWartSprzedObr As Double      'AWARTSPRZED      Double
                'Public oAIloSprzedObr As Double       'AILOSPRZED       Double
                'Public oAMARSprzedObr As Double       'AMARSPRZED       Double

                'Public oLOWartSprzedObr As Double      'LOWARTSPRZED      Double
                'Public oLOIloSprzedObr As Double       'LOILOPRZED        Double
                'Public oLOMARSprzedObr As Double       'LOMARPRZED        Double

                'Public oAOWartSprzedObr As Double      'AOWARTSPRZED      Double
                'Public oAOIloSprzedObr As Double       'AOILOSPRZED       Double
                'Public oAOMARSprzedObr As Double       'AOMARSPRZED       Double

                'Public oNAJEMWartSprzedObr As Double      'NAJEMWARTSPRZED      Double
                'Public oNAJEMIloSprzedObr As Double       'NAJEMILOPRZED        Double
                'Public oNAJEMMARSprzedObr As Double       'NAJEMMARPRZED        Double

                'Public oSTRONYWartSprzedObr As Double      'STRONYWARTSPRZED      Double
                'Public oSTRONYIloSprzedObr As Double       'STRONYILOPRZED        Double
                'Public oSTRONYMARSprzedObr As Double       'STRONYMARPRZED        Double

                'Public oDRUKARKIWartSprzedObr As Double      'DRUKARKIWARTSPRZED      Double
                'Public oDRUKARKIIloSprzedObr As Double       'DRUKARKIILOPRZED        Double
                'Public oDRUKARKIMARSprzedObr As Double       'DRUKARKIMARPRZED        Double

                'Public oSKUPWartSprzedObr As Double      'SKUPWARTSPRZED      Double
                'Public oSKUPIloSprzedObr As Double       'SKUPILOPRZED        Double
                'Public oSKUPMARSprzedObr As Double       'SKUPMARPRZED        Double

                'Public oWersjaObr As Integer          'WERSJA           Integer








                DataViewObroty.RowFilter = "(IDENTKLIENTA='" & SymbKlienta & "') AND (DATA>='" & OdDaty1 & "') AND (DATA<='" & DoDaty1 & "')"
                If DataViewObroty.Count > 0 Then
                    For io = 0 To DataViewObroty.Count - 1
                        drvObroty = DataViewObroty.Item(io)

                        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne")
                        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Iloœci")
                        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne")
                        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Wartoœci")
                        'cbxCoAnalizowac.Items.Add("Analizuj wartoœæ Mar¿y")
                        'cbxCoAnalizowac.Items.Add("Analizuj procent Mar¿y")

                        Select Case _CoAnalizowac
                            Case "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne", "Analizuj Iloœciowo-Sumuj Iloœci"
                                WSprzed1 += GetDblDRVField(drvObroty, "LOILOSPRZED") +
                                    GetDblDRVField(drvObroty, "AOILOSPRZED") +
                                    GetDblDRVField(drvObroty, "LILOSPRZED") +
                                    GetDblDRVField(drvObroty, "AILOSPRZED") +
                                    GetDblDRVField(drvObroty, "NAJEMILOSPRZED") +
                                    GetDblDRVField(drvObroty, "STRONYILOSPRZED") +
                                    GetDblDRVField(drvObroty, "DRUKARKIILOSPRZED") +
                                    GetDblDRVField(drvObroty, "SKUPILOSPRZED")
                            Case "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne", "Analizuj Wartoœciowo-Sumuj Wartoœci"
                                WSprzed1 += GetDblDRVField(drvObroty, "LOWARTSPRZED") +
                                    GetDblDRVField(drvObroty, "AOWARTSPRZED") +
                                    GetDblDRVField(drvObroty, "LWARTSPRZED") +
                                    GetDblDRVField(drvObroty, "AWARTSPRZED") +
                                    GetDblDRVField(drvObroty, "NAJEMWARTSPRZED") +
                                    GetDblDRVField(drvObroty, "STRONYWARTSPRZED") +
                                    GetDblDRVField(drvObroty, "DRUKARKIWARTSPRZED") +
                                    GetDblDRVField(drvObroty, "SKUPWARTSPRZED")
                            Case "Analizuj wartoœæ Mar¿y", "Analizuj procent Mar¿y"
                                WSprzed1 += GetDblDRVField(drvObroty, "LOMARSPRZED") +
                                    GetDblDRVField(drvObroty, "AOMARSPRZED") +
                                    GetDblDRVField(drvObroty, "LMARSPRZED") +
                                    GetDblDRVField(drvObroty, "AMARSPRZED") +
                                    GetDblDRVField(drvObroty, "NAJEMMARSPRZED") +
                                    GetDblDRVField(drvObroty, "STRONYMARSPRZED") +
                                    GetDblDRVField(drvObroty, "DRUKARKIMARSPRZED") +
                                    GetDblDRVField(drvObroty, "SKUPMARSPRZED")
                        End Select
                        'System.Windows.Forms.Application.DoEvents()
                    Next
                End If

                DataViewObroty.RowFilter = "(IDENTKLIENTA='" & SymbKlienta & "') AND (DATA>='" & OdDaty2 & "') AND (DATA<='" & DoDaty2 & "')"
                If DataViewObroty.Count > 0 Then
                    For io = 0 To DataViewObroty.Count - 1
                        drvObroty = DataViewObroty.Item(io)

                        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne")
                        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Iloœci")
                        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne")
                        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Wartoœci")
                        'cbxCoAnalizowac.Items.Add("Analizuj wartoœæ Mar¿y")
                        'cbxCoAnalizowac.Items.Add("Analizuj procent Mar¿y")

                        Select Case _CoAnalizowac
                            Case "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne", "Analizuj Iloœciowo-Sumuj Iloœci"
                                WSprzed2 += GetDblDRVField(drvObroty, "LOILOSPRZED") +
                                    GetDblDRVField(drvObroty, "AOILOSPRZED") +
                                    GetDblDRVField(drvObroty, "LILOSPRZED") +
                                    GetDblDRVField(drvObroty, "AILOSPRZED") +
                                    GetDblDRVField(drvObroty, "NAJEMILOSPRZED") +
                                    GetDblDRVField(drvObroty, "STRONYILOSPRZED") +
                                    GetDblDRVField(drvObroty, "DRUKARKIILOSPRZED") +
                                    GetDblDRVField(drvObroty, "SKUPILOSPRZED")
                            Case "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne", "Analizuj Wartoœciowo-Sumuj Wartoœci"
                                WSprzed2 += GetDblDRVField(drvObroty, "LOWARTSPRZED") +
                                    GetDblDRVField(drvObroty, "AOWARTSPRZED") +
                                    GetDblDRVField(drvObroty, "LWARTSPRZED") +
                                    GetDblDRVField(drvObroty, "AWARTSPRZED") +
                                    GetDblDRVField(drvObroty, "NAJEMWARTSPRZED") +
                                    GetDblDRVField(drvObroty, "STRONYWARTSPRZED") +
                                    GetDblDRVField(drvObroty, "DRUKARKIWARTSPRZED") +
                                    GetDblDRVField(drvObroty, "SKUPWARTSPRZED")
                            Case "Analizuj wartoœæ Mar¿y", "Analizuj procent Mar¿y"
                                WSprzed2 += GetDblDRVField(drvObroty, "LOMARSPRZED") +
                                    GetDblDRVField(drvObroty, "AOMARSPRZED") +
                                    GetDblDRVField(drvObroty, "LMARSPRZED") +
                                    GetDblDRVField(drvObroty, "AMARSPRZED") +
                                    GetDblDRVField(drvObroty, "NAJEMMARSPRZED") +
                                    GetDblDRVField(drvObroty, "STRONYMARSPRZED") +
                                    GetDblDRVField(drvObroty, "DRUKARKIMARSPRZED") +
                                    GetDblDRVField(drvObroty, "SKUPMARSPRZED")
                        End Select
                        'System.Windows.Forms.Application.DoEvents()
                    Next
                End If
                DataViewObroty.RowFilter = ""

                ''---------------------------------------------------------------------
                ''ObrotyMies
                'Public oIdentMies As String            'IDENTKLIENTA     Text, 6
                'Public oMiesiacMies As String          'MIESIAC          Text,7

                'Public oLWartSprzedMies As Double      'LWARTSPRZED      Double
                'Public oLIloSprzedMies As Double       'LILOSPRZED       Double
                'Public oLMARSprzedMies As Double       'LMARSPRZED       Double

                'Public oAWartSprzedMies As Double      'AWARTSPRZED      Double
                'Public oAIloSprzedMies As Double       'AILOSPRZED       Double
                'Public oAMARSprzedMies As Double       'AMARSPRZED       Double

                'Public oLOWartSprzedMies As Double      'LOWARTSPRZED      Double
                'Public oLOIloSprzedMies As Double       'LOILOSPRZED       Double
                'Public oLOMARSprzedMies As Double       'LOMARSPRZED       Double

                'Public oAOWartSprzedMies As Double      'AOWARTSPRZED      Double
                'Public oAOIloSprzedMies As Double       'AOILOSPRZED       Double
                'Public oAOMARSprzedMies As Double       'AOMARSPRZED       Double

                'Public oNAJEMWartSprzedMies As Double      'NAJEMWARTSPRZED      Double
                'Public oNAJEMIloSprzedMies As Double       'NAJEMILOPRZED        Double
                'Public oNAJEMMARSprzedMies As Double       'NAJEMMARPRZED        Double

                'Public oSTRONYWartSprzedMies As Double      'STRONYWARTSPRZED      Double
                'Public oSTRONYIloSprzedMies As Double       'STRONYILOPRZED        Double
                'Public oSTRONYMARSprzedMies As Double       'STRONYMARPRZED        Double

                'Public oDRUKARKIWartSprzedMies As Double      'DRUKARKIWARTSPRZED      Double
                'Public oDRUKARKIIloSprzedMies As Double       'DRUKARKIILOPRZED        Double
                'Public oDRUKARKIMARSprzedMies As Double       'DRUKARKIMARPRZED        Double

                'Public oSKUPWartSprzedMies As Double      'SKUPWARTSPRZED      Double
                'Public oSKUPIloSprzedMies As Double       'SKUPILOPRZED        Double
                'Public oSKUPMARSprzedMies As Double       'SKUPMARPRZED        Double

                'Public oWersjaMies As Integer          'WERSJA           Integer

                DataViewObrotyMies.RowFilter = "(IDENTKLIENTA='" & SymbKlienta & "') AND (MIESIAC>='" & Mid(OdDaty1, 1, 7) & "') AND (MIESIAC<='" & Mid(DoDaty1, 1, 7) & "')"
                If DataViewObrotyMies.Count > 0 Then
                    For io = 0 To DataViewObrotyMies.Count - 1
                        drvObrotyMies = DataViewObrotyMies.Item(io)

                        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne")
                        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Iloœci")
                        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne")
                        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Wartoœci")
                        'cbxCoAnalizowac.Items.Add("Analizuj wartoœæ Mar¿y")
                        'cbxCoAnalizowac.Items.Add("Analizuj procent Mar¿y")

                        Select Case _CoAnalizowac
                            Case "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne", "Analizuj Iloœciowo-Sumuj Iloœci"
                                WSprzed1 += GetDblDRVField(drvObrotyMies, "LOILOSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "AOILOSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "LILOSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "AILOSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "NAJEMILOSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "STRONYILOSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "DRUKARKIILOSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "SKUPILOSPRZED")
                            Case "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne", "Analizuj Wartoœciowo-Sumuj Wartoœci"
                                WSprzed1 += GetDblDRVField(drvObrotyMies, "LOWARTSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "AOWARTSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "LWARTSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "AWARTSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "NAJEMWARTSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "STRONYWARTSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "DRUKARKIWARTSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "SKUPWARTSPRZED")
                            Case "Analizuj wartoœæ Mar¿y", "Analizuj procent Mar¿y"
                                WSprzed1 += GetDblDRVField(drvObrotyMies, "LOMARSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "AOMARSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "LMARSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "AMARSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "NAJEMMARSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "STRONYMARSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "DRUKARKIMARSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "SKUPMARSPRZED")
                        End Select
                        'System.Windows.Forms.Application.DoEvents()
                    Next
                End If


                DataViewObrotyMies.RowFilter = "(IDENTKLIENTA='" & SymbKlienta & "') AND (MIESIAC>='" & Mid(OdDaty2, 1, 7) & "') AND (MIESIAC<='" & Mid(DoDaty2, 1, 7) & "')"
                If DataViewObrotyMies.Count > 0 Then
                    For io = 0 To DataViewObrotyMies.Count - 1
                        drvObrotyMies = DataViewObrotyMies.Item(io)

                        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne")
                        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Iloœci")
                        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne")
                        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Wartoœci")
                        'cbxCoAnalizowac.Items.Add("Analizuj wartoœæ Mar¿y")
                        'cbxCoAnalizowac.Items.Add("Analizuj procent Mar¿y")

                        Select Case _CoAnalizowac
                            Case "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne", "Analizuj Iloœciowo-Sumuj Iloœci"
                                WSprzed2 += GetDblDRVField(drvObrotyMies, "LOILOSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "AOILOSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "LILOSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "AILOSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "NAJEMILOSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "STRONYILOSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "DRUKARKIILOSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "SKUPILOSPRZED")
                            Case "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne", "Analizuj Wartoœciowo-Sumuj Wartoœci"
                                WSprzed2 += GetDblDRVField(drvObrotyMies, "LOWARTSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "AOWARTSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "LWARTSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "AWARTSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "NAJEMWARTSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "STRONYWARTSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "DRUKARKIWARTSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "SKUPWARTSPRZED")
                            Case "Analizuj wartoœæ Mar¿y", "Analizuj procent Mar¿y"
                                WSprzed2 += GetDblDRVField(drvObrotyMies, "LOMARSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "AOMARSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "LMARSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "AMARSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "NAJEMMARSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "STRONYMARSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "DRUKARKIMARSPRZED") +
                                    GetDblDRVField(drvObrotyMies, "SKUPMARSPRZED")
                        End Select
                        'System.Windows.Forms.Application.DoEvents()
                    Next
                End If









                'Dim robCol1 As DataColumn = DataTableRobo.Columns.Add("IDENTKLI", GetType(System.String))
                'Dim robCol2 As DataColumn = DataTableRobo.Columns.Add("WARTSPRZED1", GetType(System.Double))
                'Dim robCol2 As DataColumn = DataTableRobo.Columns.Add("WARTSPRZED2", GetType(System.Double))

                'cbxJakSieZmienilo.Items.Add("Zwiêkszyli sprzeda¿")
                'cbxJakSieZmienilo.Items.Add("Sprzeda¿ taka sama")
                'cbxJakSieZmienilo.Items.Add("Zmniejszyli sprzeda¿")
                Select Case cbxJakSieZmienilo.Text
                    Case "Zwiêkszyli sprzeda¿"
                        If WSprzed2 - WSprzed1 >= KwotaMin Then
                            'dopisz tego klienta do ROBOCZY
                            drRoboczy = DataSetRobo.Tables(0).NewRow()
                            drRoboczy.Item("IDENTKLI") = SymbKlienta
                            drRoboczy.Item("WARTSPRZED1") = WSprzed1
                            drRoboczy.Item("WARTSPRZED2") = WSprzed2
                            DataSetRobo.Tables(0).Rows.Add(drRoboczy)
                        End If

                    Case "Sprzeda¿ taka sama"
                        If WSprzed1 = WSprzed2 Then
                            'dopisz tego klienta do ROBOCZY
                            drRoboczy = DataSetRobo.Tables(0).NewRow()
                            drRoboczy.Item("IDENTKLI") = SymbKlienta
                            drRoboczy.Item("WARTSPRZED1") = WSprzed1
                            drRoboczy.Item("WARTSPRZED2") = WSprzed2
                            DataSetRobo.Tables(0).Rows.Add(drRoboczy)
                        End If

                    Case "Zmniejszyli sprzeda¿"
                        If WSprzed1 - WSprzed2 >= KwotaMin Then
                            'dopisz tego klienta do ROBOCZY
                            drRoboczy = DataSetRobo.Tables(0).NewRow()
                            drRoboczy.Item("IDENTKLI") = SymbKlienta
                            drRoboczy.Item("WARTSPRZED1") = WSprzed1
                            drRoboczy.Item("WARTSPRZED2") = WSprzed2
                            DataSetRobo.Tables(0).Rows.Add(drRoboczy)
                        End If

                End Select


                System.Windows.Forms.Application.DoEvents()
            Next
            '-------------------------------
            If DataViewRobo.Count > 0 Then
                '...........................
                'definiuj delegat
                Dim deleg As delegateAktualizuj = _AktualizujKartKlientow
                Dim FormRaport As New RaportWybierzAkcjeMarketingowa("",
                                                                     _dsKlenci,
                                                                     DataViewRobo,
                                                                     deleg,
                                                                     True)
                FormRaport.ShowDialog()
                oWybranoAkcjeMarketingowa = True
                '...........................
                MessageBox.Show("Pobra³em informacje o klientach (" & Trim(Str(DataViewRobo.Count)) & ")",
                        "OK - skoñczy³em",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                If DataViewRobo.Count > 0 Then
                    Me.Close()
                Else
                End If


            End If
        End If
    End Sub

    Private Sub txtKwotaMin_TextChanged(sender As Object, e As EventArgs) Handles txtKwotaMin.TextChanged
        If TestNum(txtKwotaMin, "KWOTA MINIMALNA") Then
        End If
    End Sub
    Private Sub txtKwotaMin_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKwotaMin.GotFocus
        StartEdycjiTxt(txtKwotaMin)
    End Sub
    Private Sub txtKwotaMin_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKwotaMin.LostFocus
        KoniecEdycjiTxt(txtKwotaMin)
    End Sub


End Class
