Public Class AkcjaMarketingowaWybierzWgUslug

    Dim OdDaty As String = ""
    Dim DoDaty As String = ""
    Dim TakenDate As Date = Nothing

    Private Sub WybierzAkcjeMarketingowa_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        oWybranoAkcjeMarketingowa = False
        dtpOdDaty.Value = Mid(SysData(), 1, 4) & "-01-01"
        dtpDoDaty.Value = SysData()

        chkU1.Checked = False
        chkU2.Checked = False
        chkU3.Checked = False
        chkU4.Checked = False
        chkU5.Checked = False
        chkU6.Checked = False
        chkU7.Checked = False
        chkU8.Checked = False
        chkU9.Checked = False
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
        TakenDate = dtpOdDaty.Value
        OdDaty = TakenDate.ToString("yyyy-MM-dd")
        TakenDate = dtpDoDaty.Value
        DoDaty = TakenDate.ToString("yyyy-MM-dd")
        '-------------------------------
        If _dsKlenci.Tables(0).Rows.Count = 0 Then
            MessageBox.Show("Specyfikacja klientów jest pusta.", _
                "Uwaga:", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            'zakladamy strukture robocz¹ do zapisywania wynikow testowania
            '-------------------------------
            'Roboczy Dataset
            DataSetRobo = New System.Data.DataSet
            DataTableRobo = New System.Data.DataTable
            DataViewRobo = New System.Data.DataView

            Dim robCol1 As DataColumn = DataTableRobo.Columns.Add("IDENTKLI", GetType(System.String))
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





            ' analizuj poszczegolnych klientow i wybrane przez nich us³ugi RZU
            ' dopisuj do bazy robo
            Dim drKlienci As DataRow
            Dim drRoboczy As DataRow

            Dim ik As Integer = 0
            Dim io As Integer = 0
            Dim ir As Integer = 0
            Dim SymbKlienta As String = ""
            Dim IleRecKlienci As Integer = 0
            Dim TotalSprzedaz As Double = 0
            Dim NarastSprzedaz As Double = 0

            Dim findObrotyMies(1) As Object
            Dim foundRow As DataRow = Nothing
            Dim biezMiesiac As String = ""

            Dim iRok As Integer = 0
            Dim iMies As Integer = 0
            Dim i As Integer = 0
            Dim KlientOK As Boolean = False

            IleRecKlienci = DataSetKlienci.Tables(0).Rows.Count
            ProgressBar1.Maximum = IleRecKlienci
            For ik = 0 To IleRecKlienci - 1
                ProgressBar1.Value = ik
                drKlienci = DataSetKlienci.Tables(0).Rows.Item(ik)

                SymbKlienta = DataSetKlienci.Tables(0).Rows.Item(ik).Item("NRKARTY")
                oUslugiDodKli = DataSetKlienci.Tables(0).Rows.Item(ik).Item("USLUGIDOD")
                oZakupPopRokKli = DataSetKlienci.Tables(0).Rows.Item(ik).Item("ZAKUPPOPROK")

                'chkU1.Checked = (Mid(oUslugiDodKli, 1, 1) = "T")
                'chkU2.Checked = (Mid(oUslugiDodKli, 2, 1) = "T")
                'chkU3.Checked = (Mid(oUslugiDodKli, 3, 1) = "T")
                'chkU4.Checked = (Mid(oUslugiDodKli, 4, 1) = "T")
                'chkU5.Checked = (Mid(oUslugiDodKli, 5, 1) = "T")
                'chkU6.Checked = (Mid(oUslugiDodKli, 6, 1) = "T")
                'chkU7.Checked = (Mid(oUslugiDodKli, 7, 1) = "T")
                'chkU8.Checked = (Mid(oUslugiDodKli, 8, 1) = "T")
                'chkU9.Checked = (Mid(oUslugiDodKli, 9, 1) = "T")

                'dtpU1.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 11, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 11, 10))
                'dtpU2.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 21, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 21, 10))
                'dtpU3.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 31, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 31, 10))
                'dtpU4.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 41, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 41, 10))
                'dtpU5.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 51, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 51, 10))
                'dtpU6.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 61, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 61, 10))
                'dtpU7.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 71, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 71, 10))
                'dtpU8.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 81, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 81, 10))
                'dtpU9.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 91, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 91, 10))



                'If oZakupPopRokKli <= 1000 Then
                '    chkU1.Enabled = False
                '    chkU2.Enabled = False
                '    chkU3.Enabled = False
                '    chkU4.Enabled = False
                '    chkU5.Enabled = False
                '    chkU6.Enabled = False
                '    chkU7.Enabled = False
                '    chkU8.Enabled = False
                '    chkU9.Enabled = False

                'ElseIf oZakupPopRokKli <= 3000 Then
                '    chkU1.Enabled = True
                '    chkU2.Enabled = True
                '    chkU3.Enabled = True
                '    chkU4.Enabled = True
                '    chkU5.Enabled = False
                '    chkU6.Enabled = False
                '    chkU7.Enabled = False
                '    chkU8.Enabled = False
                '    chkU9.Enabled = False

                'ElseIf oZakupPopRokKli <= 6000 Then
                '    chkU1.Enabled = True
                '    chkU2.Enabled = True
                '    chkU3.Enabled = True
                '    chkU4.Enabled = True
                '    chkU5.Enabled = True
                '    chkU6.Enabled = True
                '    chkU7.Enabled = False
                '    chkU8.Enabled = False
                '    chkU9.Enabled = False

                'Else
                '    chkU1.Enabled = True
                '    chkU2.Enabled = True
                '    chkU3.Enabled = True
                '    chkU4.Enabled = True
                '    chkU5.Enabled = True
                '    chkU6.Enabled = True
                '    chkU7.Enabled = True
                '    chkU8.Enabled = True
                '    chkU9.Enabled = True
                'End If



                'KlientOK = False
                'If (oZakupPopRokKli > 1000) And chkU1.Checked Then
                '    If (Mid(oUslugiDodKli, 1, 1) = "T") And (Mid(oUslugiDodKli, 11, 10) >= OdDaty) And (Mid(oUslugiDodKli, 11, 10) <= DoDaty) Then
                '        KlientOK = True
                '    End If
                'End If
                'If (oZakupPopRokKli > 1000) And chkU2.Checked Then
                '    If (Mid(oUslugiDodKli, 2, 1) = "T") And (Mid(oUslugiDodKli, 21, 10) >= OdDaty) And (Mid(oUslugiDodKli, 21, 10) <= DoDaty) Then
                '        KlientOK = True
                '    End If
                'End If
                'If (oZakupPopRokKli > 1000) And chkU3.Checked Then
                '    If (Mid(oUslugiDodKli, 3, 1) = "T") And (Mid(oUslugiDodKli, 31, 10) >= OdDaty) And (Mid(oUslugiDodKli, 31, 10) <= DoDaty) Then
                '        KlientOK = True
                '    End If
                'End If
                'If (oZakupPopRokKli > 1000) And chkU4.Checked Then
                '    If (Mid(oUslugiDodKli, 4, 1) = "T") And (Mid(oUslugiDodKli, 41, 10) >= OdDaty) And (Mid(oUslugiDodKli, 41, 10) <= DoDaty) Then
                '        KlientOK = True
                '    End If
                'End If
                'If (oZakupPopRokKli > 3000) And chkU5.Checked Then
                '    If (Mid(oUslugiDodKli, 5, 1) = "T") And (Mid(oUslugiDodKli, 51, 10) >= OdDaty) And (Mid(oUslugiDodKli, 51, 10) <= DoDaty) Then
                '        KlientOK = True
                '    End If
                'End If
                'If (oZakupPopRokKli > 3000) And chkU6.Checked Then
                '    If (Mid(oUslugiDodKli, 6, 1) = "T") And (Mid(oUslugiDodKli, 61, 10) >= OdDaty) And (Mid(oUslugiDodKli, 61, 10) <= DoDaty) Then
                '        KlientOK = True
                '    End If
                'End If
                'If (oZakupPopRokKli > 6000) And chkU7.Checked Then
                '    If (Mid(oUslugiDodKli, 7, 1) = "T") And (Mid(oUslugiDodKli, 71, 10) >= OdDaty) And (Mid(oUslugiDodKli, 71, 10) <= DoDaty) Then
                '        KlientOK = True
                '    End If
                'End If
                'If (oZakupPopRokKli > 6000) And chkU8.Checked Then
                '    If (Mid(oUslugiDodKli, 8, 1) = "T") And (Mid(oUslugiDodKli, 81, 10) >= OdDaty) And (Mid(oUslugiDodKli, 81, 10) <= DoDaty) Then
                '        KlientOK = True
                '    End If
                'End If
                'If (oZakupPopRokKli > 6000) And chkU9.Checked Then
                '    If (Mid(oUslugiDodKli, 9, 1) = "T") And (Mid(oUslugiDodKli, 91, 10) >= OdDaty) And (Mid(oUslugiDodKli, 91, 10) <= DoDaty) Then
                '        KlientOK = True
                '    End If
                'End If




                'nie bierzemyu pod uwagê limitów sprzeda¿owych
                KlientOK = False
                If chkU1.Checked Then
                    If (Mid(oUslugiDodKli, 1, 1) = "T") And (Mid(oUslugiDodKli, 11, 10) >= OdDaty) And (Mid(oUslugiDodKli, 11, 10) <= DoDaty) Then
                        KlientOK = True
                    End If
                End If
                If chkU2.Checked Then
                    If (Mid(oUslugiDodKli, 2, 1) = "T") And (Mid(oUslugiDodKli, 21, 10) >= OdDaty) And (Mid(oUslugiDodKli, 21, 10) <= DoDaty) Then
                        KlientOK = True
                    End If
                End If
                If chkU3.Checked Then
                    If (Mid(oUslugiDodKli, 3, 1) = "T") And (Mid(oUslugiDodKli, 31, 10) >= OdDaty) And (Mid(oUslugiDodKli, 31, 10) <= DoDaty) Then
                        KlientOK = True
                    End If
                End If
                If chkU4.Checked Then
                    If (Mid(oUslugiDodKli, 4, 1) = "T") And (Mid(oUslugiDodKli, 41, 10) >= OdDaty) And (Mid(oUslugiDodKli, 41, 10) <= DoDaty) Then
                        KlientOK = True
                    End If
                End If
                If chkU5.Checked Then
                    If (Mid(oUslugiDodKli, 5, 1) = "T") And (Mid(oUslugiDodKli, 51, 10) >= OdDaty) And (Mid(oUslugiDodKli, 51, 10) <= DoDaty) Then
                        KlientOK = True
                    End If
                End If
                If chkU6.Checked Then
                    If (Mid(oUslugiDodKli, 6, 1) = "T") And (Mid(oUslugiDodKli, 61, 10) >= OdDaty) And (Mid(oUslugiDodKli, 61, 10) <= DoDaty) Then
                        KlientOK = True
                    End If
                End If
                If chkU7.Checked Then
                    If (Mid(oUslugiDodKli, 7, 1) = "T") And (Mid(oUslugiDodKli, 71, 10) >= OdDaty) And (Mid(oUslugiDodKli, 71, 10) <= DoDaty) Then
                        KlientOK = True
                    End If
                End If
                If chkU8.Checked Then
                    If (Mid(oUslugiDodKli, 8, 1) = "T") And (Mid(oUslugiDodKli, 81, 10) >= OdDaty) And (Mid(oUslugiDodKli, 81, 10) <= DoDaty) Then
                        KlientOK = True
                    End If
                End If
                If chkU9.Checked Then
                    If (Mid(oUslugiDodKli, 9, 1) = "T") And (Mid(oUslugiDodKli, 91, 10) >= OdDaty) And (Mid(oUslugiDodKli, 91, 10) <= DoDaty) Then
                        KlientOK = True
                    End If
                End If




                If KlientOK Then
                    'dopisz tego klienta do ROBOCZY
                    drRoboczy = DataSetRobo.Tables(0).NewRow()
                    drRoboczy.Item("IDENTKLI") = SymbKlienta
                    DataSetRobo.Tables(0).Rows.Add(drRoboczy)
                End If
                System.Windows.Forms.Application.DoEvents()
            Next

            '-------------------------------
            If DataViewRobo.Count > 0 Then
                'definiuj delegat
                Dim deleg As delegateAktualizuj = _AktualizujKartKlientow
                Dim FormRaport As New RaportWybierzAkcjeMarketingowa("", _
                                                                     _dsKlenci, _
                                                                     DataViewRobo, _
                                                                     deleg, _
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


End Class