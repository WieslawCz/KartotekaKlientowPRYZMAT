Public Class AkcjaMarketingowaWybierzWgSprzedazy

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
        InicjujListeProcent(cbxProcent)
        cbxProcent.Text = "100"
        '---------------------------------------
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
            Dim robCol2 As DataColumn = DataTableRobo.Columns.Add("WARTSPRZED", GetType(System.Double))
            Dim robCol3 As DataColumn = DataTableRobo.Columns.Add("NARASTAJACO", GetType(System.Double))
            Dim robCol4 As DataColumn = DataTableRobo.Columns.Add("PROCENT", GetType(System.Double))
            DataSetRobo.Tables.Add(DataTableRobo)
            'zdefiniuj unikalny klucz wyszukiwania w tabeli
            DataSetRobo.Tables(0).PrimaryKey = New DataColumn() {DataSetRobo.Tables(0).Columns("IDENTKLI")}
            'definiuj DataView
            DataViewRobo = New DataView(DataSetRobo.Tables(0))

            '-----------------------------
            'analizuj klientow
            '-----------------------------

            dbSelectSQLKlienci = "SELECT * FROM " &
                             "(SELECT IDENTKLIENTA," &
                                    "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED+" &
                                                        "NAJEMWARTSPRZED+STRONYWARTSPRZED+DRUKARKIWARTSPRZED+SKUPWARTSPRZED) " &
                                    " AS SPRZEDO FROM Obroty     WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA)     And (SUBSTRING(Obroty.DATA,1,10)>='" & Mid(OdDaty, 1, 10) & "') AND (SUBSTRING(Obroty.DATA,1,10)<='" & Mid(DoDaty, 1, 10) & "') ),0) + " &
                                    " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED+" &
                                                        "NAJEMWARTSPRZED+STRONYWARTSPRZED+DRUKARKIWARTSPRZED+SKUPWARTSPRZED) " &
                                    " AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) And (SUBSTRING(ObrotyMies.MIESIAC,1,07)>='" & Mid(OdDaty, 1, 7) & "')  AND (SUBSTRING(ObrotyMies.MIESIAC,1,07)<='" & Mid(DoDaty, 1, 7) & "')),0)) AS SPRZEDAZ " &
                            " FROM Klienci) AS KL " &
                            " ORDER BY KL.SPRZEDAZ DESC"



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
                sqlCommandSelectKlienci = New SqlClient.SqlCommand(dbSelectSQLKlienci, sqlConnectionKlienci)
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









            ' analizuj poszczegolnych klientow
            ' a dla klientow analizuj obroty miesieczne
            ' dopisuj do bazy AnalizyZakupu
            Dim drKlienci As DataRow = Nothing
            Dim drRoboczy As DataRow = Nothing
            Dim drvObroty As DataRowView = Nothing
            Dim drvRoboczy As DataRowView = Nothing

            Dim ik As Integer = 0
            Dim io As Integer = 0
            Dim ir As Integer = 0
            Dim SymbKlienta As String = ""
            Dim IleRecKlienci As Integer = 0
            Dim TotalSprzedaz As Double = 0
            Dim GranSprzedaz As Double = 0
            Dim NarastSprzedaz As Double = 0

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
                SymbKlienta = DataSetKlienci.Tables(0).Rows.Item(ik).Item("IDENTKLIENTA")

                'Dim robCol1 As DataColumn = DataTableRobo.Columns.Add("IDENTKLI", GetType(System.String))
                'Dim robCol2 As DataColumn = DataTableRobo.Columns.Add("WARTSPRZED", GetType(System.Double))
                'Dim robCol3 As DataColumn = DataTableRobo.Columns.Add("NARASTAJACO", GetType(System.Double))
                'Dim robCol4 As DataColumn = DataTableRobo.Columns.Add("PROCENT", GetType(System.Double))

                'Sprawdz czyjest taki zapis

                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = SymbKlienta
                drRoboczy = DataSetRobo.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If (drRoboczy Is Nothing) Then
                    'dopisz tego klienta do ROBOCZY
                    drRoboczy = DataSetRobo.Tables(0).NewRow()

                    drRoboczy.Item("IDENTKLI") = SymbKlienta
                    drRoboczy.Item("WARTSPRZED") = DataSetKlienci.Tables(0).Rows.Item(ik).Item("SPRZEDAZ")
                    drRoboczy.Item("NARASTAJACO") = 0
                    drRoboczy.Item("PROCENT") = 0

                    'dodaj rekord do DataSet
                    DataSetRobo.Tables(0).Rows.Add(drRoboczy)
                End If
                TotalSprzedaz += drRoboczy.Item("WARTSPRZED")

                System.Windows.Forms.Application.DoEvents()
            Next
            '-------------------------------
            GranSprzedaz = Math.Round((CInt(cbxProcent.Text) / 100) * TotalSprzedaz, 2)
            '-------------------------------
            If DataViewRobo.Count > 0 Then
                'sortuj malej¹co wg wartosci sprzedazy
                DataViewRobo.Sort = "WARTSPRZED DESC"
                'wyliczamy wartosc sprzedazy narastajaco
                NarastSprzedaz = 0
                For ir = 0 To DataViewRobo.Count - 1
                    drvRoboczy = DataViewRobo.Item(ir)
                    NarastSprzedaz += drvRoboczy("WARTSPRZED")

                    drvRoboczy("NARASTAJACO") = NarastSprzedaz
                    If TotalSprzedaz <> 0 Then
                        drvRoboczy("PROCENT") = 100 * NarastSprzedaz / TotalSprzedaz
                    Else
                        drvRoboczy("PROCENT") = 0
                    End If

                    If NarastSprzedaz >= GranSprzedaz Then
                        Exit For
                    End If
                Next
                'pozosta³e zapisy z ROBO - usun
                'jeœli usuniesz do ir - tzn wartoœæ graniczn¹ okreœla pierwszy klient który PRZEKROCZY 80% sprzeda¿y
                'jeœli usuniesz do (ir-1) - tzn wartoœæ graniczn¹ okreœla ostatni klient który NIE PRZEKROCZY 80% sprzeda¿y
                For ik = DataViewRobo.Count - 1 To ir + 1 Step -1
                    DataViewRobo.Item(ik).Delete()
                Next
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

End Class