Public Class AkcjaMarketingowaWybierzWgOsobKontaktowych

    Dim Osoba As String = ""

    Private Sub WybierzAkcjeMarketingowa_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        oWybranoAkcjeMarketingowa = False
        txtImieINazwisko.Text = ""
    End Sub

    Private Sub WybierzAkcjeMarketingowa_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        RysujGradient(Me)
    End Sub


    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub


    Private Sub txtImieINazwisko_TextChanged(sender As Object, e As EventArgs) Handles txtImieINazwisko.TextChanged
        'TestLen(txtImieINazwisko, "Imie i Nazwisko", 50)
    End Sub
    Private Sub txtImieINazwisko_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtImieINazwisko.GotFocus
        StartEdycjiTxt(txtImieINazwisko)
    End Sub
    Private Sub txtImieINazwisko_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtImieINazwisko.LostFocus
        KoniecEdycjiTxt(txtImieINazwisko)
    End Sub

    Private Sub txtImieINazwisko_KeyDown(sender As Object, e As KeyEventArgs) Handles txtImieINazwisko.KeyDown
        'sprawdzaj znacznik konca kodu kreskowego - CrLf
        Select Case e.KeyCode
            Case Keys.Enter
                cmdZapamietaj_Click(sender, e)
            Case Keys.Tab
            Case Else
        End Select
    End Sub




    '**************************************************
    '** identyfikacja klientów wg osób kontaktowych
    '**************************************************

    ''---------------------------------------------------------------------
    ''Osoby Kontaktowe
    ''---------------------------------------------------------------------
    'Public oIdentOso As String            'IDENTKLIENTA     Text, 6
    'Public oOsobaOso As String            'OSOBA            Text, 100
    'Public oVIPOso As Boolean             'VIP              boolean
    'Public oeMailOso As String            'EMAIL            Text, 100
    'Public oTelefonOso As String          'TELEFON          Text, 100
    'Public oTelefonKomOso As String       'TELEFONKOM       Text, 100
    'Public oStanowiskoOso As String       'STANOWISKO       Text, 100
    'Public oKompetencjeOso As String      'KOMPETENCJE      Text, 100
    'Public oWersjaOso As Integer          'WERSJA           Integer

    Dim dbSelectSQLRobo As String = ""       'sSelectSQLRobo & " ORDER BY IDENTKLIENTA"

    Dim dbConnectionRobo As OleDb.OleDbConnection
    Dim dbCommandSelectRobo As OleDb.OleDbCommand
    Dim dbCommandDeleteRobo As OleDb.OleDbCommand
    Dim dbCommandUpdateRobo As OleDb.OleDbCommand
    Dim dbCommandInsertRobo As OleDb.OleDbCommand
    Dim dbDataAdapterRobo As OleDb.OleDbDataAdapter

    Dim sqlConnectionRobo As SqlClient.SqlConnection
    Dim sqlCommandSelectRobo As SqlClient.SqlCommand
    Dim sqlCommandDeleteRobo As SqlClient.SqlCommand
    Dim sqlCommandUpdateRobo As SqlClient.SqlCommand
    Dim sqlCommandInsertRobo As SqlClient.SqlCommand
    Dim sqlDataAdapterRobo As SqlClient.SqlDataAdapter

    Dim DataSetRobo As New DataSet
    Dim DataViewRobo As New DataView
    '---------------------------
    Dim ConnectionState As System.Data.ConnectionState



    'Select Case KLIENCI.IDENTKLIENTA FROM KLIENCI INNER JOIN Osoby On (Osoby.IDENTKLIENTA=Klienci.IDENTKLIENTA) WHERE (Osoby.IDENTKLIENTA=Klienci.IDENTKLIENTA) And (OSOBY.Osoba Like '%JAN')

    'Select Case KLIENCI.IDENTKLIENTA FROM KLIENCI INNER JOIN Osoby On (Osoby.IDENTKLIENTA=Klienci.IDENTKLIENTA) WHERE (Osoby.IDENTKLIENTA=Klienci.IDENTKLIENTA) And (OSOBY.Osoba Like '%NOWAK')

    Private Sub cmdZapamietaj_Click(sender As Object, e As EventArgs) Handles cmdZapamietaj.Click
        Osoba = txtImieINazwisko.Text
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



            '-----------------------------
            'analizuj klientow
            '-----------------------------
            dbSelectSQLRobo = "Select Klienci.IDENTKLIENTA AS IDENTKLI FROM Klienci INNER JOIN Osoby On (Osoby.IDENTKLIENTA=Klienci.IDENTKLIENTA) WHERE (Osoby.IDENTKLIENTA=Klienci.IDENTKLIENTA) And (OSOBY.Osoba Like '%" & Trim(txtImieINazwisko.Text) & "%')"

            'DataSet
            DataSetRobo.Locale = New System.Globalization.CultureInfo("pl-PL")
            If ParL_DataBaseType = "MS ACCESS" Then
                'dbConnectionRobo = New OleDb.OleDbConnection(sConnectionRobo)
                'dbCommandSelectRobo = New OleDb.OleDbCommand(dbSelectSQLRobo, dbConnectionRobo)
                ''dbCommandDeleteRobo = New OleDb.OleDbCommand(sDeleteSQLRobo, dbConnectionRobo)
                ''dbCommandUpdateRobo = New OleDb.OleDbCommand(sUpdateSQLRobo, dbConnectionRobo)
                ''dbCommandInsertRobo = New OleDb.OleDbCommand(sInsertSQLRobo, dbConnectionRobo)
                'dbDataAdapterRobo = New OleDb.OleDbDataAdapter(dbCommandSelectRobo)

                ''mapujemy domyslna nazwe tabeli
                'Dim MapowanieTabeliRobo As System.Data.Common.DataTableMapping
                'MapowanieTabeliRobo = dbDataAdapterRobo.TableMappings.Add("Table", "TABELA_Robo")

                '' Add the RowUpdated event handler.
                'AddHandler dbDataAdapterRobo.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

                ''------------------------------------------
                ''wypelnij DATASET
                'Try
                '    dbConnectionRobo.Open()
                'Catch Ex As System.Exception
                '    ViewInLog(Ex, Me.Name(), Nothing)
                '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    '    MessageBoxIcon.Information, _
                '    '    MessageBoxDefaultButton.Button1)
                'Finally
                '    ConnectionState = dbConnectionRobo.State
                'End Try

            Else
                sqlConnectionRobo = New SqlClient.SqlConnection(sConnectionKlienci)
                sqlCommandSelectRobo = New SqlClient.SqlCommand(dbSelectSQLRobo, sqlConnectionRobo)
                'sqlCommandDeleteRobo = New SqlClient.SqlCommand(sDeleteSQLRobo, sqlConnectionRobo)
                'sqlCommandUpdateRobo = New SqlClient.SqlCommand(sUpdateSQLRobo, sqlConnectionRobo)
                'sqlCommandInsertRobo = New SqlClient.SqlCommand(sInsertSQLRobo, sqlConnectionRobo)
                sqlDataAdapterRobo = New SqlClient.SqlDataAdapter(sqlCommandSelectRobo)

                'mapujemy domyslna nazwe tabeli
                Dim MapowanieTabeliRobo As System.Data.Common.DataTableMapping
                MapowanieTabeliRobo = sqlDataAdapterRobo.TableMappings.Add("Table", "TABELA_Robo")

                ' Add the RowUpdated event handler.
                AddHandler sqlDataAdapterRobo.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

                '------------------------------------------
                'wypelnij DATASET
                Try
                    sqlConnectionRobo.Open()
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                Finally
                    ConnectionState = sqlConnectionRobo.State
                End Try
            End If

            If ConnectionState = ConnectionState.Open Then
                Try
                    DataSetRobo.EnforceConstraints = False
                    If ParL_DataBaseType = "MS ACCESS" Then
                        dbDataAdapterRobo.FillSchema(DataSetRobo, SchemaType.Mapped, "TABELA_Robo")
                        dbDataAdapterRobo.Fill(DataSetRobo)
                        dbConnectionRobo.Close()
                    Else
                        sqlDataAdapterRobo.FillSchema(DataSetRobo, SchemaType.Mapped, "TABELA_Robo")
                        sqlDataAdapterRobo.Fill(DataSetRobo)
                        sqlConnectionRobo.Close()
                    End If

                    DataViewRobo = New DataView(DataSetRobo.Tables(0))

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

                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                End Try
                '---------------------------------
            End If
        End If
    End Sub

End Class