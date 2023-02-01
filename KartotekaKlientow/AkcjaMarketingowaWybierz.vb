Public Class AkcjaMarketingowaWybierz

    Dim OdDaty As String = ""
    Dim DoDaty As String = ""
    Dim TakenDate As Date = Nothing

    Private Sub WybierzAkcjeMarketingowa_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        Me.PictureBox1.Image = My.Resources.logomini
        '-------------------------------
        txtIleZap.Text = 0.ToString("N0")
        txtIdent.Text = ""
        txtOpis.Text = ""
        txtData.Text = SysData()
        txtUwagi.Text = ""

        oWybranoAkcjeMarketingowa = False
        '-------------------------------
        dtpOdDaty.Value = Mid(SysData(), 1, 7) & "-01"
        dtpDoDaty.Value = SysData()
        rbtWszyscy.Checked = True
        '---------------------------------------

    End Sub

    Private Sub WybierzAkcjeMarketingowa_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        RysujGradient(Me)
    End Sub

    '*******************************************
    '** procedury obslugi edycji
    '*******************************************

    'Private Sub txtOpis_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOpis.TextChanged
    '    TestLen(txtOpis, "OPIS", 100)
    'End Sub
    'Private Sub txtOpis_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOpis.GotFocus
    '    StartEdycjiTxt(txtOpis)
    'End Sub
    'Private Sub txtOpis_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOpis.LostFocus
    '    KoniecEdycjiTxt(txtOpis)
    'End Sub

    'Private Sub txtData_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtData.TextChanged
    '    TestLen(txtData, "DATA", 10)
    'End Sub
    'Private Sub txtData_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtData.GotFocus
    '    StartEdycjiTxt(txtData)
    'End Sub
    'Private Sub txtData_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtData.LostFocus
    '    KoniecEdycjiTxt(txtData)
    'End Sub

    Private Sub txtIdent_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIdent.TextChanged
        TestLen(txtIdent, "IDENTYFIKATOR", 20)
    End Sub
    Private Sub txtIdent_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIdent.GotFocus
        StartEdycjiTxt(txtIdent)
    End Sub
    Private Sub txtIdent_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIdent.LostFocus
        KoniecEdycjiTxt(txtIdent)
    End Sub

    'Private Sub txtUwagi_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUwagi.GotFocus
    '    StartEdycjiTxt(txtUwagi)
    'End Sub
    'Private Sub txtUwagi_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUwagi.LostFocus
    '    KoniecEdycjiTxt(txtUwagi)
    'End Sub

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub

    Private Sub cmdAkcja_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAkcja.Click
        oCoMamRobic = "WYBIERAJ"
        oAkcja = txtIdent.Text

        Dim Proc As New KatalogAkcjiOpis
        Proc.ShowDialog()

        txtIdent.Text = oAkcja
        txtIleZap.Text = IleKlientowObejmujeAkcja(oAkcja).ToString("N0")
        txtData.Text = aoDataAkcji
        txtOpis.Text = aoOpisAkcji
        txtUwagi.Text = aoUwagiAkcji
    End Sub

    '**************************************************
    '** zapamietaj informacje o Akcjach...
    '**************************************************

    'Akcje Handlowe - Opis
    'Public aoIdentAkcji As String            'IDENT     Text 20
    'Public aoDataAkcji As String             'DATA      Text,10
    'Public aoOpisAkcji As String             'OPIS      Text,100
    'Public aoUwagiAkcji As String            'UWAGI     Text,10
    'Public aoMarkerAkcji As String           'MARKER    logical
    'Public aoWersjaAkcji As Integer           'WERSJA    Integer

    Dim dbConnectionAkcjeOpis As OleDb.OleDbConnection
    Dim dbCommandSelectAkcjeOpis As OleDb.OleDbCommand
    Dim dbCommandDeleteAkcjeOpis As OleDb.OleDbCommand
    Dim dbCommandUpdateAkcjeOpis As OleDb.OleDbCommand
    Dim dbCommandInsertAkcjeOpis As OleDb.OleDbCommand
    Dim dbDataAdapterAkcjeOpis As OleDb.OleDbDataAdapter

    Dim sqlConnectionAkcjeOpis As SqlClient.SqlConnection
    Dim sqlCommandSelectAkcjeOpis As SqlClient.SqlCommand
    Dim sqlCommandDeleteAkcjeOpis As SqlClient.SqlCommand
    Dim sqlCommandUpdateAkcjeOpis As SqlClient.SqlCommand
    Dim sqlCommandInsertAkcjeOpis As SqlClient.SqlCommand
    Dim sqlDataAdapterAkcjeOpis As SqlClient.SqlDataAdapter

    Dim DataSetAkcjeOpis As DataSet
    Dim DataViewAkcjeOpis As DataView

    '------------------------------------------------
    'AkcJe handlowe - spec
    Dim dbSelectSQLAkcjeSpec As String = sSelectSQLAkcjeSpec

    Dim dbConnectionAkcjeSpec As OleDb.OleDbConnection
    Dim dbCommandSelectAkcjeSpec As OleDb.OleDbCommand
    Dim dbCommandDeleteAkcjeSpec As OleDb.OleDbCommand
    Dim dbCommandUpdateAkcjeSpec As OleDb.OleDbCommand
    Dim dbCommandInsertAkcjeSpec As OleDb.OleDbCommand
    Dim dbDataAdapterAkcjeSpec As OleDb.OleDbDataAdapter

    Dim sqlConnectionAkcjeSpec As SqlClient.SqlConnection
    Dim sqlCommandSelectAkcjeSpec As SqlClient.SqlCommand
    Dim sqlCommandDeleteAkcjeSpec As SqlClient.SqlCommand
    Dim sqlCommandUpdateAkcjeSpec As SqlClient.SqlCommand
    Dim sqlCommandInsertAkcjeSpec As SqlClient.SqlCommand
    Dim sqlDataAdapterAkcjeSpec As SqlClient.SqlDataAdapter

    Dim DataSetAkcjeSpec As DataSet
    Dim DataViewAkcjeSpec As DataView

    Dim ConnectionState As System.Data.ConnectionState



    Private Sub cmdZapamietaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdZapamietaj.Click
        If rbtWszyscy.Checked Then
            OdDaty = ""
            DoDaty = ""
        ElseIf rbtObsluzeni.Checked Or rbtNieObsluzeni.Checked Then
            TakenDate = dtpOdDaty.Value
            OdDaty = TakenDate.ToString("yyyy-MM-dd")
            TakenDate = dtpDoDaty.Value
            DoDaty = TakenDate.ToString("yyyy-MM-dd")
        Else        'def branze
            OdDaty = ""
            DoDaty = ""
        End If

        '-------------------------------
        If _dsKlenci.Tables(0).Rows.Count = 0 Then
            MessageBox.Show("Specyfikacja klientów jest pusta.",
                "Uwaga:",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            If Len(txtIdent.Text) = 0 Then
                MessageBox.Show("Proszê wpisaæ identyfikator akcji marketingowej",
                    "Uwaga:",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            Else
                If Not CzyJestTakaAkcja(txtIdent.Text) Then
                    MessageBox.Show("Nie istnieje akcja marketingowa o takim identyfikatorze",
                        "Uwaga:",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    DataSetAkcjeSpec = New DataSet
                    DataViewAkcjeSpec = New DataView
                    DataSetAkcjeSpec.Locale = New System.Globalization.CultureInfo("pl-PL")


                    'Public asIdentAkcji As String            'IDENTAKCJI     Text 20
                    'Public asIdentKlienta As String          'IDENTKLI       Text 6
                    'Public asWersjaAkcji As Integer           'WERSJA    Integer
                    '---------------------------------------------------------------------
                    'Kontakty
                    'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
                    'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
                    'Public oPracownikKon As String        'PRACOWNIK        Text, 2     uzytkownik
                    'Public oTematKon As String            'TEMAT            Text, 50
                    'Public oUwagiKon As String            'UWAGI            Memo
                    'Public oWersjaKon As Integer          'WERSJA           Integer

                    'Dim dbSelectSQL As String = "SELECT * FROM AkcjeSpec WHERE IDENTAKCJI='" & Ident & "'"

                    If rbtWszyscy.Checked Then
                        dbSelectSQLAkcjeSpec = "SELECT * FROM AkcjeSpec WHERE IDENTAKCJI='" & txtIdent.Text & "'"
                    ElseIf rbtObsluzeni.Checked Then
                        dbSelectSQLAkcjeSpec = "SELECT * FROM AkcjeSpec WHERE IDENTAKCJI='" & txtIdent.Text & "' " &
                            " AND ((SELECT count(*) FROM KontaktyHandlowe WHERE (DATAKONTAKTU>='" & OdDaty & "') AND " &
                                                                        "(DATAKONTAKTU<='" & DoDaty & "') AND " &
                                                                        "(KontaktyHandlowe.IDENTKLIENTA=AkcjeSpec.IDENTKLI)) > 0) "
                    ElseIf rbtnieObsluzeni.Checked Then
                        dbSelectSQLAkcjeSpec = "SELECT * FROM AkcjeSpec WHERE IDENTAKCJI='" & txtIdent.Text & "' " &
                            " AND ((SELECT count(*) FROM KontaktyHandlowe WHERE (DATAKONTAKTU>='" & OdDaty & "') AND " &
                                                                        "(DATAKONTAKTU<='" & DoDaty & "') AND " &
                                                                        "(KontaktyHandlowe.IDENTKLIENTA=AkcjeSpec.IDENTKLI)) = 0) "

                    Else
                        'definiuj branze
                        dbSelectSQLAkcjeSpec = "SELECT * FROM AkcjeSpec WHERE IDENTAKCJI='" & txtIdent.Text & "'"
                    End If

                    If ParL_DataBaseType = "MS ACCESS" Then
                        'OK - mo¿na zapisaæ....
                        dbConnectionAkcjeSpec = New OleDb.OleDbConnection(sConnectionAkcjeSpec)
                        dbCommandSelectAkcjeSpec = New OleDb.OleDbCommand(dbSelectSQLAkcjeSpec, dbConnectionAkcjeSpec)
                        'dbCommandDeleteAkcjeSpec = New OleDb.OleDbCommand(sDeleteSQLAkcjeSpec, dbConnectionAkcjeSpec)
                        'dbCommandUpdateAkcjeSpec = New OleDb.OleDbCommand(sUpdateSQLAkcjeSpec, dbConnectionAkcjeSpec)
                        'dbCommandInsertAkcjeSpec = New OleDb.OleDbCommand(sInsertSQLAkcjeSpec, dbConnectionAkcjeSpec)
                        dbDataAdapterAkcjeSpec = New OleDb.OleDbDataAdapter(dbCommandSelectAkcjeSpec)

                        Dim dbMapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
                        dbMapowanieTabeliAkcjeSpec = dbDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_dbAkcjeSpec")

                        Try
                            dbConnectionAkcjeSpec.Open()
                        Catch Ex As System.Exception
                            ViewInLog(Ex, Me.Name(), Nothing)
                            'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                            '    System.Windows.Forms.MessageBoxButtons.OK, _
                            '    MessageBoxIcon.Information, _
                            '    MessageBoxDefaultButton.Button1)
                        Finally
                            ConnectionState = dbConnectionAkcjeSpec.State
                        End Try
                    Else
                        'OK - mo¿na zapisaæ....
                        sqlConnectionAkcjeSpec = New SqlClient.SqlConnection(sConnectionAkcjeSpec)
                        sqlCommandSelectAkcjeSpec = New SqlClient.SqlCommand(dbSelectSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                        'sqlCommandDeleteAkcjeSpec = New SqlClient.SqlCommand(sDeleteSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                        'sqlCommandUpdateAkcjeSpec = New SqlClient.SqlCommand(sUpdateSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                        'sqlCommandInsertAkcjeSpec = New SqlClient.SqlCommand(sInsertSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                        sqlDataAdapterAkcjeSpec = New SqlClient.SqlDataAdapter(sqlCommandSelectAkcjeSpec)

                        Dim sqlMapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
                        sqlMapowanieTabeliAkcjeSpec = sqlDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_sqlAkcjeSpec")

                        Try
                            sqlConnectionAkcjeSpec.Open()
                        Catch Ex As System.Exception
                            ViewInLog(Ex, Me.Name(), Nothing)
                            'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                            '    System.Windows.Forms.MessageBoxButtons.OK, _
                            '    MessageBoxIcon.Information, _
                            '    MessageBoxDefaultButton.Button1)
                        Finally
                            ConnectionState = sqlConnectionAkcjeSpec.State
                        End Try
                    End If

                    If ConnectionState = ConnectionState.Open Then
                        Try
                            If ParL_DataBaseType = "MS ACCESS" Then
                                dbDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_dbAkcjeSpec")
                                dbDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                                dbConnectionAkcjeSpec.Close()
                            Else
                                sqlDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_sqlAkcjeSpec")
                                sqlDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                                sqlConnectionAkcjeSpec.Close()
                            End If

                            'zdefiniuj unikalny klucz wyszukiwania w tabeli
                            DataSetAkcjeSpec.Tables(0).PrimaryKey = New DataColumn() {DataSetAkcjeSpec.Tables(0).Columns("IDENTAKCJI"), DataSetAkcjeSpec.Tables(0).Columns("IDENTKLI")}

                            'definiuj DataView
                            DataViewAkcjeSpec = New DataView(DataSetAkcjeSpec.Tables(0))
                            DataViewAkcjeSpec.AllowDelete = False
                            DataViewAkcjeSpec.AllowEdit = False
                            DataViewAkcjeSpec.AllowNew = False

                            If DataViewAkcjeSpec.Count > 0 Then
                                If rbtWszyscy.Checked Or rbtObsluzeni.Checked Or rbtNieObsluzeni.Checked Then
                                    '...........................
                                    'definiuj delegat
                                    Dim deleg As delegateAktualizuj = _AktualizujKartKlientow
                                    Dim FormRaport As New RaportWybierzAkcjeMarketingowa(txtIdent.Text,
                                                                             _dsKlenci,
                                                                             DataViewAkcjeSpec,
                                                                             deleg,
                                                                             True)
                                    FormRaport.ShowDialog()
                                    oWybranoAkcjeMarketingowa = True
                                    '...........................
                                    MessageBox.Show("Pobra³em informacje o akcji marketingowej" & vbCrLf & txtIdent.Text,
                                    "OK - skoñczy³em",
                                    System.Windows.Forms.MessageBoxButtons.OK,
                                    MessageBoxIcon.Information,
                                    MessageBoxDefaultButton.Button1)
                                Else        'def branze
                                    '...........................
                                    'definiuj delegat
                                    Dim deleg As delegateAktualizuj = _AktualizujKartKlientow
                                    Dim FormRaport As New RaportWybierzAkcjeMarketingowaIDefBranze(txtIdent.Text,
                                                                             _dsKlenci,
                                                                             DataViewAkcjeSpec,
                                                                             deleg,
                                                                             txtBranza.Text,
                                                                             txtPodbranza.Text)
                                    FormRaport.ShowDialog()
                                    oWybranoAkcjeMarketingowa = False
                                    '...........................
                                    MessageBox.Show("Zdefiniowa³em Bran¿e na podstawie akcji marketingowej" & vbCrLf & txtIdent.Text,
                                            "OK - skoñczy³em",
                                            System.Windows.Forms.MessageBoxButtons.OK,
                                            MessageBoxIcon.Information,
                                            MessageBoxDefaultButton.Button1)
                                End If
                            End If




                        Catch Ex As System.Exception
                            ViewInLog(Ex, Me.Name(), Nothing)
                            'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                            '    System.Windows.Forms.MessageBoxButtons.OK, _
                            '    MessageBoxIcon.Information, _
                            '    MessageBoxDefaultButton.Button1)
                        End Try
                    End If

                End If
            End If
        End If
    End Sub


    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        If rbtWszyscy.Checked Then
            OdDaty = ""
            DoDaty = ""
        Else
            TakenDate = dtpOdDaty.Value
            OdDaty = TakenDate.ToString("yyyy-MM-dd")
            TakenDate = dtpDoDaty.Value
            DoDaty = TakenDate.ToString("yyyy-MM-dd")
        End If
        '-------------------------------
        If _dsKlenci.Tables(0).Rows.Count = 0 Then
            MessageBox.Show("Specyfikacja klientów jest pusta.",
                "Uwaga:",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            If Len(txtIdent.Text) = 0 Then
                MessageBox.Show("Proszê wpisaæ identyfikator akcji marketingowej",
                    "Uwaga:",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            Else
                If Not CzyJestTakaAkcja(txtIdent.Text) Then
                    MessageBox.Show("Nie istnieje akcja marketingowa o takim identyfikatorze",
                        "Uwaga:",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    DataSetAkcjeSpec = New DataSet
                    DataViewAkcjeSpec = New DataView
                    DataSetAkcjeSpec.Locale = New System.Globalization.CultureInfo("pl-PL")
                    dbSelectSQLAkcjeSpec = "SELECT * FROM AkcjeSpec WHERE IDENTAKCJI='" & txtIdent.Text & "'"

                    If ParL_DataBaseType = "MS ACCESS" Then
                        'OK - mo¿na zapisaæ....
                        dbConnectionAkcjeSpec = New OleDb.OleDbConnection(sConnectionAkcjeSpec)
                        dbCommandSelectAkcjeSpec = New OleDb.OleDbCommand(dbSelectSQLAkcjeSpec, dbConnectionAkcjeSpec)
                        'dbCommandDeleteAkcjeSpec = New OleDb.OleDbCommand(sDeleteSQLAkcjeSpec, dbConnectionAkcjeSpec)
                        'dbCommandUpdateAkcjeSpec = New OleDb.OleDbCommand(sUpdateSQLAkcjeSpec, dbConnectionAkcjeSpec)
                        'dbCommandInsertAkcjeSpec = New OleDb.OleDbCommand(sInsertSQLAkcjeSpec, dbConnectionAkcjeSpec)
                        dbDataAdapterAkcjeSpec = New OleDb.OleDbDataAdapter(dbCommandSelectAkcjeSpec)

                        Dim dbMapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
                        dbMapowanieTabeliAkcjeSpec = dbDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_dbAkcjeSpec")

                        Try
                            dbConnectionAkcjeSpec.Open()
                        Catch Ex As System.Exception
                            ViewInLog(Ex, Me.Name(), Nothing)
                            'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                            '    System.Windows.Forms.MessageBoxButtons.OK, _
                            '    MessageBoxIcon.Information, _
                            '    MessageBoxDefaultButton.Button1)
                        Finally
                            ConnectionState = dbConnectionAkcjeSpec.State
                        End Try
                    Else
                        'OK - mo¿na zapisaæ....
                        sqlConnectionAkcjeSpec = New SqlClient.SqlConnection(sConnectionAkcjeSpec)
                        sqlCommandSelectAkcjeSpec = New SqlClient.SqlCommand(dbSelectSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                        'sqlCommandDeleteAkcjeSpec = New SqlClient.SqlCommand(sDeleteSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                        'sqlCommandUpdateAkcjeSpec = New SqlClient.SqlCommand(sUpdateSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                        'sqlCommandInsertAkcjeSpec = New SqlClient.SqlCommand(sInsertSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                        sqlDataAdapterAkcjeSpec = New SqlClient.SqlDataAdapter(sqlCommandSelectAkcjeSpec)

                        Dim sqlMapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
                        sqlMapowanieTabeliAkcjeSpec = sqlDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_sqlAkcjeSpec")

                        Try
                            sqlConnectionAkcjeSpec.Open()
                        Catch Ex As System.Exception
                            ViewInLog(Ex, Me.Name(), Nothing)
                            'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                            '    System.Windows.Forms.MessageBoxButtons.OK, _
                            '    MessageBoxIcon.Information, _
                            '    MessageBoxDefaultButton.Button1)
                        Finally
                            ConnectionState = sqlConnectionAkcjeSpec.State
                        End Try
                    End If

                    If ConnectionState = ConnectionState.Open Then
                        Try
                            If ParL_DataBaseType = "MS ACCESS" Then
                                dbDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_dbAkcjeSpec")
                                dbDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                                dbConnectionAkcjeSpec.Close()
                            Else
                                sqlDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_sqlAkcjeSpec")
                                sqlDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                                sqlConnectionAkcjeSpec.Close()
                            End If

                            'zdefiniuj unikalny klucz wyszukiwania w tabeli
                            DataSetAkcjeSpec.Tables(0).PrimaryKey = New DataColumn() {DataSetAkcjeSpec.Tables(0).Columns("IDENTAKCJI"), DataSetAkcjeSpec.Tables(0).Columns("IDENTKLI")}

                            'definiuj DataView
                            DataViewAkcjeSpec = New DataView(DataSetAkcjeSpec.Tables(0))
                            DataViewAkcjeSpec.AllowDelete = False
                            DataViewAkcjeSpec.AllowEdit = False
                            DataViewAkcjeSpec.AllowNew = False

                        Catch Ex As System.Exception
                            ViewInLog(Ex, Me.Name(), Nothing)
                            'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                            '    System.Windows.Forms.MessageBoxButtons.OK, _
                            '    MessageBoxIcon.Information, _
                            '    MessageBoxDefaultButton.Button1)
                        End Try
                    End If

                    '...........................
                    'definiuj delegat
                    Dim deleg As delegateAktualizuj = _AktualizujKartKlientow
                    Dim FormRaport As New RaportWybierzAkcjeMarketingowa(txtIdent.Text,
                                                                         _dsKlenci,
                                                                         DataViewAkcjeSpec,
                                                                         deleg,
                                                                         False)
                    FormRaport.ShowDialog()
                    '...........................
                    MessageBox.Show("Pobra³em informacje o akcji marketingowej" & vbCrLf & txtIdent.Text,
                        "OK - skoñczy³em",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)


                End If
            End If
        End If
    End Sub








    Private Sub rbtWszyscy_CheckedChanged(sender As Object, e As EventArgs) Handles rbtWszyscy.CheckedChanged, rbtObsluzeni.CheckedChanged, rbtNieObsluzeni.CheckedChanged, rbtDefiniujBranze.CheckedChanged
        If rbtWszyscy.Checked Then
            lblOdDaty.Visible = False
            dtpOdDaty.Visible = False
            lblDoDaty.Visible = False
            dtpDoDaty.Visible = False

            LabelBranza.Visible = False
            txtBranza.Visible = False
            cmdWybierzBranze.Visible = False
            LabelPodBranza.Visible = False
            txtPodbranza.Visible = False
            cmdWybierzPodbranze.Visible = False

            cmdZapamietaj.Text = "Pobierz akcje"
            cmdDodaj.Visible = True

        ElseIf rbtObsluzeni.Checked Or rbtNieObsluzeni.Checked Then
            lblOdDaty.Visible = True
            dtpOdDaty.Visible = True
            lblDoDaty.Visible = True
            dtpDoDaty.Visible = True

            LabelBranza.Visible = False
            txtBranza.Visible = False
            cmdWybierzBranze.Visible = False
            LabelPodBranza.Visible = False
            txtPodbranza.Visible = False
            cmdWybierzPodbranze.Visible = False

            cmdZapamietaj.Text = "Pobierz akcje"
            cmdDodaj.Visible = True

        Else        'def branze
            lblOdDaty.Visible = False
            dtpOdDaty.Visible = False
            lblDoDaty.Visible = False
            dtpDoDaty.Visible = False

            LabelBranza.Visible = True
            txtBranza.Visible = True
            cmdWybierzBranze.Visible = True
            LabelPodBranza.Visible = True
            txtPodbranza.Visible = True
            cmdWybierzPodbranze.Visible = True

            cmdZapamietaj.Text = "Pobierz bran¿e"
            cmdDodaj.Visible = False
        End If
    End Sub

    'Private Sub rbtObsluzeni_CheckedChanged(sender As Object, e As EventArgs) Handles rbtObsluzeni.CheckedChanged
    '    If rbtWszyscy.Checked Then
    '        lblOdDaty.Visible = False
    '        dtpOdDaty.Visible = False
    '        lblDoDaty.Visible = False
    '        dtpDoDaty.Visible = False
    '    Else
    '        lblOdDaty.Visible = True
    '        dtpOdDaty.Visible = True
    '        lblDoDaty.Visible = True
    '        dtpDoDaty.Visible = True
    '    End If
    'End Sub

    'Private Sub rbtNieObsluzeni_CheckedChanged(sender As Object, e As EventArgs) Handles rbtNieObsluzeni.CheckedChanged
    '    If rbtWszyscy.Checked Then
    '        lblOdDaty.Visible = False
    '        dtpOdDaty.Visible = False
    '        lblDoDaty.Visible = False
    '        dtpDoDaty.Visible = False
    '    Else
    '        lblOdDaty.Visible = True
    '        dtpOdDaty.Visible = True
    '        lblDoDaty.Visible = True
    '        dtpDoDaty.Visible = True
    '    End If
    'End Sub
















    Private Sub txtBranza_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBranza.TextChanged
        TestLen(txtBranza, "BRAN¯A", 100)
    End Sub
    Private Sub txtBranza_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBranza.GotFocus
        StartEdycjiTxt(txtBranza)
    End Sub
    Private Sub txtBranza_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBranza.LostFocus
        KoniecEdycjiTxt(txtBranza)
    End Sub


    Private Sub txtPodBranza_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPodbranza.TextChanged
        TestLen(txtPodbranza, "PODBRAN¯A", 100)
    End Sub
    Private Sub txtPodBranza_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPodbranza.GotFocus
        StartEdycjiTxt(txtPodbranza)
    End Sub
    Private Sub txtPodBranza_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPodbranza.LostFocus
        KoniecEdycjiTxt(txtPodbranza)
    End Sub

    Private Sub cmdWybierzBranze_Click(sender As Object, e As EventArgs) Handles cmdWybierzBranze.Click
        oCoMamRobic = "WYBIERAJ"
        oIdBranzy = Trim(txtBranza.Text)
        Dim form As New KatalogBranze
        form.ShowDialog()
        If Len(Trim(oIdBranzy)) > 0 Then
            txtBranza.Text = oIdBranzy
        End If
    End Sub

    Private Sub cmdWybierzPodBranze_Click(sender As Object, e As EventArgs) Handles cmdWybierzPodbranze.Click
        oCoMamRobic = "WYBIERAJ"
        oIdBranzy = Trim(txtBranza.Text)
        oIdPodBranzy = Trim(txtPodbranza.Text)
        Dim form As New KatalogPodBranze
        form.ShowDialog()
        If Len(Trim(oIdBranzy)) > 0 Then
            txtBranza.Text = oIdBranzy
        End If
        If Len(Trim(oIdPodBranzy)) > 0 Then
            txtPodbranza.Text = oIdPodBranzy
        End If
    End Sub

End Class
