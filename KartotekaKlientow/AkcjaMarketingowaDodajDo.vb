Public Class AkcjaMarketingowaDodajDo

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
    End Sub

    Private Sub WybierzAkcjeMarketingowa_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        RysujGradient(Me)
    End Sub

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


    '*******************************************
    '** procedury obslugi edycji
    '*******************************************
    Private Sub txtIdent_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIdent.TextChanged
        TestLen(txtIdent, "IDENTYFIKATOR", 20)
    End Sub
    Private Sub txtIdent_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIdent.GotFocus
        StartEdycjiTxt(txtIdent)
    End Sub
    Private Sub txtIdent_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIdent.LostFocus
        KoniecEdycjiTxt(txtIdent)
    End Sub


    '**************************************************
    '** Dodaj Info do Akcji ...
    '**************************************************

    '=====================Akcje Handlowe - Opis
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

    '====================Akcje Handlowe - Specyfikacja
    'Public asIdentAkcji As String            'IDENTAKCJI     Text 20
    'Public asIdentKlienta As String          'IDENTKLI       Text 6
    'Public asWersjaAkcji As Integer           'WERSJA    Integer

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
        Dim IleKlientow As Long = 0

        If _dvKlenci.Count = 0 Then
            MessageBox.Show("Specyfikacja klientów jest pusta.", _
                "Uwaga:", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            If Len(txtIdent.Text) = 0 Then
                MessageBox.Show("Proszê wybraæ akcjê marketingow¹", _
                    "Uwaga:", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            Else
                If Not CzyJestTakaAkcja(txtIdent.Text) Then
                    MessageBox.Show("Nie istnieje akcja marketingowa o takim identyfikatorze", _
                        "Uwaga:", _
                        System.Windows.Forms.MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, _
                        MessageBoxDefaultButton.Button1)
                Else
                    IleKlientow = IleKlientowObejmujeAkcja(txtIdent.Text)
                    If MessageBox.Show("Czy chcesz dopisaæ bie¿¹c¹ listê klientów w ilosci " & Trim(Str(_dvKlenci.Count)) & vbCrLf & _
                                       "do klientów z akcji marketingowej " & txtIdent.Text & _
                                       "licz¹cej obecnie " & Trim(Str(IleKlientow)) & " klientów ?", _
                        "Proszê o potwierdzenie", _
                        System.Windows.Forms.MessageBoxButtons.YesNo, _
                        MessageBoxIcon.Question, _
                        MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then

                        'Dopisujê...
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
                            dbCommandInsertAkcjeSpec = New OleDb.OleDbCommand(sInsertSQLAkcjeSpec, dbConnectionAkcjeSpec)
                            dbDataAdapterAkcjeSpec = New OleDb.OleDbDataAdapter(dbCommandSelectAkcjeSpec)

                            Dim objParamAkcjeSpec As OleDb.OleDbParameter
                            '------------------------------------------
                            'komenda INSERT
                            objParamAkcjeSpec = dbCommandInsertAkcjeSpec.Parameters.Add("@oIdentAkcji", OleDb.OleDbType.Char, 20, "IDENTAKCJI")
                            objParamAkcjeSpec.SourceVersion = DataRowVersion.Current
                            objParamAkcjeSpec = dbCommandInsertAkcjeSpec.Parameters.Add("@oIdentKli", OleDb.OleDbType.Char, 6, "IDENTKLI")
                            objParamAkcjeSpec.SourceVersion = DataRowVersion.Current
                            objParamAkcjeSpec = dbCommandInsertAkcjeSpec.Parameters.Add("@oWersjaAkcji", OleDb.OleDbType.Integer, Nothing, "WERSJA")
                            objParamAkcjeSpec.SourceVersion = DataRowVersion.Current
                            dbDataAdapterAkcjeSpec.InsertCommand = dbCommandInsertAkcjeSpec

                            Dim dbMapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
                            dbMapowanieTabeliAkcjeSpec = dbDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_AkcjeSpec")

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
                            sqlCommandInsertAkcjeSpec = New SqlClient.SqlCommand(sInsertSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                            sqlDataAdapterAkcjeSpec = New SqlClient.SqlDataAdapter(sqlCommandSelectAkcjeSpec)

                            Dim objParamAkcjeSpec As SqlClient.SqlParameter
                            '------------------------------------------
                            'komenda INSERT
                            objParamAkcjeSpec = sqlCommandInsertAkcjeSpec.Parameters.Add("@oIdentAkcji", SqlDbType.Char, 20, "IDENTAKCJI")
                            objParamAkcjeSpec.SourceVersion = DataRowVersion.Current
                            objParamAkcjeSpec = sqlCommandInsertAkcjeSpec.Parameters.Add("@oIdentKli", SqlDbType.Char, 6, "IDENTKLI")
                            objParamAkcjeSpec.SourceVersion = DataRowVersion.Current
                            objParamAkcjeSpec = sqlCommandInsertAkcjeSpec.Parameters.Add("@oWersjaAkcji", SqlDbType.Int, Nothing, "WERSJA")
                            objParamAkcjeSpec.SourceVersion = DataRowVersion.Current
                            sqlDataAdapterAkcjeSpec.InsertCommand = sqlCommandInsertAkcjeSpec

                            Dim sqlMapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
                            sqlMapowanieTabeliAkcjeSpec = sqlDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_AkcjeSpec")

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
                                    dbDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_AkcjeSpec")
                                    dbDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                                    dbConnectionAkcjeSpec.Close()
                                Else
                                    sqlDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_AkcjeSpec")
                                    sqlDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                                    sqlConnectionAkcjeSpec.Close()
                                End If

                                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                                DataSetAkcjeSpec.Tables(0).PrimaryKey = New DataColumn() {DataSetAkcjeSpec.Tables(0).Columns("IDENTAKCJI"), _
                                                                                          DataSetAkcjeSpec.Tables(0).Columns("IDENTKLI")}

                                'definiuj DataView
                                DataViewAkcjeSpec = New DataView(DataSetAkcjeSpec.Tables(0))
                                DataViewAkcjeSpec.AllowDelete = True
                                DataViewAkcjeSpec.AllowEdit = True
                                DataViewAkcjeSpec.AllowNew = True

                            Catch Ex As System.Exception
                                ViewInLog(Ex, Me.Name(), Nothing)
                                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                                '    System.Windows.Forms.MessageBoxButtons.OK, _
                                '    MessageBoxIcon.Information, _
                                '    MessageBoxDefaultButton.Button1)
                            End Try
                        End If
                        ''...........................
                        'definiuj delegat
                        Dim deleg As delegateAktualizuj = AddressOf AktualizujBazeAkcjeSpec

                        Dim FormRaport As New RaportNowaAkcjaMarketingowa(txtIdent.Text, _
                                                                             _dvKlenci, _
                                                                             DataSetAkcjeSpec, _
                                                                             deleg)
                        FormRaport.ShowDialog()
                        '...........................
                        MessageBox.Show("Doda³em wybranych klientów do akcji marketingowej" & vbCrLf & txtIdent.Text, _
                            "OK - skoñczy³em", _
                            System.Windows.Forms.MessageBoxButtons.OK, _
                            MessageBoxIcon.Information, _
                            MessageBoxDefaultButton.Button1)

                    Else
                    End If
                End If
            End If
        End If
    End Sub




    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujBazeAkcjeSpec()
        If ParL_DataBaseType = "MS ACCESS" Then
            Try
                dbConnectionAkcjeSpec.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If dbConnectionAkcjeSpec.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    dbDataAdapterAkcjeSpec.Update(DataSetAkcjeSpec, "TABELA_AkcjeSpec")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    dbDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                End If
                dbConnectionAkcjeSpec.Close()
            End If
        Else
            Try
                sqlConnectionAkcjeSpec.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionAkcjeSpec.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterAkcjeSpec.Update(DataSetAkcjeSpec, "TABELA_AkcjeSpec")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                End If
                sqlConnectionAkcjeSpec.Close()
            End If
        End If
    End Sub


End Class