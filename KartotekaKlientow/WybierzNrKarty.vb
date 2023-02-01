Public Class WybierzNrKarty

    Dim RowHeightSetter As Softart.myDataGridRowHeightSetter
    Dim pointInCell00 As Point
    Dim StartFormy As Boolean = True

    Dim dbSelectNrKarty As String = "SELECT IDENTKLIENTA FROM Klienci"

    Dim dbConnection As OleDb.OleDbConnection
    Dim dbCommandSelect As OleDb.OleDbCommand
    Dim dbDataAdapter As OleDb.OleDbDataAdapter

    Dim sqlConnection As SqlClient.SqlConnection
    Dim sqlCommandSelect As SqlClient.SqlCommand
    Dim sqlDataAdapter As SqlClient.SqlDataAdapter

    Dim DataSetNrKarty As New DataSet
    Dim DataViewNrKarty As New DataView
    Dim ConnectionState As System.Data.ConnectionState


    Private Sub WybierzNrKarty_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        DataSetNrKarty.Locale = New System.Globalization.CultureInfo("pl-PL")
        dbSelectNrKarty = "SELECT IDENTKLIENTA FROM Klienci ORDER BY IDENTKLIENTA"

        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnection = New OleDb.OleDbConnection(sConnectionKlienci)
            dbCommandSelect = New OleDb.OleDbCommand(dbSelectNrKarty, dbConnection)
            dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim DBMapowanieTabeli As System.Data.Common.DataTableMapping
            DBMapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_dbNrKarty")
            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnection.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = dbConnection.State
            End Try
        Else
            sqlConnection = New SqlClient.SqlConnection(sConnectionKlienci)
            sqlCommandSelect = New SqlClient.SqlCommand(dbSelectNrKarty, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeli As System.Data.Common.DataTableMapping
            sqlMapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_sqlNrKarty")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnection.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnection.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    dbDataAdapter.FillSchema(DataSetNrKarty, SchemaType.Mapped, "TABELA_dbNrKarty")
                    dbDataAdapter.Fill(DataSetNrKarty)
                    dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(DataSetNrKarty, SchemaType.Mapped, "TABELA_sqlNrKarty")
                    sqlDataAdapter.Fill(DataSetNrKarty)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                'DataSetNrKarty.Tables(0).PrimaryKey = New DataColumn() {DataSetNrKarty.Tables(0).Columns("NRKARTY")}

                'definiuj DataView
                DataViewNrKarty = New DataView(DataSetNrKarty.Tables(0))
                DataViewNrKarty.AllowDelete = False
                DataViewNrKarty.AllowEdit = False
                DataViewNrKarty.AllowNew = False

                Dim dr As DataRowView
                Dim Nr As Long = 0
                Dim KolejnyNr As Long = 0
                Dim ii As Long = 0
                For Each dr In DataViewNrKarty
                    Nr = CInt(dr.Item("IDENTKLIENTA"))
                    If Nr = KolejnyNr + 1 Then
                    Else
                        'luka w numeracji od KolejnyNr+1 do Nr-1
                        For ii = KolejnyNr + 1 To Nr - 1
                            cbxNrKarty.Items.Add(Microsoft.VisualBasic.Right("000000" + Trim(Str(ii)), 6))
                        Next
                    End If
                    KolejnyNr = Nr
                Next
                If cbxNrKarty.Items.Count > 0 Then
                    cbxNrKarty.SelectedItem = 0
                End If

            Catch Ex As System.Exception
                'ViewInLog(ex, Me.Name(), Nothing)
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            End Try
        End If
    End Sub

    Private Sub WybierzNrKarty_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdWybierz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWybierz.Click
        If cbxNrKarty.Items.Count > 0 Then
            oNumer = cbxNrKarty.Text
        End If
        Me.Close()
    End Sub

End Class