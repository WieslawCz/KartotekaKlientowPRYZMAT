Module modWyliczenia


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

    Dim DataSetKlienci As DataSet
    Dim DataViewKlienci As DataView
    '------------------------------------------------------------------
    Dim ConnectionState As System.Data.ConnectionState


    '-----------------------------------
    ' wylicza wartosc graniczna poprzedniego roku
    '-----------------------------------

    Public Function WyliczWartoscGraniczna(ByVal pRok As String,
                                           ByVal pProcentG As Double) As Double
        Dim CalkSprzedaz As Double = 0
        Dim GranSprzedaz As Double = 0
        Dim WartGraniczna As Double = 0

        ''klienci o sprzedaży poniżej 1000 zł w roku
        'dbSelectSQLKlienci = "SELECT * FROM " & _
        '                     "(SELECT IDENTKLIENTA," & _
        '                            "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pRok & "')),0) + " & _
        '                            " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pRok & "')),0)) AS SPRZEDAZ " & _
        '                    " FROM Klienci) AS KL " & _
        '                    " WHERE (KL.SPRZEDAZ>0) AND (KL.SPRZEDAZ<=1000) ORDER BY KL.SPRZEDAZ ASC"


        dbSelectSQLKlienci = "SELECT * FROM " &
                             "(SELECT IDENTKLIENTA," &
                                    "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED+" &
                                                        "NAJEMWARTSPRZED+STRONYWARTSPRZED+DRUKARKIWARTSPRZED+SKUPWARTSPRZED) AS SPRZEDO " &
                                    " FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) And (SUBSTRING(Obroty.DATA,1,4)='" & pRok & "')),0) + " &
                                    " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED+" &
                                                        "NAJEMWARTSPRZED+STRONYWARTSPRZED+DRUKARKIWARTSPRZED+SKUPWARTSPRZED) AS SPRZEDM " &
                                    " FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) And (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pRok & "')),0)) AS SPRZEDAZ " &
                            " FROM Klienci) AS KL " &
                            " ORDER BY KL.SPRZEDAZ DESC"

        DataSetKlienci = New DataSet
        DataViewKlienci = New DataView

        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionKlienci = New OleDb.OleDbConnection(sConnectionKlienci)
            dbCommandSelectKlienci = New OleDb.OleDbCommand(dbSelectSQLKlienci, dbConnectionKlienci)
            'dbCommandDeleteKlienci = New OleDb.OleDbCommand(sDeleteSQLKlienci, dbConnectionKlienci)
            'dbCommandUpdateKlienci = New OleDb.OleDbCommand(sUpdateSQLKlienci, dbConnectionKlienci)
            'dbCommandInsertKlienci = New OleDb.OleDbCommand(sInsertSQLKlienci, dbConnectionKlienci)
            dbDataAdapterKlienci = New OleDb.OleDbDataAdapter(dbCommandSelectKlienci)

            Dim DBMapowanieTabeliKlienci As System.Data.Common.DataTableMapping
            DBMapowanieTabeliKlienci = dbDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionKlienci.Open()
            Catch Ex As System.Exception
                'ViewInLog(Ex, Me.Name(), Nothing)
                MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & Ex.Message, "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = dbConnectionKlienci.State
            End Try
        Else
            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
            sqlCommandSelectKlienci = New SqlClient.SqlCommand(dbSelectSQLKlienci, sqlConnectionKlienci)
            'sqlCommandDeleteKlienci = New sqlclient.sqlCommand(sDeleteSQLKlienci, sqlconnectionKlienci)
            'sqlCommandUpdateKlienci = New sqlclient.sqlCommand(sUpdateSQLKlienci, sqlconnectionKlienci)
            'sqlCommandInsertKlienci = New sqlclient.sqlCommand(sInsertSQLKlienci, sqlconnectionKlienci)
            sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

            Dim sqlMapowanieTabeliKlienci As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliKlienci = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienci.Open()
            Catch Ex As System.Exception
                'ViewInLog(Ex, Me.Name(), Nothing)
                MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & Ex.Message, "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
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
                DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("NRKARTY")}
                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))

                Dim i As Integer = 0
                If DataViewKlienci.Count > 0 Then
                    'liczymy całkowitą sprzedaz w tej grupie
                    CalkSprzedaz = 0
                    For i = 0 To DataViewKlienci.Count - 1
                        CalkSprzedaz += DataViewKlienci.Item(i).Item("SPRZEDAZ")
                    Next

                    'liczymy granice wartosci pProcentG % sprzedazy w tej grupie
                    GranSprzedaz = Math.Round((pProcentG / 100) * CalkSprzedaz, 2)

                    'liczymy wartosc graniczną - sprzedaż klienta na którym przekropczono pProcentG calkowitej sprzedaży
                    CalkSprzedaz = 0
                    For i = 0 To DataViewKlienci.Count - 1
                        CalkSprzedaz += DataViewKlienci.Item(i).Item("SPRZEDAZ")
                        WartGraniczna = DataViewKlienci.Item(i).Item("SPRZEDAZ")
                        If Math.Round(CalkSprzedaz, 2) >= GranSprzedaz Then
                            Exit For
                        Else
                        End If
                    Next

                    ''liczymy wartosc graniczną - sprzedaż ost klienta na którym nie przekroczono pProcentG % - nast klient przekracza...
                    'CalkSprzedaz = 0
                    'For i = 0 To DataViewKlienci.Count - 1
                    '    CalkSprzedaz += DataViewKlienci.Item(i).Item("SPRZEDAZ")
                    '    If CalkSprzedaz >= GranSprzedaz Then
                    '        'ne doliczamy więc - poprzedni klient
                    '        WartGraniczna = DataViewKlienci.Item(i).Item("SPRZEDAZ")
                    '        Exit For
                    '    Else
                    '    End If
                    'Next





                End If

            Catch Ex As System.Exception
                MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & Ex.Message, "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)

            End Try
        End If

        Return WartGraniczna
    End Function



End Module
