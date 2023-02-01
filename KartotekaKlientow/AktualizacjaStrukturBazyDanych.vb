Public Class AktualizacjaStrukturBazyDanych

    Dim sConnectionStr As String = ""       'ConnectionStrings()

    Dim dbConnection As OleDb.OleDbConnection
    Dim dbCommandSelect As OleDb.OleDbCommand
    Dim dbCommandDelete As OleDb.OleDbCommand
    Dim dbCommandUpdate As OleDb.OleDbCommand
    Dim dbCommandInsert As OleDb.OleDbCommand
    Dim dbDataAdapter As OleDb.OleDbDataAdapter
    Dim dbCommand As OleDb.OleDbCommand = New OleDb.OleDbCommand

    Dim sqlConnectionKlienci As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienci As SqlClient.SqlCommand
    Dim sqlCommandDeleteKlienci As SqlClient.SqlCommand
    Dim sqlCommandUpdateKlienci As SqlClient.SqlCommand
    Dim sqlCommandInsertKlienci As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienci As SqlClient.SqlDataAdapter
    Dim sqlCommand As SqlClient.SqlCommand = New SqlClient.SqlCommand

    Dim ConnectionState As System.Data.ConnectionState
    '---------------------------
    Dim dbConnectionKontakty As OleDb.OleDbConnection
    Dim dbCommandSelectKontakty As OleDb.OleDbCommand
    Dim dbCommandDeleteKontakty As OleDb.OleDbCommand
    Dim dbCommandUpdateKontakty As OleDb.OleDbCommand
    Dim dbCommandInsertKontakty As OleDb.OleDbCommand
    Dim dbDataAdapterKontakty As OleDb.OleDbDataAdapter
    Dim dbCommandKontakty As OleDb.OleDbCommand = New OleDb.OleDbCommand

    Dim sqlConnectionKlienciKontakty As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienciKontakty As SqlClient.SqlCommand
    Dim sqlCommandDeleteKlienciKontakty As SqlClient.SqlCommand
    Dim sqlCommandUpdateKlienciKontakty As SqlClient.SqlCommand
    Dim sqlCommandInsertKlienciKontakty As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienciKontakty As SqlClient.SqlDataAdapter
    Dim sqlCommandKontakty As SqlClient.SqlCommand = New SqlClient.SqlCommand

    Dim DataSetKontakty As New DataSet
    Dim DataViewKontakty As New DataView
    '------------------------
    Dim dbConnectionKontaktyOld As OleDb.OleDbConnection
    Dim dbCommandSelectKontaktyOld As OleDb.OleDbCommand
    Dim dbCommandDeleteKontaktyOld As OleDb.OleDbCommand
    Dim dbCommandUpdateKontaktyOld As OleDb.OleDbCommand
    Dim dbCommandInsertKontaktyOld As OleDb.OleDbCommand
    Dim dbDataAdapterKontaktyOld As OleDb.OleDbDataAdapter
    Dim dbCommandKontaktyOld As OleDb.OleDbCommand = New OleDb.OleDbCommand

    Dim sqlConnectionKlienciKontaktyOld As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienciKontaktyOld As SqlClient.SqlCommand
    Dim sqlCommandDeleteKlienciKontaktyOld As SqlClient.SqlCommand
    Dim sqlCommandUpdateKlienciKontaktyOld As SqlClient.SqlCommand
    Dim sqlCommandInsertKlienciKontaktyOld As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienciKontaktyOld As SqlClient.SqlDataAdapter
    Dim sqlCommandKontaktyOld As SqlClient.SqlCommand = New SqlClient.SqlCommand

    Dim DataSetKontaktyOld As New DataSet
    Dim DataViewKontaktyOld As New DataView
    '------------------------
    Dim dbConnectionKlienci As OleDb.OleDbConnection
    Dim dbCommandSelectKlienci As OleDb.OleDbCommand
    Dim dbCommandDeleteKlienci As OleDb.OleDbCommand
    Dim dbCommandUpdateKlienci As OleDb.OleDbCommand
    Dim dbCommandInsertKlienci As OleDb.OleDbCommand
    Dim dbDataAdapterKlienci As OleDb.OleDbDataAdapter
    Dim dbCommandKlienci As OleDb.OleDbCommand = New OleDb.OleDbCommand

    Dim sqlConnectionKlienciKlienci As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienciKlienci As SqlClient.SqlCommand
    Dim sqlCommandDeleteKlienciKlienci As SqlClient.SqlCommand
    Dim sqlCommandUpdateKlienciKlienci As SqlClient.SqlCommand
    Dim sqlCommandInsertKlienciKlienci As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienciKlienci As SqlClient.SqlDataAdapter
    Dim sqlCommandKlienci As SqlClient.SqlCommand = New SqlClient.SqlCommand

    Dim DataSetKlienci As New DataSet
    Dim DataViewKlienci As New DataView
    '---------------------------
    Dim dbConnectionCoDalej As OleDb.OleDbConnection
    Dim dbCommandSelectCoDalej As OleDb.OleDbCommand
    Dim dbCommandDeleteCoDalej As OleDb.OleDbCommand
    Dim dbCommandUpdateCoDalej As OleDb.OleDbCommand
    Dim dbCommandInsertCoDalej As OleDb.OleDbCommand
    Dim dbDataAdapterCoDalej As OleDb.OleDbDataAdapter
    Dim dbCommandCoDalej As OleDb.OleDbCommand = New OleDb.OleDbCommand

    Dim sqlConnectionKlienciCoDalej As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienciCoDalej As SqlClient.SqlCommand
    Dim sqlCommandDeleteKlienciCoDalej As SqlClient.SqlCommand
    Dim sqlCommandUpdateKlienciCoDalej As SqlClient.SqlCommand
    Dim sqlCommandInsertKlienciCoDalej As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienciCoDalej As SqlClient.SqlDataAdapter
    Dim sqlCommandCoDalej As SqlClient.SqlCommand = New SqlClient.SqlCommand

    Dim DataSetCoDalej As New DataSet
    Dim DataViewCoDalej As New DataView
    '------------------------

    Dim NowaWersja As Integer = oAktualnaWersjaBazyDanych
    Dim AktualnaWersja As Integer = oWersjaBazyDanych




    Private Sub AktualizacjaStrukturBazyDanych_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        If _CoAktualizujemy = "HQ" Then
            AktualnaWersja = HQWersjaBazyDanych()
            NowaWersja = oAktualnaWersjaBazyDanych
            txtRaport.Text = "Instalacja : " & ParL_OpisInstalacji & vbCrLf
            txtRaport.Text &= "Rodzaj Bazy Danych : " & ParL_DataBaseType & vbCrLf
            txtRaport.Text &= "Nazwa Bazy Danych Importowanych : " & ParL_HQDataBase & vbCrLf
            txtRaport.Text &= "Aktualna wersja Bazy Danych Importowanych : " & (AktualnaWersja / 100).ToString("N2") & vbCrLf
            txtRaport.Text &= "Mamy aktualizowaæ Bazê Danych Importowanych do wersji " & (NowaWersja / 100).ToString("N2") & vbCrLf
            txtRaport.SelectionStart = txtRaport.Text.Length
            txtRaport.ScrollToCaret()
            System.Windows.Forms.Application.DoEvents()
            sConnectionStr = HQConnectionStrings()
        Else
            AktualnaWersja = WersjaBazyDanych()
            NowaWersja = oAktualnaWersjaBazyDanych
            txtRaport.Text = "Instalacja : " & ParL_OpisInstalacji & vbCrLf
            txtRaport.Text &= "Rpdzaj Bazy Danych : " & ParL_DataBaseType & vbCrLf
            txtRaport.Text &= "Nazwa Bazy Danych : " & ParL_DataBase & vbCrLf
            txtRaport.Text &= "Aktualna wersja Bazy Danych : " & (AktualnaWersja / 100).ToString("N2") & vbCrLf
            txtRaport.Text &= "Mamy aktualizowaæ Bazê Danych do wersji " & (NowaWersja / 100).ToString("N2") & vbCrLf
            txtRaport.SelectionStart = txtRaport.Text.Length
            txtRaport.ScrollToCaret()
            System.Windows.Forms.Application.DoEvents()
            sConnectionStr = ConnectionStrings()
        End If
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        System.Windows.Forms.Application.DoEvents()
        Me.Close()
    End Sub

    Private Sub AktualizacjaStrukturBazyDanych_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        ' gradient w tle formy
        Dim G As Graphics
        G = Me.CreateGraphics
        Dim R As New RectangleF(0, 0, Me.Width, Me.Height)
        Dim startColor As Color = System.Drawing.SystemColors.Control
        Dim EndColor As Color = System.Drawing.SystemColors.ControlDark

        Dim LGBrush As New System.Drawing.Drawing2D.LinearGradientBrush _
                (R, startColor, EndColor, System.Drawing.Drawing2D.LinearGradientMode.Horizontal)
        G.FillRectangle(LGBrush, New Rectangle(0, 0, Me.Width, Me.Height))
    End Sub

    Private Sub cmdAktualizuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAktualizuj.Click
        If ParL_DataBaseType = "MS ACCESS" Then
            'AktualizujBazeACCESS()
        Else
            AktualizujBazeSQL()
        End If
    End Sub






















    Private Sub AktDB(ByVal Komunikat As String, ByVal Kwerenda As String)
        Dim Wynik As Integer = 0

        txtRaport.Text &= "Wykonujê aktualizacje tabeli " & Komunikat & vbCrLf
        txtRaport.SelectionStart = txtRaport.Text.Length
        txtRaport.ScrollToCaret()
        System.Windows.Forms.Application.DoEvents()

        dbCommand.Connection = dbConnection
        dbCommand.CommandType = CommandType.Text
        dbCommand.CommandText = Kwerenda
        dbConnection.Open()
        Try
            Wynik = dbCommand.ExecuteNonQuery()
            If Wynik = -1 Then
            End If
            'txtRaport.Text &= "OK, aktualizacja wykonana poprawnie" & vbCrLf
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
            txtRaport.Text &= "B³¹d wykonania : " & ex.Message & vbCrLf
            txtRaport.SelectionStart = txtRaport.Text.Length
            txtRaport.ScrollToCaret()
            System.Windows.Forms.Application.DoEvents()
        Finally
            dbConnection.Close()
        End Try
    End Sub





    '*************************************************
    ' Aktualizuja baze ACCESS
    '*************************************************

    Private Sub AktualizujBazeACCESS()
        'Dim Upgrade As Boolean = False
        'Dim Wynik As Integer = 0

        'If AktualnaWersja = NowaWersja Then
        '    If MessageBox.Show("Baza danych jest juz zaktualizowana do wersji " & (NowaWersja / 100).ToString("N2") & vbCrLf & _
        '        "Nie ma potrzeby wykonywania aktualizacji." & _
        '        "Czy chcesz wykonaæ aktualizacje mimo wszystko ?", _
        '        "Uwaga :", _
        '        System.Windows.Forms.MessageBoxButtons.YesNo, _
        '        MessageBoxIcon.Information, _
        '        MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then
        '        Upgrade = True
        '    Else
        '        Upgrade = False
        '    End If
        'ElseIf AktualnaWersja > NowaWersja Then
        '    If MessageBox.Show("Baza danych jest juz zaktualizowana do wersji " & (AktualnaWersja / 100).ToString("N2") & vbCrLf & _
        '        "Nie mogê wykonaæ aktualizacji wstecznej do wersji " & (NowaWersja / 100).ToString("N2"), _
        '        "Uwaga :", _
        '        System.Windows.Forms.MessageBoxButtons.OK, _
        '        MessageBoxIcon.Information, _
        '        MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.OK Then
        '        Upgrade = False
        '    End If
        'Else
        '    Upgrade = True
        'End If

        'If Upgrade Then
        '    txtRaport.Text &= vbCrLf & Now & "  Rozpoczynam aktualizacje" & vbCrLf
        '    txtRaport.SelectionStart = txtRaport.Text.Length
        '    txtRaport.ScrollToCaret()
        '    System.Windows.Forms.Application.DoEvents()

        '    dbConnection = New OleDb.OleDbConnection(ConnStringAccess())
        '    dbCommand = New OleDb.OleDbCommand
        '    '--------------------------------------
        '    'If AktualnaWersja <= 100 Then
        '    '    'aktualizacje do wersji 1.00
        '    '    AktDB("Wykonujê aktualizacje tabeli Wersja", oCreateWersja100_1)
        '    'End If
        '    '--------------------------------------
        '    If AktualnaWersja <= 210 Then
        '        'wersja 2.10
        '        'Public dbAlterKlienci210_1 As String = "ALTER TABLE Klienci ADD MARKER BIT;"

        '        'Baza danych  SQL istnieje od wersji programu 2.00
        '        'aktualizacje do wersji 2.10
        '        '--------------------------------------
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD MARKER BIT;")
        '        '--------------------------------------
        '    End If


        '    '--------------------------------------
        '    If AktualnaWersja <= 220 Then
        '        'aktualizacje do wersji 2.20
        '        'Max dlugosc pola CHAR w ACCES=255 znakow



        '        '--------------------------------------
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDPLAN CHAR(250);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN ILOSPRZETU CHAR(250);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN NRTOFI INTEGER;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD NRTOFITXT VARCHAR(100);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD IDDOMU VARCHAR(10);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD NUMNRDOMU INTEGER;")
        '        '--------------------------------------
        '        '--------------------------------------
        '        txtRaport.Text &= "Przepisujê zmienione pola w tabeli Klienci" & vbCrLf
        '        txtRaport.SelectionStart = txtRaport.Text.Length
        '        txtRaport.ScrollToCaret()
        '        System.Windows.Forms.Application.DoEvents()

        '        Dim dbSelectKli220 As String = "SELECT " & _
        '                                               "IDENTKLIENTA, " & _
        '                                               "NRTOFI, " & _
        '                                               "NRTOFITXT, " & _
        '                                               "ILOSPRZETU, " & _
        '                                               "ILOSPRZETU2, " & _
        '                                               "NUMNRDOMU, " & _
        '                                               "NRDOMU, " & _
        '                                               "IDDOMU, " & _
        '                                               "WERSJA " & _
        '                                            "FROM Klienci"

        '        Dim dbUpdateKli220 As String = "UPDATE Klienci SET " & _
        '                                                     "NRTOFI=@oNrTofiKli, " & _
        '                                                     "NRTOFITXT=@oNrTofiTxtKli, " & _
        '                                                     "ILOSPRZETU=@oIloSprzetuKli, " & _
        '                                                     "ILOSPRZETU2=@oIloSprzetu2Kli, " & _
        '                                                     "NUMNRDOMU=@oNumNrDomuKli, " & _
        '                                                     "NRDOMU=@oNrDomuKli, " & _
        '                                                     "IDDOMU=@oIDDomuKli, " & _
        '                                                     "WERSJA=@oWersjaKli " & _
        '                                            "WHERE (IDENTKLIENTA=@orygSymbol) AND " & _
        '                                                  "(WERSJA=@orygWersja)"

        '        dbCommandSelect = New OleDb.OleDbCommand(dbSelectKli220, dbConnection)
        '        dbCommandUpdate = New OleDb.OleDbCommand(dbUpdateKli220, dbConnection)
        '        dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

        '        'mapujemy domyslna nazwe tabeli
        '        Dim MapowanieTabeli As System.Data.Common.DataTableMapping
        '        MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Klienci")

        '        Dim objParam As OleDb.OleDbParameter

        '        '------------------------------------------
        '        'komenda UPDATE
        '        'parametry aktualizacji
        '        'dbcommandUpdate.Parameters.Add("@oIdentKli", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
        '        dbCommandUpdate.Parameters.Add("@oNrTofiKli", OleDb.OleDbType.Integer, Nothing, "NRTOFI")
        '        dbCommandUpdate.Parameters.Add("@oNrTofiTxtKli", OleDb.OleDbType.Char, 100, "NRTOFITXT")
        '        dbCommandUpdate.Parameters.Add("@oIloSprzetuKli", OleDb.OleDbType.Char, 300, "ILOSPRZETU")
        '        dbCommandUpdate.Parameters.Add("@oIloSprzetu2Kli", OleDb.OleDbType.WChar, Nothing, "ILOSPRZETU2")
        '        dbCommandUpdate.Parameters.Add("@oNumNrDomuKli", OleDb.OleDbType.Integer, Nothing, "NUMNRDOMU")
        '        dbCommandUpdate.Parameters.Add("@oNrDomuKli", OleDb.OleDbType.Char, 10, "NRDOMU")
        '        dbCommandUpdate.Parameters.Add("@oIDDomuKli", OleDb.OleDbType.Char, 10, "IDDOMU")
        '        dbCommandUpdate.Parameters.Add("@oWersjaKli", OleDb.OleDbType.Integer, Nothing, "WERSJA")

        '        'parametry filtrowania
        '        objParam = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 10, "IDENTKLIENTA")
        '        objParam.SourceVersion = DataRowVersion.Original
        '        objParam = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
        '        objParam.SourceVersion = DataRowVersion.Original
        '        dbDataAdapter.UpdateCommand = dbCommandUpdate

        '        ' Add the RowUpdated event handler.
        '        AddHandler dbDataAdapter.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)
        '        '------------------------------------------
        '        'wypelnij DATASET
        '        Try
        '            dbConnection.Open()
        '        Catch Ex As System.Exception
        '            ViewInLog(Ex, Me.Name(), Nothing)
        '            'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
        '            '    System.Windows.Forms.MessageBoxButtons.OK, _
        '            '    MessageBoxIcon.Information, _
        '            '    MessageBoxDefaultButton.Button1)
        '        Finally
        '            ConnectionState = dbConnection.State
        '        End Try
        '        If ConnectionState = ConnectionState.Open Then
        '            Dim IloSp As String = ""
        '            Dim NTofi As Integer = 0
        '            Dim PopTofiTxt As String = ""
        '            Dim NTofiTxt As String = ""
        '            Dim Wers As Integer = 0
        '            Dim i As Integer = 0
        '            Dim dr As DataRow

        '            Dim NumNrD As Integer = 0
        '            Dim NrD As String = ""
        '            Dim IdD As String = ""
        '            Dim j As Integer = 0
        '            Dim ch As Char = ""

        '            Try
        '                dbDataAdapter.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
        '                dbDataAdapter.Fill(DataSetKlienci)
        '                dbConnection.Close()

        '                'zdefiniuj unikalny klucz wyszukiwania w tabeli
        '                'DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("NRKARTY")}
        '                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))


        '                If DataSetKlienci.Tables(0).Rows.Count > 0 Then
        '                    For i = 0 To DataSetKlienci.Tables(0).Rows.Count - 1
        '                        dr = DataSetKlienci.Tables(0).Rows.Item(i)
        '                        IloSp = IIf(IsDBNull(dr.Item("ILOSPRZETU")), "", dr.Item("ILOSPRZETU"))
        '                        NTofi = IIf(IsDBNull(dr.Item("NRTOFI")), 0, dr.Item("NRTOFI"))
        '                        NTofiTxt = IIf(IsDBNull(dr.Item("NRTOFITXT")), "", dr.Item("NRTOFITXT"))
        '                        NumNrD = IIf(IsDBNull(dr.Item("NUMNRDOMU")), 0, dr.Item("NUMNRDOMU"))
        '                        NrD = IIf(IsDBNull(dr.Item("NRDOMU")), "", dr.Item("NRDOMU"))
        '                        IdD = IIf(IsDBNull(dr.Item("IDDOMU")), "", dr.Item("IDDOMU"))
        '                        Wers = dr.Item("WERSJA")

        '                        'analizuj nr domu
        '                        NrD = Trim(NrD)
        '                        IdD = Trim(IdD)
        '                        If Len(NrD) > 0 And NumNrD = 0 Then
        '                            For j = 1 To Len(NrD)
        '                                ch = Mid(NrD, j, 1)
        '                                If InStr("0123456789", ch) = 0 Then
        '                                    If Len(IdD) > 0 Then
        '                                        IdD = Trim(Mid(NrD, j)) & " " & IdD
        '                                    Else
        '                                        IdD = Trim(Mid(NrD, j))
        '                                    End If
        '                                    NrD = Mid(NrD, 1, j - 1)
        '                                    Exit For
        '                                End If
        '                            Next
        '                            If Len(NrD) > 0 Then
        '                                NumNrD = CInt(NrD)
        '                            Else
        '                                NumNrD = 0
        '                            End If
        '                        End If

        '                        'analizuj numery TOFI -  maj¹ byc 6 znakowe z zerami na poczatku
        '                        'dopisujemy z pola NRTOFI(integer) do NRTOFITXT(txt)
        '                        If CInt(NTofi) = 0 Then
        '                            dr.Item("NRTOFITXT") = ""
        '                        Else
        '                            If Len(NTofiTxt) > 0 Then
        '                                If InStr(NTofiTxt, Microsoft.VisualBasic.Right("00000" + Trim(Str(NTofi)), 5)) = 0 Then
        '                                    PopTofiTxt = dr.Item("NRTOFITXT")
        '                                    dr.Item("NRTOFITXT") = Trim(PopTofiTxt) & "," & Microsoft.VisualBasic.Right("00000" + Trim(Str(NTofi)), 5)
        '                                Else
        '                                    'juz jest dopisane
        '                                End If
        '                            Else
        '                                dr.Item("NRTOFITXT") = Microsoft.VisualBasic.Right("00000" + Trim(Str(NTofi)), 5)
        '                            End If
        '                        End If

        '                        dr.Item("ILOSPRZETU2") = Trim(IloSp)
        '                        dr.Item("NUMNRDOMU") = NumNrD
        '                        dr.Item("NRDOMU") = NrD
        '                        dr.Item("IDDOMU") = IdD
        '                        dr.Item("WERSJA") = ZmienWersje(Wers)
        '                    Next
        '                    AktualizujBazeKlienci()
        '                End If



        '            Catch Ex As System.Exception
        '                'ViewInLog(ex, Me.Name(), Nothing)
        '                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :", _
        '                '    System.Windows.Forms.MessageBoxButtons.OK, _
        '                '    MessageBoxIcon.Information, _
        '                '    MessageBoxDefaultButton.Button1)
        '                MessageBox.Show(Str(NTofi) & vbCrLf & NTofiTxt)
        '                txtRaport.Text &= "B³¹d wykonania : " & Ex.Message & vbCrLf
        '                txtRaport.SelectionStart = txtRaport.Text.Length
        '                txtRaport.ScrollToCaret()
        '                System.Windows.Forms.Application.DoEvents()
        '            End Try
        '        End If
        '        '--------------------------------------

        '        AktDB("Obroty", "ALTER TABLE Obroty ADD LOWARTSPRZED double;")
        '        AktDB("Obroty", "ALTER TABLE Obroty ADD LOILOSPRZED double;")
        '        AktDB("Obroty", "ALTER TABLE Obroty ADD AOWARTSPRZED double;")
        '        AktDB("Obroty", "ALTER TABLE Obroty ADD AOILOSPRZED double;")

        '        AktDB("ObrotyMies", "ALTER TABLE ObrotyMies ADD LOWARTSPRZED double;")
        '        AktDB("ObrotyMies", "ALTER TABLE ObrotyMies ADD LOILOSPRZED double;")
        '        AktDB("ObrotyMies", "ALTER TABLE ObrotyMies ADD AOWARTSPRZED double;")
        '        AktDB("ObrotyMies", "ALTER TABLE ObrotyMies ADD AOILOSPRZED double;")


        '        AktDB("AnalizyZakupu", "CREATE TABLE AnalizyZakupu " & _
        '                                            "(IDENTKLIENTA CHAR(6) NOT NULL, " & _
        '                                                "WARTZA00 double NOT NULL, " & _
        '                                                "ILOSZA00 double NOT NULL, " & _
        '                                                "WARTZA01 double NOT NULL, " & _
        '                                                "ILOSZA01 double NOT NULL, " & _
        '                                                "WARTZA02 double NOT NULL, " & _
        '                                                "ILOSZA02 double NOT NULL, " & _
        '                                                "WARTZA03 double NOT NULL, " & _
        '                                                "ILOSZA03 double NOT NULL, " & _
        '                                                "WARTZA04 double NOT NULL, " & _
        '                                                "ILOSZA04 double NOT NULL, " & _
        '                                                "WARTZA05 double NOT NULL, " & _
        '                                                "ILOSZA05 double NOT NULL, " & _
        '                                                "WARTZA06 double NOT NULL, " & _
        '                                                "ILOSZA06 double NOT NULL, " & _
        '                                                "WARTZA07 double NOT NULL, " & _
        '                                                "ILOSZA07 double NOT NULL, " & _
        '                                                "WARTZA08 double NOT NULL, " & _
        '                                                "ILOSZA08 double NOT NULL, " & _
        '                                                "WARTZA09 double NOT NULL, " & _
        '                                                "ILOSZA09 double NOT NULL, " & _
        '                                                "WARTZA10 double NOT NULL, " & _
        '                                                "ILOSZA10 double NOT NULL, " & _
        '                                                "WARTZA11 double NOT NULL, " & _
        '                                                "ILOSZA11 double NOT NULL, " & _
        '                                                "WARTZA12 double NOT NULL, " & _
        '                                                "ILOSZA12 double NOT NULL, " & _
        '                                                "WARTZA13 double NOT NULL, " & _
        '                                                "ILOSZA13 double NOT NULL, " & _
        '                                                "WARTZA14 double NOT NULL, " & _
        '                                                "ILOSZA14 double NOT NULL, " & _
        '                                                "WARTZA15 double NOT NULL, " & _
        '                                                "ILOSZA15 double NOT NULL, " & _
        '                                                "WARTZA16 double NOT NULL, " & _
        '                                                "ILOSZA16 double NOT NULL, " & _
        '                                                "WARTZA17 double NOT NULL, " & _
        '                                                "ILOSZA17 double NOT NULL, " & _
        '                                                "WARTZA18 double NOT NULL, " & _
        '                                                "ILOSZA18 double NOT NULL, " & _
        '                                                "WARTZA19 double NOT NULL, " & _
        '                                                "ILOSZA19 double NOT NULL, " & _
        '                                                "WARTZA20 double NOT NULL, " & _
        '                                                "ILOSZA20 double NOT NULL, " & _
        '                                                "WARTZA21 double NOT NULL, " & _
        '                                                "ILOSZA21 double NOT NULL, " & _
        '                                                "WARTZA22 double NOT NULL, " & _
        '                                                "ILOSZA22 double NOT NULL, " & _
        '                                                "WARTZA23 double NOT NULL, " & _
        '                                                "ILOSZA23 double NOT NULL, " & _
        '                                                "WARTZA24 double NOT NULL, " & _
        '                                                "ILOSZA24 double NOT NULL, " & _
        '                                            "WERSJA INTEGER NOT NULL);")
        '        AktDB("AnalizyZakupu", "ALTER TABLE AnalizyZakupu " & _
        '                                                    "ADD CONSTRAINT PK_AnalizyZakupu " & _
        '                                                    "PRIMARY KEY (IDENTKLIENTA);")


        '    End If
        '    '--------------------------------------





        '    If AktualnaWersja <= 230 Then
        '        'aktualizacje do wersji 2.30
        '        '--------------------------------------
        '        If AktualnaWersja < 230 Then
        '            AktDB("Parametry", "DROP TABLE Parametry;")
        '        End If
        '        AktDB("Parametry", "CREATE TABLE Parametry " & _
        '                                            "(IDENT CHAR(10) NOT NULL, " & _
        '                                             "IDENTUZYTKOWNIKA CHAR(50) NOT NULL, " & _
        '                                             "NAZWAUZYTKOWNIKA CHAR(50) NOT NULL, " & _
        '                                             "ADRESUZYTKOWNIKA CHAR(50) NOT NULL, " & _
        '                                             "KONTOUZYTKOWNIKA CHAR(50) NOT NULL, " & _
        '                                             "BANKUZYTKOWNIKA CHAR(50) NOT NULL, " & _
        '                                             "MIEJSCOWOSCUZYTKOWNIKA CHAR(50) NOT NULL, " & _
        '                                             "NIPUZYTKOWNIKA CHAR(50) NOT NULL, " & _
        '                                             "DATAAKTANALIZY CHAR(10) NOT NULL, " & _
        '                                             "OSTOKRESANALIZY CHAR(10) NOT NULL, " & _
        '                                             "WERSJA INTEGER NOT NULL);")
        '        AktDB("Parametry", "ALTER TABLE Parametry " & _
        '                                                    "ADD CONSTRAINT PK_Parametry " & _
        '                                                    "PRIMARY KEY (IDENT);")

        '        AktDB("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD WARTZABM double;")
        '        AktDB("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD ILOSZABM double;")
        '        '--------------------------------------
        '    End If



        '    If AktualnaWersja <= 240 Then
        '        'aktualizacje do wersji 2.40
        '        '--------------------------------------
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD KARTAPAYBACK BIT;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD NRYKARTPAYBACK NOTE;")

        '        AktDB("Parametry", "ALTER TABLE Parametry ADD DANEDOANALIZY VARCHAR(10);")
        '        '--------------------------------------
        '    End If



        '    If AktualnaWersja <= 250 Then
        '        'aktualizacje do wersji 2.50
        '        '--------------------------------------
        '        AktDB("RodzajeKontaktow", "CREATE TABLE RodzajeKontaktow " & _
        '                                     "(UNIQID VARCHAR(40) NOT NULL, " & _
        '                                      "RODZAJKONTAKTU VARCHAR(50) NOT NULL, " & _
        '                                      "WERSJA INTEGER NOT NULL);")
        '        AktDB("RodzajeKontaktow", "ALTER TABLE RodzajeKontaktow " & _
        '                                             "ADD CONSTRAINT PK_RodzajeKontaktow " & _
        '                                             "PRIMARY KEY (UNIQID);")

        '        AktDB("Parametry", "ALTER TABLE Parametry ADD KATALOGOFERTY VARCHAR(100);")

        '        AktDB("Klienci", "ALTER TABLE Klienci ADD OFERTATZLOZENIA VARCHAR(4);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD OFERTAPLIK VARCHAR(100);")
        '        '--------------------------------------
        '        InicjuPolaKK250()
        '    End If


        '    If AktualnaWersja <= 251 Then
        '        'wersja 2.51
        '        '--------------------------------------
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD MARKERP BIT;")
        '        '--------------------------------------
        '        InicjuPolaKK251()
        '    End If


        '    If AktualnaWersja <= 252 Then
        '        'wersja 2.52
        '        '---------------------------------------------------------------------
        '        'DaneDoRaportu
        '        '---------------------------------------------------------------------
        '        'Public ddrPracownik As String        'PRACOWNIK   Text, 10
        '        'Public ddrDataRaportu As String      'DATARAPORTU Text, 7   RRRR-MM
        '        'Public ddrKlBezDrukarki As Integer   'KLBEZDRUKARKI  Integer
        '        'Public ddrOferty As Integer          'OFERTY         Integer
        '        'Public ddrTestyWstawionw As Integer  'TESTYWSTAWIONE Integer
        '        'Public ddrCzyszczenie As Integer     'CZYSZCZENIE    Integer
        '        'Public ddrDostawy As Integer         'DOSTAWY        Integer
        '        'Public ddrSkup As Integer            'SKUP           Integer
        '        'Public ddrWersja As Integer          'WERSJA         Integer
        '        AktDB("RodzajeKontaktow", "CREATE TABLE DaneDoRaportu (" & vbCrLf & _
        '                   "   PRACOWNIK varchar(10) NOT NULL ," & vbCrLf & _
        '                   "   DATARAPORTU varchar(10) NOT NULL ," & vbCrLf & _
        '                   "   KLBEZDRUKARKI Integer NOT NULL ," & vbCrLf & _
        '                   "   OFERTY Integer NOT NULL ," & vbCrLf & _
        '                   "   TESTYWSTAWIONE Integer NOT NULL ," & vbCrLf & _
        '                   "   CZYSZCZENIE Integer NOT NULL ," & vbCrLf & _
        '                   "   DOSTAWY Integer NOT NULL ," & vbCrLf & _
        '                   "   SKUP Integer NOT NULL ," & vbCrLf & _
        '                   "   WERSJA Integer NOT NULL " & vbCrLf & _
        '                   ");")
        '        AktDB("RodzajeKontaktow", "ALTER TABLE DaneDoRaportu " & _
        '                                             "ADD CONSTRAINT PK_DaneDoRaportu " & _
        '                                             "PRIMARY KEY (PRACOWNIK,DATARAPORTU);")
        '        '--------------------------------------
        '    End If



        '    If AktualnaWersja <= 252 Then
        '        'wersja 2.52
        '        '---------------------------------------------------------------------
        '        'DaneDoRaportu
        '        '---------------------------------------------------------------------
        '        'Public ddrPracownik As String        'PRACOWNIK   Text, 10
        '        'Public ddrDataRaportu As String      'DATARAPORTU Text, 7   RRRR-MM
        '        'Public ddrKlBezDrukarki As Integer   'KLBEZDRUKARKI  Integer
        '        'Public ddrOferty As Integer          'OFERTY         Integer
        '        'Public ddrTestyWstawionw As Integer  'TESTYWSTAWIONE Integer
        '        'Public ddrCzyszczenie As Integer     'CZYSZCZENIE    Integer
        '        'Public ddrDostawy As Integer         'DOSTAWY        Integer
        '        'Public ddrSkup As Integer            'SKUP           Integer
        '        'Public ddrWersja As Integer          'WERSJA         Integer
        '        AktDB("RodzajeKontaktow", "CREATE TABLE DaneDoRaportu (" & vbCrLf & _
        '                   "   PRACOWNIK varchar(10) NOT NULL ," & vbCrLf & _
        '                   "   DATARAPORTU varchar(10) NOT NULL ," & vbCrLf & _
        '                   "   KLBEZDRUKARKI Integer NOT NULL ," & vbCrLf & _
        '                   "   OFERTY Integer NOT NULL ," & vbCrLf & _
        '                   "   TESTYWSTAWIONE Integer NOT NULL ," & vbCrLf & _
        '                   "   CZYSZCZENIE Integer NOT NULL ," & vbCrLf & _
        '                   "   DOSTAWY Integer NOT NULL ," & vbCrLf & _
        '                   "   SKUP Integer NOT NULL ," & vbCrLf & _
        '                   "   WERSJA Integer NOT NULL " & vbCrLf & _
        '                   ");")
        '        AktDB("RodzajeKontaktow", "ALTER TABLE DaneDoRaportu " & _
        '                                             "ADD CONSTRAINT PK_DaneDoRaportu " & _
        '                                             "PRIMARY KEY (PRACOWNIK,DATARAPORTU);")
        '        '--------------------------------------
        '    End If

        '    '--------------------------------------
        '    If AktualnaWersja <= 253 Then
        '        '--------------------------------------
        '        'Max dlugosc pola CHAR w ACCES=255 znakow
        '        'AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDPLAN CHAR(250);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDWAGA CHAR(20);")

        '        '---------------------------------------------------------------------
        '        'Akcje Handlowe - Opis
        '        'Public aoIdentAkcji As String            'IDENT     Text 20
        '        'Public aoDataAkcji As String             'DATA      Text,10
        '        'Public aoOpisAkcji As String             'OPIS      Text,100
        '        'Public aoUwagiAkcji As String            'UWAGI     Text,10
        '        'Public aoMarkerAkcji As String           'MARKER    logical
        '        'Public aoWersjaAkcji As Integer           'WERSJA    Integer
        '        AktDB("AkcjeOpis", "ALTER TABLE AkcjeOpis ADD MARKER BIT;")

        '        AktDB("Parametry", "ALTER TABLE Parametry ALTER COLUMN OSTOKRESANALIZY CHAR(14);")
        '    End If


        '    '--------------------------------------
        '    If AktualnaWersja <= 260 Then
        '        'Public oZakupPopRokKli As Double       'ZAKUPPOPROK        Double
        '        'Public oUslugiDodKli As String         'USLUGIDOD          Text, 200
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD ZAKUPPOPROK double;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD USLUGIDOD VARCHAR(200);")



        '        'Public Par_wwwPAYBACK As String = ""
        '        'Public Par_MiesAnalizy1 As String = ""
        '        'Public Par_MiesAnalizy2 As String = ""
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD WWWPAYBACK VARCHAR(200);")
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD MIESANALIZY1 VARCHAR(12);")
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD MIESANALIZY2 VARCHAR(12);")




        '        '---------------------------------------------------------------------
        '        'SlownikCoDalej
        '        '---------------------------------------------------------------------
        '        'Public scdIDENT As String             'IDENT       Text, 20
        '        'Public scdOPIS As String              'OPIS        memo
        '        'Public scdWersja As Integer           'WERSJA      Integer

        '        AktDB("SlownikCoDalej", "CREATE TABLE SlownikCoDalej (" & vbCrLf & _
        '                   "   IDENT varchar(20) NOT NULL ," & vbCrLf & _
        '                   "   OPIS wchar NOT NULL ," & vbCrLf & _
        '                   "   WERSJA integer NOT NULL " & vbCrLf & _
        '                   ");")

        '        AktDB("SlownikCoDalej", "ALTER TABLE SlownikCoDalej] " & vbCrLf & _
        '                   "   ADD CONSTRAINT PK_SlownikCoDalej " & vbCrLf & _
        '                   "   PRIMARY KEY (IDENT);" )




        '        '---------------------------------------------------------------------
        '        'CoDalej
        '        '---------------------------------------------------------------------
        '        'Public cdUNIQID As String            'UNIQID Text, 40
        '        'Public cdIDENTKLIENTA As String      'IDENTKLIENTA Text, 6
        '        'Public cdIDENT As String             'IDENT        Text, 20
        '        'Public cdOPIS As String              'OPIS         memo
        '        'Public cdWersja As Integer           'WERSJA       Integer

        '        AktDB("CoDalej", "CREATE TABLE CoDalejPlan (" & vbCrLf & _
        '                   "   UNIQID varchar(40) NOT NULL ," & vbCrLf & _
        '                   "   IDENTKLIENTA varchar(6) NOT NULL ," & vbCrLf & _
        '                   "   IDENT varchar(20) NOT NULL ," & vbCrLf & _
        '                   "   OPIS wchar NOT NULL ," & vbCrLf & _
        '                   "   WERSJA integer NOT NULL " & vbCrLf & _
        '                   ");")

        '        AktDB("CoDalej", "ALTER TABLE CoDalejPlan " & vbCrLf & _
        '                   "   ADD CONSTRAINT PK_CoDalej " & vbCrLf & _
        '                   "   PRIMARY KEY (UNIQID,IDENTKLIENTA);")

        '        InicjujCoDalej()





        '        '---------------------------------------------------------------------
        '        'Kontakty HANDLOWE - nowe 05.09.2015
        '        '-----------------------------------------
        '        'Public oUniqIdKon As String           'UNIQID            varchar(40)
        '        'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
        '        'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
        '        'Public oPracownikKon As String        'PRACOWNIK        Text, 10     uzytkownik
        '        'Public oTematKon As String            'TEMAT            Text, 50
        '        'Public oUwagiKon As String            'UWAGI            Memo
        '        'Public oWersjaKon As Integer          'WERSJA           Integer

        '        AktDB("KontaktyHandlowe", "CREATE TABLE KontaktyHandlowe (" & vbCrLf & _
        '                   "   UNIQID varchar(40) NOT NULL ," & vbCrLf & _
        '                   "   IDENTKLIENTA varchar(6) NOT NULL ," & vbCrLf & _
        '                   "   DATAKONTAKTU varchar(10) NOT NULL ," & vbCrLf & _
        '                   "   PRACOWNIK varchar(10) NOT NULL ," & vbCrLf & _
        '                   "   TEMAT varchar(50) NOT NULL ," & vbCrLf & _
        '                   "   UWAGI wchar NOT NULL ," & vbCrLf & _
        '                   "   WERSJA integer NOT NULL " & vbCrLf & _
        '                   ");")

        '        AktDB("KontaktyHandlowe", "ALTER TABLE KontaktyHandlowe " & vbCrLf & _
        '                   "   ADD CONSTRAINT PK_KontaktyHandlowe " & vbCrLf & _
        '                   "   PRIMARY KEY (UNIQID);")

        '        InicjujUNIQIDKontakty()




        '        '**************************************
        '        '** pobierz informacje o UZYTKOWNIKU
        '        '**************************************
        '        '...by³o
        '        'Public oIdentUzytkownika As String          'IDENT          Text    10
        '        'Public oNazwaUzytkownika As String          'NAZWA          Text    50
        '        'Public oOpisUzytkownika As String           'OPIS           Text    50
        '        'Public oWersjaUzytkownika As Integer        'WERSJA         Integer


        '        '   ma byc
        '        'Public U_IdentUzytkownika As String           'IDENT             Text    10
        '        'Public U_NazwaUzytkownika As String           'NAZWA             Text    100
        '        'Public U_FunkcjaUzytkownika As String         'FUNKCJA           Text    100
        '        'Public U_HasloUzytkownika As String           'HASLO             Text    100
        '        'Public U_UprawnieniaUzytkownika As String     'UPRAWNIENIA       Text    100
        '        'Public U_UwagiUzytkownika As String           'UWAGI             Text
        '        'Public U_WersjaUzytkownika As Integer         'WERSJA            Integer

        '        'AktDB("Uzytkownicy", "ALTER TABLE Uzytkownicy ALTER COLUMN IDENT VARCHAR(10);")
        '        AktDB("Uzytkownicy", "ALTER TABLE Uzytkownicy ALTER COLUMN NAZWA VARCHAR(100);")
        '        AktDB("Uzytkownicy", "ALTER TABLE Uzytkownicy ADD FUNKCJA VARCHAR(100);")
        '        AktDB("Uzytkownicy", "ALTER TABLE Uzytkownicy ADD HASLO VARCHAR(100);")
        '        AktDB("Uzytkownicy", "ALTER TABLE Uzytkownicy ADD UPRAWNIENIA VARCHAR(100);")
        '        AktDB("Uzytkownicy", "ALTER TABLE Uzytkownicy ADD UWAGI VARCHAR(200);")
        '        AktDB("Uzytkownicy", "ALTER TABLE Uzytkownicy DROP COLUMN OPIS;")



        '        'Public ddrKolExcel02 As String            'KolExCEL02           String
        '        'Public ddrKolExcel03 As String            'KolExCEL03           String
        '        'Public ddrKolExcel04 As String            'KolExCEL04           String
        '        'Public ddrKolExcel05 As String            'KolExCEL05           String
        '        'Public ddrKolExcel06 As String            'KolExCEL06           String
        '        'Public ddrKolExcel07 As String            'KolExCEL07           String
        '        'Public ddrKolExcel08 As String            'KolExCEL08           String
        '        'Public ddrKolExcel09 As String            'KolExCEL09           String
        '        'Public ddrKolExcel10 As String            'KolExCEL10           String
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL01 VARCHAR(50);")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL02 VARCHAR(50);")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL03 VARCHAR(50);")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL04 VARCHAR(50);")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL05 VARCHAR(50);")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL06 VARCHAR(50);")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL07 VARCHAR(50);")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL08 VARCHAR(50);")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL09 VARCHAR(50);")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL10 VARCHAR(50);")


        '        'Public ddrExcel01 As Integer            'EXCEL01           Integer
        '        'Public ddrExcel02 As Integer            'EXCEL02           Integer
        '        'Public ddrExcel03 As Integer            'EXCEL03           Integer
        '        'Public ddrExcel04 As Integer            'EXCEL04           Integer
        '        'Public ddrExcel05 As Integer            'EXCEL05           Integer
        '        'Public ddrExcel06 As Integer            'EXCEL06           Integer
        '        'Public ddrExcel07 As Integer            'EXCEL07           Integer
        '        'Public ddrExcel08 As Integer            'EXCEL08           Integer
        '        'Public ddrExcel09 As Integer            'EXCEL09           Integer
        '        'Public ddrExcel10 As Integer            'EXCEL10           Integer
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL01 INTEGER;")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL02 INTEGER;")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL03 INTEGER;")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL04 INTEGER;")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL05 INTEGER;")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL06 INTEGER;")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL07 INTEGER;")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL08 INTEGER;")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL09 INTEGER;")
        '        AktDB("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL10 INTEGER;")

        '    End If


        '    '--------------------------------------
        '    If AktualnaWersja <= 270 Then
        '        'Tabela Parametry
        '        'Public Par_RAKolumna01 As String = ""
        '        'Public Par_RAKolumna02 As String = ""
        '        'Public Par_RAKolumna03 As String = ""
        '        'Public Par_RAKolumna04 As String = ""
        '        'Public Par_RAKolumna05 As String = ""
        '        'Public Par_RAKolumna06 As String = ""
        '        'Public Par_RAKolumna07 As String = ""
        '        'Public Par_RAKolumna08 As String = ""
        '        'Public Par_RAKolumna09 As String = ""
        '        'Public Par_RAKolumna10 As String = ""
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA01 VARCHAR(100);")
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA02 VARCHAR(100);")
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA03 VARCHAR(100);")
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA04 VARCHAR(100);")
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA05 VARCHAR(100);")
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA06 VARCHAR(100);")
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA07 VARCHAR(100);")
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA08 VARCHAR(100);")
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA09 VARCHAR(100);")
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA10 VARCHAR(100);")
        '        'Public Par_IdentOddzialu As String = ""
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD IDENTODDZIALU VARCHAR(50);")




        '        'Public oRZURap01 As String              'RZURAP01     Text, 100
        '        'Public oRZURap02 As String              'RZURAP02     Text, 100
        '        'Public oRZURap03 As String              'RZURAP03     Text, 100
        '        'Public oRZURap04 As String              'RZURAP04     Text, 100
        '        'Public oRZURap05 As String              'RZURAP05     Text, 100
        '        'Public oRZURap06 As String              'RZURAP06     Text, 100
        '        'Public oRZURap07 As String              'RZURAP07     Text, 100
        '        'Public oRZURap08 As String              'RZURAP08     Text, 100
        '        'Public oRZURap09 As String              'RZURAP09     Text, 100
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD RZURAP01 VARCHAR(100);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD RZURAP02 VARCHAR(100);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD RZURAP03 VARCHAR(100);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD RZURAP04 VARCHAR(100);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD RZURAP05 VARCHAR(100);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD RZURAP06 VARCHAR(100);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD RZURAP07 VARCHAR(100);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD RZURAP08 VARCHAR(100);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD RZURAP09 VARCHAR(100);")


        '        'Public oPotencjalBR As String           'POTENCJALBR   Text, 4
        '        'Public oPotencjalPR As String           'POTENCJALPR   Text, 4
        '        'Public oPotencjalPPR As String          'POTENCJALPPR   Text, 4
        '        'Public oRokowaniaBR As String           'RokowaniaBR   Text, 4
        '        'Public oRokowaniaPR As String           'RokowaniaPR   Text, 4
        '        'Public oRokowaniaPPR As String          'RokowaniaPPR   Text, 4
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD POTENCJALBR VARCHAR(2);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD POTENCJALPR VARCHAR(2);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD POTENCJALPPR VARCHAR(2);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD ROKOWANIABR VARCHAR(2);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD ROKOWANIAPR VARCHAR(2);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD ROKOWANIAPPR VARCHAR(2);")

        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN POTENCJALBR VARCHAR(4);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN POTENCJALPR VARCHAR(4);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN POTENCJALPPR VARCHAR(4);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN ROKOWANIABR VARCHAR(4);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN ROKOWANIAPR VARCHAR(4);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN ROKOWANIAPPR VARCHAR(4);")



        '        'UUsun kolumny SKUP
        '        ''---------------------------------------------------------------------
        '        ''Obroty
        '        'Public oIdentObr As String            'IDENTKLIENTA     Text, 6
        '        'Public oDataObr As String             'DATA             Text,10
        '        'Public oLWartSprzedObr As Double      'LWARTSPRZED      Double
        '        'Public oLIloSprzedObr As Double       'LILOPRZED        Double
        '        'Public oLMarSprzedObr As Double       'LMARPRZED        Double
        '        'Public oAWartSprzedObr As Double      'AWARTSPRZED      Double
        '        'Public oAIloSprzedObr As Double       'AILOSPRZED       Double
        '        'Public oAMARSprzedObr As Double       'AMARSPRZED       Double
        '        'Public oLOWartSprzedObr As Double      'LOWARTSPRZED      Double
        '        'Public oLOIloSprzedObr As Double       'LOILOPRZED        Double
        '        'Public oLOMARSprzedObr As Double       'LOMARPRZED        Double
        '        'Public oAOWartSprzedObr As Double      'AOWARTSPRZED      Double
        '        'Public oAOIloSprzedObr As Double       'AOILOSPRZED       Double
        '        'Public oAOMARSprzedObr As Double       'AOMARSPRZED       Double
        '        'Public oWersjaObr As Integer          'WERSJA           Integer
        '        AktDB("Obroty", "ALTER TABLE Obroty DROP COLUMN LWARTSKUP;")
        '        AktDB("Obroty", "ALTER TABLE Obroty DROP COLUMN LILOSKUP;")
        '        AktDB("Obroty", "ALTER TABLE Obroty DROP COLUMN AWARTSKUP;")
        '        AktDB("Obroty", "ALTER TABLE Obroty DROP COLUMN AILOSKUP;")
        '        AktDB("Obroty", "ALTER TABLE Obroty DROP COLUMN LOWARTSKUP;")
        '        AktDB("Obroty", "ALTER TABLE Obroty DROP COLUMN LOILOSKUP;")
        '        AktDB("Obroty", "ALTER TABLE Obroty DROP COLUMN AOWARTSKUP;")
        '        AktDB("Obroty", "ALTER TABLE Obroty DROP COLUMN AOILOSKUP;")

        '        AktDB("Obroty", "ALTER TABLE Obroty ADD LMARSPRZED double;")
        '        AktDB("Obroty", "ALTER TABLE Obroty ADD AMARSPRZED double;")
        '        AktDB("Obroty", "ALTER TABLE Obroty ADD LOMARSPRZED double;")
        '        AktDB("Obroty", "ALTER TABLE Obroty ADD AOMARSPRZED double;")




        '        ''---------------------------------------------------------------------
        '        ''ObrotyMies
        '        'Public oIdentMies As String            'IDENTKLIENTA     Text, 6
        '        'Public oMiesiacMies As String          'MIESIAC          Text,7
        '        'Public oLWartSprzedMies As Double      'LWARTSPRZED      Double
        '        'Public oLIloSprzedMies As Double       'LILOSPRZED       Double
        '        'Public oLMARSprzedMies As Double       'LMARSPRZED       Double
        '        'Public oAWartSprzedMies As Double      'AWARTSPRZED      Double
        '        'Public oAIloSprzedMies As Double       'AILOSPRZED       Double
        '        'Public oAMARSprzedMies As Double       'AMARSPRZED       Double
        '        'Public oLOWartSprzedMies As Double      'LOWARTSPRZED      Double
        '        'Public oLOIloSprzedMies As Double       'LOILOSPRZED       Double
        '        'Public oLOMARSprzedMies As Double       'LOMARSPRZED       Double
        '        'Public oAOWartSprzedMies As Double      'AOWARTSPRZED      Double
        '        'Public oAOIloSprzedMies As Double       'AOILOSPRZED       Double
        '        'Public oAOMARSprzedMies As Double       'AOMARSPRZED       Double
        '        'Public oWersjaMies As Integer          'WERSJA           Integer

        '        AktDB("ObrotyMies", "ALTER TABLE ObrotyMies DROP COLUMN LWARTSKUP;")
        '        AktDB("ObrotyMies", "ALTER TABLE ObrotyMies DROP COLUMN LILOSKUP;")
        '        AktDB("ObrotyMies", "ALTER TABLE ObrotyMies DROP COLUMN AWARTSKUP;")
        '        AktDB("ObrotyMies", "ALTER TABLE ObrotyMies DROP COLUMN AILOSKUP;")
        '        AktDB("ObrotyMies", "ALTER TABLE ObrotyMies DROP COLUMN LOWARTSKUP;")
        '        AktDB("ObrotyMies", "ALTER TABLE ObrotyMies DROP COLUMN LOILOSKUP;")
        '        AktDB("ObrotyMies", "ALTER TABLE ObrotyMies DROP COLUMN AOWARTSKUP;")
        '        AktDB("ObrotyMies", "ALTER TABLE ObrotyMies DROP COLUMN AOILOSKUP;")

        '        AktDB("ObrotyMies", "ALTER TABLE ObrotyMies ADD LMARSPRZED double;")
        '        AktDB("ObrotyMies", "ALTER TABLE ObrotyMies ADD AMARSPRZED double;")
        '        AktDB("ObrotyMies", "ALTER TABLE ObrotyMies ADD LOMARSPRZED double;")
        '        AktDB("ObrotyMies", "ALTER TABLE ObrotyMies ADD AOMARSPRZED double;")


        '    End If



        '    If AktualnaWersja <= 271 Then
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN NRTOFITXT VARCHAR(255);")

        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN KODPOCZTOWY VARCHAR(10);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN NRDOMU VARCHAR(10);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN IDDOMU VARCHAR(10);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN NIP VARCHAR(15);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN KTOWPISAL VARCHAR(10);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN REJKLIENTA VARCHAR(20);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN PKONTAKT VARCHAR(10);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN SKUPOPIEKUN VARCHAR(10);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN SKUPOKONTAKTD VARCHAR(12);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN SKUPNKONTAKTD VARCHAR(12);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDOPIEKUN VARCHAR(10);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDOKONTAKTD VARCHAR(12);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDNKONTAKTD VARCHAR(12);")
        '        AktDB("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDWAGA VARCHAR(20);")
        '        '--------------------------------------
        '        InicjuPolaKK271()
        '    End If


        '    If AktualnaWersja <= 300 Then
        '        'Public Par_WartGranWartosc As Double = ""
        '        'Public Par_WartGranProcent As Double = ""
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD WARTGRANWARTOSC DOUBLE;")
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD WARTGRANPROCENT DOUBLE;")

        '        'KatalogKlientow
        '        'Public oPrognoza01 As Double       'PROGNOZA01        Double
        '        'Public oPrognoza02 As Double       'PROGNOZA02        Double
        '        'Public oPrognoza03 As Double       'PROGNOZA03        Double
        '        'Public oPrognozaK1 As Double       'PROGNOZAK1        Double
        '        'Public oPrognoza04 As Double       'PROGNOZA04        Double
        '        'Public oPrognoza05 As Double       'PROGNOZA05        Double
        '        'Public oPrognoza06 As Double       'PROGNOZA06        Double
        '        'Public oPrognozaK2 As Double       'PROGNOZAK2        Double
        '        'Public oPrognozaP1 As Double       'PROGNOZAP1        Double
        '        'Public oPrognoza07 As Double       'PROGNOZA07        Double
        '        'Public oPrognoza08 As Double       'PROGNOZA08        Double
        '        'Public oPrognoza09 As Double       'PROGNOZA09        Double
        '        'Public oPrognozaK3 As Double       'PROGNOZAK3        Double
        '        'Public oPrognoza10 As Double       'PROGNOZA10        Double
        '        'Public oPrognoza11 As Double       'PROGNOZA11        Double
        '        'Public oPrognoza12 As Double       'PROGNOZA12        Double
        '        'Public oPrognozaK4 As Double       'PROGNOZAK4        Double
        '        'Public oPrognozaP2 As Double       'PROGNOZAP2        Double
        '        'Public oPrognozaRR As Double       'PROGNOZARR        Double
        '        'Public oPrognozaRRPLAN As Double   'PROGNOZARRPLAN    Double

        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZA01 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZA02 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZA03 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZAK1 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZA04 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZA05 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZA06 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZAK2 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZAP1 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZA07 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZA08 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZA09 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZAK3 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZA10 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZA11 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZA12 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZAK4 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZAP2 DOUBLE;")
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZARR DOUBLE;")
        '    End If



        '    If AktualnaWersja <= 301 Then
        '        'Public Par_MiesPrognozy As string = ""
        '        AktDB("Parametry", "ALTER TABLE Parametry ADD MIESPROGNOZY VARCHAR(12);")

        '        'KatalogKlientow
        '        'Public oPrognozaRRPLAN As Double   'PROGNOZARRPLAN    Double
        '        AktDB("Klienci", "ALTER TABLE Klienci ADD PROGNOZARRPLAN DOUBLE;")
        '    End If

        '    '--------------------------------------
        '    ZapiszWersjeBazyDanych(NowaWersja)
        '    AktualnaWersja = NowaWersja
        '    '--------------------------------------
        '    txtRaport.Text &= Now & "   KONIEC aktualizacji" & vbCrLf
        '    txtRaport.SelectionStart = txtRaport.Text.Length
        '    txtRaport.ScrollToCaret()
        '    System.Windows.Forms.Application.DoEvents()
        'End If
    End Sub





    '**************************************
    '** sprawdzanie wersji
    '**************************************


    Dim sConnectionWersja As String = ConnStringAccess()

    Dim dbConnectionWersja As New OleDb.OleDbConnection(sConnectionWersja)
    Dim dbCommandSelectWersja As New OleDb.OleDbCommand(sSelectSQLWersja, dbConnectionWersja)
    Dim dbCommandUpdateWersja As New OleDb.OleDbCommand(sUpdateSQLWersja, dbConnectionWersja)
    Dim dbCommandInsertWersja As New OleDb.OleDbCommand(sInsertSQLWersja, dbConnectionWersja)

    Dim dbDataAdapterWersja As New OleDb.OleDbDataAdapter(dbCommandSelectWersja)
    Dim DataSetWersja As New DataSet
    Dim DataViewWersja As New DataView

    Public Function WersjaBazyDanychxxx() As Integer
        Dim oWersjaDB As Integer = 0

        DataSetWersja.Locale = New System.Globalization.CultureInfo("pl-PL")
        Dim DBMapowanieTabeli As System.Data.Common.DataTableMapping
        dbDataAdapterWersja.TableMappings.Clear()
        DBMapowanieTabeli = dbDataAdapterWersja.TableMappings.Add("Table", "TABELA_Wersja")
        '------------------------------------------
        'wypelnij DATASET
        oWersjaBazyDanych = 0
        Try
            dbConnectionWersja.Open()
        Catch Ex As System.Exception
        End Try

        If dbConnectionWersja.State = ConnectionState.Open Then
            Try
                dbDataAdapterWersja.FillSchema(DataSetWersja, SchemaType.Mapped, "TABELA_Wersja")
                dbDataAdapterWersja.Fill(DataSetWersja)
                dbConnectionWersja.Close()
                'definiuj DataView
                DataViewWersja = New DataView(DataSetWersja.Tables(0))
                If DataViewWersja.Count > 0 Then
                    Dim ii As Integer
                    For ii = 0 To DataViewWersja.Count - 1
                        oWersjaDB = DataViewWersja.Item(ii).Item(0)
                    Next
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (oWersjaDB)
    End Function

    Public Sub ZapiszWersjeBazyDanych(ByVal DBver As Integer)
        DataSetWersja.Locale = New System.Globalization.CultureInfo("pl-PL")
        Dim DBMapowanieTabeli As System.Data.Common.DataTableMapping
        dbDataAdapterWersja.TableMappings.Clear()
        DBMapowanieTabeli = dbDataAdapterWersja.TableMappings.Add("Table", "TABELA_Wersja")

        'komendy DataAdapter
        Dim objParam As OleDb.OleDbParameter
        '------------------------------------------
        'komenda UPDATE
        dbCommandUpdateWersja.Parameters.Add("@oident", OleDb.OleDbType.Integer, Nothing, "IDENT")
        objParam = dbCommandUpdateWersja.Parameters.Add("@orygSymbol", OleDb.OleDbType.Integer, Nothing, "IDENT")
        objParam.SourceVersion = DataRowVersion.Original
        dbDataAdapterWersja.UpdateCommand = dbCommandUpdateWersja

        '------------------------------------------
        'komenda INSERT
        objParam = dbCommandInsertWersja.Parameters.Add("@oIdent", OleDb.OleDbType.Integer, Nothing, "IDENT")
        objParam.SourceVersion = DataRowVersion.Current
        dbDataAdapterWersja.InsertCommand = dbCommandInsertWersja

        ' Add the RowUpdated event handler.
        AddHandler dbDataAdapterWersja.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

        '------------------------------------------
        'wypelnij DATASET
        oWersjaBazyDanych = 0
        Try
            dbConnectionWersja.Open()
        Catch Ex As System.Exception
        End Try

        If dbConnectionWersja.State = ConnectionState.Open Then
            Try
                dbDataAdapterWersja.FillSchema(DataSetWersja, SchemaType.Mapped, "TABELA_Wersja")
                dbDataAdapterWersja.Fill(DataSetWersja)
                dbConnectionWersja.Close()

                'definiuj DataView
                Dim NewRow As DataRow
                DataViewWersja = New DataView(DataSetWersja.Tables(0))
                If DataViewWersja.Count > 0 Then
                    For Each NewRow In DataViewWersja.Table.Rows
                        NewRow("IDENT") = DBver
                    Next
                Else
                    NewRow = DataSetWersja.Tables(0).NewRow()
                    NewRow("IDENT") = DBver
                    DataSetWersja.Tables(0).Rows.Add(NewRow)
                End If
                AktualizujBaze()
            Catch Ex As System.Exception
            End Try
        End If
    End Sub

    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujBaze()
        Try
            dbConnectionWersja.Open()
        Catch Ex As System.Exception
            ViewInLog(Ex, "Aktualizacja", Nothing)
        End Try

        If dbConnectionWersja.State = ConnectionState.Open Then
            oWystapilConcurrencyException = False
            '------------------------------------------------------------
            Try
                dbDataAdapterWersja.Update(DataSetWersja, "TABELA_Wersja")
            Catch Ex As System.Exception
                ViewInLog(Ex, "Aktualizacja", Nothing)
            End Try
            '------------------------------------------------------------
            If oWystapilConcurrencyException = True Then
                dbDataAdapterWersja.Fill(DataSetWersja)
            End If
            dbConnectionWersja.Close()
        End If
    End Sub
























































    '---------------------------------
    ' inicjowanie pola NRYKARTPAYBACK - zamiana Nothing na ""
    '--------------------------------

    Private Sub InicjuPolaKK250()
        Dim xKartaPayback As Boolean = False
        Dim xNryKartPayback As String = ""
        Dim xOfertaTZlozenia As String = ""
        Dim xOfertaPlik As String = ""
        Dim xOferta As String = ""

        Dim Wers As Integer = 0
        Dim i As Integer = 0
        Dim dr As DataRow

        '--------------------------------------
        'AktDB("Klienci", "ALTER TABLE Klienci ADD KARTAPAYBACK BIT;")
        'AktDB("Klienci", "ALTER TABLE Klienci ADD NRYKARTPAYBACK NOTE;")

        txtRaport.Text &= "Inicjujê Nowe kolumny Kartoteki Klientow" & vbCrLf
        txtRaport.SelectionStart = txtRaport.Text.Length
        txtRaport.ScrollToCaret()
        System.Windows.Forms.Application.DoEvents()

        Dim dbSelectKli As String = "SELECT " &
                                               "IDENTKLIENTA, " &
                                               "KARTAPAYBACK, " &
                                               "NRYKARTPAYBACK, " &
                                               "OFERTATZLOZENIA, " &
                                               "OFERTAPLIK, " &
                                               "OFERTA, " &
                                               "WERSJA " &
                                            "FROM Klienci"

        Dim dbUpdateKli As String = "UPDATE Klienci SET " &
                                                     "KARTAPAYBACK=@oKARTAPAYBACK, " &
                                                     "NRYKARTPAYBACK=@oNRYKARTPAYBACK, " &
                                                     "OFERTATZLOZENIA=@oOFERTATZLOZENIA, " &
                                                     "OFERTAPLIK=@oOFERTAPLIK, " &
                                                     "OFERTA=@oOFERTA, " &
                                                     "WERSJA=@oWersjaKli " &
                                            "WHERE (IDENTKLIENTA=@orygSymbol) AND " &
                                                  "(WERSJA=@orygWersja)"


        DataSetKlienci = New DataSet
        DataViewKlienci = New DataView
        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionKlienci)
            'dbCommandSelect = New OleDb.OleDbCommand(dbSelectKli, dbConnection)
            ''dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLKlienci, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(dbUpdateKli, dbConnection)
            ''dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLKlienci, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Klienci")

            'Dim objParam As OleDb.OleDbParameter
            ''------------------------------------------
            ''komenda UPDATE
            ''parametry aktualizacji
            ''dbcommandUpdate.Parameters.Add("@oIdentKli", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
            'dbCommandUpdate.Parameters.Add("@oKARTAPAYBACK", OleDb.OleDbType.Boolean, Nothing, "KARTAPAYBACK")
            'dbCommandUpdate.Parameters.Add("@oNRYKARTPAYBACK", OleDb.OleDbType.WChar, Nothing, "NRYKARTPAYBACK")
            'dbCommandUpdate.Parameters.Add("@oOFERTATZLOZENIA", OleDb.OleDbType.VarChar, 4, "OFERTATZLOZENIA")
            'dbCommandUpdate.Parameters.Add("@oOFERTAPLIK", OleDb.OleDbType.VarChar, 100, "OFERTAPLIK")
            'dbCommandUpdate.Parameters.Add("@oOFERTA", OleDb.OleDbType.VarChar, 50, "OFERTA")
            'dbCommandUpdate.Parameters.Add("@oWersjaKli", OleDb.OleDbType.Integer, Nothing, "WERSJA")

            ''parametry filtrowania
            'objParam = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 10, "IDENTKLIENTA")
            'objParam.SourceVersion = DataRowVersion.Original
            'objParam = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            'objParam.SourceVersion = DataRowVersion.Original
            'dbDataAdapter.UpdateCommand = dbCommandUpdate

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapter.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnection.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnection.State
            'End Try


        Else
            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionStr)
            sqlCommandSelectKlienci = New SqlClient.SqlCommand(dbSelectKli, sqlConnectionKlienci)
            'sqlCommandDeleteKlienci = New SqlClient.SqlCommand(sDeleteSQLKlienci, sqlConnectionKlienci)
            sqlCommandUpdateKlienci = New SqlClient.SqlCommand(dbUpdateKli, sqlConnectionKlienci)
            'sqlCommandInsertKlienci = New SqlClient.SqlCommand(sInsertSQLKlienci, sqlConnectionKlienci)
            sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            Dim objParam As SqlClient.SqlParameter
            '------------------------------------------
            '------------------------------------------
            'komenda UPDATE
            'parametry aktualizacji
            'sqlCommandUpdateKlienci.Parameters.Add("@oIdentKli", sqldbtype.Char, 6, "IDENTKLIENTA")
            sqlCommandUpdateKlienci.Parameters.Add("@oKARTAPAYBACK", SqlDbType.Bit, Nothing, "KARTAPAYBACK")
            sqlCommandUpdateKlienci.Parameters.Add("@oNRYKARTPAYBACK", SqlDbType.Text, Nothing, "NRYKARTPAYBACK")
            sqlCommandUpdateKlienci.Parameters.Add("@oOFERTATZLOZENIA", SqlDbType.VarChar, 4, "OFERTATZLOZENIA")
            sqlCommandUpdateKlienci.Parameters.Add("@oOFERTAPLIK", SqlDbType.VarChar, 100, "OFERTAPLIK")
            sqlCommandUpdateKlienci.Parameters.Add("@oOFERTA", SqlDbType.VarChar, 50, "OFERTA")
            sqlCommandUpdateKlienci.Parameters.Add("@oWersjaKli", SqlDbType.Int, Nothing, "WERSJA")

            'parametry filtrowania
            objParam = sqlCommandUpdateKlienci.Parameters.Add("@orygSymbol", SqlDbType.Char, 10, "IDENTKLIENTA")
            objParam.SourceVersion = DataRowVersion.Original
            objParam = sqlCommandUpdateKlienci.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Original
            sqlDataAdapterKlienci.UpdateCommand = sqlCommandUpdateKlienci

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
                    'dbDataAdapter.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    'dbDataAdapter.Fill(DataSetKlienci)
                    'dbConnection.Close()
                Else
                    sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                'DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("NRKARTY")}
                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))


                If DataSetKlienci.Tables(0).Rows.Count > 0 Then
                    For i = 0 To DataSetKlienci.Tables(0).Rows.Count - 1
                        dr = DataSetKlienci.Tables(0).Rows.Item(i)

                        xKartaPayback = GetLogDRField(dr, "KARTAPAYBACK")
                        xNryKartPayback = GetTxtDRField(dr, "NRYKARTPAYBACK")
                        xOfertaTZlozenia = GetTxtDRField(dr, "OFERTATZLOZENIA")
                        xOfertaPlik = GetTxtDRField(dr, "OFERTAPLIK")
                        xOferta = GetTxtDRField(dr, "OFERTA")
                        Wers = dr.Item("WERSJA")

                        'analizuj kod terminu zlozenia oferty
                        If IsNumeric(Mid(xOferta, 1, 1)) And
                            IsNumeric(Mid(xOferta, 2, 1)) And
                            IsNumeric(Mid(xOferta, 3, 1)) And
                            IsNumeric(Mid(xOferta, 4, 1)) And
                            (Not IsNumeric(Mid(xOferta, 5, 1))) Then
                            xOfertaTZlozenia = Mid(xOferta, 1, 4)
                        End If

                        dr.Item("KARTAPAYBACK") = xKartaPayback
                        dr.Item("NRYKARTPAYBACK") = xNryKartPayback
                        dr.Item("OFERTATZLOZENIA") = xOfertaTZlozenia
                        dr.Item("OFERTAPLIK") = xOfertaPlik
                        dr.Item("OFERTA") = xOferta
                        dr.Item("WERSJA") = ZmienWersje(Wers)
                    Next
                    AktualizujBazeKlienci()
                End If



            Catch Ex As System.Exception
                'ViewInLog(ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
                txtRaport.Text &= "B³¹d wykonania : " & Ex.Message & vbCrLf
                txtRaport.SelectionStart = txtRaport.Text.Length
                txtRaport.ScrollToCaret()
                System.Windows.Forms.Application.DoEvents()
            End Try
        End If
        '--------------------------------------
    End Sub




    '---------------------------------
    ' inicjowanie pola MARKERP - zamiana Nothing na FAlse
    '--------------------------------

    Private Sub InicjuPolaKK251()
        Dim xMarkerP As Boolean = False
        Dim Wers As Integer = 0
        Dim i As Integer = 0
        Dim dr As DataRow

        '--------------------------------------
        'AktDB("Klienci", "ALTER TABLE Klienci ADD MARKERP BIT;")

        txtRaport.Text &= "Inicjujê Nowe kolumny Kartoteki Klientow" & vbCrLf
        txtRaport.SelectionStart = txtRaport.Text.Length
        txtRaport.ScrollToCaret()
        System.Windows.Forms.Application.DoEvents()

        Dim dbSelectKli As String = "SELECT " &
                                               "IDENTKLIENTA, " &
                                               "MARKERP, " &
                                               "WERSJA " &
                                            "FROM Klienci"

        Dim dbUpdateKli As String = "UPDATE Klienci SET " &
                                                     "MARKERP=@oMARKERP, " &
                                                     "WERSJA=@oWersjaKli " &
                                            "WHERE (IDENTKLIENTA=@orygSymbol) AND " &
                                                  "(WERSJA=@orygWersja)"


        DataSetKlienci = New DataSet
        DataViewKlienci = New DataView
        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionKlienci)
            'dbCommandSelect = New OleDb.OleDbCommand(dbSelectKli, dbConnection)
            ''dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLKlienci, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(dbUpdateKli, dbConnection)
            ''dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLKlienci, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Klienci")

            'Dim objParam As OleDb.OleDbParameter
            ''------------------------------------------
            ''komenda UPDATE
            ''parametry aktualizacji
            ''dbcommandUpdate.Parameters.Add("@oIdentKli", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
            'dbCommandUpdate.Parameters.Add("@oMARKERP", OleDb.OleDbType.Boolean, Nothing, "MARKERP")
            'dbCommandUpdate.Parameters.Add("@oWersjaKli", OleDb.OleDbType.Integer, Nothing, "WERSJA")

            ''parametry filtrowania
            'objParam = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 10, "IDENTKLIENTA")
            'objParam.SourceVersion = DataRowVersion.Original
            'objParam = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            'objParam.SourceVersion = DataRowVersion.Original
            'dbDataAdapter.UpdateCommand = dbCommandUpdate

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapter.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnection.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnection.State
            'End Try


        Else
            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionStr)
            sqlCommandSelectKlienci = New SqlClient.SqlCommand(dbSelectKli, sqlConnectionKlienci)
            'sqlCommandDeleteKlienci = New SqlClient.SqlCommand(sDeleteSQLKlienci, sqlConnectionKlienci)
            sqlCommandUpdateKlienci = New SqlClient.SqlCommand(dbUpdateKli, sqlConnectionKlienci)
            'sqlCommandInsertKlienci = New SqlClient.SqlCommand(sInsertSQLKlienci, sqlConnectionKlienci)
            sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            Dim objParam As SqlClient.SqlParameter
            '------------------------------------------
            '------------------------------------------
            'komenda UPDATE
            'parametry aktualizacji
            'sqlCommandUpdateKlienci.Parameters.Add("@oIdentKli", sqldbtype.Char, 6, "IDENTKLIENTA")
            sqlCommandUpdateKlienci.Parameters.Add("@oMARKERP", SqlDbType.Bit, Nothing, "MARKERP")
            sqlCommandUpdateKlienci.Parameters.Add("@oWersjaKli", SqlDbType.Int, Nothing, "WERSJA")

            'parametry filtrowania
            objParam = sqlCommandUpdateKlienci.Parameters.Add("@orygSymbol", SqlDbType.Char, 10, "IDENTKLIENTA")
            objParam.SourceVersion = DataRowVersion.Original
            objParam = sqlCommandUpdateKlienci.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Original
            sqlDataAdapterKlienci.UpdateCommand = sqlCommandUpdateKlienci

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
                    'dbDataAdapter.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    'dbDataAdapter.Fill(DataSetKlienci)
                    'dbConnection.Close()
                Else
                    sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("IDENTKLIENTA")}
                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))


                If DataSetKlienci.Tables(0).Rows.Count > 0 Then
                    For i = 0 To DataSetKlienci.Tables(0).Rows.Count - 1
                        dr = DataSetKlienci.Tables(0).Rows.Item(i)

                        xMarkerP = GetLogDRField(dr, "MARKERP")
                        Wers = dr.Item("WERSJA")

                        dr.Item("MARKERP") = False
                        dr.Item("WERSJA") = ZmienWersje(Wers)
                    Next
                    AktualizujBazeKlienci()
                End If



            Catch Ex As System.Exception
                'ViewInLog(ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
                txtRaport.Text &= "B³¹d wykonania : " & Ex.Message & vbCrLf
                txtRaport.SelectionStart = txtRaport.Text.Length
                txtRaport.ScrollToCaret()
                System.Windows.Forms.Application.DoEvents()
            End Try
        End If
        '--------------------------------------
    End Sub





    Private Sub InicjuPolaKK271()
        Dim xMarkerP As Boolean = False
        Dim Wers As Integer = 0
        Dim i As Integer = 0
        Dim dr As DataRow

        txtRaport.Text &= "Inicjujê Nowe kolumny Kartoteki Klientow" & vbCrLf
        txtRaport.SelectionStart = txtRaport.Text.Length
        txtRaport.ScrollToCaret()
        System.Windows.Forms.Application.DoEvents()

        DataSetKlienci = New DataSet
        DataViewKlienci = New DataView
        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionKlienci)
            'dbCommandSelect = New OleDb.OleDbCommand(sSelectSQLKlienci, dbConnection)
            ''dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLKlienci, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLKlienci, dbConnection)
            ''dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLKlienci, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Klienci")

            ''DBDeleteKlienci(dbCommandDelete, dbDataAdapter)
            'DBUpdateKlienci(dbCommandUpdate, dbDataAdapter)
            ''DBInsertKlienci(dbCommandInsert, dbDataAdapter)

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapter.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnection.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnection.State
            'End Try


        Else
            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionStr)
            sqlCommandSelectKlienci = New SqlClient.SqlCommand(sSelectSQLKlienci, sqlConnectionKlienci)
            'sqlCommandDeleteKlienci = New SqlClient.SqlCommand(sDeleteSQLKlienci, sqlConnectionKlienci)
            sqlCommandUpdateKlienci = New SqlClient.SqlCommand(sUpdateSQLKlienci, sqlConnectionKlienci)
            'sqlCommandInsertKlienci = New SqlClient.SqlCommand(sInsertSQLKlienci, sqlConnectionKlienci)
            sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            'SQLDeleteKlienci(sqlCommandDeleteKlienci, sqlDataAdapterKlienci)
            SQLUpdateKlienci(sqlCommandUpdateKlienci, sqlDataAdapterKlienci)
            'SQLInsertKlienci(sqlCommandInsertKlienci, sqlDataAdapterKlienci)

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
                    'dbDataAdapter.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    'dbDataAdapter.Fill(DataSetKlienci)
                    'dbConnection.Close()
                Else
                    sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("IDENTKLIENTA")}
                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))


                '--------------------------------------
                'AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN KODPOCZTOWY VARCHAR(10);")
                'AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN NRDOMU VARCHAR(10);")
                'AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN IDDOMU VARCHAR(10);")
                'AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN NIP VARCHAR(15);")
                'AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN KTOWPISAL VARCHAR(10);")
                'AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN REJKLIENTA VARCHAR(20);")
                'AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN PKONTAKT VARCHAR(10);")
                'AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SKUPOPIEKUN VARCHAR(10);")
                'AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SKUPOKONTAKTD VARCHAR(12);")
                'AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SKUPNKONTAKTD VARCHAR(12);")
                'AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDOPIEKUN VARCHAR(10);")
                'AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDOKONTAKTD VARCHAR(12);")
                'AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDNKONTAKTD VARCHAR(12);")
                'AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDWAGA VARCHAR(20);")

                If DataSetKlienci.Tables(0).Rows.Count > 0 Then
                    For i = 0 To DataSetKlienci.Tables(0).Rows.Count - 1
                        dr = DataSetKlienci.Tables(0).Rows.Item(i)

                        dr.Item("KODPOCZTOWY") = Trim(dr.Item("KODPOCZTOWY"))
                        dr.Item("NRDOMU") = Trim(dr.Item("NRDOMU"))
                        dr.Item("IDDOMU") = Trim(dr.Item("IDDOMU"))
                        dr.Item("NIP") = Trim(dr.Item("NIP"))
                        dr.Item("KTOWPISAL") = Trim(dr.Item("KTOWPISAL"))
                        dr.Item("REJONKLIENTA") = Trim(dr.Item("REJONKLIENTA"))
                        dr.Item("PKONTAKT") = Trim(dr.Item("PKONTAKT"))
                        dr.Item("PROMOTOR") = Trim(dr.Item("PROMOTOR"))
                        dr.Item("PROMOTOROSTKONTAKTDZIEN") = Trim(dr.Item("PROMOTOROSTKONTAKTDZIEN"))
                        dr.Item("PROMOTORKOLKONTAKTDZIEN") = Trim(dr.Item("PROMOTORKOLKONTAKTDZIEN"))
                        dr.Item("OPIEKUN") = Trim(dr.Item("OPIEKUN"))
                        dr.Item("OPIEKUNOSTKONTAKTDZIEN") = Trim(dr.Item("OPIEKUNOSTKONTAKTDZIEN"))
                        dr.Item("OPIEKUNKOLKONTAKTDZIEN") = Trim(dr.Item("OPIEKUNKOLKONTAKTDZIEN"))

                        dr.Item("WAGAKLIENTA") = Trim(dr.Item("WAGAKLIENTA"))


                        '"SKUPOPIEKUN AS PROMOTOR," & _
                        ' "SKUPOKONTAKTR AS PROMOTOROSTKONTAKTROK," & _
                        ' "SKUPOKONTAKTT AS PROMOTOROSTKONTAKTTYDZ," & _
                        ' "SKUPOKONTAKTD AS PROMOTOROSTKONTAKTDZIEN," & _
                        ' "SKUPNKONTAKTR AS PROMOTORKOLKONTAKTROK," & _
                        ' "SKUPNKONTAKTT AS PROMOTORKOLKONTAKTTYDZ," & _
                        ' "SKUPNKONTAKTD AS PROMOTORKOLKONTAKTDZIEN," & _
                        ' "SKUPPLAN AS PROMOTORCODALEJ," & _
                        ' "SKUPUWAGI AS PROMOTORUWAGI," & _
                        ' "GDZIEKUPUJE," & _
                        ' "SPRZEDOPIEKUN AS OPIEKUN," & _
                        ' "SPRZEDWARTOSC AS SPRZEDAZ," & _
                        ' "SPRZEDOKONTAKTR AS OPIEKUNOSTKONTAKTROK," & _
                        ' "SPRZEDOKONTAKTT AS OPIEKUNOSTKONTAKTTYDZ," & _
                        ' "SPRZEDOKONTAKTD AS OPIEKUNOSTKONTAKTDZIEN," & _
                        ' "SPRZEDNKONTAKTR AS OPIEKUNKOLKONTAKTROK," & _
                        ' "SPRZEDNKONTAKTT AS OPIEKUNKOLKONTAKTTYDZ," & _
                        ' "SPRZEDNKONTAKTD AS OPIEKUNKOLKONTAKTDZIEN," & _

                        ' "SPRZEDWAGA AS WAGAKLIENTA," & _
                    Next
                    AktualizujBazeKlienci()
                End If



            Catch Ex As System.Exception
                'ViewInLog(ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
                txtRaport.Text &= "B³¹d wykonania : " & Ex.Message & vbCrLf
                txtRaport.SelectionStart = txtRaport.Text.Length
                txtRaport.ScrollToCaret()
                System.Windows.Forms.Application.DoEvents()
            End Try
        End If
        '--------------------------------------
    End Sub



    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujBazeKlienci()
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnection.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnection.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapter.Update(DataSetKlienci, "TABELA_Klienci")
            '    Catch Ex As System.Exception
            '        'ViewInLog(ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapter.Fill(DataSetKlienci)
            '    End If
            '    dbConnection.Close()
            'End If
        Else
            Try
                sqlConnectionKlienci.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionKlienci.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterKlienci.Update(DataSetKlienci, "TABELA_Klienci")
                Catch Ex As System.Exception
                    'ViewInLog(ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                End If
                sqlConnectionKlienci.Close()
            End If
        End If
    End Sub













    '---------------------------------
    ' inicjowanie pola UNIQID w kat KONTAKTY
    '--------------------------------

    Private Sub InicjujUNIQIDKontakty()
        Dim i As Integer = 0
        Dim ileRec As Integer = 0
        Dim dr As DataRow = Nothing
        Dim NewRow As DataRow = Nothing
        Dim LenTXT As Integer = 0

        txtRaport.Text &= "Inicjujê dane Kartoteki Kontakty" & vbCrLf
        txtRaport.SelectionStart = txtRaport.Text.Length
        txtRaport.ScrollToCaret()
        System.Windows.Forms.Application.DoEvents()



        DataSetKontaktyOld = New DataSet
        DataViewKontaktyOld = New DataView
        DataSetKontaktyOld.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionKontaktyOld = New OleDb.OleDbConnection(sConnectionKontakty)
            'dbCommandSelectKontaktyOld = New OleDb.OleDbCommand(sSelectSQLKontaktyOld, dbConnectionKontaktyOld)
            'dbCommandDeleteKontaktyOld = New OleDb.OleDbCommand(sDeleteSQLKontaktyOld, dbConnectionKontaktyOld)
            ''dbCommandUpdateKontaktyOld = New OleDb.OleDbCommand(sUpdateSQLKontaktyOld, dbConnectionKontaktyOld)
            ''dbCommandInsertKontaktyOld = New OleDb.OleDbCommand(sInsertSQLKontaktyOld, dbConnectionKontaktyOld)
            'dbDataAdapterKontaktyOld = New OleDb.OleDbDataAdapter(dbCommandSelectKontaktyOld)

            ''mapujemy domyslna nazwe tabeli
            'Dim dbMapowanieTabeliKontaktyOld As System.Data.Common.DataTableMapping
            'dbMapowanieTabeliKontaktyOld = dbDataAdapterKontaktyOld.TableMappings.Add("Table", "TABELA_KontaktyOld")

            'DBDeleteKontaktyOld(dbCommandDeleteKontaktyOld, dbDataAdapterKontaktyOld)
            ''DBUpdateKontaktyOld(dbCommandUpdateKontaktyOld, dbDataAdapterKontaktyOld)
            ''DBInsertKontaktyOld(dbCommandInsertKontaktyOld, dbDataAdapterKontaktyOld)

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapterKontaktyOld.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionKontaktyOld.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnectionKontaktyOld.State
            'End Try


        Else
            sqlConnectionKlienciKontaktyOld = New SqlClient.SqlConnection(sConnectionStr)
            sqlCommandSelectKlienciKontaktyOld = New SqlClient.SqlCommand(sSelectSQLKontaktyOld, sqlConnectionKlienciKontaktyOld)
            sqlCommandDeleteKlienciKontaktyOld = New SqlClient.SqlCommand(sDeleteSQLKontaktyOld, sqlConnectionKlienciKontaktyOld)
            'sqlCommandUpdateKlienciKontaktyOld = New SqlClient.SqlCommand(sUpdateSQLKontaktyOld, sqlConnectionKlienciKontaktyOld)
            'sqlCommandInsertKlienciKontaktyOld = New SqlClient.SqlCommand(sInsertSQLKontaktyOld, sqlConnectionKlienciKontaktyOld)
            sqlDataAdapterKlienciKontaktyOld = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienciKontaktyOld)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliKontaktyOld As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliKontaktyOld = sqlDataAdapterKlienciKontaktyOld.TableMappings.Add("Table", "TABELA_KontaktyOld")

            SQLDeleteKontaktyOld(sqlCommandDeleteKlienciKontaktyOld, sqlDataAdapterKlienciKontaktyOld)
            'SQLUpdateKontaktyOld(sqlCommandUpdateKlienciKontaktyOld, sqlDataAdapterKlienciKontaktyOld)
            'SQLInsertKontaktyOld(sqlCommandInsertKlienciKontaktyOld, sqlDataAdapterKlienciKontaktyOld)

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterKlienciKontaktyOld.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienciKontaktyOld.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionKlienciKontaktyOld.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterKontaktyOld.FillSchema(DataSetKontaktyOld, SchemaType.Mapped, "TABELA_KontaktyOld")
                    'dbDataAdapterKontaktyOld.Fill(DataSetKontaktyOld)
                    'dbConnectionKontaktyOld.Close()
                Else
                    sqlDataAdapterKlienciKontaktyOld.FillSchema(DataSetKontaktyOld, SchemaType.Mapped, "TABELA_KontaktyOld")
                    sqlDataAdapterKlienciKontaktyOld.Fill(DataSetKontaktyOld)
                    sqlConnectionKlienciKontaktyOld.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetKontaktyOld.Tables(0).PrimaryKey = New DataColumn() {DataSetKontaktyOld.Tables(0).Columns("IDENTKLIENTA"),
                                                                         DataSetKontaktyOld.Tables(0).Columns("DATAKONTAKTU")}
                DataViewKontaktyOld = New DataView(DataSetKontaktyOld.Tables(0))

                'jesli nie ma nic do przenoszwenia - koniec
                If DataViewKontaktyOld.Count = 0 Then Return

            Catch Ex As System.Exception
                'ViewInLog(ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
                txtRaport.Text &= "B³¹d wykonania : " & Ex.Message & vbCrLf
                txtRaport.SelectionStart = txtRaport.Text.Length
                txtRaport.ScrollToCaret()
                System.Windows.Forms.Application.DoEvents()

            End Try
        End If
        '--------------------------------------




        DataSetKontakty = New DataSet
        DataViewKontakty = New DataView
        DataSetKontakty.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionKontakty = New OleDb.OleDbConnection(sConnectionKontakty)
            'dbCommandSelectKontakty = New OleDb.OleDbCommand(sSelectSQLKontakty, dbConnectionKontakty)
            'dbCommandDeleteKontakty = New OleDb.OleDbCommand(sDeleteSQLKontakty, dbConnectionKontakty)
            'dbCommandUpdateKontakty = New OleDb.OleDbCommand(sUpdateSQLKontakty, dbConnectionKontakty)
            'dbCommandInsertKontakty = New OleDb.OleDbCommand(sInsertSQLKontakty, dbConnectionKontakty)
            'dbDataAdapterKontakty = New OleDb.OleDbDataAdapter(dbCommandSelectKontakty)

            ''mapujemy domyslna nazwe tabeli
            'Dim dbMapowanieTabeliKontakty As System.Data.Common.DataTableMapping
            'dbMapowanieTabeliKontakty = dbDataAdapterKontakty.TableMappings.Add("Table", "TABELA_Kontakty")

            'DBDeleteKontaktyHandlowe(dbCommandDeleteKontakty, dbDataAdapterKontakty)
            'DBUpdateKontaktyHandlowe(dbCommandUpdateKontakty, dbDataAdapterKontakty)
            'DBInsertKontaktyHandlowe(dbCommandInsertKontakty, dbDataAdapterKontakty)

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapterKontakty.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionKontakty.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnectionKontakty.State
            'End Try


        Else
            sqlConnectionKlienciKontakty = New SqlClient.SqlConnection(sConnectionStr)
            sqlCommandSelectKlienciKontakty = New SqlClient.SqlCommand(sSelectSQLKontakty, sqlConnectionKlienciKontakty)
            sqlCommandDeleteKlienciKontakty = New SqlClient.SqlCommand(sDeleteSQLKontakty, sqlConnectionKlienciKontakty)
            sqlCommandUpdateKlienciKontakty = New SqlClient.SqlCommand(sUpdateSQLKontakty, sqlConnectionKlienciKontakty)
            sqlCommandInsertKlienciKontakty = New SqlClient.SqlCommand(sInsertSQLKontakty, sqlConnectionKlienciKontakty)
            sqlDataAdapterKlienciKontakty = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienciKontakty)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliKontakty As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliKontakty = sqlDataAdapterKlienciKontakty.TableMappings.Add("Table", "TABELA_Kontakty")

            SQLDeleteKontaktyHandlowe(sqlCommandDeleteKlienciKontakty, sqlDataAdapterKlienciKontakty)
            SQLUpdateKontaktyHandlowe(sqlCommandUpdateKlienciKontakty, sqlDataAdapterKlienciKontakty)
            SQLInsertKontaktyHandlowe(sqlCommandInsertKlienciKontakty, sqlDataAdapterKlienciKontakty)

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterKlienciKontakty.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienciKontakty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionKlienciKontakty.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterKontakty.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_Kontakty")
                    'dbDataAdapterKontakty.Fill(DataSetKontakty)
                    'dbConnectionKontakty.Close()
                Else
                    sqlDataAdapterKlienciKontakty.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_Kontakty")
                    sqlDataAdapterKlienciKontakty.Fill(DataSetKontakty)
                    sqlConnectionKlienciKontakty.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                'DataSetKontakty.Tables(0).PrimaryKey = New DataColumn() {DataSetKontakty.Tables(0).Columns("NRKARTY")}
                DataViewKontakty = New DataView(DataSetKontakty.Tables(0))

                ileRec = DataSetKontaktyOld.Tables(0).Rows.Count
                LenTXT = txtRaport.Text.Length

                If DataSetKontaktyOld.Tables(0).Rows.Count > 0 Then
                    For i = DataSetKontaktyOld.Tables(0).Rows.Count - 1 To 0 Step -1
                        dr = DataSetKontaktyOld.Tables(0).Rows.Item(i)

                        'przepisz rekord do KontaktyHandlowe i dodaj UniqID
                        NewRow = DataSetKontakty.Tables(0).NewRow()
                        NewRow("UNIQID") = Guid.NewGuid.ToString
                        NewRow("IDENTKLIENTA") = dr("IDENTKLIENTA")
                        NewRow("DATAKONTAKTU") = dr("DATAKONTAKTU")
                        NewRow("PRACOWNIK") = dr("PRACOWNIK")
                        NewRow("TEMAT") = dr("TEMAT")
                        NewRow("UWAGI") = dr("UWAGI")
                        NewRow("WERSJA") = dr("WERSJA")
                        'dodaj rekord do DataSet
                        DataSetKontakty.Tables(0).Rows.Add(NewRow)
                        AktualizujBazeKontakty()

                        dr.Delete()
                        AktualizujBazeKontaktyOld()


                        txtRaport.Text = Mid(txtRaport.Text, 1, LenTXT) & "Przepisuje Rekord : " & Str(ileRec - i) & " / " & Str(ileRec) & vbCrLf
                        txtRaport.SelectionStart = txtRaport.Text.Length
                        txtRaport.ScrollToCaret()
                        System.Windows.Forms.Application.DoEvents()



                    Next
                End If

            Catch Ex As System.Exception
                'ViewInLog(ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
                txtRaport.Text &= "B³¹d wykonania : " & Ex.Message & vbCrLf
                txtRaport.SelectionStart = txtRaport.Text.Length
                txtRaport.ScrollToCaret()
                System.Windows.Forms.Application.DoEvents()

            End Try
        End If
        '--------------------------------------
    End Sub

    Private Sub AktualizujBazeKontakty()
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionKontakty.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnectionKontakty.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapterKontakty.Update(DataSetKontakty, "TABELA_Kontakty")
            '    Catch Ex As System.Exception
            '        'ViewInLog(ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapterKontakty.Fill(DataSetKontakty)
            '    End If
            '    dbConnectionKontakty.Close()
            'End If
        Else
            Try
                sqlConnectionKlienciKontakty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionKlienciKontakty.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterKlienciKontakty.Update(DataSetKontakty, "TABELA_Kontakty")
                Catch Ex As System.Exception
                    'ViewInLog(ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterKlienciKontakty.Fill(DataSetKontakty)
                End If
                sqlConnectionKlienciKontakty.Close()
            End If
        End If
    End Sub



    Private Sub AktualizujBazeKontaktyOld()
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionKontaktyOld.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnectionKontaktyOld.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapterKontaktyOld.Update(DataSetKontaktyOld, "TABELA_KontaktyOld")
            '    Catch Ex As System.Exception
            '        'ViewInLog(ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapterKontaktyOld.Fill(DataSetKontaktyOld)
            '    End If
            '    dbConnectionKontaktyOld.Close()
            'End If
        Else
            Try
                sqlConnectionKlienciKontaktyOld.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionKlienciKontaktyOld.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterKlienciKontaktyOld.Update(DataSetKontaktyOld, "TABELA_KontaktyOld")
                Catch Ex As System.Exception
                    'ViewInLog(ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterKlienciKontaktyOld.Fill(DataSetKontaktyOld)
                End If
                sqlConnectionKlienciKontaktyOld.Close()
            End If
        End If
    End Sub














    '---------------------------------
    ' inicjowanie pola UNIQID w kat KONTAKTY
    '--------------------------------

    Private Sub InicjujCoDalej()
        Dim i As Integer = 0
        Dim ileRec As Integer = 0
        Dim dr As DataRow = Nothing
        Dim NewRow As DataRow = Nothing
        Dim FoundRow As DataRow = Nothing
        Dim LenTXT As Integer = 0
        Dim findKeys(1) As Object


        txtRaport.Text &= "Inicjujê dane w tabeli CoDalej" & vbCrLf
        txtRaport.SelectionStart = txtRaport.Text.Length
        txtRaport.ScrollToCaret()
        System.Windows.Forms.Application.DoEvents()



        DataSetKlienci = New DataSet
        DataViewKlienci = New DataView
        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionKlienci = New OleDb.OleDbConnection(sConnectionKontakty)
            'dbCommandSelectKlienci = New OleDb.OleDbCommand(sSelectSQLKlienci, dbConnectionKlienci)
            ''dbCommandDeleteKlienci = New OleDb.OleDbCommand(sDeleteSQLKlienci, dbConnectionKlienci)
            ''dbCommandUpdateKlienci = New OleDb.OleDbCommand(sUpdateSQLKlienci, dbConnectionKlienci)
            ''dbCommandInsertKlienci = New OleDb.OleDbCommand(sInsertSQLKlienci, dbConnectionKlienci)
            'dbDataAdapterKlienci = New OleDb.OleDbDataAdapter(dbCommandSelectKlienci)

            ''mapujemy domyslna nazwe tabeli
            'Dim dbMapowanieTabeliKlienci As System.Data.Common.DataTableMapping
            'dbMapowanieTabeliKlienci = dbDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            ''DBDeleteKlienci(dbCommandDeleteKlienci, dbDataAdapterKlienci)
            ''DBUpdateKlienci(dbCommandUpdateKlienci, dbDataAdapterKlienci)
            ''DBInsertKlienci(dbCommandInsertKlienci, dbDataAdapterKlienci)

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
            sqlConnectionKlienciKlienci = New SqlClient.SqlConnection(sConnectionStr)
            sqlCommandSelectKlienciKlienci = New SqlClient.SqlCommand(sSelectSQLKlienci, sqlConnectionKlienciKlienci)
            'sqlCommandDeleteKlienciKlienci = New SqlClient.SqlCommand(sDeleteSQLKlienci, sqlConnectionKlienciKlienci)
            'sqlCommandUpdateKlienciKlienci = New SqlClient.SqlCommand(sUpdateSQLKlienci, sqlConnectionKlienciKlienci)
            'sqlCommandInsertKlienciKlienci = New SqlClient.SqlCommand(sInsertSQLKlienci, sqlConnectionKlienciKlienci)
            sqlDataAdapterKlienciKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienciKlienci)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliKlienci As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliKlienci = sqlDataAdapterKlienciKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            'SQLDeleteKlienci(sqlCommandDeleteKlienciKlienci, sqlDataAdapterKlienciKlienci)
            'SQLUpdateKlienci(sqlCommandUpdateKlienciKlienci, sqlDataAdapterKlienciKlienci)
            'SQLInsertKlienci(sqlCommandInsertKlienciKlienci, sqlDataAdapterKlienciKlienci)

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterKlienciKlienci.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienciKlienci.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionKlienciKlienci.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    'dbDataAdapterKlienci.Fill(DataSetKlienci)
                    'dbConnectionKlienci.Close()
                Else
                    sqlDataAdapterKlienciKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapterKlienciKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienciKlienci.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("IDENTKLIENTA")}
                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))

                'jesli nie ma nic do przenoszwenia - koniec
                If DataViewKlienci.Count = 0 Then Return

            Catch Ex As System.Exception
                'ViewInLog(ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
                txtRaport.Text &= "B³¹d wykonania : " & Ex.Message & vbCrLf
                txtRaport.SelectionStart = txtRaport.Text.Length
                txtRaport.ScrollToCaret()
                System.Windows.Forms.Application.DoEvents()

            End Try
        End If
        '--------------------------------------




        DataSetCoDalej = New DataSet
        DataViewCoDalej = New DataView
        DataSetCoDalej.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionCoDalej = New OleDb.OleDbConnection(sConnectionCoDalej)
            'dbCommandSelectCoDalej = New OleDb.OleDbCommand(sSelectSQLCoDalej, dbConnectionCoDalej)
            'dbCommandDeleteCoDalej = New OleDb.OleDbCommand(sDeleteSQLCoDalej, dbConnectionCoDalej)
            'dbCommandUpdateCoDalej = New OleDb.OleDbCommand(sUpdateSQLCoDalej, dbConnectionCoDalej)
            'dbCommandInsertCoDalej = New OleDb.OleDbCommand(sInsertSQLCoDalej, dbConnectionCoDalej)
            'dbDataAdapterCoDalej = New OleDb.OleDbDataAdapter(dbCommandSelectCoDalej)

            ''mapujemy domyslna nazwe tabeli
            'Dim dbMapowanieTabeliCoDalej As System.Data.Common.DataTableMapping
            'dbMapowanieTabeliCoDalej = dbDataAdapterCoDalej.TableMappings.Add("Table", "TABELA_CoDalej")

            'DBDeleteCoDalej(dbCommandDeleteCoDalej, dbDataAdapterCoDalej)
            'DBUpdateCoDalej(dbCommandUpdateCoDalej, dbDataAdapterCoDalej)
            'DBInsertCoDalej(dbCommandInsertCoDalej, dbDataAdapterCoDalej)

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapterCoDalej.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)
            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionCoDalej.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnectionCoDalej.State
            'End Try


        Else
            sqlConnectionKlienciCoDalej = New SqlClient.SqlConnection(sConnectionStr)
            sqlCommandSelectKlienciCoDalej = New SqlClient.SqlCommand(sSelectSQLCoDalej, sqlConnectionKlienciCoDalej)
            sqlCommandDeleteKlienciCoDalej = New SqlClient.SqlCommand(sDeleteSQLCoDalej, sqlConnectionKlienciCoDalej)
            sqlCommandUpdateKlienciCoDalej = New SqlClient.SqlCommand(sUpdateSQLCoDalej, sqlConnectionKlienciCoDalej)
            sqlCommandInsertKlienciCoDalej = New SqlClient.SqlCommand(sInsertSQLCoDalej, sqlConnectionKlienciCoDalej)
            sqlDataAdapterKlienciCoDalej = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienciCoDalej)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliCoDalej As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliCoDalej = sqlDataAdapterKlienciCoDalej.TableMappings.Add("Table", "TABELA_CoDalej")

            SQLDeleteCoDalej(sqlCommandDeleteKlienciCoDalej, sqlDataAdapterKlienciCoDalej)
            SQLUpdateCoDalej(sqlCommandUpdateKlienciCoDalej, sqlDataAdapterKlienciCoDalej)
            SQLInsertCoDalej(sqlCommandInsertKlienciCoDalej, sqlDataAdapterKlienciCoDalej)

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterKlienciCoDalej.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienciCoDalej.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionKlienciCoDalej.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterCoDalej.FillSchema(DataSetCoDalej, SchemaType.Mapped, "TABELA_CoDalej")
                    'dbDataAdapterCoDalej.Fill(DataSetCoDalej)
                    'dbConnectionCoDalej.Close()
                Else
                    sqlDataAdapterKlienciCoDalej.FillSchema(DataSetCoDalej, SchemaType.Mapped, "TABELA_CoDalej")
                    sqlDataAdapterKlienciCoDalej.Fill(DataSetCoDalej)
                    sqlConnectionKlienciCoDalej.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetCoDalej.Tables(0).PrimaryKey = New DataColumn() {DataSetCoDalej.Tables(0).Columns("UNIQID"),
                                                                        DataSetCoDalej.Tables(0).Columns("IDENTKLIENTA")}
                DataViewCoDalej = New DataView(DataSetCoDalej.Tables(0))

                ileRec = DataSetKlienci.Tables(0).Rows.Count
                LenTXT = txtRaport.Text.Length

                If DataSetKlienci.Tables(0).Rows.Count > 0 Then
                    For i = 0 To DataSetKlienci.Tables(0).Rows.Count - 1
                        dr = DataSetKlienci.Tables(0).Rows.Item(i)
                        cdOPIS = ""
                        cdIDENTKLIENTA = dr("NRKARTY")
                        cdIDENT = ""

                        ''CoDalej
                        'Public cdUNIQID As String            'UNIQID       Text, 40
                        'Public cdIDENTKLIENTA As String      'IDENTKLIENTA Text, 40
                        'Public cdIDENT As String             'IDENT        Text, 40
                        'Public cdOPIS As String              'OPIS         memo
                        'Public cdWersja As Integer           'WERSJA       Integer

                        If Len(cdOPIS) > 0 Then
                            'sprawdz czy ju¿ przepisano
                            DataViewCoDalej.RowFilter = "(IDENTKLIENTA='" & cdIDENTKLIENTA & "') AND (IDENT='')"
                            If DataViewCoDalej.Count > 0 Then
                                ' juz jest taki zapis
                            Else
                                ' nie ma - dopisz
                                'przepisz rekord do CoDalej i dodaj UniqID
                                NewRow = DataSetCoDalej.Tables(0).NewRow()
                                NewRow("UNIQID") = Guid.NewGuid.ToString
                                NewRow("IDENTKLIENTA") = cdIDENTKLIENTA
                                NewRow("IDENT") = cdIDENT
                                NewRow("OPIS") = cdOPIS
                                NewRow("WERSJA") = 1
                                'dodaj rekord do DataSet
                                DataSetCoDalej.Tables(0).Rows.Add(NewRow)
                                AktualizujBazeCoDalej()
                            End If
                            'DataViewCoDalej.RowFilter = ""
                        End If

                        txtRaport.Text = Mid(txtRaport.Text, 1, LenTXT) & "Przepisuje Rekord : " & Str(i) & " / " & Str(ileRec) & vbCrLf
                        txtRaport.SelectionStart = txtRaport.Text.Length
                        txtRaport.ScrollToCaret()
                        System.Windows.Forms.Application.DoEvents()
                    Next
                End If

            Catch Ex As System.Exception
                'ViewInLog(ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
                txtRaport.Text &= "B³¹d wykonania : " & Ex.Message & vbCrLf
                txtRaport.SelectionStart = txtRaport.Text.Length
                txtRaport.ScrollToCaret()
                System.Windows.Forms.Application.DoEvents()

            End Try
        End If
        '--------------------------------------
    End Sub

    Private Sub AktualizujBazeCoDalej()
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionCoDalej.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnectionCoDalej.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapterCoDalej.Update(DataSetCoDalej, "TABELA_CoDalej")
            '    Catch Ex As System.Exception
            '        'ViewInLog(ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapterCoDalej.Fill(DataSetCoDalej)
            '    End If
            '    dbConnectionCoDalej.Close()
            'End If
        Else
            Try
                sqlConnectionKlienciCoDalej.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionKlienciCoDalej.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterKlienciCoDalej.Update(DataSetCoDalej, "TABELA_CoDalej")
                Catch Ex As System.Exception
                    'ViewInLog(ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterKlienciCoDalej.Fill(DataSetCoDalej)
                End If
                sqlConnectionKlienciCoDalej.Close()
            End If
        End If
    End Sub

















































    Private Sub AktSQL(ByVal Komunikat As String, ByVal Kwerenda As String)
        Dim Wynik As Integer = 0

        txtRaport.Text &= "Wykonujê aktualizacje tabeli " & Komunikat & vbCrLf
        txtRaport.SelectionStart = txtRaport.Text.Length
        txtRaport.ScrollToCaret()
        System.Windows.Forms.Application.DoEvents()

        Try
            sqlCommand.Connection = sqlConnectionKlienci
            sqlCommand.CommandType = CommandType.Text
            'sqlCommand.CommandText = sqlAlterParametry123_1
            sqlCommand.CommandText = Kwerenda
            sqlConnectionKlienci.Open()

            Wynik = sqlCommand.ExecuteNonQuery()
            If Wynik = -1 Then
            End If
            'txtRaport.Text &= "OK, aktualizacja wykonana poprawnie" & vbCrLf
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
            txtRaport.Text &= "B³¹d wykonania aktualizacji tabeli : " & ex.Message & vbCrLf
            txtRaport.SelectionStart = txtRaport.Text.Length
            txtRaport.ScrollToCaret()
            System.Windows.Forms.Application.DoEvents()

        Finally
            sqlConnectionKlienci.Close()
        End Try
    End Sub



    '*************************************************
    ' Aktualizuj baze SQL
    '*************************************************

    Private Sub AktualizujBazeSQL()
        Dim Upgrade As Boolean = False
        Dim Wynik As Integer = 0

        If AktualnaWersja = NowaWersja Then
            If MessageBox.Show("Baza danych jest juz zaktualizowana do wersji " & (NowaWersja / 100).ToString("N2") & vbCrLf &
                "Nie ma potrzeby wykonywania aktualizacji." &
                "Czy chcesz wykonaæ aktualizacje mimo wszystko ?",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then
                Upgrade = True
            Else
                Upgrade = False
            End If
        ElseIf AktualnaWersja > NowaWersja Then
            If MessageBox.Show("Baza danych jest juz zaktualizowana do wersji " & (AktualnaWersja / 100).ToString("N2") & vbCrLf &
                "Nie mogê wykonaæ aktualizacji wstecznej do wersji " & (NowaWersja / 100).ToString("N2"),
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.OK Then
                Upgrade = False
            End If
        Else
            Upgrade = True
        End If

        If Upgrade Then
            txtRaport.Text &= vbCrLf & Now & "  Rozpoczynam aktualizacje" & vbCrLf
            txtRaport.SelectionStart = txtRaport.Text.Length
            txtRaport.ScrollToCaret()
            System.Windows.Forms.Application.DoEvents()

            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionStr)
            sqlCommand = New SqlClient.SqlCommand

            '--------------------------------------
            'If AktualnaWersja <= 123 Then
            '    'aktualizacje do wersji 1.23
            '    AktSQL("Parametry", "ALTER TABLE Parametry ADD IDENTUZYTKOWNIKA VARCHAR(50);")
            'End If
            '--------------------------------------
            If AktualnaWersja <= 210 Then
                'wersja 2.10
                'Public dbAlterKlienci210_1 As String = "ALTER TABLE Klienci ADD MARKER BIT;"
                'Public sqlAlterKlienci210_1 As String = "ALTER TABLE Klienci ADD MARKER BIT;"

                'Baza danych  SQL istnieje od wersji programu 2.00
                'aktualizacje do wersji 2.10
                '--------------------------------------
                AktSQL("Klienci", "ALTER TABLE Klienci ADD MARKER BIT;")
                '--------------------------------------
            End If


            '--------------------------------------
            If AktualnaWersja <= 220 Then
                'aktualizacje do wersji 2.20
                'Max dlugosc pola CHAR w ACCES=255 znakow
                '--------------------------------------
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDPLAN VARCHAR(300);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN ILOSPRZETU VARCHAR(300);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN NRTOFI INT;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD NRTOFITXT VARCHAR(100);")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD IDDOMU VARCHAR(10);")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD NUMNRDOMU INT;")
                '--------------------------------------
                txtRaport.Text &= "Przepisujê zmienione pola w tabeli Klienci" & vbCrLf
                txtRaport.SelectionStart = txtRaport.Text.Length
                txtRaport.ScrollToCaret()
                System.Windows.Forms.Application.DoEvents()

                Dim dbSelectKli220 As String = "SELECT " &
                                                       "IDENTKLIENTA, " &
                                                       "NRTOFI, " &
                                                       "NRTOFITXT, " &
                                                       "ILOSPRZETU, " &
                                                       "ILOSPRZETU2, " &
                                                       "NUMNRDOMU, " &
                                                       "NRDOMU, " &
                                                       "IDDOMU, " &
                                                       "WERSJA " &
                                                    "FROM Klienci"

                Dim dbUpdateKli220 As String = "UPDATE Klienci SET " &
                                                             "NRTOFI=@oNrTofiKli, " &
                                                             "NRTOFITXT=@oNrTofiTxtKli, " &
                                                             "ILOSPRZETU=@oIloSprzetuKli, " &
                                                             "ILOSPRZETU2=@oIloSprzetu2Kli, " &
                                                             "NUMNRDOMU=@oNumNrDomuKli, " &
                                                             "NRDOMU=@oNrDomuKli, " &
                                                             "IDDOMU=@oIDDomuKli, " &
                                                             "WERSJA=@oWersjaKli " &
                                                    "WHERE (IDENTKLIENTA=@orygSymbol) AND " &
                                                          "(WERSJA=@orygWersja)"

                sqlCommandSelectKlienci = New SqlClient.SqlCommand(dbSelectKli220, sqlConnectionKlienci)
                sqlCommandUpdateKlienci = New SqlClient.SqlCommand(dbUpdateKli220, sqlConnectionKlienci)
                sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

                'mapujemy domyslna nazwe tabeli
                Dim MapowanieTabeli As System.Data.Common.DataTableMapping
                MapowanieTabeli = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

                Dim objParam As SqlClient.SqlParameter

                '------------------------------------------
                'komenda UPDATE
                'parametry aktualizacji
                'sqlCommandUpdateKlienci.Parameters.Add("@oIdentKli", sqldbtype.Char, 6, "IDENTKLIENTA")
                sqlCommandUpdateKlienci.Parameters.Add("@oNrTofiKli", SqlDbType.Int, Nothing, "NRTOFI")
                sqlCommandUpdateKlienci.Parameters.Add("@oNrTofiTxtKli", SqlDbType.VarChar, 100, "NRTOFITXT")
                sqlCommandUpdateKlienci.Parameters.Add("@oIloSprzetuKli", SqlDbType.VarChar, 300, "ILOSPRZETU")
                sqlCommandUpdateKlienci.Parameters.Add("@oIloSprzetu2Kli", SqlDbType.Text, Nothing, "ILOSPRZETU2")
                sqlCommandUpdateKlienci.Parameters.Add("@oNumNrDomuKli", SqlDbType.Int, Nothing, "NUMNRDOMU")
                sqlCommandUpdateKlienci.Parameters.Add("@oNrDomuKli", SqlDbType.VarChar, 10, "NRDOMU")
                sqlCommandUpdateKlienci.Parameters.Add("@oIDDomuKli", SqlDbType.VarChar, 10, "IDDOMU")
                sqlCommandUpdateKlienci.Parameters.Add("@oWersjaKli", SqlDbType.Int, Nothing, "WERSJA")

                'parametry filtrowania
                objParam = sqlCommandUpdateKlienci.Parameters.Add("@orygSymbol", SqlDbType.Char, 10, "IDENTKLIENTA")
                objParam.SourceVersion = DataRowVersion.Original
                objParam = sqlCommandUpdateKlienci.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
                objParam.SourceVersion = DataRowVersion.Original
                sqlDataAdapterKlienci.UpdateCommand = sqlCommandUpdateKlienci

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
                If ConnectionState = ConnectionState.Open Then
                    Try
                        sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                        sqlDataAdapterKlienci.Fill(DataSetKlienci)
                        sqlConnectionKlienci.Close()

                        'zdefiniuj unikalny klucz wyszukiwania w tabeli
                        'DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("NRKARTY")}
                        DataViewKlienci = New DataView(DataSetKlienci.Tables(0))

                        Dim IloSp As String = ""
                        Dim NTofi As Integer = 0
                        Dim NTofiTxt As String = ""
                        Dim Wers As Integer = 0
                        Dim i As Integer = 0
                        Dim dr As DataRow

                        Dim NumNrD As Integer = 0
                        Dim NrD As String = ""
                        Dim IdD As String = ""
                        Dim j As Integer = 0
                        Dim ch As Char = ""

                        If DataSetKlienci.Tables(0).Rows.Count > 0 Then
                            For i = 0 To DataSetKlienci.Tables(0).Rows.Count - 1
                                dr = DataSetKlienci.Tables(0).Rows.Item(i)

                                IloSp = IIf(IsDBNull(dr.Item("ILOSPRZETU")), "", dr.Item("ILOSPRZETU"))
                                NTofi = IIf(IsDBNull(dr.Item("NRTOFI")), 0, dr.Item("NRTOFI"))
                                NTofiTxt = IIf(IsDBNull(dr.Item("NRTOFITXT")), 0, dr.Item("NRTOFITXT"))
                                NumNrD = IIf(IsDBNull(dr.Item("NUMNRDOMU")), 0, dr.Item("NUMNRDOMU"))
                                NrD = IIf(IsDBNull(dr.Item("NRDOMU")), "", dr.Item("NRDOMU"))
                                IdD = IIf(IsDBNull(dr.Item("IDDOMU")), "", dr.Item("IDDOMU"))
                                Wers = dr.Item("WERSJA")

                                'analizuj nr domu
                                NrD = Trim(NrD)
                                IdD = Trim(IdD)
                                If Len(NrD) > 0 And NumNrD = 0 Then
                                    For j = 1 To Len(NrD)
                                        ch = Mid(NrD, j, 1)
                                        If InStr("0123456789", ch) = 0 Then
                                            If Len(IdD) > 0 Then
                                                IdD = Trim(Mid(NrD, j)) & " " & IdD
                                            Else
                                                IdD = Trim(Mid(NrD, j))
                                            End If
                                            NrD = Mid(NrD, 1, j - 1)
                                            Exit For
                                        End If
                                    Next
                                    NumNrD = CInt(NrD)
                                End If

                                'analizuj numery TOFI -  maj¹ byc 6 znakowe z zerami na poczatku
                                'dopisujemy z pola NRTOFI(integer) do NRTOFITXT(txt)
                                If CInt(NTofi) = 0 Then
                                    dr.Item("NRTOFITXT") = ""
                                Else
                                    If Len(NTofiTxt) > 0 Then
                                        If InStr(NTofiTxt, Microsoft.VisualBasic.Right("00000" + Trim(Str(NTofi)), 5)) = 0 Then
                                            dr.Item("NRTOFITXT") &= "," & Microsoft.VisualBasic.Right("00000" + Trim(Str(NTofi)), 5)
                                        Else
                                            'juz jest dopisane
                                        End If
                                    Else
                                        dr.Item("NRTOFITXT") = Microsoft.VisualBasic.Right("00000" + Trim(Str(NTofi)), 5)
                                    End If
                                End If

                                dr.Item("ILOSPRZETU2") = Trim(IloSp)
                                dr.Item("NUMNRDOMU") = NumNrD
                                dr.Item("NRDOMU") = NrD
                                dr.Item("IDDOMU") = IdD
                                dr.Item("WERSJA") = ZmienWersje(Wers)
                                AktualizujBazeKlienci()
                            Next
                            'AktualizujBazeKlienci()
                        End If



                    Catch Ex As System.Exception
                        'ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                        txtRaport.Text &= "B³¹d wykonania : " & Ex.Message & vbCrLf
                        txtRaport.SelectionStart = txtRaport.Text.Length
                        txtRaport.ScrollToCaret()
                        System.Windows.Forms.Application.DoEvents()
                    End Try
                End If
                '--------------------------------------



                '--------------------------------------
                AktSQL("Obroty", "ALTER TABLE Obroty ADD LOWARTSPRZED float;")
                AktSQL("Obroty", "ALTER TABLE Obroty ADD LOILOSPRZED float;")
                AktSQL("Obroty", "ALTER TABLE Obroty ADD AOWARTSPRZED float;")
                AktSQL("Obroty", "ALTER TABLE Obroty ADD AOILOSPRZED float;")

                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD LOWARTSPRZED float;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD LOILOSPRZED float;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD AOWARTSPRZED float;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD AOILOSPRZED float;")

                AktSQL("AnalizyZakupu", "CREATE TABLE AnalizyZakupu " &
                                                    "(IDENTKLIENTA CHAR(6) NOT NULL, " &
                                                        "WARTZA00 float NOT NULL, " &
                                                        "ILOSZA00 float NOT NULL, " &
                                                        "WARTZA01 float NOT NULL, " &
                                                        "ILOSZA01 float NOT NULL, " &
                                                        "WARTZA02 float NOT NULL, " &
                                                        "ILOSZA02 float NOT NULL, " &
                                                        "WARTZA03 float NOT NULL, " &
                                                        "ILOSZA03 float NOT NULL, " &
                                                        "WARTZA04 float NOT NULL, " &
                                                        "ILOSZA04 float NOT NULL, " &
                                                        "WARTZA05 float NOT NULL, " &
                                                        "ILOSZA05 float NOT NULL, " &
                                                        "WARTZA06 float NOT NULL, " &
                                                        "ILOSZA06 float NOT NULL, " &
                                                        "WARTZA07 float NOT NULL, " &
                                                        "ILOSZA07 float NOT NULL, " &
                                                        "WARTZA08 float NOT NULL, " &
                                                        "ILOSZA08 float NOT NULL, " &
                                                        "WARTZA09 float NOT NULL, " &
                                                        "ILOSZA09 float NOT NULL, " &
                                                        "WARTZA10 float NOT NULL, " &
                                                        "ILOSZA10 float NOT NULL, " &
                                                        "WARTZA11 float NOT NULL, " &
                                                        "ILOSZA11 float NOT NULL, " &
                                                        "WARTZA12 float NOT NULL, " &
                                                        "ILOSZA12 float NOT NULL, " &
                                                        "WARTZA13 float NOT NULL, " &
                                                        "ILOSZA13 float NOT NULL, " &
                                                        "WARTZA14 float NOT NULL, " &
                                                        "ILOSZA14 float NOT NULL, " &
                                                        "WARTZA15 float NOT NULL, " &
                                                        "ILOSZA15 float NOT NULL, " &
                                                        "WARTZA16 float NOT NULL, " &
                                                        "ILOSZA16 float NOT NULL, " &
                                                        "WARTZA17 float NOT NULL, " &
                                                        "ILOSZA17 float NOT NULL, " &
                                                        "WARTZA18 float NOT NULL, " &
                                                        "ILOSZA18 float NOT NULL, " &
                                                        "WARTZA19 float NOT NULL, " &
                                                        "ILOSZA19 float NOT NULL, " &
                                                        "WARTZA20 float NOT NULL, " &
                                                        "ILOSZA20 float NOT NULL, " &
                                                        "WARTZA21 float NOT NULL, " &
                                                        "ILOSZA21 float NOT NULL, " &
                                                        "WARTZA22 float NOT NULL, " &
                                                        "ILOSZA22 float NOT NULL, " &
                                                        "WARTZA23 float NOT NULL, " &
                                                        "ILOSZA23 float NOT NULL, " &
                                                        "WARTZA24 float NOT NULL, " &
                                                        "ILOSZA24 float NOT NULL, " &
                                                    "WERSJA INTEGER NOT NULL);")
                AktSQL("AnalizyZakupu", "ALTER TABLE AnalizyZakupu " &
                                                            "ADD CONSTRAINT PK_AnalizyZakupu " &
                                                            "PRIMARY KEY (IDENTKLIENTA);")


            End If
            '--------------------------------------





            If AktualnaWersja <= 230 Then
                'aktualizacje do wersji 2.30

                '--------------------------------------
                If AktualnaWersja < 230 Then
                    AktSQL("Parametry", "DROP TABLE Parametry;")
                End If
                AktSQL("Parametry", "CREATE TABLE Parametry " &
                                                    "(IDENT CHAR(10) NOT NULL, " &
                                                     "IDENTUZYTKOWNIKA VARCHAR(50) NOT NULL, " &
                                                     "NAZWAUZYTKOWNIKA VARCHAR(50) NOT NULL, " &
                                                     "ADRESUZYTKOWNIKA VARCHAR(50) NOT NULL, " &
                                                     "KONTOUZYTKOWNIKA VARCHAR(50) NOT NULL, " &
                                                     "BANKUZYTKOWNIKA VARCHAR(50) NOT NULL, " &
                                                     "MIEJSCOWOSCUZYTKOWNIKA VARCHAR(50) NOT NULL, " &
                                                     "NIPUZYTKOWNIKA VARCHAR(50) NOT NULL, " &
                                                     "DATAAKTANALIZY VARCHAR(10) NOT NULL, " &
                                                     "OSTOKRESANALIZY VARCHAR(10) NOT NULL, " &
                                                     "WERSJA INTEGER NOT NULL);")
                AktSQL("Parametry", "ALTER TABLE Parametry " &
                                                     "ADD CONSTRAINT PK_Parametry " &
                                                     "PRIMARY KEY (IDENT);")

                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD WARTZABM float;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD ILOSZABM float;")
                '--------------------------------------
            End If



            If AktualnaWersja <= 240 Then
                'aktualizacje do wersji 2.40
                '--------------------------------------
                AktSQL("Klienci", "ALTER TABLE Klienci ADD KARTAPAYBACK BIT;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD NRYKARTPAYBACK TEXT;")

                AktSQL("Parametry", "ALTER TABLE Parametry ADD DANEDOANALIZY VARCHAR(10);")
                '--------------------------------------
            End If


            If AktualnaWersja <= 250 Then
                'aktualizacje do wersji 2.50
                '--------------------------------------
                AktSQL("RodzajeKontaktow", "CREATE TABLE RodzajeKontaktow " &
                                             "(UNIQID VARCHAR(40) NOT NULL, " &
                                              "RODZAJKONTAKTU VARCHAR(50) NOT NULL, " &
                                              "WERSJA INTEGER NOT NULL);")
                AktSQL("RodzajeKontaktow", "ALTER TABLE RodzajeKontaktow " &
                                                     "ADD CONSTRAINT PK_RodzajeKontaktow " &
                                                     "PRIMARY KEY (UNIQID);")

                AktSQL("Parametry", "ALTER TABLE Parametry ADD KATALOGOFERTY VARCHAR(100);")

                AktSQL("Klienci", "ALTER TABLE Klienci ADD OFERTATZLOZENIA VARCHAR(4);")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD OFERTAPLIK VARCHAR(100);")
                '--------------------------------------
                InicjuPolaKK250()
            End If

            '--------------------------------------
            If AktualnaWersja <= 251 Then
                '--------------------------------------
                AktSQL("Klienci", "ALTER TABLE Klienci ADD MARKERP BIT;")
                '--------------------------------------
                InicjuPolaKK251()
            End If


            '--------------------------------------
            If AktualnaWersja <= 252 Then
                '--------------------------------------
                AktSQL("Klienci", "ALTER TABLE Klienci ADD MARKERP BIT;")
                '---------------------------------------------------------------------
                'DaneDoRaportu
                '---------------------------------------------------------------------
                'Public ddrPracownik As String        'PRACOWNIK   Text, 10
                'Public ddrDataRaportu As String      'DATARAPORTU Text, 7   RRRR-MM
                'Public ddrKlBezDrukarki As Integer   'KLBEZDRUKARKI  Integer
                'Public ddrOferty As Integer          'OFERTY         Integer
                'Public ddrTestyWstawionw As Integer  'TESTYWSTAWIONE Integer
                'Public ddrCzyszczenie As Integer     'CZYSZCZENIE    Integer
                'Public ddrDostawy As Integer         'DOSTAWY        Integer
                'Public ddrSkup As Integer            'SKUP           Integer
                'Public ddrWersja As Integer          'WERSJA         Integer
                AktSQL("DaneDoRaportu", "CREATE TABLE [dbo].[DaneDoRaportu] (" & vbCrLf &
                           "   [PRACOWNIK] [varchar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                           "   [DATARAPORTU] [varchar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                           "   [KLBEZDRUKARKI] [int] NOT NULL ," & vbCrLf &
                           "   [OFERTY] [int] NOT NULL ," & vbCrLf &
                           "   [TESTYWSTAWIONE] [int] NOT NULL ," & vbCrLf &
                           "   [CZYSZCZENIE] [int] NOT NULL ," & vbCrLf &
                           "   [DOSTAWY] [int] NOT NULL ," & vbCrLf &
                           "   [SKUP] [int] NOT NULL ," & vbCrLf &
                           "   [WERSJA] [int] NOT NULL " & vbCrLf &
                           ") ON [PRIMARY]")

                AktSQL("DaneDoRaportu", "ALTER TABLE [dbo].[DaneDoRaportu] WITH NOCHECK ADD " & vbCrLf &
                           "   CONSTRAINT [PK_DaneDoRaportu] PRIMARY KEY  CLUSTERED " & vbCrLf &
                           "   (" & vbCrLf &
                           "   [PRACOWNIK], " & vbCrLf &
                           "   [DATARAPORTU]" & vbCrLf &
                           "   )  ON [PRIMARY] ")
                '--------------------------------------
            End If




            '--------------------------------------
            If AktualnaWersja <= 253 Then
                '--------------------------------------
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDPLAN VARCHAR(500);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDWAGA VARCHAR(20);")


                '---------------------------------------------------------------------
                'Akcje Handlowe - Opis
                'Public aoIdentAkcji As String            'IDENT     Text 20
                'Public aoDataAkcji As String             'DATA      Text,10
                'Public aoOpisAkcji As String             'OPIS      Text,100
                'Public aoUwagiAkcji As String            'UWAGI     Text,10
                'Public aoMarkerAkcji As String           'MARKER    logical
                'Public aoWersjaAkcji As Integer           'WERSJA    Integer
                AktSQL("AkcjeOpis", "ALTER TABLE AkcjeOpis ADD MARKER BIT DEFAULT 0 WITH VALUES;")

                AktSQL("Parametry", "ALTER TABLE Parametry ALTER COLUMN OSTOKRESANALIZY CHAR(14);")

            End If



            '--------------------------------------
            If AktualnaWersja <= 260 Then
                'Public oZakupPopRokKli As Double       'ZAKUPPOPROK        Double
                'Public oUslugiDodKli As String         'USLUGIDOD          Text, 200
                AktSQL("Klienci", "ALTER TABLE Klienci ADD ZAKUPPOPROK FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD USLUGIDOD VARCHAR(200) DEFAULT '' WITH VALUES;")




                'Public Par_wwwPAYBACK As String = ""
                'Public Par_MiesAnalizy1 As String = ""
                'Public Par_MiesAnalizy2 As String = ""
                AktSQL("Parametry", "ALTER TABLE Parametry ADD WWWPAYBACK VARCHAR(200) DEFAULT 'https://www.payback.pl/pb/id/36624/index.html/?utm_source=pryzmatwww&utm_medium=homepage&utm_campaign=ekupony' WITH VALUES;")
                AktSQL("Parametry", "ALTER TABLE Parametry ADD MIESANALIZY1 VARCHAR(12) DEFAULT 'TTTTTTTTTTTT' WITH VALUES;")
                AktSQL("Parametry", "ALTER TABLE Parametry ADD MIESANALIZY2 VARCHAR(12) DEFAULT 'TTTTTTTTTTTT' WITH VALUES;")



                '---------------------------------------------------------------------
                'SlownikCoDalej
                '---------------------------------------------------------------------
                'Public scdIDENT As String             'IDENT       Text, 40
                'Public scdOPIS As String              'OPIS        memo
                'Public scdWersja As Integer           'WERSJA      Integer

                AktSQL("SlownikCoDalej", "CREATE TABLE [dbo].[SlownikCoDalej] (" & vbCrLf &
                           "   [IDENT] [varchar] (20) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                           "   [OPIS] [text] COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                           "   [WERSJA] [int] NOT NULL " & vbCrLf &
                           ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]")

                AktSQL("SlownikCoDalej", "ALTER TABLE [dbo].[SlownikCoDalej] WITH NOCHECK ADD " & vbCrLf &
                           "   CONSTRAINT [PK_SlownikCoDalej] PRIMARY KEY  CLUSTERED " & vbCrLf &
                           "   (" & vbCrLf &
                           "   [IDENT]" & vbCrLf &
                           "   )  ON [PRIMARY] ")









                '---------------------------------------------------------------------
                'CoDalej
                '---------------------------------------------------------------------
                'Public cdUNIQID As String            'UNIQID Text, 40
                'Public cdIDENTKLIENTA As String      'IDENTKLIENTA Text, 40
                'Public cdIDENT As String             'IDENT        Text, 40
                'Public cdOPIS As String              'OPIS         memo
                'Public cdWersja As Integer           'WERSJA       Integer

                AktSQL("CoDalej", "CREATE TABLE [dbo].[CoDalejPlan] (" & vbCrLf &
                           "   [UNIQID] [varchar] (40) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                           "   [IDENTKLIENTA] [varchar] (6) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                           "   [IDENT] [varchar] (20) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                           "   [OPIS] [text] COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                           "   [WERSJA] [int] NOT NULL " & vbCrLf &
                           ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]")

                AktSQL("CoDalej", "ALTER TABLE [dbo].[CoDalejPlan] WITH NOCHECK ADD " & vbCrLf &
                           "   CONSTRAINT [PK_CoDalej] PRIMARY KEY  CLUSTERED " & vbCrLf &
                           "   (" & vbCrLf &
                           "   [UNIQID], " & vbCrLf &
                           "   [IDENTKLIENTA]" & vbCrLf &
                           "   )  ON [PRIMARY] ")



                InicjujCoDalej()





                '---------------------------------------------------------------------
                'Kontakty HANDLOWE - nowe 05.09.2015
                '-----------------------------------------
                'Public oUniqIdKon As String           'UNIQID            varchar(40)
                'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
                'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
                'Public oPracownikKon As String        'PRACOWNIK        Text, 10     uzytkownik
                'Public oTematKon As String            'TEMAT            Text, 50
                'Public oUwagiKon As String            'UWAGI            Memo
                'Public oWersjaKon As Integer          'WERSJA           Integer

                AktSQL("KontaktyHandlowe", "CREATE TABLE [dbo].[KontaktyHandlowe] (" & vbCrLf &
                           "   [UNIQID] [VarChar] (40) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                           "   [IDENTKLIENTA] [VarChar] (6) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                           "   [DATAKONTAKTU] [VarChar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                           "   [PRACOWNIK] [VarChar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                           "   [TEMAT] [varchar] (50) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                           "   [UWAGI] [text] COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                           "   [WERSJA] [int] NOT NULL " & vbCrLf &
                           ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]")

                AktSQL("KontaktyHandlowe", "ALTER TABLE [dbo].[KontaktyHandlowe] WITH NOCHECK ADD " & vbCrLf &
                           "   CONSTRAINT [PK_KontaktyHandlowe] PRIMARY KEY  CLUSTERED " & vbCrLf &
                           "   (" & vbCrLf &
                           "   [UNIQID] " & vbCrLf &
                           "   )  ON [PRIMARY] ")

                InicjujUNIQIDKontakty()








                '**************************************
                '** pobierz informacje o UZYTKOWNIKU
                '**************************************
                '...by³o
                'Public oIdentUzytkownika As String          'IDENT          Text    10
                'Public oNazwaUzytkownika As String          'NAZWA          Text    50
                'Public oOpisUzytkownika As String           'OPIS           Text    50
                'Public oWersjaUzytkownika As Integer        'WERSJA         Integer


                '   ma byc
                'Public U_IdentUzytkownika As String           'IDENT             Text    10
                'Public U_NazwaUzytkownika As String           'NAZWA             Text    100
                'Public U_FunkcjaUzytkownika As String         'FUNKCJA           Text    100
                'Public U_HasloUzytkownika As String           'HASLO             Text    100
                'Public U_UprawnieniaUzytkownika As String     'UPRAWNIENIA       Text    100
                'Public U_UwagiUzytkownika As String           'UWAGI             Text
                'Public U_WersjaUzytkownika As Integer         'WERSJA            Integer

                'AktSQL("Uzytkownicy", "ALTER TABLE Uzytkownicy ALTER COLUMN IDENT VARCHAR(20);")
                AktSQL("Uzytkownicy", "ALTER TABLE Uzytkownicy ALTER COLUMN NAZWA VARCHAR(100);")
                AktSQL("Uzytkownicy", "ALTER TABLE Uzytkownicy ADD FUNKCJA VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Uzytkownicy", "ALTER TABLE Uzytkownicy ADD HASLO VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Uzytkownicy", "ALTER TABLE Uzytkownicy ADD UPRAWNIENIA VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Uzytkownicy", "ALTER TABLE Uzytkownicy ADD UWAGI [text] COLLATE Polish_CI_AS NOT NULL DEFAULT '' WITH VALUES;")
                AktSQL("Uzytkownicy", "ALTER TABLE Uzytkownicy DROP COLUMN OPIS;")



                'Public ddrKolExcel02 As String            'KolExCEL02           String
                'Public ddrKolExcel03 As String            'KolExCEL03           String
                'Public ddrKolExcel04 As String            'KolExCEL04           String
                'Public ddrKolExcel05 As String            'KolExCEL05           String
                'Public ddrKolExcel06 As String            'KolExCEL06           String
                'Public ddrKolExcel07 As String            'KolExCEL07           String
                'Public ddrKolExcel08 As String            'KolExCEL08           String
                'Public ddrKolExcel09 As String            'KolExCEL09           String
                'Public ddrKolExcel10 As String            'KolExCEL10           String
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL01 VARCHAR(50) DEFAULT '' WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL02 VARCHAR(50) DEFAULT '' WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL03 VARCHAR(50) DEFAULT '' WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL04 VARCHAR(50) DEFAULT '' WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL05 VARCHAR(50) DEFAULT '' WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL06 VARCHAR(50) DEFAULT '' WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL07 VARCHAR(50) DEFAULT '' WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL08 VARCHAR(50) DEFAULT '' WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL09 VARCHAR(50) DEFAULT '' WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD KOLEXCEL10 VARCHAR(50) DEFAULT '' WITH VALUES;")


                'Public ddrExcel01 As Integer            'EXCEL01           Integer
                'Public ddrExcel02 As Integer            'EXCEL02           Integer
                'Public ddrExcel03 As Integer            'EXCEL03           Integer
                'Public ddrExcel04 As Integer            'EXCEL04           Integer
                'Public ddrExcel05 As Integer            'EXCEL05           Integer
                'Public ddrExcel06 As Integer            'EXCEL06           Integer
                'Public ddrExcel07 As Integer            'EXCEL07           Integer
                'Public ddrExcel08 As Integer            'EXCEL08           Integer
                'Public ddrExcel09 As Integer            'EXCEL09           Integer
                'Public ddrExcel10 As Integer            'EXCEL10           Integer
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL01 INTEGER DEFAULT 0 WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL02 INTEGER DEFAULT 0 WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL03 INTEGER DEFAULT 0 WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL04 INTEGER DEFAULT 0 WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL05 INTEGER DEFAULT 0 WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL06 INTEGER DEFAULT 0 WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL07 INTEGER DEFAULT 0 WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL08 INTEGER DEFAULT 0 WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL09 INTEGER DEFAULT 0 WITH VALUES;")
                AktSQL("DaneDoRaportu", "ALTER TABLE DaneDoRaportu ADD EXCEL10 INTEGER DEFAULT 0 WITH VALUES;")
            End If

            '--------------------------------------
            If AktualnaWersja <= 270 Then
                'Tabela Parametry
                'Public Par_RAKolumna01 As String = ""
                'Public Par_RAKolumna02 As String = ""
                'Public Par_RAKolumna03 As String = ""
                'Public Par_RAKolumna04 As String = ""
                'Public Par_RAKolumna05 As String = ""
                'Public Par_RAKolumna06 As String = ""
                'Public Par_RAKolumna07 As String = ""
                'Public Par_RAKolumna08 As String = ""
                'Public Par_RAKolumna09 As String = ""
                'Public Par_RAKolumna10 As String = ""
                AktSQL("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA01 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA02 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA03 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA04 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA05 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA06 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA07 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA08 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA09 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Parametry", "ALTER TABLE Parametry ADD RAKOLUMNA10 VARCHAR(100) DEFAULT '' WITH VALUES;")
                'Public Par_IdentOddzialu As String = ""
                AktSQL("Parametry", "ALTER TABLE Parametry ADD IDENTODDZIALU VARCHAR(50) DEFAULT '' WITH VALUES;")





                'Public oRZURap01 As String              'RZURAP01     Text, 100
                'Public oRZURap02 As String              'RZURAP02     Text, 100
                'Public oRZURap03 As String              'RZURAP03     Text, 100
                'Public oRZURap04 As String              'RZURAP04     Text, 100
                'Public oRZURap05 As String              'RZURAP05     Text, 100
                'Public oRZURap06 As String              'RZURAP06     Text, 100
                'Public oRZURap07 As String              'RZURAP07     Text, 100
                'Public oRZURap08 As String              'RZURAP08     Text, 100
                'Public oRZURap09 As String              'RZURAP09     Text, 100
                AktSQL("Klienci", "ALTER TABLE Klienci ADD RZURAP01 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD RZURAP02 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD RZURAP03 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD RZURAP04 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD RZURAP05 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD RZURAP06 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD RZURAP07 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD RZURAP08 VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD RZURAP09 VARCHAR(100) DEFAULT '' WITH VALUES;")

                'Public oPotencjalBR As String           'POTENCJALBR   Text, 4
                'Public oPotencjalPR As String           'POTENCJALPR   Text, 4
                'Public oPotencjalPPR As String          'POTENCJALPPR   Text, 4
                'Public oRokowaniaBR As String           'RokowaniaBR   Text, 4
                'Public oRokowaniaPR As String           'RokowaniaPR   Text, 4
                'Public oRokowaniaPPR As String          'RokowaniaPPR   Text, 4
                AktSQL("Klienci", "ALTER TABLE Klienci ADD POTENCJALBR VARCHAR(2) DEFAULT '' WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD POTENCJALPR VARCHAR(2) DEFAULT '' WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD POTENCJALPPR VARCHAR(2) DEFAULT '' WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD ROKOWANIABR VARCHAR(2) DEFAULT '' WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD ROKOWANIAPR VARCHAR(2) DEFAULT '' WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD ROKOWANIAPPR VARCHAR(2) DEFAULT '' WITH VALUES;")

                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN POTENCJALBR VARCHAR(4);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN POTENCJALPR VARCHAR(4);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN POTENCJALPPR VARCHAR(4);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN ROKOWANIABR VARCHAR(4);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN ROKOWANIAPR VARCHAR(4);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN ROKOWANIAPPR VARCHAR(4);")




                'UUsun kolumny SKUP
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
                'Public oWersjaObr As Integer          'WERSJA           Integer
                AktSQL("Obroty", "ALTER TABLE Obroty DROP COLUMN LWARTSKUP;")
                AktSQL("Obroty", "ALTER TABLE Obroty DROP COLUMN LILOSKUP;")
                AktSQL("Obroty", "ALTER TABLE Obroty DROP COLUMN AWARTSKUP;")
                AktSQL("Obroty", "ALTER TABLE Obroty DROP COLUMN AILOSKUP;")
                AktSQL("Obroty", "ALTER TABLE Obroty DROP COLUMN LOWARTSKUP;")
                AktSQL("Obroty", "ALTER TABLE Obroty DROP COLUMN LOILOSKUP;")
                AktSQL("Obroty", "ALTER TABLE Obroty DROP COLUMN AOWARTSKUP;")
                AktSQL("Obroty", "ALTER TABLE Obroty DROP COLUMN AOILOSKUP;")

                AktSQL("Obroty", "ALTER TABLE Obroty ADD LMARSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Obroty", "ALTER TABLE Obroty ADD AMARSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Obroty", "ALTER TABLE Obroty ADD LOMARSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Obroty", "ALTER TABLE Obroty ADD AOMARSPRZED FLOAT DEFAULT 0 WITH VALUES;")




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
                'Public oWersjaMies As Integer          'WERSJA           Integer

                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies DROP COLUMN LWARTSKUP;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies DROP COLUMN LILOSKUP;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies DROP COLUMN AWARTSKUP;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies DROP COLUMN AILOSKUP;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies DROP COLUMN LOWARTSKUP;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies DROP COLUMN LOILOSKUP;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies DROP COLUMN AOWARTSKUP;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies DROP COLUMN AOILOSKUP;")

                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD LMARSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD AMARSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD LOMARSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD AOMARSPRZED FLOAT DEFAULT 0 WITH VALUES;")

            End If




            If AktualnaWersja <= 271 Then
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN NRTOFITXT VARCHAR(500);")

                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN KODPOCZTOWY VARCHAR(10);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN NRDOMU VARCHAR(10);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN IDDOMU VARCHAR(10);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN NIP VARCHAR(15);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN KTOWPISAL VARCHAR(10);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN REJKLIENTA VARCHAR(20);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN PKONTAKT VARCHAR(10);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SKUPOPIEKUN VARCHAR(10);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SKUPOKONTAKTD VARCHAR(12);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SKUPNKONTAKTD VARCHAR(12);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDOPIEKUN VARCHAR(10);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDOKONTAKTD VARCHAR(12);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDNKONTAKTD VARCHAR(12);")
                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN SPRZEDWAGA VARCHAR(20);")
                '--------------------------------------
                InicjuPolaKK271()
            End If


            If AktualnaWersja <= 300 Then
                'Public Par_WartGranWartosc As Double = ""
                'Public Par_WartGranProcent As Double = ""
                AktSQL("Parametry", "ALTER TABLE Parametry ADD WARTGRANWARTOSC FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Parametry", "ALTER TABLE Parametry ADD WARTGRANPROCENT FLOAT DEFAULT 0 WITH VALUES;")

                'KatalogKlientow
                'Public oPrognoza01 As Double       'PROGNOZA01        Double
                'Public oPrognoza02 As Double       'PROGNOZA02        Double
                'Public oPrognoza03 As Double       'PROGNOZA03        Double
                'Public oPrognozaK1 As Double       'PROGNOZAK1        Double
                'Public oPrognoza04 As Double       'PROGNOZA04        Double
                'Public oPrognoza05 As Double       'PROGNOZA05        Double
                'Public oPrognoza06 As Double       'PROGNOZA06        Double
                'Public oPrognozaK2 As Double       'PROGNOZAK2        Double
                'Public oPrognozaP1 As Double       'PROGNOZAP1        Double
                'Public oPrognoza07 As Double       'PROGNOZA07        Double
                'Public oPrognoza08 As Double       'PROGNOZA08        Double
                'Public oPrognoza09 As Double       'PROGNOZA09        Double
                'Public oPrognozaK3 As Double       'PROGNOZAK3        Double
                'Public oPrognoza10 As Double       'PROGNOZA10        Double
                'Public oPrognoza11 As Double       'PROGNOZA11        Double
                'Public oPrognoza12 As Double       'PROGNOZA12        Double
                'Public oPrognozaK4 As Double       'PROGNOZAK4        Double
                'Public oPrognozaP2 As Double       'PROGNOZAP2        Double
                'Public oPrognozaRR As Double       'PROGNOZARR        Double
                'Public oPrognozaRRPLAN As Double   'PROGNOZARRPLAN    Double

                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZA01 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZA02 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZA03 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZAK1 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZA04 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZA05 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZA06 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZAK2 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZAP1 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZA07 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZA08 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZA09 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZAK3 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZA10 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZA11 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZA12 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZAK4 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZAP2 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZARR FLOAT DEFAULT 0 WITH VALUES;")
            End If



            If AktualnaWersja <= 301 Then
                'Public Par_MiesPrognozy As string = ""
                AktSQL("Parametry", "ALTER TABLE Parametry ADD MIESPROGNOZY VARCHAR(12) DEFAULT 'TTTTTTTTTTTT' WITH VALUES;")

                'KatalogKlientow
                'Public oPrognozaRRPLAN As Double   'PROGNOZARRPLAN    Double
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROGNOZARRPLAN FLOAT DEFAULT 0 WITH VALUES;")

            End If

            If AktualnaWersja <= 302 Then
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

                AktSQL("Obroty", "ALTER TABLE Obroty ADD NAJEMWARTSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Obroty", "ALTER TABLE Obroty ADD NAJEMILOSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Obroty", "ALTER TABLE Obroty ADD NAJEMMARSPRZED FLOAT DEFAULT 0 WITH VALUES;")

                AktSQL("Obroty", "ALTER TABLE Obroty ADD STRONYWARTSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Obroty", "ALTER TABLE Obroty ADD STRONYILOSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Obroty", "ALTER TABLE Obroty ADD STRONYMARSPRZED FLOAT DEFAULT 0 WITH VALUES;")

                AktSQL("Obroty", "ALTER TABLE Obroty ADD DRUKARKIWARTSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Obroty", "ALTER TABLE Obroty ADD DRUKARKIILOSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Obroty", "ALTER TABLE Obroty ADD DRUKARKIMARSPRZED FLOAT DEFAULT 0 WITH VALUES;")

                AktSQL("Obroty", "ALTER TABLE Obroty ADD SKUPWARTSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Obroty", "ALTER TABLE Obroty ADD SKUPILOSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Obroty", "ALTER TABLE Obroty ADD SKUPMARSPRZED FLOAT DEFAULT 0 WITH VALUES;")





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

                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD NAJEMWARTSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD NAJEMILOSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD NAJEMMARSPRZED FLOAT DEFAULT 0 WITH VALUES;")

                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD STRONYWARTSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD STRONYILOSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD STRONYMARSPRZED FLOAT DEFAULT 0 WITH VALUES;")

                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD DRUKARKIWARTSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD DRUKARKIILOSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD DRUKARKIMARSPRZED FLOAT DEFAULT 0 WITH VALUES;")

                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD SKUPWARTSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD SKUPILOSPRZED FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("ObrotyMies", "ALTER TABLE ObrotyMies ADD SKUPMARSPRZED FLOAT DEFAULT 0 WITH VALUES;")



                ''Nowa kategoria klienta
                'Public oKatKlientaPRYZMATKli As Boolean   'KATKLIENTAPRYZMAT    Logical
                'Public oKatKlientaORYGKli As Boolean      'KATKLIENTAORYG       Logical
                'Public oKatKlientaNAJEMKli As Boolean     'KATKLIENTANAJEM      Logical
                'Public oKatKlientaSTRONYKli As Boolean    'KATKLIENTASTRONY     Logical
                'Public oKatKlientaDRUKARKIKli As Boolean  'KATKLIENTADRUKARKI   Logical
                AktSQL("Klienci", "ALTER TABLE Klienci ADD KATKLIENTAPRYZMAT BIT DEFAULT 0 WITH VALUES")     'FALSE=0 / True=-1
                AktSQL("Klienci", "ALTER TABLE Klienci ADD KATKLIENTAORYG BIT DEFAULT 0 WITH VALUES")     'FALSE=0 / True=-1
                AktSQL("Klienci", "ALTER TABLE Klienci ADD KATKLIENTANAJEM BIT DEFAULT 0 WITH VALUES")     'FALSE=0 / True=-1
                AktSQL("Klienci", "ALTER TABLE Klienci ADD KATKLIENTASTRONY BIT DEFAULT 0 WITH VALUES")     'FALSE=0 / True=-1
                AktSQL("Klienci", "ALTER TABLE Klienci ADD KATKLIENTADRUKARKI BIT DEFAULT 0 WITH VALUES")     'FALSE=0 / True=-1

                'Public oKatOpisPRYZMATKli As String      'KATOPISPRYZMAT       Memo
                'Public oKatOpisORYGKli As String         'KATOPISORYG          Memo
                'Public oKatOpisNAJEMKli As String        'KATOPISNAJEM         Memo
                'Public oKatOpisSTRONYKli As String       'KATOPISSTRONY        Memo
                'Public oKatOpisDRUKARKIKli As String     'KATOPISDRUKARKI      Memo
                AktSQL("Klienci", "ALTER TABLE Klienci ADD KATOPISPRYZMAT [text] COLLATE Polish_CI_AS NOT NULL DEFAULT '' WITH VALUES")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD KATOPISORYG [text] COLLATE Polish_CI_AS NOT NULL DEFAULT '' WITH VALUES")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD KATOPISNAJEM [text] COLLATE Polish_CI_AS NOT NULL DEFAULT '' WITH VALUES")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD KATOPISSTRONY [text] COLLATE Polish_CI_AS NOT NULL DEFAULT '' WITH VALUES")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD KATOPISDRUKARKI [text] COLLATE Polish_CI_AS NOT NULL DEFAULT '' WITH VALUES")

                'Pobranie danych do kolumny OPISPRYZMAT
                AktSQL("Klienci", "Update [dbo].[Klienci] SET [KATOPISPRYZMAT] = [ILOSPRZETU2] WHERE Convert(varchar, KATOPISPRYZMAT) =''")

                'dr.Item("POTENCJALPPR") = dr.Item("POTENCJALPR")    'PrzedPoprzedni=Poprzedni
                'dr.Item("POTENCJALPR") = dr.Item("POTENCJALBR")     'Poprzedni=Bie¿¹cy
                'dr.Item("POTENCJALBR") = ""

                'dr.Item("ROKOWANIAPPR") = dr.Item("ROKOWANIAPR")    'PrzedPoprzedni=Poprzedni
                'dr.Item("ROKOWANIAPR") = dr.Item("ROKOWANIABR")     'Poprzedni=Bie¿¹cy
                'dr.Item("ROKOWANIABR") = ""







                ''prognozowanie sprzeda¿y uslug
                'Public oProUslugi01 As Double       'ProUslugi01        Double
                'Public oProUslugi02 As Double       'ProUslugi02        Double
                'Public oProUslugi03 As Double       'ProUslugi03        Double
                'Public oProUslugiK1 As Double       'ProUslugiK1        Double
                'Public oProUslugi04 As Double       'ProUslugi04        Double
                'Public oProUslugi05 As Double       'ProUslugi05        Double
                'Public oProUslugi06 As Double       'ProUslugi06        Double
                'Public oProUslugiK2 As Double       'ProUslugiK2        Double
                'Public oProUslugiP1 As Double       'ProUslugiP1        Double
                'Public oProUslugi07 As Double       'ProUslugi07        Double
                'Public oProUslugi08 As Double       'ProUslugi08        Double
                'Public oProUslugi09 As Double       'ProUslugi09        Double
                'Public oProUslugiK3 As Double       'ProUslugiK3        Double
                'Public oProUslugi10 As Double       'ProUslugi10        Double
                'Public oProUslugi11 As Double       'ProUslugi11        Double
                'Public oProUslugi12 As Double       'ProUslugi12        Double
                'Public oProUslugiK4 As Double       'ProUslugiK4        Double
                'Public oProUslugiP2 As Double       'ProUslugiP2        Double
                'Public oProUslugiRR As Double       'ProUslugiRR        Double
                'Public oProUslugiRRPlan As Double       'ProUslugiRRPlan        Double

                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGIRR FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGIRRPLAN FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGIP2 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGIP1 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGIK1 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGIK2 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGIK3 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGIK4 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGI12 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGI11 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGI10 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGI09 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGI08 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGI07 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGI06 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGI05 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGI04 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGI03 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGI02 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROUSLUGI01 FLOAT DEFAULT 0 WITH VALUES;")

                'Public oProSkupRRPlan As Double       'ProSkupRRPlan        Double
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PROSKUPRRPLAN FLOAT DEFAULT 0 WITH VALUES;")

                'Public Par_DaneDoPrognozy As String = ""
                AktSQL("Parametry", "ALTER TABLE Parametry ADD DANEDOPROGNOZY VARCHAR(10) NOT NULL DEFAULT '' WITH VALUES;")
                'Public Par_WartGranMatWartosc As Double = ""
                'Public Par_WartGranMatProcent As Double = ""
                AktSQL("Parametry", "ALTER TABLE Parametry ADD WARTGRANMATWARTOSC FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("Parametry", "ALTER TABLE Parametry ADD WARTGRANMATPROCENT FLOAT DEFAULT 0 WITH VALUES;")


            End If



            If AktualnaWersja <= 303 Then
                AktSQL("Parametry", "ALTER TABLE Parametry ALTER COLUMN DANEDOANALIZY CHAR(20);")

                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA00 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZABM FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA01 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA02 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA03 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA04 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA05 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA06 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA07 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA08 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA09 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA10 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA11 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA12 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA13 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA14 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA15 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA16 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA17 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA18 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA19 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA20 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA21 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA22 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA23 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD MARZA24 FLOAT DEFAULT 0 WITH VALUES;")

                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM00 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCMBM FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM01 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM02 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM03 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM04 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM05 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM06 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM07 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM08 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM09 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM10 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM11 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM12 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM13 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM14 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM15 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM16 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM17 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM18 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM19 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM20 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM21 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM22 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM23 FLOAT DEFAULT 0 WITH VALUES;")
                AktSQL("AnalizaZakupu", "ALTER TABLE AnalizyZakupu ADD PROCM24 FLOAT DEFAULT 0 WITH VALUES;")

            End If


            If AktualnaWersja <= 400 Then

                ''---------------------------------------------------------------------
                ''Osoby Kontaktowe
                ''---------------------------------------------------------------------
                'Public oIdentOso As String            'IDENTKLIENTA     Text, 6
                'Public oOsobaOso As String            'OSOBA            Text, 100
                'Public oStanowiskoOso As String       'STANOWISKO       Text, 100
                'Public oKompetencjeOso As String      'KOMPETENCJE      Text, 100
                'Public oTelefonOso As String          'TELEFON          Text, 100
                'Public oTelefonKomOso As String       'TELEFONKOM       Text, 100
                'Public oeMailOso As String            'EMAIL            Text, 100
                'Public oVIPOso As Boolean             'VIP              boolean
                'Public oWersjaOso As Integer          'WERSJA           Integer
                AktSQL("Osoby", "ALTER TABLE Osoby ALTER COLUMN STANOWISKO VARCHAR(100);")
                AktSQL("Osoby", "ALTER TABLE Osoby ALTER COLUMN KOMPETENCJE VARCHAR(100);")
                AktSQL("Osoby", "ALTER TABLE Osoby ALTER COLUMN TELEFON VARCHAR(100);")
                AktSQL("Osoby", "ALTER TABLE Osoby ADD TELEFONKOM VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Osoby", "ALTER TABLE Osoby ALTER COLUMN EMAIL VARCHAR(100);")
                'zmiana d³ugosci PRIMARY KEY
                AktSQL("Osoby", "ALTER TABLE Osoby DROP CONSTRAINT PK_Osoby;")
                AktSQL("Osoby", "ALTER TABLE Osoby ALTER COLUMN OSOBA VARCHAR(100) COLLATE Polish_CI_AS NOT NULL;")
                AktSQL("Osoby", "ALTER TABLE Osoby WITH NOCHECK ADD CONSTRAINT PK_Osoby PRIMARY KEY CLUSTERED (IDENTKLIENTA,OSOBA);")

                ''---------------------------------------------------------------------
                ''Katalog Klientow
                ''---------------------------------------------------------------------

                'Ustaw kolumny NULLable
                If Not CzyJestKolumnaWTabeli("Klienci", "WYKAZURZADZEN") Then
                    UstawKolumnyNULL()

                    'Zainicjowanie wartoœci pola RodzajSprzetuT (obecnie A3) na false
                    'Public oRodzSprzetuLKli As Boolean 'RODZSPRZETUL Logical
                    'Public oRodzSprzetuAKli As Boolean 'RODZSPRZETUA Logical
                    'Public oRodzSprzetuTKli As Boolean 'RODZSPRZETUT Logical
                    AktSQL("Klienci", "UPDATE Klienci SET RODZSPRZETUT = 'FALSE' ")
                End If

                'Public oWYKAZURZADZENKli As String      'WYKAZURZADZEN       Memo
                If Not CzyJestKolumnaWTabeli("Klienci", "WYKAZURZADZEN") Then
                    'AktSQL("Klienci", "ALTER TABLE Klienci ADD WYKAZURZADZEN [text] COLLATE Polish_CI_AS NOT NULL DEFAULT '' WITH VALUES;")
                    AktSQL("Klienci", "ALTER TABLE Klienci ADD WYKAZURZADZEN [text] COLLATE Polish_CI_AS NULL DEFAULT '' WITH VALUES;")
                    'aktualizacja wartoœci - pobranie zawartosci pól KatOpisPRYZMAT,KatOpisORYG,KatOpisNAJEM,KatOpisSTRONY,KatOpisDRUKARKI 

                    AktSQL("Klienci", "UPDATE Klienci " &
                                          "   SET WYKAZURZADZEN = CAST(WYKAZURZADZEN as nvarchar(1000)) + 'PRYZMAT : ' + CAST(KATOPISPRYZMAT as nvarchar(1000)) + char(10) " &
                                          "WHERE Len(cast(KATOPISPRYZMAT As nvarchar(1000))) > 0 ")
                    AktSQL("Klienci", "UPDATE Klienci " &
                                          "   SET WYKAZURZADZEN = CAST(WYKAZURZADZEN as nvarchar(1000)) + 'ORYGINA£Y : ' + CAST(KATOPISORYG as nvarchar(1000)) + char(10) " &
                                          "WHERE Len(cast(KATOPISORYG As nvarchar(1000))) > 0 ")
                    AktSQL("Klienci", "UPDATE Klienci " &
                                          "   SET WYKAZURZADZEN = CAST(WYKAZURZADZEN as nvarchar(1000)) + 'NAJEM : ' + CAST(KATOPISNAJEM as nvarchar(1000)) + char(10) " &
                                          "WHERE Len(cast(KATOPISNAJEM As nvarchar(1000))) > 0 ")
                    AktSQL("Klienci", "UPDATE Klienci " &
                                          "   SET WYKAZURZADZEN = CAST(WYKAZURZADZEN as nvarchar(1000)) + 'STRONY : ' + CAST(KATOPISSTRONY as nvarchar(1000)) + char(10) " &
                                          "WHERE Len(cast(KATOPISSTRONY As nvarchar(1000))) > 0 ")
                    AktSQL("Klienci", "UPDATE Klienci " &
                                          "   SET WYKAZURZADZEN = CAST(WYKAZURZADZEN as nvarchar(1000)) + 'DRUKARKI : ' + CAST(KATOPISDRUKARKI as nvarchar(1000)) + char(10) " &
                                          "WHERE Len(cast(KATOPISDRUKARKI As nvarchar(1000))) > 0 ")
                End If

                'Public oSposobWyboruDostawcyKli As String       'SPOSOBWYBORUDOSTAWCY    Text 30
                If Not CzyJestKolumnaWTabeli("Klienci", "SPOSOBWYBORUDOSTAWCY") Then
                    AktSQL("Klienci", "ALTER TABLE Klienci ADD SPOSOBWYBORUDOSTAWCY VARCHAR(30) DEFAULT '' WITH VALUES;")
                    'aktualizacja wartoœci - pobranie zawartosci pola Public oSkupOpiekunKli As String  SKUPOPIEKUN    Text, 10   
                    'jeœli w polu "Opis Potencja³u" (oSkupUwagikli/SKUPUWAGI) wystepuje tekst "%SEG!%"
                    AktSQL("Klienci", "UPDATE Klienci SET SPOSOBWYBORUDOSTAWCY = SUBSTRING(SKUPOPIEKUN,1,30) WHERE (SKUPUWAGI LIKE '%SEG!%') ")
                End If

                'Public oZainstalowanoMonitorKli As Boolean      'ZAINSTALOWANOMONITOR    Logical
                If Not CzyJestKolumnaWTabeli("Klienci", "ZAINSTALOWANOMONITOR") Then
                    AktSQL("Klienci", "ALTER TABLE Klienci ADD ZAINSTALOWANOMONITOR BIT DEFAULT 0 WITH VALUES;")     'FALSE=0 / True=-1
                    'aktualizacja wartoœci - pobranie zawartosci pola     Public oSprzedTestKli As Boolean      'SPRZEDTEST       Logical
                    'jeœli w polu "Opis Potencja³u" (oSkupUwagikli/SKUPUWAGI) wystepuje tekst "%SEG!%"
                    AktSQL("Klienci", "UPDATE Klienci SET ZAINSTALOWANOMONITOR = SPRZEDTEST WHERE (SKUPUWAGI LIKE '%SEG!%') ")
                End If

                'Public oLiczbaUrzadzenKli As Integer            'LICZBAURZADZEN     Integer
                If Not CzyJestKolumnaWTabeli("Klienci", "LICZBAURZADZEN") Then
                    AktSQL("Klienci", "ALTER TABLE Klienci ADD LICZBAURZADZEN INTEGER DEFAULT 0 WITH VALUES;")
                    'aktualizacja wartoœci - pobranie zawartosci pola     Public oSprzedWartoscKli As String    'SPRZEDWARTOSC    Text, 50
                    'jeœli w polu "Opis Potencja³u" (oSkupUwagikli/SKUPUWAGI) wystepuje tekst "%SEG!%"
                    'ustal domyœln¹ wartoœæ na podst pola ILOSC SPRZETU (KATKLIENAA / KATKLIENTAB / KATKLIENTAC)
                    AktSQL("Klienci", "UPDATE Klienci SET LICZBAURZADZEN = CASE WHEN KATKLIENTAA='TRUE' THEN 5 " &
                                                                           "    WHEN KATKLIENTAB='TRUE' THEN 2 " &
                                                                           "    WHEN KATKLIENTAC='TRUE' THEN 1 " &
                                                                           " 	ELSE 0 " &
                                                                           "END ")
                    'uzupe³nij iloœci predefiniowane w polu CZYSZCZENIE (SPRZEDAWARTOSC)
                    'AktSQL("Klienci", "UPDATE Klienci SET LICZBAURZADZEN = iif(TRY_CAST(SPRZEDWARTOSC AS INTEGER) IS NULL,0,cast(SPRZEDWARTOSC AS integer)) WHERE (SKUPUWAGI LIKE '%SEG!%') ")

                    'UPDATE Klienci SET LICZBAURZADZEN = (CASE WHEN ISNUMERIC(SPRZEDWARTOSC)=1 THEN cast(SPRZEDWARTOSC AS INTEGER) Else 0 End) WHERE(SKUPUWAGI Like '%SEG!%')
                    AktSQL("Klienci", "UPDATE Klienci SET LICZBAURZADZEN = (CASE WHEN ISNUMERIC(SPRZEDWARTOSC)=1 THEN cast(SPRZEDWARTOSC AS INTEGER) Else 0 End) WHERE(SKUPUWAGI Like '%SEG!%') ")
                End If

                'Public oLiczbaWydrukowKli As Integer            'LICZBAWYDRUKOW     Integer
                If Not CzyJestKolumnaWTabeli("Klienci", "LICZBAWYDRUKOW") Then
                    AktSQL("Klienci", "ALTER TABLE Klienci ADD LICZBAWYDRUKOW INTEGER DEFAULT 0 WITH VALUES;")
                    'aktualizacja wartoœci - pobranie zawartosci pola         Public oSprzedWagaKli As String       'SPRZEDWAGA       TEXT, 20
                    'jeœli w polu "Opis Potencja³u" (oSkupUwagikli/SKUPUWAGI) wystepuje tekst "%SEG!%"
                    'AktSQL("Klienci", "UPDATE Klienci SET LICZBAWYDRUKOW = iif(TRY_CAST(SPRZEDWAGA AS INTEGER) IS NULL,0,cast(SPRZEDWAGA AS integer)) WHERE (SKUPUWAGI LIKE '%SEG!%') ")

                    'UPDATE Klienci SET LICZBAURZADZEN = (CASE WHEN ISNUMERIC(SPRZEDWARTOSC)=1 THEN cast(SPRZEDWARTOSC AS INTEGER) Else 0 End) WHERE(SKUPUWAGI Like '%SEG!%')
                    AktSQL("Klienci", "UPDATE Klienci SET LICZBAWYDRUKOW = (CASE WHEN ISNUMERIC(SPRZEDWAGA)=1 THEN cast(SPRZEDWAGA AS INTEGER) Else 0 End) WHERE(SKUPUWAGI Like '%SEG!%') ")
                End If

                'Public oPotencjalDrukuKli As String             'POTENCJALDRUKU     Text 2
                If Not CzyJestKolumnaWTabeli("Klienci", "POTENCJALDRUKU") Then
                    AktSQL("Klienci", "ALTER TABLE Klienci ADD POTENCJALDRUKU VARCHAR(2) DEFAULT ' ' WITH VALUES;")

                    AktSQL("Klienci", "Update Klienci SET POTENCJALDRUKU = CASE WHEN (LICZBAWYDRUKOW = 0) AND (LICZBAURZADZEN = 0) then '' " &
                                                                               "WHEN LICZBAWYDRUKOW > 0 then " &
                                                                                    "(CASE WHEN LICZBAWYDRUKOW<15 THEN 'W0' " &
                                                                                    "      WHEN LICZBAWYDRUKOW<50 THEN 'W1' " &
                                                                                    "      WHEN LICZBAWYDRUKOW<100 THEN 'W2' " &
                                                                                    "      WHEN LICZBAWYDRUKOW<200 THEN 'W3' " &
                                                                                    "      ELSE 'W4' " &
                                                                                    "END) " &
                                                                               "Else (CASE WHEN LICZBAURZADZEN<=2 THEN 'U1' " &
                                                                                    "      WHEN LICZBAURZADZEN<=5 THEN 'U2' " &
                                                                                    "      WHEN LICZBAURZADZEN<=24 THEN 'U3' " &
                                                                                    "      ELSE 'U4' " &
                                                                                    "END) " &
                                                                          "End; ")
                End If

                'Public oAktZakupMatEkspKli As Boolean           'AKTZAKUPMATEKSP    Logical
                If Not CzyJestKolumnaWTabeli("Klienci", "AKTZAKUPMATEKSP") Then
                    AktSQL("Klienci", "ALTER TABLE Klienci ADD AKTZAKUPMATEKSP BIT DEFAULT 0 WITH VALUES;")     'FALSE=0 / True=-1
                End If

                'Public oAktRozliczaStronyDrukuKli As Boolean    'AKTROZLICZASTRONYDRUKU    Logical
                If Not CzyJestKolumnaWTabeli("Klienci", "AKTROZLICZASTRONYDRUKU") Then
                    AktSQL("Klienci", "ALTER TABLE Klienci ADD AKTROZLICZASTRONYDRUKU BIT DEFAULT 0 WITH VALUES;")     'FALSE=0 / True=-1
                    'aktualizacja wartoœci - pobranie zawartosci pola     Public oRokowaniaBR As String           'RokowaniaBR   Text, 4
                    'jeœli w polu "Opis Potencja³u" (oSkupUwagikli/SKUPUWAGI) wystepuje tekst "%SEG!%"
                    AktSQL("Klienci", "UPDATE Klienci SET AKTROZLICZASTRONYDRUKU = (CASE WHEN ROKOWANIABR IN ('T','TAK','Y','TES') THEN 'TRUE' ELSE 'FALSE' END) WHERE (SKUPUWAGI LIKE '%SEG!%') ")
                End If

                'Public oAktKorzystaZNajmuKli As Boolean         'AKTKORZYSTAZNAJMU  Logical
                If Not CzyJestKolumnaWTabeli("Klienci", "AKTKORZYSTAZNAJMU") Then
                    AktSQL("Klienci", "ALTER TABLE Klienci ADD AKTKORZYSTAZNAJMU BIT DEFAULT 0 WITH VALUES;")     'FALSE=0 / True=-1
                    'aktualizacja wartoœci - pobranie zawartosci pola     Public oPotencjalBR As String           'POTENCJALBR   Text, 4
                    'jeœli w polu "Opis Potencja³u" (oSkupUwagikli/SKUPUWAGI) wystepuje tekst "%SEG!%"
                    AktSQL("Klienci", "UPDATE Klienci SET AKTKORZYSTAZNAJMU = (CASE WHEN POTENCJALBR IN ('T','TAK','Y','TES') THEN 'TRUE' ELSE 'FALSE' END) WHERE (SKUPUWAGI LIKE '%SEG!%') ")
                End If

                'Public oZaintRozliczaniemStronDrukuKli As Boolean   'ZAINTROZLICZANIEMSTRONDRUKU    Logical
                If Not CzyJestKolumnaWTabeli("Klienci", "ZAINTROZLICZANIEMSTRONDRUKU") Then
                    AktSQL("Klienci", "ALTER TABLE Klienci ADD ZAINTROZLICZANIEMSTRONDRUKU BIT DEFAULT 0 WITH VALUES;")     'FALSE=0 / True=-1
                    'aktualizacja wartoœci - pobranie zawartosci pola     Public oRokowaniaPR As String           'RokowaniaPR   Text, 4
                    'jeœli w polu "Opis Potencja³u" (oSkupUwagikli/SKUPUWAGI) wystepuje tekst "%SEG!%"
                    AktSQL("Klienci", "UPDATE Klienci SET ZAINTROZLICZANIEMSTRONDRUKU = (CASE WHEN ROKOWANIAPR IN ('T','TAK','Y','TES') THEN 'TRUE' ELSE 'FALSE' END) WHERE (SKUPUWAGI LIKE '%SEG!%') ")
                End If

                'Public oZaintKorzystaniemZNajmuKli As Boolean   'ZAINTKORZYSTANIEMZNAJMU    Logical
                If Not CzyJestKolumnaWTabeli("Klienci", "ZAINTKORZYSTANIEMZNAJMU") Then
                    AktSQL("Klienci", "ALTER TABLE Klienci ADD ZAINTKORZYSTANIEMZNAJMU BIT DEFAULT 0 WITH VALUES;")     'FALSE=0 / True=-1
                    'aktualizacja wartoœci - pobranie zawartosci pola     Public oPotencjalPR As String           'POTENCJALPR   Text, 4
                    'jeœli w polu "Opis Potencja³u" (oSkupUwagikli/SKUPUWAGI) wystepuje tekst "%SEG!%"
                    AktSQL("Klienci", "UPDATE Klienci SET ZAINTKORZYSTANIEMZNAJMU = (CASE WHEN POTENCJALPR IN ('T','TAK','Y','TES') THEN 'TRUE' ELSE 'FALSE' END) WHERE (SKUPUWAGI LIKE '%SEG!%') ")
                End If

                'Public oDataWeryfSegmentacji As String          'DATAWERYFSEGMENTACJI  Text 10
                If Not CzyJestKolumnaWTabeli("Klienci", "DATAWERYFSEGMENTACJI") Then
                    AktSQL("Klienci", "ALTER TABLE Klienci ADD DATAWERYFSEGMENTACJI VARCHAR(10) DEFAULT '' WITH VALUES;")
                End If

                'Public oPlanDlugookresowyKli As Integer         'PLANDLUGOOKRESOWY     Integer
                If Not CzyJestKolumnaWTabeli("Klienci", "PLANDLUGOOKRESOWY") Then
                    AktSQL("Klienci", "ALTER TABLE Klienci ADD PLANDLUGOOKRESOWY INTEGER DEFAULT 0 WITH VALUES;")
                End If

                'Public oPlanKrotkookresowyKli As Integer        'PLANKROTKOOKRESOWY    Integer
                If Not CzyJestKolumnaWTabeli("Klienci", "PLANKROTKOOKRESOWY") Then
                    AktSQL("Klienci", "ALTER TABLE Klienci ADD PLANKROTKOOKRESOWY INTEGER DEFAULT 0 WITH VALUES;")
                End If
                'Public oAktywnyKli As Boolean         'AKTYWNY          boolean
                If Not CzyJestKolumnaWTabeli("Klienci", "AKTYWNY") Then
                    AktSQL("Klienci", "ALTER TABLE Klienci ADD AKTYWNY BIT DEFAULT 1 WITH VALUES;")     'FALSE=0 / True=-1
                End If


            End If






            If AktualnaWersja <= 500 Then
                'usuwamy nieuzywane kolumny...

                'usuwamy tabele KONTAKTY
                If CzyJestTabela("Kontakty") Then AktSQL("Kontakty", "DROP TABLE Kontakty;")

                'usuwamy tabele HISTORIAZAKUPI
                If CzyJestTabela("HistoriaZakupu") Then AktSQL("Historiazakupu", "DROP TABLE HistoriaZakupu;")

                'usuwamy wszystkie CONSTRAINTS oprócz PRIMARY KEY
                If CzyJestTabela("Klienci") Then AktSQL("Klienci", "Declare @sql NVARCHAR(MAX) = N'';" &
                                " Select @sql += N'ALTER TABLE ' + QUOTENAME(OBJECT_SCHEMA_NAME(parent_object_id)) " &
                                       "+ '.' + QUOTENAME(OBJECT_NAME(parent_object_id)) + ' DROP CONSTRAINT ' + QUOTENAME(name) + " & "';' " &
                "From sys.objects WHERE (type_desc Like '%CONSTRAINT') And (type_desc<>'PRIMARY_KEY_CONSTRAINT') And (OBJECT_NAME(PARENT_OBJECT_ID) Like '%Klienci'); " &
                "EXEC sp_executesql @sql; ")

                If CzyJestKolumnaWTabeli("Klienci", "WWW") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN WWW;")
                If CzyJestKolumnaWTabeli("Klienci", "OSOBAKONTAKTOWA") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN OSOBAKONTAKTOWA;")
                If CzyJestKolumnaWTabeli("Klienci", "KATKLIENTAA") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN KATKLIENTAA;")
                If CzyJestKolumnaWTabeli("Klienci", "KATKLIENTAB") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN KATKLIENTAB;")
                If CzyJestKolumnaWTabeli("Klienci", "KATKLIENTAC") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN KATKLIENTAC;")
                If CzyJestKolumnaWTabeli("Klienci", "OFERTA") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN OFERTA;")
                If CzyJestKolumnaWTabeli("Klienci", "GDZIEKUPUJE") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN GDZIEKUPUJE;")
                If CzyJestKolumnaWTabeli("Klienci", "ZUZYCIE") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN ZUZYCIE;")
                If CzyJestKolumnaWTabeli("Klienci", "SKUPOPIEKUN") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN SKUPOPIEKUN;")
                If CzyJestKolumnaWTabeli("Klienci", "SKUPWARTOSC") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN SKUPWARTOSC;")
                If CzyJestKolumnaWTabeli("Klienci", "SKUPOKONTAKTR") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN SKUPOKONTAKTR;")
                If CzyJestKolumnaWTabeli("Klienci", "SKUPOKONTAKTT") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN SKUPOKONTAKTT;")
                If CzyJestKolumnaWTabeli("Klienci", "SKUPOKONTAKTD") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN SKUPOKONTAKTD;")
                If CzyJestKolumnaWTabeli("Klienci", "SKUPNKONTAKTR") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN SKUPNKONTAKTR;")
                If CzyJestKolumnaWTabeli("Klienci", "SKUPNKONTAKTT") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN SKUPNKONTAKTT;")
                If CzyJestKolumnaWTabeli("Klienci", "SKUPNKONTAKTD") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN SKUPNKONTAKTD;")
                If CzyJestKolumnaWTabeli("Klienci", "SKUPPLAN") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN SKUPPLAN;")
                If CzyJestKolumnaWTabeli("Klienci", "SPRZEDWARTOSC") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN SPRZEDWARTOSC;")
                If CzyJestKolumnaWTabeli("Klienci", "SPRZEDTEST") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN SPRZEDTEST;")
                If CzyJestKolumnaWTabeli("Klienci", "SPRZEDPLAN") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN SPRZEDPLAN;")
                If CzyJestKolumnaWTabeli("Klienci", "SPRZEDWAGA") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN SPRZEDWAGA;")
                If CzyJestKolumnaWTabeli("Klienci", "DZIALANIAAKCYJNE") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN DZIALANIAAKCYJNE;")
                If CzyJestKolumnaWTabeli("Klienci", "POTENCJALBR") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN POTENCJALBR;")
                If CzyJestKolumnaWTabeli("Klienci", "POTENCJALPR") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN POTENCJALPR;")
                If CzyJestKolumnaWTabeli("Klienci", "POTENCJALPPR") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN POTENCJALPPR;")
                If CzyJestKolumnaWTabeli("Klienci", "ROKOWANIABR") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN ROKOWANIABR;")
                If CzyJestKolumnaWTabeli("Klienci", "ROKOWANIAPR") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN ROKOWANIAPR;")
                If CzyJestKolumnaWTabeli("Klienci", "ROKOWANIAPPR") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN ROKOWANIAPPR;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZA01") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZA01;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZA02") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZA02;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZA03") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZA03;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZAK1") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZAK1;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZA04") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZA04;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZA05") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZA05;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZA06") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZA06;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZAK2") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZAK2;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZAP1") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZAP1;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZA07") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZA07;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZA08") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZA08;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZA09") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZA09;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZAK3") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZAK3;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZA10") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZA10;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZA11") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZA11;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZA12") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZA12;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZAK4") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZAK4;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZAP2") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZAP2;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZARR") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZARR;")
                If CzyJestKolumnaWTabeli("Klienci", "PROGNOZARRPLAN") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROGNOZARRPLAN;")
                If CzyJestKolumnaWTabeli("Klienci", "KATKLIENTAPRYZMAT") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN KATKLIENTAPRYZMAT;")
                If CzyJestKolumnaWTabeli("Klienci", "KATKLIENTAORYG") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN KATKLIENTAORYG;")
                If CzyJestKolumnaWTabeli("Klienci", "KATKLIENTANAJEM") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN KATKLIENTANAJEM;")
                If CzyJestKolumnaWTabeli("Klienci", "KATKLIENTASTRONY") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN KATKLIENTASTRONY;")
                If CzyJestKolumnaWTabeli("Klienci", "KATKLIENTADRUKARKI") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN KATKLIENTADRUKARKI;")
                If CzyJestKolumnaWTabeli("Klienci", "KATOPISPRYZMAT") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN KATOPISPRYZMAT;")
                If CzyJestKolumnaWTabeli("Klienci", "KATOPISORYG") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN KATOPISORYG;")
                If CzyJestKolumnaWTabeli("Klienci", "KATOPISNAJEM") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN KATOPISNAJEM;")
                If CzyJestKolumnaWTabeli("Klienci", "KATOPISSTRONY") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN KATOPISSTRONY;")
                If CzyJestKolumnaWTabeli("Klienci", "KATOPISDRUKARKI") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN KATOPISDRUKARKI;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGI01") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGI01;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGI02") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGI02;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGI03") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGI03;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGIK1") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGIK1;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGI04") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGI04;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGI05") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGI05;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGI06") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGI06;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGIK2") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGIK2;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGIP1") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGIP1;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGI07") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGI07;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGI08") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGI08;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGI09") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGI09;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGIK3") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGIK3;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGI10") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGI10;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGI11") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGI11;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGI12") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGI12;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGIK4") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGIK4;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGIP2") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGIP2;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGIRR") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGIRR;")
                If CzyJestKolumnaWTabeli("Klienci", "PROUSLUGIRRPLAN") Then AktSQL("Klienci", "ALTER TABLE Klienci DROP COLUMN PROUSLUGIRRPLAN;")








            End If



            If AktualnaWersja <= 510 Then

                ''----------------------------------------------------------
                ''Katalog LogAktywnosci
                ''----------------------------------------------------------
                'Public LG_UniqID As String       'UNIQID     STRING 40
                'Public LG_Temat As String        'TEMAT      STRING 100
                'Public LG_Katalog As String      'KATALOG    STRING 100
                'Public LG_DataZapisu As String   'DATAZAPISU STRING 30
                'Public LG_Pracownik As String    'PRACOWNIK  STRING 10
                'Public LG_Operacja As String     'OPERACJA   STRING 20
                'Public LG_Uwagi As String        'UWAGI      Text
                'Public LG_Wersja As Integer      'WERSJA     Integer

                AktSQL("LogAktywnosci", "CREATE TABLE [dbo].[LogAktywnosci] (" & vbCrLf &
                                   "   [UNIQID] [varchar] (40) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                                   "   [TEMAT] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                                   "   [KATALOG] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                                   "   [DATAZAPISU] [varchar] (30) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                                   "   [PRACOWNIK] [varchar] (10) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                                   "   [OPERACJA] [varchar] (20) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                                   "   [UWAGI] [text] COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                                   "   [WERSJA] [int] NOT NULL " & vbCrLf &
                                   ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]")

                AktSQL("LogAktywnosci", "ALTER TABLE [dbo].[LogAktywnosci] WITH NOCHECK ADD " & vbCrLf &
                                   "   CONSTRAINT [PK_LogAktywnosci] PRIMARY KEY  CLUSTERED " & vbCrLf &
                                   "   (" & vbCrLf &
                                   "   [UNIQID] " & vbCrLf &
                                   "   )  ON [PRIMARY] ")


            End If


            If AktualnaWersja <= 520 Then

                '---------------------------------------------------------------------
                'Branze
                '---------------------------------------------------------------------
                'Public brIdBranzy As String            'IDBRANZY         Text, 
                'Public brWersja As Integer             'WERSJA           Integer

                AktSQL("Branze", "CREATE TABLE [dbo].[Branze] (" & vbCrLf &
                                   "   [IDBRANZY] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                                   "   [WERSJA] [int] NOT NULL " & vbCrLf &
                                   ") ON [PRIMARY] ")

                AktSQL("Branze", "ALTER TABLE [dbo].[Branze] WITH NOCHECK ADD " & vbCrLf &
                                   "   CONSTRAINT [PK_Branze] PRIMARY KEY  CLUSTERED " & vbCrLf &
                                   "   (" & vbCrLf &
                                   "   [IDBRANZY]" & vbCrLf &
                                   "   )  ON [PRIMARY] ")






                '---------------------------------------------------------------------
                'PodBranze
                '---------------------------------------------------------------------
                'Public pbrIdBranzy As String            'IDBRANZY         Text, 
                'Public pbrIdPodBranzy As String         'IDPODBRANZY         Text, 
                'Public pbrWersja As Integer             'WERSJA           Integer

                AktSQL("PodBranze", "CREATE TABLE [dbo].[PodBranze] (" & vbCrLf &
                                   "   [IDBRANZY] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                                   "   [IDpodBRANZY] [varchar] (100) COLLATE Polish_CI_AS NOT NULL ," & vbCrLf &
                                   "   [WERSJA] [int] NOT NULL " & vbCrLf &
                                   ") ON [PRIMARY] ")

                AktSQL("PodBranze", "ALTER TABLE [dbo].[PodBranze] WITH NOCHECK ADD " & vbCrLf &
                                   "   CONSTRAINT [PK_PodBranze] PRIMARY KEY  CLUSTERED " & vbCrLf &
                                   "   (" & vbCrLf &
                                   "   [IDBRANZY], " & vbCrLf &
                                   "   [IDPODBRANZY]" & vbCrLf &
                                   "   )  ON [PRIMARY] ")


                '---------------------------------------------------------------------
                'Klienci
                '---------------------------------------------------------------------
                'Public oBranzaKli As String        'BRANZA         Text, 100
                'Public oPodBranzaKli As String     'PODBRANZA      Text, 100
                'Public oLiczbaZaktrudnionychKli As Integer   'LICZBAZATRUDNIONYCH       INTEGER
                'Public oLiczbaPracZdalnieKli As Integer      'LICZBAPRACZDALNIE         INTEGER
                AktSQL("Klienci", "ALTER TABLE Klienci ADD BRANZA VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD PODBRANZA VARCHAR(100) DEFAULT '' WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD LICZBAZATRUDNIONYCH INTEGER DEFAULT 0 WITH VALUES;")
                AktSQL("Klienci", "ALTER TABLE Klienci ADD LICZBAPRACZDALNIE INTEGER DEFAULT 0 WITH VALUES;")
            End If














            '--Returns one row for each CHECK, UNIQUE, PRIMARY KEY, And/Or FOREIGN KEY
            'Select Case*
            '    From INFORMATION_SCHEMA.TABLE_CONSTRAINTS
            '    Where CONSTRAINT_NAME ='XYZ'  
            '
            '--Returns one row for each FOREIGN KEY constrain
            'Select Case*
            '    From INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS 
            '    Where CONSTRAINT_NAME ='XYZ'
            '
            '--Returns one row for each CHECK constraint 
            'Select Case*
            '    From INFORMATION_SCHEMA.CHECK_CONSTRAINTS
            '    Where CONSTRAINT_NAME ='XYZ'

            'sprawdz czy mamy PrimaryKey dla AkcjeSpec
            'Select Case count(*) From INFORMATION_SCHEMA.TABLE_CONSTRAINTS Where CONSTRAINT_NAME ='PK_AkcjeSpec'

            'If SQLExecuteScalarInt("Select Case count(*) From INFORMATION_SCHEMA.TABLE_CONSTRAINTS Where CONSTRAINT_NAME ='PK_AkcjeSpec'") = 0 Then
            '    AktSQL("AkcjeSpec", "ALTER TABLE [dbo].[AkcjeSpec] WITH NOCHECK ADD " & vbCrLf &
            '                   "   CONSTRAINT [PK_AkcjeSpec] PRIMARY KEY  CLUSTERED " & vbCrLf &
            '                   "   (" & vbCrLf &
            '                   "   [IDENTAKCJI], " & vbCrLf &
            '                   "   [IDENTKLI]" & vbCrLf &
            '                   "   )  ON [PRIMARY] ")
            'End If









            '--------------------------------------
            'aktualizacje numeru wersji bazy danych
            txtRaport.Text &= "Wykonujê aktualizacje numeru wersji" & vbCrLf
            txtRaport.SelectionStart = txtRaport.Text.Length
            txtRaport.ScrollToCaret()
            System.Windows.Forms.Application.DoEvents()

            Try
                sqlCommand.Connection = sqlConnectionKlienci
                sqlCommand.CommandType = CommandType.Text
                'sqlCommand.CommandText = sqlAlterKlienci210_1
                sqlCommand.CommandText = "/****** Object:  Stored Procedure dbo.softart_wersjaprogramu    Script Date: 2007-03-14 10:30:15 ******/" & vbCrLf &
                                                                 "/*  Podaje wersje programu" & vbCrLf &
                                                                 "*/ " & vbCrLf &
                                                                 "alter procedure dbo.softart_wersjaprogramu" & vbCrLf &
                                                                 "@wersja int output" & vbCrLf &
                                                                 "as " & vbCrLf &
                                                                 "	set @wersja = " & NowaWersja.ToString & vbCrLf &
                                                                 "  Return"
                sqlConnectionKlienci.Open()

                Wynik = sqlCommand.ExecuteNonQuery()
                If Wynik = -1 Then
                End If
                'txtRaport.Text &= "OK, aktualizacja wykonana poprawnie" & vbCrLf
            Catch ex As Exception
                'MessageBox.Show(ex.Message)
                txtRaport.Text &= "B³¹d wykonania : " & ex.Message & vbCrLf
                txtRaport.SelectionStart = txtRaport.Text.Length
                txtRaport.ScrollToCaret()
                System.Windows.Forms.Application.DoEvents()

            Finally
                sqlConnectionKlienci.Close()
            End Try

            AktualnaWersja = NowaWersja
            '--------------------------------------
            txtRaport.Text &= Now & "   KONIEC aktualizacji" & vbCrLf
            txtRaport.SelectionStart = txtRaport.Text.Length
            txtRaport.ScrollToCaret()
            System.Windows.Forms.Application.DoEvents()
        End If
    End Sub







    Public Function CzyJestTabela(ByVal NazwaTabeli As String) As Boolean
        Dim sqlConnection As SqlClient.SqlConnection
        Dim sqlCommand As SqlClient.SqlCommand
        Dim Wynik As Integer = 0

        'sprawdz czy istnieje kolumna...
        'SELECT count(*) FROM INFORMATION_SCHEMA.columns WHERE (table_name = 'Klienci') AND (column_name = 'KatKlientaPRYZMAT')

        sqlConnection = New SqlClient.SqlConnection(sConnectionStr)
        sqlCommand = New SqlClient.SqlCommand
        Try
            sqlCommand.Connection = sqlConnection
            sqlCommand.CommandType = CommandType.Text
            sqlCommand.CommandText = "SELECT count(*) FROM INFORMATION_SCHEMA.columns WHERE (table_name = '" & NazwaTabeli & "')"
            sqlConnection.Open()

            Wynik = sqlCommand.ExecuteScalar()

        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        Finally
            sqlConnection.Close()
        End Try
        Return IIf(Wynik = Nothing, False, (Wynik > 0))
    End Function




    Public Function CzyJestKolumnaWTabeli(ByVal NazwaTabeli As String, ByVal NazwaKolumny As String) As Boolean
        Dim sqlConnection As SqlClient.SqlConnection
        Dim sqlCommand As SqlClient.SqlCommand
        Dim Wynik As Integer = 0

        'sprawdz czy istnieje kolumna...
        'SELECT count(*) FROM INFORMATION_SCHEMA.columns WHERE (table_name = 'Klienci') AND (column_name = 'KatKlientaPRYZMAT')

        sqlConnection = New SqlClient.SqlConnection(sConnectionStr)
        sqlCommand = New SqlClient.SqlCommand
        Try
            sqlCommand.Connection = sqlConnection
            sqlCommand.CommandType = CommandType.Text
            sqlCommand.CommandText = "SELECT count(*) FROM INFORMATION_SCHEMA.columns WHERE (table_name = '" & NazwaTabeli & "') AND (column_name = '" & NazwaKolumny & "')"
            sqlConnection.Open()

            Wynik = sqlCommand.ExecuteScalar()

        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        Finally
            sqlConnection.Close()
        End Try
        Return IIf(Wynik = Nothing, False, (Wynik > 0))
    End Function






    '****************************
    ' Ustaw cechê NULL dla kazdej kolumny
    '****************************

    Private Sub UstawKolumnyNULL()
        Dim sConnection As String = sConnectionStr      'DataBaseConnection()
        Dim sSelectSQL As String = "SELECT * FROM Klienci"

        'Dim dbConnection As OleDb.OleDbConnection
        'Dim dbCommandSelect As OleDb.OleDbCommand
        'Dim dbDataAdapter As OleDb.OleDbDataAdapter

        Dim sqlConnection As New SqlClient.SqlConnection
        Dim sqlCommandSelect As New SqlClient.SqlCommand
        Dim sqlDataAdapter As New SqlClient.SqlDataAdapter

        Dim DataSetTabela As New System.Data.DataSet
        Dim DataTableTabela As System.Data.DataTable
        Dim DataViewTabela As New System.Data.DataView

        Dim ConnectionState As System.Data.ConnectionState

        DataSetTabela = New DataSet
        If ParL_DataBaseType = "MS ACCESS" Then
        Else
            sqlConnection = New SqlClient.SqlConnection(sConnection)
            sqlCommandSelect = New SqlClient.SqlCommand(sSelectSQL, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnection.Open()
            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnection.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                Else
                    sqlDataAdapter.FillSchema(DataSetTabela, SchemaType.Mapped, "TABELA")
                    sqlDataAdapter.Fill(DataSetTabela)
                    sqlConnection.Close()
                End If

                DataTableTabela = DataSetTabela.Tables("TABELA")
                Dim objColumn As DataColumn
                Dim i As Integer = 0

                i = 0
                For Each objColumn In DataTableTabela.Columns
                    If (objColumn.AllowDBNull = False) And (objColumn.Unique = False) Then
                        'nie pozwala na NULL
                        'If objColumn.MaxLength > 0 Then
                        Select Case objColumn.DataType.Name
                            Case "Int32"
                                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN " & objColumn.ColumnName & " INTEGER NULL;")
                            Case "Boolean"
                                AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN " & objColumn.ColumnName & " BIT NULL;")
                            Case "String"
                                If objColumn.MaxLength > 1000 Then
                                    AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN " & objColumn.ColumnName & " [text] NULL;")
                                Else
                                    AktSQL("Klienci", "ALTER TABLE Klienci ALTER COLUMN " & objColumn.ColumnName & " VARCHAR(" & Trim(Str(objColumn.MaxLength)) & ") NULL;")
                                End If
                            Case Else
                        End Select
                        'End If
                    End If
                    i += 1
                Next

            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
            End Try
        End If
    End Sub














End Class