'Imports System.Data.Odbc
Imports System.Data.Oledb

Public Class ZmienNazwyKatalogOsob
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cmdZmien As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblIleAkt As System.Windows.Forms.Label
    Friend WithEvents lblIlePrzed As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar As System.Windows.Forms.ProgressBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ZmienNazwyKatalogOsob))
        Me.cmdZmien = New System.Windows.Forms.Button()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lblIleAkt = New System.Windows.Forms.Label()
        Me.lblIlePrzed = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdZmien
        '
        Me.cmdZmien.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdZmien.Image = CType(resources.GetObject("cmdZmien.Image"), System.Drawing.Image)
        Me.cmdZmien.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdZmien.Location = New System.Drawing.Point(192, 152)
        Me.cmdZmien.Name = "cmdZmien"
        Me.cmdZmien.Size = New System.Drawing.Size(79, 23)
        Me.cmdZmien.TabIndex = 32
        Me.cmdZmien.Text = "Zmie�"
        Me.cmdZmien.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(277, 152)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 23)
        Me.cmdPowrot.TabIndex = 31
        Me.cmdPowrot.Text = "Powr�t"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(9, 152)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 23)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 30
        Me.PictureBox2.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.ProgressBar)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.lblIleAkt)
        Me.Panel1.Controls.Add(Me.lblIlePrzed)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(9, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(348, 137)
        Me.Panel1.TabIndex = 29
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(5, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(327, 52)
        Me.Label1.TabIndex = 51
        Me.Label1.Text = "Ta funkcja zmienia NAZWISKA I IMIONA WSZYSTKICH OS�B w Katalogu Os�b Kontaktowych" &
    " NIEODWRACALNIE (aby nie mo�na by�o rozpozna� tej osoby...)."
        '
        'ProgressBar
        '
        Me.ProgressBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar.Location = New System.Drawing.Point(8, 120)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(324, 8)
        Me.ProgressBar.TabIndex = 50
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(0, 64)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(388, 1)
        Me.Panel2.TabIndex = 48
        '
        'lblIleAkt
        '
        Me.lblIleAkt.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblIleAkt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIleAkt.ForeColor = System.Drawing.Color.Purple
        Me.lblIleAkt.Location = New System.Drawing.Point(220, 89)
        Me.lblIleAkt.Name = "lblIleAkt"
        Me.lblIleAkt.Size = New System.Drawing.Size(112, 17)
        Me.lblIleAkt.TabIndex = 28
        Me.lblIleAkt.Text = "0"
        Me.lblIleAkt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblIlePrzed
        '
        Me.lblIlePrzed.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblIlePrzed.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIlePrzed.ForeColor = System.Drawing.Color.Purple
        Me.lblIlePrzed.Location = New System.Drawing.Point(220, 72)
        Me.lblIlePrzed.Name = "lblIlePrzed"
        Me.lblIlePrzed.Size = New System.Drawing.Size(112, 17)
        Me.lblIlePrzed.TabIndex = 26
        Me.lblIlePrzed.Text = "0"
        Me.lblIlePrzed.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(8, 91)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(304, 16)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "Ilo�� zapis�w zmienionych . . . . . . . . . . . . . . . . . ."
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(304, 16)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Ilo�� zapis�w w katalogu . . . . . . . . . . . . . . . . . "
        '
        'ZmienNazwyKatalogOsob
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(365, 178)
        Me.Controls.Add(Me.cmdZmien)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ZmienNazwyKatalogOsob"
        Me.ShowInTaskbar = False
        Me.Text = "ANONIMIZACJA - Zmiana nazwisk w Katalogu Os�b Kontaktowych"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim dbSelectSQLOsoby As String = "SELECT * FROM Osoby"

    Dim dbConnectionOsoby As OleDb.OleDbConnection
    Dim dbCommandSelectOsoby As OleDb.OleDbCommand
    Dim dbCommandDeleteOsoby As OleDb.OleDbCommand
    Dim dbCommandUpdateOsoby As OleDb.OleDbCommand
    Dim dbCommandInsertOsoby As OleDb.OleDbCommand
    Dim dbDataAdapterOsoby As OleDb.OleDbDataAdapter

    Dim sqlConnectionOsoby As SqlClient.SqlConnection
    Dim sqlCommandSelectOsoby As SqlClient.SqlCommand
    Dim sqlCommandDeleteOsoby As SqlClient.SqlCommand
    Dim sqlCommandUpdateOsoby As SqlClient.SqlCommand
    Dim sqlCommandInsertOsoby As SqlClient.SqlCommand
    Dim sqlDataAdapterOsoby As SqlClient.SqlDataAdapter

    'DataSet
    Dim DataSetOsoby As New DataSet
    Dim DataViewOsoby As New DataView
    Dim ConnectionState As System.Data.ConnectionState
    '-------------------------------------
    Dim IlePrzed As Long
    Dim IleAkt As Long

    Private Sub ZmienNazwyKatalogOsob_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub




    Private Sub dbSQLUpdateOsoby(ByRef sqlCommandUpdate As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda Update
        'parametry aktualizacji
        'sqlCommandUpdate.Parameters.Add("@oIdentOso", sqlDbType.Char, 6, "IDENTKLIENTA")
        sqlCommandUpdate.Parameters.Add("@oOsobaOso", SqlDbType.VarChar, 100, "OSOBA")
        sqlCommandUpdate.Parameters.Add("@oVIPOso", SqlDbType.Bit, Nothing, "VIP")
        sqlCommandUpdate.Parameters.Add("@oeMailOso", SqlDbType.VarChar, 100, "EMAIL")
        sqlCommandUpdate.Parameters.Add("@oTelefonOso", SqlDbType.VarChar, 100, "TELEFON")
        sqlCommandUpdate.Parameters.Add("@oTelefonKomOso", SqlDbType.VarChar, 100, "TELEFONKOM")
        sqlCommandUpdate.Parameters.Add("@oStanowiskoOso", SqlDbType.VarChar, 100, "STANOWISKO")
        sqlCommandUpdate.Parameters.Add("@oKompetencjeOso", SqlDbType.VarChar, 100, "KOMPETENCJE")
        sqlCommandUpdate.Parameters.Add("@oWersjaOso", SqlDbType.Int, Nothing, "WERSJA")

        'parametry filtrowania
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygSymbol", SqlDbType.Char, 6, "IDENTKLIENTA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygOsoba", SqlDbType.VarChar, 100, "OSOBA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.UpdateCommand = sqlCommandUpdate
    End Sub


    'Public sSelectSQLOsoby As String = "SELECT " &
    '                                       "IDENTKLIENTA, " &
    '                                       "OSOBA," &
    '                                       "VIP," &
    '                                       "EMAIL," &
    '                                       "TELEFON," &
    '                                       "TELEFONKOM," &
    '                                       "STANOWISKO," &
    '                                       "KOMPETENCJE," &
    '                                       "WERSJA " &
    '                                    "FROM Osoby"

    Dim dbUpdateSQLOsoby As String = "UPDATE Osoby SET " &
                                                 "OSOBA=@oOsobaOso," &
                                                 "VIP=@oVIPOso," &
                                                 "EMAIL=@oeMailOso," &
                                                 "TELEFON=@oTelefonOso," &
                                                 "TELEFONKOM=@oTelefonKomOso," &
                                                 "STANOWISKO=@oStanowiskoOso," &
                                                 "KOMPETENCJE=@oKompetencjeOso," &
                                                 "WERSJA=@oWersjaOso " &
                                        "WHERE (IDENTKLIENTA=@orygSymbol) AND " &
                                              "(OSOBA=@orygOsoba) AND " &
                                              "(WERSJA=@orygWersja)"



    Private Sub ImportOsoby_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        IlePrzed = 0
        IleAkt = 0

        lblIlePrzed.Text = IlePrzed.ToString
        lblIleAkt.Text = IleAkt.ToString

        ProgressBar.Visible = False
        '---------------------------------------------
        'katalog Osob
        DataSetOsoby.Locale = New System.Globalization.CultureInfo("pl-PL")
        '---------------------------------------------
        If ParL_DataBaseType = "MS ACCESS" Then
        Else
            sqlConnectionOsoby = New SqlClient.SqlConnection(sConnectionOsoby)
            sqlCommandSelectOsoby = New SqlClient.SqlCommand(sSelectSQLOsoby, sqlConnectionOsoby)
            'sqlCommandDeleteOsoby = New SqlClient.SqlCommand(sDeleteSQLOsoby, sqlConnectionOsoby)
            sqlCommandUpdateOsoby = New SqlClient.SqlCommand(dbUpdateSQLOsoby, sqlConnectionOsoby)
            'sqlCommandInsertOsoby = New SqlClient.SqlCommand(sInsertSQLOsoby, sqlConnectionOsoby)
            sqlDataAdapterOsoby = New SqlClient.SqlDataAdapter(sqlCommandSelectOsoby)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliOsoby As System.Data.Common.DataTableMapping
            MapowanieTabeliOsoby = sqlDataAdapterOsoby.TableMappings.Add("Table", "TABELA_Osoby")

            'SQLDeleteOsoby(sqlCommandDeleteOsoby, sqlDataAdapterOsoby)
            dbSQLUpdateOsoby(sqlCommandUpdateOsoby, sqlDataAdapterOsoby)
            'SQLInsertOsoby(sqlCommandInsertOsoby, sqlDataAdapterOsoby)

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterOsoby.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionOsoby.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst�pi� b��d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionOsoby.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                Else
                    sqlDataAdapterOsoby.FillSchema(DataSetOsoby, SchemaType.Mapped, "TABELA_Osoby")
                    sqlDataAdapterOsoby.Fill(DataSetOsoby)
                    sqlConnectionOsoby.Close()
                End If

                ''zdefiniuj unikalny klucz wyszukiwania w tabeli
                'DataSetOsoby.Tables(0).PrimaryKey = New DataColumn() {DataSetOsoby.Tables(0).Columns("IDENTKLIENTA"),
                '                                                      DataSetOsoby.Tables(0).Columns("OSOBA")}

                DataViewOsoby = New DataView(DataSetOsoby.Tables(0))

                IlePrzed = DataViewOsoby.Count
                IleAkt = 0

            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst�pi� b��d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try
            '---------------------------------
        End If
        '---------------------
        lblIlePrzed.Text = IlePrzed.ToString
        lblIleAkt.Text = IleAkt.ToString
        System.Windows.Forms.Application.DoEvents()
    End Sub

    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujBaze()
        If ParL_DataBaseType = "MS ACCESS" Then
        Else
            Try
                sqlConnectionOsoby.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionOsoby.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterOsoby.Update(DataSetOsoby, "TABELA_Osoby")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterOsoby.Fill(DataSetOsoby)
                End If
                sqlConnectionOsoby.Close()
            End If
        End If
    End Sub


    '*****************************************************
    '** Sprawdz dane
    '*****************************************************

    Private Sub cmdSprawdz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdZmien.Click
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim IdKlienta As String = ""
        Dim NazwaKlientaPrzed As String = ""
        Dim NazwaKlientaPo As String = ""
        Dim Litera As String = ""
        Dim ch As String = ""

        Try
            If DataViewOsoby.Count > 0 Then
                IlePrzed = DataViewOsoby.Count
                IleAkt = 0

                ProgressBar.Maximum = IlePrzed
                ProgressBar.Visible = True
                For i = 0 To IlePrzed - 1
                    ProgressBar.Value = i
                    NazwaKlientaPrzed = DataViewOsoby.Item(i).Item("OSOBA")
                    NazwaKlientaPo = ""
                    '---------------------------------------
                    'Zmiana nazwy klienta
                    If Len(NazwaKlientaPrzed) > 0 Then
                        For j = 0 To Len(NazwaKlientaPrzed) - 1
                            ch = Mid(NazwaKlientaPrzed, j + 1, 1)
                            Select Case ch
                                Case " ", ".", ",", ":", ";", "/", "\", "-", "_", "'"
                                    NazwaKlientaPo &= ch
                                    Litera = ""

                                Case Else
                                    If Len(Litera) = 0 Then Litera = ch
                                    NazwaKlientaPo &= Litera
                            End Select

                        Next
                    End If
                    '---------------------------------------
                    DataViewOsoby.Item(i).Item("OSOBA") = NazwaKlientaPo
                    lblIleAkt.Text += 1
                    System.Windows.Forms.Application.DoEvents()
                Next
                AktualizujBaze()
            End If
            '---------------------
            lblIlePrzed.Text = IlePrzed.ToString
            lblIleAkt.Text = IleAkt.ToString
            System.Windows.Forms.Application.DoEvents()

            MessageBox.Show("Zmieni�em nieodwracalnie nazwiska i imiona klient�w w Katalogu Os�b Kontaktowych.", "Ok, sko�czy�em.",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
            ProgressBar.Visible = False


        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst�pi� b��d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
        End Try
    End Sub


End Class

