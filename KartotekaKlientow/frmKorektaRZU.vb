'Imports System.Data.Odbc
Imports System.Data.Oledb

Public Class frmKorektaRZU
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
    Friend WithEvents cmdSprawdz As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblIleSkorygowano As System.Windows.Forms.Label
    Friend WithEvents lblIlePrzed As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar As System.Windows.Forms.ProgressBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmKorektaRZU))
        Me.cmdSprawdz = New System.Windows.Forms.Button()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lblIleSkorygowano = New System.Windows.Forms.Label()
        Me.lblIlePrzed = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdSprawdz
        '
        Me.cmdSprawdz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSprawdz.Image = CType(resources.GetObject("cmdSprawdz.Image"), System.Drawing.Image)
        Me.cmdSprawdz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSprawdz.Location = New System.Drawing.Point(347, 244)
        Me.cmdSprawdz.Name = "cmdSprawdz"
        Me.cmdSprawdz.Size = New System.Drawing.Size(79, 23)
        Me.cmdSprawdz.TabIndex = 32
        Me.cmdSprawdz.Text = "Koryguj"
        Me.cmdSprawdz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(432, 244)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 23)
        Me.cmdPowrot.TabIndex = 31
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(9, 244)
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
        Me.Panel1.Controls.Add(Me.lblIleSkorygowano)
        Me.Panel1.Controls.Add(Me.lblIlePrzed)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(9, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(501, 229)
        Me.Panel1.TabIndex = 29
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(5, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(480, 129)
        Me.Label1.TabIndex = 51
        Me.Label1.Text = resources.GetString("Label1.Text")
        '
        'ProgressBar
        '
        Me.ProgressBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar.Location = New System.Drawing.Point(8, 206)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(477, 8)
        Me.ProgressBar.TabIndex = 50
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(0, 140)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(501, 1)
        Me.Panel2.TabIndex = 48
        '
        'lblIleSkorygowano
        '
        Me.lblIleSkorygowano.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIleSkorygowano.ForeColor = System.Drawing.Color.Purple
        Me.lblIleSkorygowano.Location = New System.Drawing.Point(309, 176)
        Me.lblIleSkorygowano.Name = "lblIleSkorygowano"
        Me.lblIleSkorygowano.Size = New System.Drawing.Size(112, 17)
        Me.lblIleSkorygowano.TabIndex = 27
        Me.lblIleSkorygowano.Text = "0"
        Me.lblIleSkorygowano.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblIlePrzed
        '
        Me.lblIlePrzed.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIlePrzed.ForeColor = System.Drawing.Color.Purple
        Me.lblIlePrzed.Location = New System.Drawing.Point(309, 152)
        Me.lblIlePrzed.Name = "lblIlePrzed"
        Me.lblIlePrzed.Size = New System.Drawing.Size(112, 17)
        Me.lblIlePrzed.TabIndex = 26
        Me.lblIlePrzed.Text = "0"
        Me.lblIlePrzed.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(5, 176)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(304, 16)
        Me.Label3.TabIndex = 22
        Me.Label3.Text = "Iloœæ zapisów skorygowanych . . . . . . . . . . . . . . . . . . . . . . . . . . ." &
    " . . . . . . . . . . . . . . ."
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(5, 152)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(304, 16)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Iloœæ zapisów w Katalogu Klientów . . . . . . . . . . . . . . . . . . . . . . . ." &
    " . . . "
        '
        'frmKorektaRZU
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(518, 270)
        Me.Controls.Add(Me.cmdSprawdz)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmKorektaRZU"
        Me.ShowInTaskbar = False
        Me.Text = "Korekta b³êdnych zapisów Rozszerzonego Zakresu Us³ug"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region



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

    'DataSet
    Dim DataSetKlienci As New DataSet
    Dim DataViewKlienci As New DataView
    Dim ConnectionState As System.Data.ConnectionState
    '-------------------------------------
    Dim IlePrzed As Long
    Dim IleSkorygowano As Long

    Private Sub frmKorektaRZU_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub ImportKlienci_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        IlePrzed = 0
        IleSkorygowano = 0

        lblIlePrzed.Text = IlePrzed.ToString
        lblIleSkorygowano.Text = IleSkorygowano.ToString

        ProgressBar.Visible = False
        '---------------------------------------------
        'katalog kontaktow
        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")
        '---------------------------------------------
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionKlienci = New OleDb.OleDbConnection(sConnectionKlienci)
            'dbCommandSelectKlienci = New OleDb.OleDbCommand(dbSelectSQLKlienci, dbConnectionKlienci)
            ''dbCommandDeleteKlienci = New OleDb.OleDbCommand(sDeleteSQLKlienci, dbConnectionKlienci)
            'dbCommandUpdateKlienci = New OleDb.OleDbCommand(sUpdateSQLKlienci, dbConnectionKlienci)
            ''dbCommandInsertKlienci = New OleDb.OleDbCommand(sInsertSQLKlienci, dbConnectionKlienci)
            'dbDataAdapterKlienci = New OleDb.OleDbDataAdapter(dbCommandSelectKlienci)

            ''mapujemy domyslna nazwe tabeli
            'Dim dbMapowanieTabeliKlienci As System.Data.Common.DataTableMapping
            'dbMapowanieTabeliKlienci = dbDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            ''DBDeleteKlienci(dbCommandDeleteKlienci, dbDataAdapterKlienci)
            'DBUpdateKlienci(dbCommandUpdateKlienci, dbDataAdapterKlienci)
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
            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
            sqlCommandSelectKlienci = New SqlClient.SqlCommand(sSelectSQLKlienciAktywni, sqlConnectionKlienci)
            'sqlCommandDeleteKlienci = New SqlClient.SqlCommand(sDeleteSQLKlienci, sqlConnectionKlienci)
            sqlCommandUpdateKlienci = New SqlClient.SqlCommand(sUpdateSQLKlienci, sqlConnectionKlienci)
            'sqlCommandInsertKlienci = New SqlClient.SqlCommand(sInsertSQLKlienci, sqlConnectionKlienci)
            sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliKlienci As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliKlienci = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

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
                    dbDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    dbDataAdapterKlienci.Fill(DataSetKlienci)
                    dbConnectionKlienci.Close()
                Else
                    sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("IDENTKLIENTA")}

                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))

                IlePrzed = DataViewKlienci.Count
                IleSkorygowano = 0

            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try
            '---------------------------------
        End If
        '---------------------
        lblIlePrzed.Text = IlePrzed.ToString
        lblIleSkorygowano.Text = IleSkorygowano.ToString
    End Sub

    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujBaze()
        If ParL_DataBaseType = "MS ACCESS" Then
            Try
                dbConnectionKlienci.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If dbConnectionKlienci.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    dbDataAdapterKlienci.Update(DataSetKlienci, "TABELA_Klienci")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    dbDataAdapterKlienci.Fill(DataSetKlienci)
                End If
                dbConnectionKlienci.Close()
            End If
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
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                End If
                sqlConnectionKlienci.Close()
            End If
        End If
    End Sub


    '*****************************************************
    '** Sprawdz dane
    '*****************************************************

    Private Sub cmdSprawdz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSprawdz.Click
        Dim i As Long = 0
        Dim RZUKlienta As String = ""
        Dim foundRow As DataRow

        Try
            If DataSetKlienci.Tables(0).Rows.Count > 0 Then
                IlePrzed = DataSetKlienci.Tables(0).Rows.Count
                IleSkorygowano = 0

                lblIlePrzed.Text = IlePrzed.ToString
                lblIleSkorygowano.Text = IleSkorygowano.ToString

                ProgressBar.Maximum = IlePrzed
                ProgressBar.Visible = True
                For i = IlePrzed - 1 To 0 Step -1
                    ProgressBar.Value = IlePrzed - i
                    foundRow = DataSetKlienci.Tables(0).Rows.Item(i)

                    RZUKlienta = GetTxtDRField(foundRow, "USLUGIDOD")
                    If Mid(RZUKlienta, 1, 9) = "TTTTTTTTT" Then
                        'zmieniamy ten zapis...
                        foundRow("USLUGIDOD") = "NNNNNNNNN" & Mid(RZUKlienta, 10)

                        IleSkorygowano += 1
                        lblIleSkorygowano.Text = IleSkorygowano.ToString
                    End If
                    System.Windows.Forms.Application.DoEvents()
                Next
                AktualizujBaze()
            End If

            MessageBox.Show("Wykona³em korekte zapisów dotyczacych Rozszerzonego Zakresu Us³ug.", "Ok, skoñczy³em.",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
            ProgressBar.Visible = False


        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
        End Try
    End Sub


End Class

