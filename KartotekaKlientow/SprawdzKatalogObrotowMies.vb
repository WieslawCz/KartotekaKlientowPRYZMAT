'Imports System.Data.Odbc
Imports System.Data.Oledb

Public Class SprawdzKatalogObrotowMies
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
    Friend WithEvents lblIlePo As System.Windows.Forms.Label
    Friend WithEvents lblIleUsunieto As System.Windows.Forms.Label
    Friend WithEvents lblIlePrzed As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar As System.Windows.Forms.ProgressBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SprawdzKatalogObrotowMies))
        Me.cmdSprawdz = New System.Windows.Forms.Button()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lblIlePo = New System.Windows.Forms.Label()
        Me.lblIleUsunieto = New System.Windows.Forms.Label()
        Me.lblIlePrzed = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdSprawdz
        '
        Me.cmdSprawdz.Image = CType(resources.GetObject("cmdSprawdz.Image"), System.Drawing.Image)
        Me.cmdSprawdz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSprawdz.Location = New System.Drawing.Point(326, 187)
        Me.cmdSprawdz.Name = "cmdSprawdz"
        Me.cmdSprawdz.Size = New System.Drawing.Size(79, 23)
        Me.cmdSprawdz.TabIndex = 32
        Me.cmdSprawdz.Text = "Sprawdz"
        Me.cmdSprawdz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(411, 187)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 23)
        Me.cmdPowrot.TabIndex = 31
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(9, 187)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 23)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 30
        Me.PictureBox2.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.ProgressBar)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.lblIlePo)
        Me.Panel1.Controls.Add(Me.lblIleUsunieto)
        Me.Panel1.Controls.Add(Me.lblIlePrzed)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(9, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(480, 172)
        Me.Panel1.TabIndex = 29
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(5, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(459, 52)
        Me.Label1.TabIndex = 51
        Me.Label1.Text = resources.GetString("Label1.Text")
        '
        'ProgressBar
        '
        Me.ProgressBar.Location = New System.Drawing.Point(11, 149)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(456, 8)
        Me.ProgressBar.TabIndex = 50
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(0, 64)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(480, 1)
        Me.Panel2.TabIndex = 48
        '
        'lblIlePo
        '
        Me.lblIlePo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIlePo.ForeColor = System.Drawing.Color.Purple
        Me.lblIlePo.Location = New System.Drawing.Point(312, 120)
        Me.lblIlePo.Name = "lblIlePo"
        Me.lblIlePo.Size = New System.Drawing.Size(112, 17)
        Me.lblIlePo.TabIndex = 28
        Me.lblIlePo.Text = "0"
        Me.lblIlePo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblIleUsunieto
        '
        Me.lblIleUsunieto.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIleUsunieto.ForeColor = System.Drawing.Color.Purple
        Me.lblIleUsunieto.Location = New System.Drawing.Point(312, 96)
        Me.lblIleUsunieto.Name = "lblIleUsunieto"
        Me.lblIleUsunieto.Size = New System.Drawing.Size(112, 17)
        Me.lblIleUsunieto.TabIndex = 27
        Me.lblIleUsunieto.Text = "0"
        Me.lblIleUsunieto.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblIlePrzed
        '
        Me.lblIlePrzed.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIlePrzed.ForeColor = System.Drawing.Color.Purple
        Me.lblIlePrzed.Location = New System.Drawing.Point(312, 72)
        Me.lblIlePrzed.Name = "lblIlePrzed"
        Me.lblIlePrzed.Size = New System.Drawing.Size(112, 17)
        Me.lblIlePrzed.TabIndex = 26
        Me.lblIlePrzed.Text = "0"
        Me.lblIlePrzed.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(8, 120)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(304, 16)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "Iloœæ zapisów pozosta³ych po sprawdzeniu . . . . . . . . . . . . . . . . . . . . " &
    "."
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 96)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(304, 16)
        Me.Label3.TabIndex = 22
        Me.Label3.Text = "Iloœæ zapisów usuniêtych w wyniku sprawdzenia  . . . . . . . . . . . . . . . . . " &
    ". . . . . . . . ."
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(304, 16)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Iloœæ zapisów w katalogu przed sprawdzeniem . . . . . . . . . . . . . . . . . . "
        '
        'SprawdzKatalogObrotowMies
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(496, 216)
        Me.Controls.Add(Me.cmdSprawdz)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SprawdzKatalogObrotowMies"
        Me.ShowInTaskbar = False
        Me.Text = "Sprawdzanie Katalogu Obrotów Miesiêcznych"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    '*********************************************************
    '** ObrotyMies kontaktowe
    '*********************************************************

    Dim dbSelectSQLObrotyMies As String = "SELECT * FROM ObrotyMies"

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

    'DataSet
    Dim DataSetObrotyMies As New DataSet
    Dim DataViewObrotyMies As New DataView


    '*********************************************************
    '** Klienci
    '*********************************************************

    Dim dbConnection As OleDb.OleDbConnection
    Dim dbCommandSelect As OleDb.OleDbCommand
    Dim dbCommandDelete As OleDb.OleDbCommand
    Dim dbCommandUpdate As OleDb.OleDbCommand
    Dim dbCommandInsert As OleDb.OleDbCommand
    Dim dbDataAdapter As OleDb.OleDbDataAdapter

    Dim sqlConnection As SqlClient.SqlConnection
    Dim sqlCommandSelect As SqlClient.SqlCommand
    Dim sqlCommandDelete As SqlClient.SqlCommand
    Dim sqlCommandUpdate As SqlClient.SqlCommand
    Dim sqlCommandInsert As SqlClient.SqlCommand
    Dim sqlDataAdapter As SqlClient.SqlDataAdapter

    'DataSet
    Dim DataSetKlienci As New DataSet
    Dim DataViewKlienci As New DataView

    Dim ConnectionState As System.Data.ConnectionState
    '-------------------------------------
    Dim IlePrzed As Long
    Dim IleUsunieto As Long
    Dim IlePo As Long

    Private Sub SprawdzKatalogObrotowMies_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
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
        IleUsunieto = 0
        IlePo = 0

        lblIlePrzed.Text = IlePrzed.ToString
        lblIleUsunieto.Text = IleUsunieto.ToString
        lblIlePo.Text = IlePo.ToString

        ProgressBar.Visible = False
        '---------------------------------------------
        'katalog ObrotowMies
        DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")
        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")
        '---------------------------------------------
        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnection = New OleDb.OleDbConnection(sConnectionObrotyMies)
            dbCommandSelect = New OleDb.OleDbCommand(dbSelectSQLObrotyMies, dbConnection)
            dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLObrotyMies, dbConnection)
            dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLObrotyMies, dbConnection)
            dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLObrotyMies, dbConnection)
            dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_ObrotyMies")

            'komendy DataAdapter
            Dim objParam As OleDb.OleDbParameter
            '------------------------------------------
            'komenda DELETE
            'parametry filtrowania
            objParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 6, "IDENTKLIENTA")
            objParam.SourceVersion = DataRowVersion.Original
            objParam = dbCommandDelete.Parameters.Add("@orygMies", OleDb.OleDbType.VarChar, 7, "MIESIAC")
            objParam.SourceVersion = DataRowVersion.Original
            objParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Original
            dbDataAdapter.DeleteCommand = dbCommandDelete


            ' Add the RowUpdated event handler.
            AddHandler dbDataAdapter.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

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
            sqlConnection = New SqlClient.SqlConnection(sConnectionObrotyMies)
            sqlCommandSelect = New SqlClient.SqlCommand(dbSelectSQLObrotyMies, sqlConnection)
            sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLObrotyMies, sqlConnection)
            sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLObrotyMies, sqlConnection)
            sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLObrotyMies, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_ObrotyMies")

            'komendy DataAdapter
            Dim objParam As SqlClient.SqlParameter
            '------------------------------------------
            'komenda DELETE
            'parametry filtrowania
            objParam = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 6, "IDENTKLIENTA")
            objParam.SourceVersion = DataRowVersion.Original
            objParam = sqlCommandDelete.Parameters.Add("@orygMies", SqlDbType.VarChar, 7, "MIESIAC")
            objParam.SourceVersion = DataRowVersion.Original
            objParam = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Original
            sqlDataAdapter.DeleteCommand = sqlCommandDelete

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapter.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

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
                    dbDataAdapter.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    dbDataAdapter.Fill(DataSetObrotyMies)
                    dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    sqlDataAdapter.Fill(DataSetObrotyMies)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetObrotyMies.Tables(0).PrimaryKey = New DataColumn() {DataSetObrotyMies.Tables(0).Columns("IDENTKLIENTA"), DataSetObrotyMies.Tables(0).Columns("MIESIAC")}
                DataViewObrotyMies = New DataView(DataSetObrotyMies.Tables(0))

                IlePrzed = DataViewObrotyMies.Count
                IleUsunieto = 0
                IlePo = DataViewObrotyMies.Count

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
        lblIleUsunieto.Text = IleUsunieto.ToString
        lblIlePo.Text = IlePo.ToString
    End Sub

    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujBaze()
        If ParL_DataBaseType = "MS ACCESS" Then
            Try
                dbConnection.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If dbConnection.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    dbDataAdapter.Update(DataSetObrotyMies, "TABELA_ObrotyMies")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    dbDataAdapter.Fill(DataSetObrotyMies)
                End If
                dbConnection.Close()
            End If
        Else
            Try
                sqlConnection.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnection.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapter.Update(DataSetObrotyMies, "TABELA_ObrotyMies")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapter.Fill(DataSetObrotyMies)
                End If
                sqlConnection.Close()
            End If
        End If
    End Sub


    '*****************************************************
    '** Sprawdz dane
    '*****************************************************

    Private Sub cmdSprawdz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSprawdz.Click
        Dim i As Long = 0
        Dim IdKlienta As String = ""
        Try
            If DataViewObrotyMies.Count > 0 Then
                IlePrzed = DataViewObrotyMies.Count
                IleUsunieto = 0
                IlePo = DataViewObrotyMies.Count

                ProgressBar.Maximum = IlePrzed
                ProgressBar.Visible = True
                For i = IlePrzed - 1 To 0 Step -1
                    ProgressBar.Value = IlePrzed - i
                    IdKlienta = GetTxtDRVField(DataViewObrotyMies.Item(i), "IDENTKLIENTA")
                    If Not CzyJestKlient(IdKlienta) Then
                        'usuwamy ten zapis...
                        DataViewObrotyMies.Item(i).Delete()
                        IleUsunieto += 1
                        IlePo -= 1

                        lblIleUsunieto.Text = IleUsunieto.ToString
                        lblIlePo.Text = IlePo.ToString
                    End If
                    System.Windows.Forms.Application.DoEvents()
                Next
                AktualizujBaze()
            End If

            MessageBox.Show("Wykona³em sprawdzenie Katalogu Obrotów Miesiêcznych.", "Ok, skoñczy³em.", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
            ProgressBar.Visible = False


        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Finally
        End Try
    End Sub


End Class

