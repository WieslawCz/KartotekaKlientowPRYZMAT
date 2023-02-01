Public Class PobierzObrotyzTOFI
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents cmdWybierz As System.Windows.Forms.Button
    Friend WithEvents lblIleOdrzucono As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblIleZaktualizowano As System.Windows.Forms.Label
    Friend WithEvents lblIleDopisano As System.Windows.Forms.Label
    Friend WithEvents lblIlePrzeczytano As System.Windows.Forms.Label
    Friend WithEvents lblPlik As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cmdImportuj As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdPokazPlik As System.Windows.Forms.Button
    Friend WithEvents ProgressBar As System.Windows.Forms.ProgressBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PobierzObrotyzTOFI))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdPokazPlik = New System.Windows.Forms.Button()
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cmdWybierz = New System.Windows.Forms.Button()
        Me.lblIleOdrzucono = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblIleZaktualizowano = New System.Windows.Forms.Label()
        Me.lblIleDopisano = New System.Windows.Forms.Label()
        Me.lblIlePrzeczytano = New System.Windows.Forms.Label()
        Me.lblPlik = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.cmdImportuj = New System.Windows.Forms.Button()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.cmdPokazPlik)
        Me.Panel1.Controls.Add(Me.ProgressBar)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.cmdWybierz)
        Me.Panel1.Controls.Add(Me.lblIleOdrzucono)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.lblIleZaktualizowano)
        Me.Panel1.Controls.Add(Me.lblIleDopisano)
        Me.Panel1.Controls.Add(Me.lblIlePrzeczytano)
        Me.Panel1.Controls.Add(Me.lblPlik)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(488, 192)
        Me.Panel1.TabIndex = 30
        '
        'cmdPokazPlik
        '
        Me.cmdPokazPlik.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPokazPlik.ForeColor = System.Drawing.Color.Black
        Me.cmdPokazPlik.Image = CType(resources.GetObject("cmdPokazPlik.Image"), System.Drawing.Image)
        Me.cmdPokazPlik.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPokazPlik.Location = New System.Drawing.Point(251, 32)
        Me.cmdPokazPlik.Name = "cmdPokazPlik"
        Me.cmdPokazPlik.Size = New System.Drawing.Size(95, 24)
        Me.cmdPokazPlik.TabIndex = 51
        Me.cmdPokazPlik.Text = "Poka¿ plik"
        Me.cmdPokazPlik.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ProgressBar
        '
        Me.ProgressBar.Location = New System.Drawing.Point(8, 176)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(464, 8)
        Me.ProgressBar.TabIndex = 49
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(0, 64)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(480, 1)
        Me.Panel2.TabIndex = 48
        '
        'cmdWybierz
        '
        Me.cmdWybierz.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWybierz.ForeColor = System.Drawing.Color.Black
        Me.cmdWybierz.Image = CType(resources.GetObject("cmdWybierz.Image"), System.Drawing.Image)
        Me.cmdWybierz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWybierz.Location = New System.Drawing.Point(352, 32)
        Me.cmdWybierz.Name = "cmdWybierz"
        Me.cmdWybierz.Size = New System.Drawing.Size(112, 24)
        Me.cmdWybierz.TabIndex = 47
        Me.cmdWybierz.Text = "Wybierz plik"
        Me.cmdWybierz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblIleOdrzucono
        '
        Me.lblIleOdrzucono.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIleOdrzucono.ForeColor = System.Drawing.Color.Purple
        Me.lblIleOdrzucono.Location = New System.Drawing.Point(312, 144)
        Me.lblIleOdrzucono.Name = "lblIleOdrzucono"
        Me.lblIleOdrzucono.Size = New System.Drawing.Size(112, 17)
        Me.lblIleOdrzucono.TabIndex = 30
        Me.lblIleOdrzucono.Text = "0"
        Me.lblIleOdrzucono.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(72, 144)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(240, 16)
        Me.Label6.TabIndex = 29
        Me.Label6.Text = "Iloœæ rekordów odrzuconych. . . . . . . . . . . . . . . . ."
        '
        'lblIleZaktualizowano
        '
        Me.lblIleZaktualizowano.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIleZaktualizowano.ForeColor = System.Drawing.Color.Purple
        Me.lblIleZaktualizowano.Location = New System.Drawing.Point(312, 120)
        Me.lblIleZaktualizowano.Name = "lblIleZaktualizowano"
        Me.lblIleZaktualizowano.Size = New System.Drawing.Size(112, 17)
        Me.lblIleZaktualizowano.TabIndex = 28
        Me.lblIleZaktualizowano.Text = "0"
        Me.lblIleZaktualizowano.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblIleDopisano
        '
        Me.lblIleDopisano.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIleDopisano.ForeColor = System.Drawing.Color.Purple
        Me.lblIleDopisano.Location = New System.Drawing.Point(312, 96)
        Me.lblIleDopisano.Name = "lblIleDopisano"
        Me.lblIleDopisano.Size = New System.Drawing.Size(112, 17)
        Me.lblIleDopisano.TabIndex = 27
        Me.lblIleDopisano.Text = "0"
        Me.lblIleDopisano.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblIlePrzeczytano
        '
        Me.lblIlePrzeczytano.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIlePrzeczytano.ForeColor = System.Drawing.Color.Purple
        Me.lblIlePrzeczytano.Location = New System.Drawing.Point(312, 72)
        Me.lblIlePrzeczytano.Name = "lblIlePrzeczytano"
        Me.lblIlePrzeczytano.Size = New System.Drawing.Size(112, 17)
        Me.lblIlePrzeczytano.TabIndex = 26
        Me.lblIlePrzeczytano.Text = "0"
        Me.lblIlePrzeczytano.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblPlik
        '
        Me.lblPlik.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPlik.ForeColor = System.Drawing.Color.Purple
        Me.lblPlik.Location = New System.Drawing.Point(72, 8)
        Me.lblPlik.Name = "lblPlik"
        Me.lblPlik.Size = New System.Drawing.Size(392, 17)
        Me.lblPlik.TabIndex = 24
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(72, 120)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(240, 16)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "Iloœæ rekordów zaktualizowanych . . . . . . . . . . . . . . . . . . . . ."
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(72, 96)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(240, 16)
        Me.Label3.TabIndex = 22
        Me.Label3.Text = "Iloœæ rekordów dopisanych . . . . . . . . . . . . . . . . . . . . ."
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(72, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(240, 16)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Iloœæ rekordów przeczytanych . . . . . . . . . . . . . . . . . . "
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(8, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(176, 16)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Plik z TOFI . . . . . .. . ."
        '
        'cmdImportuj
        '
        Me.cmdImportuj.Image = CType(resources.GetObject("cmdImportuj.Image"), System.Drawing.Image)
        Me.cmdImportuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdImportuj.Location = New System.Drawing.Point(328, 208)
        Me.cmdImportuj.Name = "cmdImportuj"
        Me.cmdImportuj.Size = New System.Drawing.Size(75, 23)
        Me.cmdImportuj.TabIndex = 35
        Me.cmdImportuj.Text = "Importuj"
        Me.cmdImportuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(409, 208)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(82, 23)
        Me.cmdPowrot.TabIndex = 34
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(15, 207)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 33
        Me.PictureBox2.TabStop = False
        '
        'PobierzObrotyzTOFI
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(502, 237)
        Me.Controls.Add(Me.cmdImportuj)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "PobierzObrotyzTOFI"
        Me.ShowInTaskbar = False
        Me.Text = "Pobieranie pliku z obrotami z TOFI"
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Dim KlienciOK As Boolean = False
    Dim ObrotyOK As Boolean = False
    Dim AnalizyOK As Boolean = False


    'OleDbConnection-polaczenie
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

    Dim DataSetKlienci As New DataSet
    Dim DataViewKlienci As New DataView




    Dim dbSelectSQLAnalizyZakupu As String = sSelectSQLAnalizyZakupu

    Dim dbConnectionAnalizyZakupu As OleDb.OleDbConnection
    Dim dbCommandSelectAnalizyZakupu As OleDb.OleDbCommand
    Dim dbCommandDeleteAnalizyZakupu As OleDb.OleDbCommand
    Dim dbCommandUpdateAnalizyZakupu As OleDb.OleDbCommand
    Dim dbCommandInsertAnalizyZakupu As OleDb.OleDbCommand
    Dim dbDataAdapterAnalizyZakupu As OleDb.OleDbDataAdapter

    Dim sqlConnectionAnalizyZakupu As SqlClient.SqlConnection
    Dim sqlCommandSelectAnalizyZakupu As SqlClient.SqlCommand
    Dim sqlCommandDeleteAnalizyZakupu As SqlClient.SqlCommand
    Dim sqlCommandUpdateAnalizyZakupu As SqlClient.SqlCommand
    Dim sqlCommandInsertAnalizyZakupu As SqlClient.SqlCommand
    Dim sqlDataAdapterAnalizyZakupu As SqlClient.SqlDataAdapter

    Dim DataSetAnalizyZakupu As New DataSet
    Dim DataViewAnalizyZakupu As New DataView






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
    Dim ConnectionState As System.Data.ConnectionState






    '-------------------------------------
    Dim IlePrzeczytano As Long
    Dim IleDopisano As Long
    Dim IleZaktualizowano As Long
    Dim IleOdrzucono As Long
    Dim IleJest As Long

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub



    Private Sub PobierzObrotyzTOFI_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        IlePrzeczytano = 0
        IleDopisano = 0
        IleZaktualizowano = 0
        IleOdrzucono = 0

        IleJest = 0
        lblPlik.Text = ""

        lblIlePrzeczytano.Text = IlePrzeczytano.ToString
        lblIleDopisano.Text = IleDopisano.ToString
        lblIleZaktualizowano.Text = IleZaktualizowano.ToString
        lblIleOdrzucono.Text = IleOdrzucono.ToString





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

        'Public oWersjaObr As Integer          'WERSJA           Integer


        'DataSet
        DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionObroty)
            'dbCommandSelect = New OleDb.OleDbCommand(dbSelectSQLObroty, dbConnection)
            'dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLObroty, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLObroty, dbConnection)
            'dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLObroty, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            ''DataSet
            'DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")

            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeliObroty As System.Data.Common.DataTableMapping
            'MapowanieTabeliObroty = dbDataAdapter.TableMappings.Add("Table", "TABELA_Obroty")

            'DBDeleteObroty(dbCommandDelete, dbDataAdapter)
            'DBUpdateObroty(dbCommandUpdate, dbDataAdapter)
            'DBInsertObroty(dbCommandInsert, dbDataAdapter)

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
            sqlConnectionObroty = New SqlClient.SqlConnection(sConnectionObroty)
            sqlCommandSelectObroty = New SqlClient.SqlCommand(dbSelectSQLObroty, sqlConnectionObroty)
            sqlCommandDeleteObroty = New SqlClient.SqlCommand(sDeleteSQLObroty, sqlConnectionObroty)
            sqlCommandUpdateObroty = New SqlClient.SqlCommand(sUpdateSQLObroty, sqlConnectionObroty)
            sqlCommandInsertObroty = New SqlClient.SqlCommand(sInsertSQLObroty, sqlConnectionObroty)
            sqlDataAdapterObroty = New SqlClient.SqlDataAdapter(sqlCommandSelectObroty)


            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliObroty As System.Data.Common.DataTableMapping
            MapowanieTabeliObroty = sqlDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")

            SQLDeleteObroty(sqlCommandDeleteObroty, sqlDataAdapterObroty)
            SQLUpdateObroty(sqlCommandUpdateObroty, sqlDataAdapterObroty)
            SQLInsertObroty(sqlCommandInsertObroty, sqlDataAdapterObroty)

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterObroty.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObroty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionObroty.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    dbDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    dbDataAdapterObroty.Fill(DataSetObroty)
                    dbConnectionObroty.Close()
                Else
                    sqlDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    sqlDataAdapterObroty.Fill(DataSetObroty)
                    sqlConnectionObroty.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetObroty.Tables(0).PrimaryKey = New DataColumn() {DataSetObroty.Tables(0).Columns("IDENTKLIENTA"), DataSetObroty.Tables(0).Columns("DATA")}

                'definiuj DataView
                DataViewObroty = New DataView(DataSetObroty.Tables(0))
                ObrotyOK = True

            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            '---------------------------------
        End If






        If ObrotyOK Then

            'DataSet
            DataSetAnalizyZakupu.Locale = New System.Globalization.CultureInfo("pl-PL")
            If ParL_DataBaseType = "MS ACCESS" Then
                'dbConnectionAnalizyZakupu = New OleDb.OleDbConnection(sConnectionAnalizyZakupu)
                'dbCommandSelectAnalizyZakupu = New OleDb.OleDbCommand(dbSelectSQLAnalizyZakupu, dbConnectionAnalizyZakupu)
                'dbCommandDeleteAnalizyZakupu = New OleDb.OleDbCommand(sDeleteSQLAnalizyZakupu, dbConnectionAnalizyZakupu)
                'dbCommandUpdateAnalizyZakupu = New OleDb.OleDbCommand(sUpdateSQLAnalizyZakupu, dbConnectionAnalizyZakupu)
                'dbCommandInsertAnalizyZakupu = New OleDb.OleDbCommand(sInsertSQLAnalizyZakupu, dbConnectionAnalizyZakupu)
                'dbDataAdapterAnalizyZakupu = New OleDb.OleDbDataAdapter(dbCommandSelectAnalizyZakupu)

                ''mapujemy domyslna nazwe tabeli
                'Dim MapowanieTabeliAnalizyZakupu As System.Data.Common.DataTableMapping
                'MapowanieTabeliAnalizyZakupu = dbDataAdapterAnalizyZakupu.TableMappings.Add("Table", "TABELA_AnalizyZakupu")

                'DBDeleteAnalizyZakupu(dbCommandDeleteAnalizyZakupu, dbDataAdapterAnalizyZakupu)
                'DBUpdateAnalizyZakupu(dbCommandUpdateAnalizyZakupu, dbDataAdapterAnalizyZakupu)
                'DBInsertAnalizyZakupu(dbCommandInsertAnalizyZakupu, dbDataAdapterAnalizyZakupu)

                '' Add the RowUpdated event handler.
                'AddHandler dbDataAdapterAnalizyZakupu.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

                ''------------------------------------------
                ''wypelnij DATASET
                'Try
                '    dbConnectionAnalizyZakupu.Open()
                'Catch Ex As System.Exception
                '    ViewInLog(Ex, Me.Name(), Nothing)
                '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    '    MessageBoxIcon.Information, _
                '    '    MessageBoxDefaultButton.Button1)
                'Finally
                '    ConnectionState = dbConnectionAnalizyZakupu.State
                'End Try

            Else
                sqlConnectionAnalizyZakupu = New SqlClient.SqlConnection(sConnectionAnalizyZakupu)
                sqlCommandSelectAnalizyZakupu = New SqlClient.SqlCommand(dbSelectSQLAnalizyZakupu, sqlConnectionAnalizyZakupu)
                sqlCommandDeleteAnalizyZakupu = New SqlClient.SqlCommand(sDeleteSQLAnalizyZakupu, sqlConnectionAnalizyZakupu)
                sqlCommandUpdateAnalizyZakupu = New SqlClient.SqlCommand(sUpdateSQLAnalizyZakupu, sqlConnectionAnalizyZakupu)
                sqlCommandInsertAnalizyZakupu = New SqlClient.SqlCommand(sInsertSQLAnalizyZakupu, sqlConnectionAnalizyZakupu)
                sqlDataAdapterAnalizyZakupu = New SqlClient.SqlDataAdapter(sqlCommandSelectAnalizyZakupu)

                'mapujemy domyslna nazwe tabeli
                Dim MapowanieTabeliAnalizyZakupu As System.Data.Common.DataTableMapping
                MapowanieTabeliAnalizyZakupu = sqlDataAdapterAnalizyZakupu.TableMappings.Add("Table", "TABELA_AnalizyZakupu")

                SQLDeleteAnalizyZakupu(sqlCommandDeleteAnalizyZakupu, sqlDataAdapterAnalizyZakupu)
                SQLUpdateAnalizyZakupu(sqlCommandUpdateAnalizyZakupu, sqlDataAdapterAnalizyZakupu)
                SQLInsertAnalizyZakupu(sqlCommandInsertAnalizyZakupu, sqlDataAdapterAnalizyZakupu)

                ' Add the RowUpdated event handler.
                AddHandler sqlDataAdapterAnalizyZakupu.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

                '------------------------------------------
                'wypelnij DATASET
                Try
                    sqlConnectionAnalizyZakupu.Open()
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                Finally
                    ConnectionState = sqlConnectionAnalizyZakupu.State
                End Try
            End If

            If ConnectionState = ConnectionState.Open Then
                Try
                    If ParL_DataBaseType = "MS ACCESS" Then
                        dbDataAdapterAnalizyZakupu.FillSchema(DataSetAnalizyZakupu, SchemaType.Mapped, "TABELA_AnalizyZakupu")
                        dbDataAdapterAnalizyZakupu.Fill(DataSetAnalizyZakupu)
                        dbConnectionAnalizyZakupu.Close()
                    Else
                        sqlDataAdapterAnalizyZakupu.FillSchema(DataSetAnalizyZakupu, SchemaType.Mapped, "TABELA_AnalizyZakupu")
                        sqlDataAdapterAnalizyZakupu.Fill(DataSetAnalizyZakupu)
                        sqlConnectionAnalizyZakupu.Close()
                    End If

                    DataSetAnalizyZakupu.Tables(0).PrimaryKey = New DataColumn() {DataSetAnalizyZakupu.Tables(0).Columns("IDENTKLIENTA")}
                    DataViewAnalizyZakupu = New DataView(DataSetAnalizyZakupu.Tables(0))
                    AnalizyOK = True

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


    Private Sub cmdWybierz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWybierz.Click
        With OpenFileDialog1
            .Title = "Proszê wybraæ plik z TOFI..."
            .InitialDirectory = "c:\"
            .DefaultExt = "txt"
            .Filter = "Plik tekstowy (*.txt)|*.txt|Wszystkie pliki|*.*"
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                lblPlik.Text = .FileName
            End If
        End With
    End Sub


    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujBaze()
        If ParL_DataBaseType = "MS ACCESS" Then
            Try
                dbConnectionObroty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If dbConnectionObroty.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    dbDataAdapterObroty.Update(DataSetObroty, "TABELA_Obroty")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    dbDataAdapterObroty.Fill(DataSetObroty)
                End If
                dbConnectionObroty.Close()
            End If
        Else
            Try
                sqlConnectionObroty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionObroty.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterObroty.Update(DataSetObroty, "TABELA_Obroty")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterObroty.Fill(DataSetObroty)
                End If
                sqlConnectionObroty.Close()
            End If
        End If
    End Sub



    Private Sub AktualizujBazeAnalizyZakupu()
        If ParL_DataBaseType = "MS ACCESS" Then
            Try
                dbConnectionAnalizyZakupu.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If dbConnectionAnalizyZakupu.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    dbDataAdapterAnalizyZakupu.Update(DataSetAnalizyZakupu, "TABELA_AnalizyZakupu")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    dbDataAdapterAnalizyZakupu.Fill(DataSetAnalizyZakupu)
                End If
                dbConnectionAnalizyZakupu.Close()
            End If
        Else
            Try
                sqlConnectionAnalizyZakupu.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionAnalizyZakupu.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterAnalizyZakupu.Update(DataSetAnalizyZakupu, "TABELA_AnalizyZakupu")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterAnalizyZakupu.Fill(DataSetAnalizyZakupu)
                End If
                sqlConnectionAnalizyZakupu.Close()
            End If
        End If
    End Sub



    Private Sub AktualizujBazeKlienci()
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
    '** Importujemy dane
    '*****************************************************


    Private Sub cmdImportuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImportuj.Click
        Dim dbSelectSQLKlienci As String = ""
        Dim SymbolKlienta As String = ""

        Dim TxtPlik As String
        Dim FileNum As Integer
        Dim NextLine As String
        Dim BiezacyRekord As Long
        Dim IleRec As Long
        Dim ip As Integer
        Dim ObrotyRow As DataRow
        Dim AnalizyZakupuRow As DataRow
        Dim Nierozpoz As String = ""

        Dim SymbTOFI As Long = 0
        Dim IDOddzialu As Long = 0
        Dim SymbTOFITxt As String = ""
        Dim SymbKlienta As String = ""
        Dim DataTOFI As String = ""
        Dim ObriacTOFI As String = ""
        Dim DataObrotu As String = ""

        Dim AIlosc As Integer = 0
        Dim AWartosc As Double = 0
        Dim AMarza As Double = 0

        Dim LIlosc As Integer = 0
        Dim LWartosc As Double = 0
        Dim LMarza As Double = 0

        Dim AoIlosc As Integer = 0
        Dim AoWartosc As Double = 0
        Dim AoMarza As Double = 0

        Dim LoIlosc As Integer = 0
        Dim LoWartosc As Double = 0
        Dim LoMarza As Double = 0

        Dim NAJEMIlosc As Integer = 0
        Dim NAJEMWartosc As Double = 0
        Dim NAJEMMarza As Double = 0

        Dim STRONYIlosc As Integer = 0
        Dim STRONYWartosc As Double = 0
        Dim STRONYMarza As Double = 0

        Dim DRUKARKIIlosc As Integer = 0
        Dim DRUKARKIWartosc As Double = 0
        Dim DRUKARKIMarza As Double = 0

        Dim SKUPIlosc As Integer = 0
        Dim SKUPWartosc As Double = 0
        Dim SKUPMarza As Double = 0

        Dim ik As Integer = 0
        Dim AktMiesAnalizy As String = ""


        IlePrzeczytano = 0
        IleDopisano = 0
        IleZaktualizowano = 0
        IleOdrzucono = 0
        IleJest = 0

        lblIlePrzeczytano.Text = IlePrzeczytano.ToString
        lblIleDopisano.Text = IleDopisano.ToString
        lblIleZaktualizowano.Text = IleZaktualizowano.ToString
        lblIleOdrzucono.Text = IleOdrzucono.ToString

        TxtPlik = Trim(lblPlik.Text)
        If Len(TxtPlik) = 0 Then
            MessageBox.Show("Brak definicji pliku z TOFI do pobrania",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            If Len(Dir(TxtPlik)) = 0 Then
                MessageBox.Show("Nie znalaz³em na dysku pliku do pobrania" & vbCrLf & TxtPlik & vbCrLf,
                    "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            Else
                cmdImportuj.Enabled = False

                'przejdz przez katalog OBROT - usuwaj
                Dim idr As Integer = 0
                If DataSetObroty.Tables(0).Rows.Count > 0 Then
                    For idr = DataSetObroty.Tables(0).Rows.Count - 1 To 0 Step -1
                        DataSetObroty.Tables(0).Rows.Item(idr).Delete()
                    Next
                    AktualizujBaze()
                End If

                'przejdz przez katalog AnalizyZakupu- zeruj
                If DataSetAnalizyZakupu.Tables(0).Rows.Count > 0 Then
                    For idr = 0 To DataSetAnalizyZakupu.Tables(0).Rows.Count - 1
                        aWersjaAnal = DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WERSJA")

                        DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("ILOSZABM") = 0
                        DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WARTZABM") = 0
                        DataSetAnalizyZakupu.Tables(0).Rows.Item(idr).Item("WERSJA") = ZmienWersje(aWersjaAnal)
                    Next
                    AktualizujBazeAnalizyZakupu()
                End If




                'jest plik - otwieramy
                BiezacyRekord = 0
                FileNum = FreeFile()    'kolejny nr otwartego zbioru

                'wylicz ile ma wierszy
                FileOpen(FileNum, TxtPlik, OpenMode.Input)
                IleRec = 0
                Do Until EOF(FileNum)
                    NextLine = LineInput(FileNum)

                    If IleRec = 0 Then
                        'za który miesi¹c pobieramy dane...
                        'analizuj pierwszy rekord
                        ip = InStr(NextLine, ",")
                        If ip > 0 Then
                            'SymbTOFI = CLng(Mid(NextLine, 1, ip - 1))
                            NextLine = Mid(NextLine, ip + 1)
                            ip = InStr(NextLine, ",")
                            If ip > 0 Then
                                'IDOddzialu = CLng(Mid(NextLine, 1, ip - 1))
                                NextLine = Mid(NextLine, ip + 1)
                                ip = InStr(NextLine, ",")
                                If ip > 0 Then
                                    DataTOFI = Mid(NextLine, 1, ip - 1)
                                End If
                            End If
                        End If
                        AktMiesAnalizy = Mid(DataTOFI, 5, 2)
                    End If

                    IleRec += 1
                Loop
                FileClose(FileNum)










                'przejdz przez katalog klienci
                DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")
                If ParL_DataBaseType = "MS ACCESS" Then
                Else
                    sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
                    sqlCommandSelectKlienci = New SqlClient.SqlCommand(sSelectSQLKlienciAktywni, sqlConnectionKlienci)
                    sqlCommandDeleteKlienci = New SqlClient.SqlCommand(sDeleteSQLKlienci, sqlConnectionKlienci)
                    sqlCommandUpdateKlienci = New SqlClient.SqlCommand(sUpdateSQLKlienci, sqlConnectionKlienci)
                    sqlCommandInsertKlienci = New SqlClient.SqlCommand(sInsertSQLKlienci, sqlConnectionKlienci)
                    sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)


                    'mapujemy domyslna nazwe tabeli
                    Dim MapowanieTabeliKlienci As System.Data.Common.DataTableMapping
                    MapowanieTabeliKlienci = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

                    SQLDeleteKlienci(sqlCommandDeleteKlienci, sqlDataAdapterKlienci)
                    SQLUpdateKlienci(sqlCommandUpdateKlienci, sqlDataAdapterKlienci)
                    SQLInsertKlienci(sqlCommandInsertKlienci, sqlDataAdapterKlienci)

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
                            'dbDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                            'dbDataAdapterKlienci.Fill(DataSetKlienci)
                            'dbConnectionKlienci.Close()
                        Else
                            sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                            sqlDataAdapterKlienci.Fill(DataSetKlienci)
                            sqlConnectionKlienci.Close()
                        End If

                        'zdefiniuj unikalny klucz wyszukiwania w tabeli
                        DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("IDENTKLIENTA")}

                        'definiuj DataView
                        DataViewKlienci = New DataView(DataSetKlienci.Tables(0))

                        'If DataSetKlienci.Tables(0).Rows.Count > 0 Then
                        'For idr = DataSetKlienci.Tables(0).Rows.Count - 1 To 0 Step -1
                        '    Select Case AktMiesAnalizy
                        '        Case "01"
                        '            DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA01") = 0
                        '        Case "02"
                        '            DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA02") = 0
                        '        Case "03"
                        '            DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA03") = 0
                        '        Case "04"
                        '            DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA04") = 0
                        '        Case "05"
                        '            DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA05") = 0
                        '        Case "06"
                        '            DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA06") = 0
                        '        Case "07"
                        '            DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA07") = 0
                        '        Case "08"
                        '            DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA08") = 0
                        '        Case "09"
                        '            DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA09") = 0
                        '        Case "10"
                        '            DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA10") = 0
                        '        Case "11"
                        '            DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA11") = 0
                        '        Case "12"
                        '            DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA12") = 0
                        '    End Select
                        'Next
                        'AktualizujBazeKlienci()
                        'End If




                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    End Try

                    '---------------------------------
                End If









                'analizuj plik tekstowy
                FileOpen(FileNum, TxtPlik, OpenMode.Input)
                ProgressBar.Maximum = IleRec
                Do Until EOF(FileNum)
                    NextLine = LineInput(FileNum)
                    BiezacyRekord = BiezacyRekord + 1
                    ProgressBar.Value = BiezacyRekord
                    '---------------------------------
                    SymbTOFI = 0
                    DataTOFI = ""

                    AIlosc = 0
                    AWartosc = 0
                    AMarza = 0

                    LIlosc = 0
                    LWartosc = 0
                    LMarza = 0

                    AoIlosc = 0
                    AoWartosc = 0
                    AoMarza = 0

                    LoIlosc = 0
                    LoWartosc = 0
                    LoMarza = 0

                    NAJEMIlosc = 0
                    NAJEMWartosc = 0
                    NAJEMMarza = 0

                    STRONYIlosc = 0
                    STRONYWartosc = 0
                    STRONYMarza = 0

                    DRUKARKIIlosc = 0
                    DRUKARKIWartosc = 0
                    DRUKARKIMarza = 0

                    SKUPIlosc = 0
                    SKUPWartosc = 0
                    SKUPMarza = 0

                    oNryKartPaybackKli = ""

                    'NrTOFI,DataTOFI,AIlosc,AWartosc,LIlosc,LWartosc,AoIlosc,AoWartosc,LoIlosc,LoWartosc,KartyPayback
                    '-------------------------------------------------
                    '302251303,20140204,0,0.00,2,280.99,0,0.00,0,0.00,"11111.22222.33333"
                    '1151842,20140204,0,0.00,2,776.86,0,0.00,0,0.00,"11111.22222.33333"
                    '221504240,20140204,0,0.00,1,95.04,0,0.00,0,0.00,"11111.22222.33333"

                    ''analizuj pobrany rekord
                    'ip = InStr(NextLine, ",")
                    'If ip > 0 Then
                    '    SymbTOFI = CLng(Mid(NextLine, 1, ip - 1))
                    '    NextLine = Mid(NextLine, ip + 1)
                    '    ip = InStr(NextLine, ",")
                    '    If ip > 0 Then
                    '        DataTOFI = Mid(NextLine, 1, ip - 1)
                    '        NextLine = Mid(NextLine, ip + 1)
                    '        ip = InStr(NextLine, ",")
                    '        If ip > 0 Then
                    '            AIlosc = Val(Mid(NextLine, 1, ip - 1))
                    '            NextLine = Mid(NextLine, ip + 1)
                    '            ip = InStr(NextLine, ",")
                    '            If ip > 0 Then
                    '                AWartosc = Val(Mid(NextLine, 1, ip - 1))
                    '                NextLine = Mid(NextLine, ip + 1)
                    '                ip = InStr(NextLine, ",")
                    '                If ip > 0 Then
                    '                    LIlosc = Val(Mid(NextLine, 1, ip - 1))
                    '                    NextLine = Mid(NextLine, ip + 1)
                    '                    ip = InStr(NextLine, ",")
                    '                    If ip > 0 Then
                    '                        LWartosc = Val(Mid(NextLine, 1, ip - 1))
                    '                        NextLine = Mid(NextLine, ip + 1)
                    '                        ip = InStr(NextLine, ",")
                    '                        If ip > 0 Then
                    '                            AoIlosc = Val(Mid(NextLine, 1, ip - 1))
                    '                            NextLine = Mid(NextLine, ip + 1)
                    '                            ip = InStr(NextLine, ",")
                    '                            If ip > 0 Then
                    '                                AoWartosc = Val(Mid(NextLine, 1, ip - 1))
                    '                                NextLine = Mid(NextLine, ip + 1)
                    '                                ip = InStr(NextLine, ",")
                    '                                If ip > 0 Then
                    '                                    LoIlosc = Val(Mid(NextLine, 1, ip - 1))
                    '                                    NextLine = Mid(NextLine, ip + 1)
                    '                                    ip = InStr(NextLine, ",")
                    '                                    If ip > 0 Then
                    '                                        LoWartosc = Val(Mid(NextLine, 1, ip - 1))
                    '                                        oNryKartPaybackKli = Mid(NextLine, ip + 1)
                    '                                        If oNryKartPaybackKli = """" Then oNryKartPaybackKli = ""
                    '                                    End If
                    '                                End If
                    '                            End If
                    '                        End If
                    '                    Else
                    '                        LWartosc = Val(NextLine)
                    '                    End If
                    '                End If
                    '            End If
                    '        End If
                    '    End If
                    'End If




                    'Nowa STRUKTURA maj 2016 - 15 przecinków
                    'NrTOFI,IdOdzialu,DataTOFI,AIlosc,AWartosc,AMarza,LIlosc,LWartosc,LMarza,AoIlosc,AoWartosc,AoMarza,LoIlosc,LoWartosc,LoMarza,KartyPayback
                    '-------------------------------------------------
                    'Nowa struktura marzec 2018 - 26 przecinków (= 15 +3 +3 +3 +2)
                    '119,0,20180104,2,76.00,29.30,0,0.00,0.00,0,0.00,0.00,0,0.00,0.00,15,5555.55,444.44,13,13333.33,111.00,11111,999.99,111.00,702,2248.00,"0143301632"
                    '119,0,20180108,0,0.00,0.00,4,640.00,199.00,0,0.00,0.00,0,0.00,0.00,0,0.00,0.00,0,0.00,0.00,0,0.00,0.00,0,0.00,"0143301632"
                    '119,0,20180109,2,95.20,35.20,2,507.00,172.00,0,0.00,0.00,0,0.00,0.00,0,0.00,0.00,0,0.00,0.00,0,0.00,0.00,0,0.00,"0143301632"

                    '119,                     nr Tofi
                    '0,                       nr Oddzia³u
                    '20180104,                data
                    '2,76.00,29.30,           iloœæ,  wartoœæ, mar¿a dla Atrament Pryzmat 
                    '0,0.00,0.00,             iloœæ,  wartoœæ, mar¿a dla Laser Pryzmat
                    '0,0.00,0.00,             iloœæ,  wartoœæ, mar¿a dla Atrament Orygina³y
                    '0,0.00,0.00,             iloœæ,  wartoœæ, mar¿a dla Laser Orygina³y
                    '  15,5555.55,444.44,     iloœæ,  wartoœæ, mar¿a dla najmu
                    '  13,13333.33,111.00,    iloœæ,  wartoœæ, mar¿a dla drukarek                                      „tutaj 
                    '  11111,999.99,111.00,   iloœæ,  wartoœæ, mar¿a dla stron                                         odwrotnie do Pana treœci”
                    '  702,2248.00,           iloœæ,  wartoœæ dla skupu
                    '"0143301632"             karty PAYBACK
                    '-------------------------------------------------

                    'policz iloœæ przecinków w NextLine
                    Dim IlePrzec As Integer = 0
                    IlePrzec = NextLine.Split(",").Length - 1

                    If IlePrzec = 26 Then
                        'OBSLUGUJEMY STRUKTURE MARZEC 2018
                        'analizuj pobrany rekord
                        ip = InStr(NextLine, ",")
                        If ip > 0 Then
                            SymbTOFI = CLng(Mid(NextLine, 1, ip - 1))
                            NextLine = Mid(NextLine, ip + 1)
                            ip = InStr(NextLine, ",")
                            If ip > 0 Then
                                IDOddzialu = CLng(Mid(NextLine, 1, ip - 1))
                                NextLine = Mid(NextLine, ip + 1)
                                ip = InStr(NextLine, ",")
                                If ip > 0 Then
                                    DataTOFI = Mid(NextLine, 1, ip - 1)
                                    NextLine = Mid(NextLine, ip + 1)

                                    ip = InStr(NextLine, ",")
                                    If ip > 0 Then
                                        AIlosc = Val(Mid(NextLine, 1, ip - 1))
                                        NextLine = Mid(NextLine, ip + 1)
                                        ip = InStr(NextLine, ",")
                                        If ip > 0 Then
                                            AWartosc = Val(Mid(NextLine, 1, ip - 1))
                                            NextLine = Mid(NextLine, ip + 1)
                                            ip = InStr(NextLine, ",")
                                            If ip > 0 Then
                                                AMarza = Val(Mid(NextLine, 1, ip - 1))
                                                NextLine = Mid(NextLine, ip + 1)

                                                ip = InStr(NextLine, ",")
                                                If ip > 0 Then
                                                    LIlosc = Val(Mid(NextLine, 1, ip - 1))
                                                    NextLine = Mid(NextLine, ip + 1)
                                                    ip = InStr(NextLine, ",")
                                                    If ip > 0 Then
                                                        LWartosc = Val(Mid(NextLine, 1, ip - 1))
                                                        NextLine = Mid(NextLine, ip + 1)
                                                        ip = InStr(NextLine, ",")
                                                        If ip > 0 Then
                                                            LMarza = Val(Mid(NextLine, 1, ip - 1))
                                                            NextLine = Mid(NextLine, ip + 1)

                                                            ip = InStr(NextLine, ",")
                                                            If ip > 0 Then
                                                                AoIlosc = Val(Mid(NextLine, 1, ip - 1))
                                                                NextLine = Mid(NextLine, ip + 1)
                                                                ip = InStr(NextLine, ",")
                                                                If ip > 0 Then
                                                                    AoWartosc = Val(Mid(NextLine, 1, ip - 1))
                                                                    NextLine = Mid(NextLine, ip + 1)
                                                                    ip = InStr(NextLine, ",")
                                                                    If ip > 0 Then
                                                                        AoMarza = Val(Mid(NextLine, 1, ip - 1))
                                                                        NextLine = Mid(NextLine, ip + 1)

                                                                        ip = InStr(NextLine, ",")
                                                                        If ip > 0 Then
                                                                            LoIlosc = Val(Mid(NextLine, 1, ip - 1))
                                                                            NextLine = Mid(NextLine, ip + 1)
                                                                            ip = InStr(NextLine, ",")
                                                                            If ip > 0 Then
                                                                                LoWartosc = Val(Mid(NextLine, 1, ip - 1))
                                                                                NextLine = Mid(NextLine, ip + 1)
                                                                                ip = InStr(NextLine, ",")
                                                                                If ip > 0 Then
                                                                                    LoMarza = Val(Mid(NextLine, 1, ip - 1))
                                                                                    NextLine = Mid(NextLine, ip + 1)

                                                                                    ip = InStr(NextLine, ",")
                                                                                    If ip > 0 Then
                                                                                        NAJEMIlosc = Val(Mid(NextLine, 1, ip - 1))
                                                                                        NextLine = Mid(NextLine, ip + 1)
                                                                                        ip = InStr(NextLine, ",")
                                                                                        If ip > 0 Then
                                                                                            NAJEMWartosc = Val(Mid(NextLine, 1, ip - 1))
                                                                                            NextLine = Mid(NextLine, ip + 1)
                                                                                            ip = InStr(NextLine, ",")
                                                                                            If ip > 0 Then
                                                                                                NAJEMMarza = Val(Mid(NextLine, 1, ip - 1))
                                                                                                NextLine = Mid(NextLine, ip + 1)

                                                                                                ip = InStr(NextLine, ",")
                                                                                                If ip > 0 Then
                                                                                                    DRUKARKIIlosc = Val(Mid(NextLine, 1, ip - 1))
                                                                                                    NextLine = Mid(NextLine, ip + 1)
                                                                                                    ip = InStr(NextLine, ",")
                                                                                                    If ip > 0 Then
                                                                                                        DRUKARKIWartosc = Val(Mid(NextLine, 1, ip - 1))
                                                                                                        NextLine = Mid(NextLine, ip + 1)
                                                                                                        ip = InStr(NextLine, ",")
                                                                                                        If ip > 0 Then
                                                                                                            DRUKARKIMarza = Val(Mid(NextLine, 1, ip - 1))
                                                                                                            NextLine = Mid(NextLine, ip + 1)

                                                                                                            ip = InStr(NextLine, ",")
                                                                                                            If ip > 0 Then
                                                                                                                STRONYIlosc = Val(Mid(NextLine, 1, ip - 1))
                                                                                                                NextLine = Mid(NextLine, ip + 1)
                                                                                                                ip = InStr(NextLine, ",")
                                                                                                                If ip > 0 Then
                                                                                                                    STRONYWartosc = Val(Mid(NextLine, 1, ip - 1))
                                                                                                                    NextLine = Mid(NextLine, ip + 1)
                                                                                                                    ip = InStr(NextLine, ",")
                                                                                                                    If ip > 0 Then
                                                                                                                        STRONYMarza = Val(Mid(NextLine, 1, ip - 1))
                                                                                                                        NextLine = Mid(NextLine, ip + 1)

                                                                                                                        ip = InStr(NextLine, ",")
                                                                                                                        If ip > 0 Then
                                                                                                                            SKUPIlosc = Val(Mid(NextLine, 1, ip - 1))
                                                                                                                            NextLine = Mid(NextLine, ip + 1)
                                                                                                                            ip = InStr(NextLine, ",")
                                                                                                                            If ip > 0 Then
                                                                                                                                SKUPWartosc = Val(Mid(NextLine, 1, ip - 1))
                                                                                                                                NextLine = Mid(NextLine, ip + 1)

                                                                                                                                oNryKartPaybackKli = NextLine
                                                                                                                                If oNryKartPaybackKli = """" Then oNryKartPaybackKli = ""
                                                                                                                            End If
                                                                                                                        End If
                                                                                                                    End If
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If





                    Else
                        'OBSLUGUJEMY STRUKTURE MAJ 2016
                        'analizuj pobrany rekord
                        ip = InStr(NextLine, ",")
                        If ip > 0 Then
                            SymbTOFI = CLng(Mid(NextLine, 1, ip - 1))
                            NextLine = Mid(NextLine, ip + 1)
                            ip = InStr(NextLine, ",")
                            If ip > 0 Then
                                IDOddzialu = CLng(Mid(NextLine, 1, ip - 1))
                                NextLine = Mid(NextLine, ip + 1)
                                ip = InStr(NextLine, ",")
                                If ip > 0 Then
                                    DataTOFI = Mid(NextLine, 1, ip - 1)
                                    NextLine = Mid(NextLine, ip + 1)
                                    ip = InStr(NextLine, ",")
                                    If ip > 0 Then
                                        AIlosc = Val(Mid(NextLine, 1, ip - 1))
                                        NextLine = Mid(NextLine, ip + 1)
                                        ip = InStr(NextLine, ",")
                                        If ip > 0 Then
                                            AWartosc = Val(Mid(NextLine, 1, ip - 1))
                                            NextLine = Mid(NextLine, ip + 1)
                                            ip = InStr(NextLine, ",")
                                            If ip > 0 Then
                                                AMarza = Val(Mid(NextLine, 1, ip - 1))
                                                NextLine = Mid(NextLine, ip + 1)
                                                ip = InStr(NextLine, ",")
                                                If ip > 0 Then
                                                    LIlosc = Val(Mid(NextLine, 1, ip - 1))
                                                    NextLine = Mid(NextLine, ip + 1)
                                                    ip = InStr(NextLine, ",")
                                                    If ip > 0 Then
                                                        LWartosc = Val(Mid(NextLine, 1, ip - 1))
                                                        NextLine = Mid(NextLine, ip + 1)
                                                        ip = InStr(NextLine, ",")
                                                        If ip > 0 Then
                                                            LMarza = Val(Mid(NextLine, 1, ip - 1))
                                                            NextLine = Mid(NextLine, ip + 1)
                                                            ip = InStr(NextLine, ",")
                                                            If ip > 0 Then
                                                                AoIlosc = Val(Mid(NextLine, 1, ip - 1))
                                                                NextLine = Mid(NextLine, ip + 1)
                                                                ip = InStr(NextLine, ",")
                                                                If ip > 0 Then
                                                                    AoWartosc = Val(Mid(NextLine, 1, ip - 1))
                                                                    NextLine = Mid(NextLine, ip + 1)
                                                                    ip = InStr(NextLine, ",")
                                                                    If ip > 0 Then
                                                                        AoMarza = Val(Mid(NextLine, 1, ip - 1))
                                                                        NextLine = Mid(NextLine, ip + 1)
                                                                        ip = InStr(NextLine, ",")
                                                                        If ip > 0 Then
                                                                            LoIlosc = Val(Mid(NextLine, 1, ip - 1))
                                                                            NextLine = Mid(NextLine, ip + 1)
                                                                            ip = InStr(NextLine, ",")
                                                                            If ip > 0 Then
                                                                                LoWartosc = Val(Mid(NextLine, 1, ip - 1))
                                                                                NextLine = Mid(NextLine, ip + 1)
                                                                                ip = InStr(NextLine, ",")
                                                                                If ip > 0 Then
                                                                                    LoMarza = Val(Mid(NextLine, 1, ip - 1))
                                                                                    oNryKartPaybackKli = Mid(NextLine, ip + 1)
                                                                                    If oNryKartPaybackKli = """" Then oNryKartPaybackKli = ""
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If





                    End If


                    'konwersja Daty RRRRMMDD -> RRRR-MM-DD
                    DataObrotu = Mid(DataTOFI, 1, 4) & "-" & Mid(DataTOFI, 5, 2) & "-" & Mid(DataTOFI, 7, 2)
                    SymbKlienta = ""
                    IlePrzeczytano += 1

                    'przetlumacz symbol TOFI na symbol klienta
                    If SymbTOFI = 0 Then
                        'brak symbolu TOFI - nie mam co zapisywac
                        IleOdrzucono += 1
                    Else
                        DataSetKlienci = New DataSet
                        DataViewKlienci = New DataView
                        SymbolKlienta = ""
                        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

                        '2015-08-30
                        'Jeœli jest jeden klient o kilku nrach TOFI, to bowinien zostaæ obsluzony niezale¿nie od wyboru nru tofi
                        'szukany jest ka¿dy nr TOFI po kolei
                        'Jeœli jest kilku klientów z tym samym nrem TOFI, to kwerenda powinna daæ wiecej ni¿ jeden rekord
                        'i wszystkie trzeba aktualizowaæ...

                        'uzupelnienie symbolu TOFI o id oddzialu
                        If IDOddzialu <> 0 Then
                            If Len(Trim(Str(SymbTOFI))) < 5 Then
                                'symbol tofi krótszy od 5 znaków - uzupe³niamy zerami z lewej - PO STAREMU
                                SymbTOFITxt = Microsoft.VisualBasic.Right("00000" & Trim(Str(SymbTOFI)), 5) & Trim(Str(IDOddzialu))
                            Else
                                'symbol d³u¿szy lub równy 5 znaków - bierzemy wprost - DLA LITWY !!!!
                                SymbTOFITxt = Trim(Str(SymbTOFI)) & Trim(Str(IDOddzialu))
                            End If
                        Else
                            'na Litwie nie ma oddzia³ów
                            If Len(Trim(Str(SymbTOFI))) < 5 Then
                                'symbol tofi krótszy od 5 znaków - uzupe³niamy zerami z lewej - PO STAREMU
                                SymbTOFITxt = Microsoft.VisualBasic.Right("00000" & Trim(Str(SymbTOFI)), 5)
                            Else
                                'symbol d³u¿szy lub równy 5 znaków - bierzemy wprost - DLA LITWY !!!!
                                SymbTOFITxt = Trim(Str(SymbTOFI))
                            End If
                        End If


                        If ParL_DataBaseType = "MS ACCESS" Then
                            ''dbSelectSQLKlienci = "SELECT * FROM Klienci WHERE NRTOFITXT LIKE '%" & Microsoft.VisualBasic.Right("00000" & Trim(Str(SymbTOFI)), 5) & "%' "

                            ''If Len(Trim(Str(SymbTOFI))) < 5 Then
                            ''    'symbol tofi krótszy od 5 znaków - uzupe³niamy zerami z lewej - PO STAREMU
                            ''    dbSelectSQLKlienci = sSelectSQLKlienci & " WHERE NRTOFITXT LIKE '%" & Microsoft.VisualBasic.Right("00000" & Trim(Str(SymbTOFI)), 5) & "%' "
                            ''Else
                            ''    'symbol d³u¿szy lub równy 5 znaków - bierzemy wprost - DLA LITWY !!!!
                            ''    dbSelectSQLKlienci = sSelectSQLKlienci & " WHERE NRTOFITXT LIKE '%" & Trim(Str(SymbTOFI)) & "%' "
                            ''End If

                            'dbConnectionKlienci = New OleDb.OleDbConnection(sConnectionKlienci)
                            'dbCommandSelectKlienci = New OleDb.OleDbCommand(sSelectSQLKlienci & " WHERE NRTOFITXT LIKE '%" & SymbTOFITxt & "%' ", dbConnectionKlienci)
                            'dbCommandUpdateKlienci = New OleDb.OleDbCommand(sUpdateSQLKlienci, dbConnectionKlienci)
                            'dbDataAdapterKlienci = New OleDb.OleDbDataAdapter(dbCommandSelectKlienci)

                            'Dim MapowanieTabeliKlienci As System.Data.Common.DataTableMapping
                            'MapowanieTabeliKlienci = dbDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

                            ''DBDeleteKlienci(dbCommandDeleteKlienci, dbDataAdapterKlienci)
                            'DBUpdateKlienci(dbCommandUpdateKlienci, dbDataAdapterKlienci)
                            ''DBInsertKlienci(dbCommandInsertKlienci, dbDataAdapterKlienci)

                            '' Add the RowUpdated event handler.
                            'AddHandler dbDataAdapterKlienci.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

                            'Try
                            '    dbConnectionKlienci.Open()
                            'Catch Ex As System.Exception
                            'Finally
                            '    ConnectionState = dbConnectionKlienci.State
                            'End Try
                        Else
                            'dbSelectSQLKlienci = "SELECT * FROM Klienci WHERE NRTOFITXT LIKE '%" & Microsoft.VisualBasic.Right("00000" & Trim(Str(SymbTOFI)), 5) & "%' "

                            'If Len(Trim(Str(SymbTOFI))) < 5 Then
                            '    'symbol tofi krótszy od 5 znaków - uzupe³niamy zerami z lewej - PO STAREMU
                            '    dbSelectSQLKlienci = sSelectSQLKlienci & " WHERE NRTOFITXT LIKE '%" & Microsoft.VisualBasic.Right("00000" & Trim(Str(SymbTOFI)), 5) & "%' "
                            'Else
                            '    'symbol d³u¿szy lub równy 5 znaków - bierzemy wprost - DLA LITWY !!!!
                            '    dbSelectSQLKlienci = sSelectSQLKlienci & " WHERE NRTOFITXT LIKE '%" & Trim(Str(SymbTOFI)) & "%' "
                            'End If

                            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
                            sqlCommandSelectKlienci = New SqlClient.SqlCommand(sSelectSQLKlienci & " WHERE NRTOFITXT LIKE '%" & SymbTOFITxt & "%' ", sqlConnectionKlienci)
                            sqlCommandUpdateKlienci = New SqlClient.SqlCommand(sUpdateSQLKlienci, sqlConnectionKlienci)
                            sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

                            Dim MapowanieTabeliKlienci As System.Data.Common.DataTableMapping
                            MapowanieTabeliKlienci = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")
                            '------------------------------------------

                            'SQLDeleteKlienci(sqlCommandDeleteKlienci, sqlDataAdapterKlienci)
                            SQLUpdateKlienci(sqlCommandUpdateKlienci, sqlDataAdapterKlienci)
                            'SQLInsertKlienci(sqlCommandInsertKlienci, sqlDataAdapterKlienci)

                            ' Add the RowUpdated event handler.
                            AddHandler sqlDataAdapterKlienci.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)
                            '------------------------------------------
                            Try
                                sqlConnectionKlienci.Open()
                            Catch Ex As System.Exception
                                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
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

                                'definiuj DataView
                                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
                                If DataViewKlienci.Count > 0 Then
                                    Dim pos As Integer = 0
                                    Dim jestOK As Boolean = False
                                    'mamy klientów z takim symbolem TOFI
                                    For ik = 0 To DataViewKlienci.Count - 1
                                        'musimy przeanalizowaæ czy znaleziono poprawny symbol TOFI
                                        'szukaliœmy 123451 " WHERE NRTOFITXT LIKE '%" & SymbTOFITxt & "%' "
                                        'mogliœmy znaleŸæ 123451, 12345610, 12345611 itp
                                        'analizuj znak PO znalezionym symbolu - jeœli spacja, przec lub nic - ok / jeœli cyfra - b³¹d
                                        'Podobnie mogliœmy znaleŸæ 9123451, 99123451 itp
                                        'analizuj znak PRZED znalezionym symbolem - jeœli spacja, przec lub nic - ok / jeœli cyfra - b³¹d
                                        oNrTOFITxtKli = DataViewKlienci.Item(ik).Item("NRTOFITXT")
                                        Do While Len(oNrTOFITxtKli) > 0
                                            jestOK = False
                                            pos = InStr(oNrTOFITxtKli, SymbTOFITxt)
                                            If pos > 0 Then
                                                'sprawdzamy znak ZA i PRZED znalezionym nr TOFI
                                                Select Case Mid(oNrTOFITxtKli, pos + Len(SymbTOFITxt), 1)
                                                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                                                        'ten numer TOFI ma wiêcej cyfr...
                                                        'to nie jest ten NRTOFI który szukamy
                                                        jestOK = False
                                                        oNrTOFITxtKli = Mid(oNrTOFITxtKli, pos + Len(SymbTOFITxt))

                                                    Case Else
                                                        'ten nr TOFI jest OK
                                                        'sprawdŸmy jeszcze znak bezpoœrednio PRZED
                                                        If pos = 1 Then
                                                            'nie ma znaku PRZED - znaleziony nrTOFI jest OK
                                                            jestOK = True
                                                            oNrTOFITxtKli = ""
                                                        Else
                                                            Select Case Mid(oNrTOFITxtKli, pos - 1, 1)
                                                                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                                                                    'ten numer TOFI ma wiêcej cyfr...
                                                                    'to nie jest ten NRTOFI który szukamy
                                                                    jestOK = False
                                                                    oNrTOFITxtKli = Mid(oNrTOFITxtKli, pos + Len(SymbTOFITxt))

                                                                Case Else
                                                                    'ten nr TOFI jest OK
                                                                    jestOK = True
                                                                    oNrTOFITxtKli = ""
                                                            End Select
                                                        End If

                                                End Select

                                                'sprawdz czy jet OK...
                                                If jestOK Then
                                                    SymbKlienta = DataViewKlienci.Item(ik).Item(0)
                                                    oWersjaKli = DataViewKlienci.Item(ik).Item("WERSJA")
                                                    '---------------------------------
                                                    'aktualizujemy informacje o PAYBACK
                                                    If Len(Trim(oNryKartPaybackKli)) = 0 Then
                                                        DataViewKlienci.Item(ik).Item("KARTAPAYBACK") = False
                                                        DataViewKlienci.Item(ik).Item("NRYKARTPAYBACK") = ""
                                                        'ElseIf Trim(oNryKartPaybackKli) = "NIE" Then
                                                        '    DataViewKlienci.Item(ik).Item("KARTAPAYBACK") = False
                                                        '    DataViewKlienci.Item(ik).Item("NRYKARTPAYBACK") = ""
                                                    Else
                                                        DataViewKlienci.Item(ik).Item("KARTAPAYBACK") = True
                                                        DataViewKlienci.Item(ik).Item("NRYKARTPAYBACK") = PominCudzyslowy(oNryKartPaybackKli)
                                                    End If
                                                    ''---------------------------------
                                                    ''aktualizuj sprzedaz zamiast prognozy...
                                                    'Select Case AktMiesAnalizy
                                                    '    Case "01"
                                                    '        DataViewKlienci.Item(ik).Item("PROGNOZA01") += AWartosc + LWartosc + LoWartosc + AoWartosc
                                                    '    Case "02"
                                                    '        DataViewKlienci.Item(ik).Item("PROGNOZA02") += AWartosc + LWartosc + LoWartosc + AoWartosc
                                                    '    Case "03"
                                                    '        DataViewKlienci.Item(ik).Item("PROGNOZA03") += AWartosc + LWartosc + LoWartosc + AoWartosc
                                                    '    Case "04"
                                                    '        DataViewKlienci.Item(ik).Item("PROGNOZA04") += AWartosc + LWartosc + LoWartosc + AoWartosc
                                                    '    Case "05"
                                                    '        DataViewKlienci.Item(ik).Item("PROGNOZA05") += AWartosc + LWartosc + LoWartosc + AoWartosc
                                                    '    Case "06"
                                                    '        DataViewKlienci.Item(ik).Item("PROGNOZA06") += AWartosc + LWartosc + LoWartosc + AoWartosc
                                                    '    Case "07"
                                                    '        DataViewKlienci.Item(ik).Item("PROGNOZA07") += AWartosc + LWartosc + LoWartosc + AoWartosc
                                                    '    Case "08"
                                                    '        DataViewKlienci.Item(ik).Item("PROGNOZA08") += AWartosc + LWartosc + LoWartosc + AoWartosc
                                                    '    Case "09"
                                                    '        DataViewKlienci.Item(ik).Item("PROGNOZA09") += AWartosc + LWartosc + LoWartosc + AoWartosc
                                                    '    Case "10"
                                                    '        DataViewKlienci.Item(ik).Item("PROGNOZA10") += AWartosc + LWartosc + LoWartosc + AoWartosc
                                                    '    Case "11"
                                                    '        DataViewKlienci.Item(ik).Item("PROGNOZA11") += AWartosc + LWartosc + LoWartosc + AoWartosc
                                                    '    Case "12"
                                                    '        DataViewKlienci.Item(ik).Item("PROGNOZA12") += AWartosc + LWartosc + LoWartosc + AoWartosc
                                                    'End Select
                                                    ''---------------------------------
                                                    DataViewKlienci.Item(ik).Item("WERSJA") = ZmienWersje(oWersjaKli) 'Wersja         Integer, 2
                                                    AktualizujBazeKlienci()



                                                    '---------------------------------
                                                    ' Create an array for the key values to find.
                                                    Dim findObroty(1) As Object
                                                    findObroty(0) = SymbKlienta
                                                    findObroty(1) = DataObrotu
                                                    ObrotyRow = DataSetObroty.Tables(0).Rows.Find(findObroty)
                                                    ' sprawdz czy jest...
                                                    If Not (ObrotyRow Is Nothing) Then
                                                        oOperacja = "Edycja"
                                                        'oWersjaKli = KlienciRow("WERSJA")
                                                        'aktualizuj rekord

                                                        ObrotyRow("LILOSPRZED") += LIlosc
                                                        ObrotyRow("LWARTSPRZED") += LWartosc
                                                        ObrotyRow("LMARSPRZED") += LMarza

                                                        ObrotyRow("AILOSPRZED") += AIlosc
                                                        ObrotyRow("AWARTSPRZED") += AWartosc
                                                        ObrotyRow("AMARSPRZED") += AMarza

                                                        ObrotyRow("LOILOSPRZED") += LoIlosc
                                                        ObrotyRow("LOWARTSPRZED") += LoWartosc
                                                        ObrotyRow("LOMARSPRZED") += LoMarza

                                                        ObrotyRow("AOILOSPRZED") += AoIlosc
                                                        ObrotyRow("AOWARTSPRZED") += AoWartosc
                                                        ObrotyRow("AOMARSPRZED") += AoMarza

                                                        ObrotyRow("NAJEMILOSPRZED") += NAJEMIlosc
                                                        ObrotyRow("NAJEMWARTSPRZED") += NAJEMWartosc
                                                        ObrotyRow("NAJEMMARSPRZED") += NAJEMMarza

                                                        ObrotyRow("DRUKARKIILOSPRZED") += DRUKARKIIlosc
                                                        ObrotyRow("DRUKARKIWARTSPRZED") += DRUKARKIWartosc
                                                        ObrotyRow("DRUKARKIMARSPRZED") += DRUKARKIMarza

                                                        ObrotyRow("STRONYILOSPRZED") += STRONYIlosc
                                                        ObrotyRow("STRONYWARTSPRZED") += STRONYWartosc
                                                        ObrotyRow("STRONYMARSPRZED") += STRONYMarza

                                                        ObrotyRow("SKUPILOSPRZED") += SKUPIlosc
                                                        ObrotyRow("SKUPWARTSPRZED") += SKUPWartosc
                                                        ObrotyRow("SKUPMARSPRZED") += SKUPMarza

                                                        IleZaktualizowano += 1
                                                    Else
                                                        oOperacja = "Nowy"
                                                        oWersjaKli = 0
                                                        'stworz nowy rekord
                                                        ObrotyRow = DataSetObroty.Tables(0).NewRow()

                                                        ObrotyRow("IDENTKLIENTA") = SymbKlienta
                                                        ObrotyRow("DATA") = DataObrotu

                                                        ObrotyRow("LILOSPRZED") = LIlosc
                                                        ObrotyRow("LWARTSPRZED") = LWartosc
                                                        ObrotyRow("LMARSPRZED") = LMarza

                                                        ObrotyRow("AILOSPRZED") = AIlosc
                                                        ObrotyRow("AWARTSPRZED") = AWartosc
                                                        ObrotyRow("AMARSPRZED") = AMarza

                                                        ObrotyRow("LOILOSPRZED") = LoIlosc
                                                        ObrotyRow("LOWARTSPRZED") = LoWartosc
                                                        ObrotyRow("LOMARSPRZED") = LoMarza

                                                        ObrotyRow("AOILOSPRZED") = AoIlosc
                                                        ObrotyRow("AOWARTSPRZED") = AoWartosc
                                                        ObrotyRow("AOMARSPRZED") = AoMarza

                                                        ObrotyRow("NAJEMILOSPRZED") = NAJEMIlosc
                                                        ObrotyRow("NAJEMWARTSPRZED") = NAJEMWartosc
                                                        ObrotyRow("NAJEMMARSPRZED") = NAJEMMarza

                                                        ObrotyRow("DRUKARKIILOSPRZED") = DRUKARKIIlosc
                                                        ObrotyRow("DRUKARKIWARTSPRZED") = DRUKARKIWartosc
                                                        ObrotyRow("DRUKARKIMARSPRZED") = DRUKARKIMarza

                                                        ObrotyRow("STRONYILOSPRZED") = STRONYIlosc
                                                        ObrotyRow("STRONYWARTSPRZED") = STRONYWartosc
                                                        ObrotyRow("STRONYMARSPRZED") = STRONYMarza

                                                        ObrotyRow("SKUPILOSPRZED") = SKUPIlosc
                                                        ObrotyRow("SKUPWARTSPRZED") = SKUPWartosc
                                                        ObrotyRow("SKUPMARSPRZED") = SKUPMarza

                                                        ObrotyRow("WERSJA") = 1     'Wersja         Integer, 2

                                                        'dodaj rekord do DataSet
                                                        IleDopisano += 1
                                                        DataSetObroty.Tables(0).Rows.Add(ObrotyRow)
                                                    End If
                                                    AktualizujBaze()
                                                    '---------------------------------


                                                    '---------------------------------
                                                    ' Create an array for the key values to find.
                                                    Dim findAnalizyZakupu(0) As Object
                                                    findAnalizyZakupu(0) = SymbKlienta
                                                    AnalizyZakupuRow = DataSetAnalizyZakupu.Tables(0).Rows.Find(findAnalizyZakupu)
                                                    ' sprawdz czy jest...
                                                    If Not (AnalizyZakupuRow Is Nothing) Then
                                                        oOperacja = "Edycja"
                                                        aWersjaAnal = AnalizyZakupuRow("WERSJA")
                                                        aIlosBM = AnalizyZakupuRow("ILOSZABM")
                                                        aWartBM = AnalizyZakupuRow("WARTZABM")

                                                        'aktualizuj rekord
                                                        AnalizyZakupuRow("ILOSZABM") = aIlosBM + LIlosc
                                                        AnalizyZakupuRow("WARTZABM") = 0
                                                        AnalizyZakupuRow("WERSJA") = ZmienWersje(aWersjaAnal)
                                                    Else
                                                        oOperacja = "Nowy"
                                                        'stworz nowy rekord
                                                        AnalizyZakupuRow = DataSetAnalizyZakupu.Tables(0).NewRow()

                                                        AnalizyZakupuRow("IDENTKLIENTA") = SymbKlienta
                                                        AnalizyZakupuRow.Item("WARTZA00") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA00") = 0

                                                        AnalizyZakupuRow.Item("WARTZABM") = 0
                                                        AnalizyZakupuRow.Item("ILOSZABM") = LIlosc

                                                        AnalizyZakupuRow.Item("WARTZA01") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA01") = 0
                                                        AnalizyZakupuRow.Item("WARTZA02") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA02") = 0
                                                        AnalizyZakupuRow.Item("WARTZA03") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA03") = 0
                                                        AnalizyZakupuRow.Item("WARTZA04") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA04") = 0
                                                        AnalizyZakupuRow.Item("WARTZA05") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA05") = 0
                                                        AnalizyZakupuRow.Item("WARTZA06") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA06") = 0
                                                        AnalizyZakupuRow.Item("WARTZA07") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA07") = 0
                                                        AnalizyZakupuRow.Item("WARTZA08") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA08") = 0
                                                        AnalizyZakupuRow.Item("WARTZA09") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA09") = 0
                                                        AnalizyZakupuRow.Item("WARTZA10") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA10") = 0
                                                        AnalizyZakupuRow.Item("WARTZA11") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA11") = 0
                                                        AnalizyZakupuRow.Item("WARTZA12") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA12") = 0

                                                        AnalizyZakupuRow.Item("WARTZA13") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA13") = 0
                                                        AnalizyZakupuRow.Item("WARTZA14") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA14") = 0
                                                        AnalizyZakupuRow.Item("WARTZA15") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA15") = 0
                                                        AnalizyZakupuRow.Item("WARTZA16") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA16") = 0
                                                        AnalizyZakupuRow.Item("WARTZA17") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA17") = 0
                                                        AnalizyZakupuRow.Item("WARTZA18") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA18") = 0
                                                        AnalizyZakupuRow.Item("WARTZA19") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA19") = 0
                                                        AnalizyZakupuRow.Item("WARTZA20") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA20") = 0
                                                        AnalizyZakupuRow.Item("WARTZA21") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA21") = 0
                                                        AnalizyZakupuRow.Item("WARTZA22") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA22") = 0
                                                        AnalizyZakupuRow.Item("WARTZA23") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA23") = 0
                                                        AnalizyZakupuRow.Item("WARTZA24") = 0
                                                        AnalizyZakupuRow.Item("ILOSZA24") = 0
                                                        AnalizyZakupuRow.Item("WERSJA") = 1

                                                        'dodaj rekord do DataSet
                                                        DataSetAnalizyZakupu.Tables(0).Rows.Add(AnalizyZakupuRow)
                                                    End If
                                                    AktualizujBazeAnalizyZakupu()
                                                    '---------------------------------
                                                End If
                                            Else
                                                'nie znalaz³em numeru TOFI
                                                oNrTOFITxtKli = ""
                                            End If
                                        Loop
                                    Next
                                Else
                                    Nierozpoz &= SymbTOFITxt & " "
                                    IleOdrzucono += 1
                                End If

                            Catch Ex As System.Exception
                                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                                    System.Windows.Forms.MessageBoxButtons.OK, _
                                    MessageBoxIcon.Information, _
                                    MessageBoxDefaultButton.Button1)
                            End Try
                        End If

                    End If

                    lblIlePrzeczytano.Text = IlePrzeczytano.ToString
                    lblIleDopisano.Text = IleDopisano.ToString
                    lblIleZaktualizowano.Text = IleZaktualizowano.ToString
                    lblIleOdrzucono.Text = IleOdrzucono.ToString
                    System.Windows.Forms.Application.DoEvents() 'pozwol na wykonanie
                Loop
                FileClose(FileNum)
                'AktualizujWartoscObrotowwKartKli()
                cmdImportuj.Enabled = True
                '-------------------------------
                If IleDopisano > 0 Then
                    MessageBox.Show("Pobra³em informacje o obrotach klientów Oddzia³u " & Trim(Par_IdentOddzialu) & vbCrLf & _
                        "przekazane z programu TOFI " & vbCrLf & _
                        IIf(Len(Nierozpoz) > 0, "Nierozpoznane symbole TOFI :" & vbCrLf & Nierozpoz, ""), _
                        "OK, skoñczy³em", _
                        System.Windows.Forms.MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, _
                        MessageBoxDefaultButton.Button1)
                Else
                    MessageBox.Show("Nie znalaz³em w pliku przekazanym z TOFI" & vbCrLf & _
                        "zapisów dotycz¹cych klientów Oddzia³u " & Trim(Par_IdentOddzialu), _
                        "OK, skoñczy³em", _
                        System.Windows.Forms.MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, _
                        MessageBoxDefaultButton.Button1)
                End If
                Me.Close()
            End If
        End If
    End Sub

    Private Sub PobierzObrotyzTOFI_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdPokazPlik_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPokazPlik.Click
        Dim TxtPlik As String = Trim(lblPlik.Text)

        If Len(TxtPlik) = 0 Then
            MessageBox.Show("Brak definicji pliku z TOFI", _
                "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            If Len(Dir(TxtPlik)) = 0 Then
                MessageBox.Show("Nie znalaz³em na dysku pliku z TOFI" & vbCrLf & TxtPlik & vbCrLf, _
                    "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            Else
                'Shell("notepad", AppWinStyle.NormalFocus, False)
                Dim proc As New Process
                System.Windows.Forms.Application.DoEvents()
                proc.StartInfo.FileName = "Notepad.exe"
                proc.StartInfo.Arguments = TxtPlik
                proc.Start()
            End If
        End If
    End Sub



    Private Function PominCudzyslowy(ByVal inTxt As String) As String
        Dim outTxt As String = ""
        Dim i As Integer
        Dim Ch As String = ""
        If Len(intxt) > 0 Then
            For i = 1 To Len(inTxt)
                Ch = Mid(inTxt, i, 1)
                Select Case Ch
                    Case """"
                    Case Else
                        outTxt &= Ch
                End Select
            Next
        End If
        Return (outTxt)
    End Function


End Class

