Public Class PobierzDaneMiesieczne
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtMiesiac As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents cmdAnalizuj As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents chbUsun As System.Windows.Forms.CheckBox
    Friend WithEvents cmdPobierzMiesiac As System.Windows.Forms.Button
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Timer1 As Timer
    Friend WithEvents ProgressBar As System.Windows.Forms.ProgressBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PobierzDaneMiesieczne))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.chbUsun = New System.Windows.Forms.CheckBox()
        Me.cmdPobierzMiesiac = New System.Windows.Forms.Button()
        Me.TxtMiesiac = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.cmdAnalizuj = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.ProgressBar)
        Me.Panel1.Controls.Add(Me.chbUsun)
        Me.Panel1.Controls.Add(Me.cmdPobierzMiesiac)
        Me.Panel1.Controls.Add(Me.TxtMiesiac)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(403, 208)
        Me.Panel1.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(6, 86)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(379, 40)
        Me.Label4.TabIndex = 44
        Me.Label4.Text = "Kartoteka Obrotów Miesiêcznych zawiera dane z nastêpuj¹cych miesiêcy :" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & " "
        '
        'Label3
        '
        Me.Label3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(6, 46)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(379, 40)
        Me.Label3.TabIndex = 43
        Me.Label3.Text = "Obecnie, w Katalogu Obrotów Ostatniego Miesi¹ca znajduj¹ siê dane sprzeda¿owe za " &
    "miesi¹c "
        '
        'ProgressBar
        '
        Me.ProgressBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar.Location = New System.Drawing.Point(8, 192)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(387, 8)
        Me.ProgressBar.TabIndex = 42
        '
        'chbUsun
        '
        Me.chbUsun.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chbUsun.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.chbUsun.ForeColor = System.Drawing.Color.Navy
        Me.chbUsun.Location = New System.Drawing.Point(9, 156)
        Me.chbUsun.Name = "chbUsun"
        Me.chbUsun.Size = New System.Drawing.Size(376, 34)
        Me.chbUsun.TabIndex = 40
        Me.chbUsun.Text = "Usun zsumowane dane z Katalogu Obrotów Ostatniego Miesi¹ca"
        '
        'cmdPobierzMiesiac
        '
        Me.cmdPobierzMiesiac.Image = CType(resources.GetObject("cmdPobierzMiesiac.Image"), System.Drawing.Image)
        Me.cmdPobierzMiesiac.Location = New System.Drawing.Point(201, 130)
        Me.cmdPobierzMiesiac.Name = "cmdPobierzMiesiac"
        Me.cmdPobierzMiesiac.Size = New System.Drawing.Size(32, 22)
        Me.cmdPobierzMiesiac.TabIndex = 39
        '
        'TxtMiesiac
        '
        Me.TxtMiesiac.ForeColor = System.Drawing.Color.Purple
        Me.TxtMiesiac.Location = New System.Drawing.Point(123, 130)
        Me.TxtMiesiac.Name = "TxtMiesiac"
        Me.TxtMiesiac.Size = New System.Drawing.Size(72, 20)
        Me.TxtMiesiac.TabIndex = 37
        Me.TxtMiesiac.Text = "2004"
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(6, 134)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(178, 16)
        Me.Label2.TabIndex = 38
        Me.Label2.Text = "Miesi¹c do analizy . . . . . . . . ."
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(379, 40)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Program analizuje dane o sprzeda¿y z ostatniego miesi¹ca, sumuje i wpisuje do Kar" &
    "toteki Obrotów Miesiêcznych ..."
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(331, 224)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 24)
        Me.cmdPowrot.TabIndex = 12
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdAnalizuj
        '
        Me.cmdAnalizuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdAnalizuj.Image = CType(resources.GetObject("cmdAnalizuj.Image"), System.Drawing.Image)
        Me.cmdAnalizuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAnalizuj.Location = New System.Drawing.Point(243, 224)
        Me.cmdAnalizuj.Name = "cmdAnalizuj"
        Me.cmdAnalizuj.Size = New System.Drawing.Size(80, 24)
        Me.cmdAnalizuj.TabIndex = 13
        Me.cmdAnalizuj.Text = "Analizuj"
        Me.cmdAnalizuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 223)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(88, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 27
        Me.PictureBox2.TabStop = False
        '
        'Timer1
        '
        '
        'PobierzDaneMiesieczne
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(417, 252)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.cmdAnalizuj)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "PobierzDaneMiesieczne"
        Me.ShowInTaskbar = False
        Me.Text = "Pobierz Dane Miesiêczne"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim KlienciOK As Boolean = False
    Dim ObrotyOK As Boolean = False
    Dim ObrotyMiesOK As Boolean = False

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





    Dim dbSelectSQLObrotyMies As String = sSelectSQLObrotyMies

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

    Dim DataSetObrotyMies As New DataSet
    Dim DataViewObrotyMies As New DataView







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



    Private Sub PobierzDaneMiesieczne_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        chbUsun.Checked = False
        TxtMiesiac.Text = Mid(SysData(), 1, 7)
        Timer1.Interval = 200
        Timer1.Enabled = True
    End Sub

    Private Sub PobierzDaneMiesieczne_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        Label3.Text &= PobierzMiesiacDanychZKatDanychOstMiesiaca()
        Label4.Text &= PobierzMiesiaceDanychZKatDanychMiesiecznych()
    End Sub



    Public Function PobierzMiesiacDanychZKatDanychOstMiesiaca() As String
        Dim sqlConnection As SqlClient.SqlConnection
        Dim sqlCommand As SqlClient.SqlCommand
        Dim DataP As String = ""
        Dim DataK As String = ""

        sqlConnection = New SqlClient.SqlConnection(sConnectionObroty)
        sqlCommand = New SqlClient.SqlCommand
        Try
            sqlCommand.Connection = sqlConnection
            sqlCommand.CommandType = CommandType.Text
            sqlConnection.Open()

            sqlCommand.CommandText = "SELECT TOP 1 ISNULL(DATA,'') FROM Obroty ORDER BY DATA ASC "
            DataP = sqlCommand.ExecuteScalar()

            sqlCommand.CommandText = "SELECT TOP 1 ISNULL(DATA,'') FROM Obroty ORDER BY DATA DESC "
            DataK = sqlCommand.ExecuteScalar()

        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        Finally
            sqlConnection.Close()
        End Try

        If Len(Trim(Mid(DataP, 1, 7))) = 0 And Len(Trim(Mid(DataK, 1, 7))) = 0 Then
            Return "BRAK DANYCH"
        ElseIf Mid(DataP, 1, 7) = Mid(DataK, 1, 7) Then
            Return Mid(DataP, 1, 7)
        Else
            Return Mid(DataP, 1, 7) & " - " & Mid(DataK, 1, 7)
        End If
    End Function



    Public Function PobierzMiesiaceDanychZKatDanychMiesiecznych() As String
        Dim sqlConnection As SqlClient.SqlConnection
        Dim sqlCommand As SqlClient.SqlCommand
        Dim MiesP As String = ""
        Dim MiesK As String = ""

        sqlConnection = New SqlClient.SqlConnection(sConnectionObrotyMies)
        sqlCommand = New SqlClient.SqlCommand
        Try
            sqlCommand.Connection = sqlConnection
            sqlCommand.CommandType = CommandType.Text
            sqlConnection.Open()

            sqlCommand.CommandText = "SELECT TOP 1 ISNULL(MIESIAC,'') FROM ObrotyMies ORDER BY MIESIAC ASC "
            MiesP = sqlCommand.ExecuteScalar()

            sqlCommand.CommandText = "SELECT TOP 1 ISNULL(MIESIAC,'') FROM ObrotyMies ORDER BY MIESIAC DESC "
            MiesK = sqlCommand.ExecuteScalar()

        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        Finally
            sqlConnection.Close()
        End Try

        If Len(Trim(Mid(MiesP, 1, 7))) = 0 And Len(Trim(Mid(MiesK, 1, 7))) = 0 Then
            Return "BRAK DANYCH"
        ElseIf Mid(MiesP, 1, 7) = Mid(MiesK, 1, 7) Then
            Return Mid(MiesP, 1, 7)
        Else
            Return Mid(MiesP, 1, 7) & " - " & Mid(MiesK, 1, 7)
        End If
    End Function







    Private Sub cmdPobierzMiesiac_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPobierzMiesiac.Click
        oData = TxtMiesiac.Text & "-01"
        Dim Proc As New WybierzDate
        Proc.ShowDialog()
        TxtMiesiac.Text = Mid(oData, 1, 7)
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub cmdAnalizuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAnalizuj.Click

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
                'DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("IDENTKLIENTA")}
                DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("NRKARTY")}

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

                KlienciOK = True


            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            '---------------------------------
        End If






        If KlienciOK Then
            DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")
            DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")

            ''ObrotyMies
            'Public oIdentMies As String            'IDENTKLIENTA     Text, 6
            'Public oMiesiacMies As String          'MIESIAC          Text,7

            'Public oLWartSprzedMies As Double      'LWARTSPRZED      Double
            'Public oLIloSprzedMies As Double       'LILOSPRZED       Double
            'Public oLMARSprzedMies As Double       'LMARSPRZED       Double
            'Public oAWartSprzedMies As Double      'AWARTSPRZED      Double
            'Public oAIloSprzedMies As Double       'AILOSPRZED       Double
            'Public oAMARSprzedMies As Double       'AMARSPRZED       Double3

            'Public oLOWartSprzedMies As Double      'LOWARTSPRZED      Double
            'Public oLOIloSprzedMies As Double       'LOILOSPRZED       Double
            'Public oLOMARSprzedMies As Double       'LOMARSPRZED       Double
            'Public oAOWartSprzedMies As Double      'AOWARTSPRZED      Double
            'Public oAOIloSprzedMies As Double       'AOILOSPRZED       Double
            'Public oAOMARSprzedMies As Double       'AOMARSPRZED       Double

            'Public oWersjaMies As Integer          'WERSJA           Integer

            If ParL_DataBaseType = "MS ACCESS" Then
            Else
                sqlConnectionObrotyMies = New SqlClient.SqlConnection(sConnectionObrotyMies)
                sqlCommandSelectObrotyMies = New SqlClient.SqlCommand(sSelectSQLObrotyMies, sqlConnectionObrotyMies)
                sqlCommandDeleteObrotyMies = New SqlClient.SqlCommand(sDeleteSQLObrotyMies, sqlConnectionObrotyMies)
                sqlCommandUpdateObrotyMies = New SqlClient.SqlCommand(sUpdateSQLObrotyMies, sqlConnectionObrotyMies)
                sqlCommandInsertObrotyMies = New SqlClient.SqlCommand(sInsertSQLObrotyMies, sqlConnectionObrotyMies)
                sqlDataAdapterObrotyMies = New SqlClient.SqlDataAdapter(sqlCommandSelectObrotyMies)

                'mapujemy domyslna nazwe tabeli
                Dim MapowanieTabeli As System.Data.Common.DataTableMapping
                MapowanieTabeli = sqlDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")

                SQLDeleteObrotyMies(sqlCommandDeleteObrotyMies, sqlDataAdapterObrotyMies)
                SQLUpdateObrotyMies(sqlCommandUpdateObrotyMies, sqlDataAdapterObrotyMies)
                SQLInsertObrotyMies(sqlCommandInsertObrotyMies, sqlDataAdapterObrotyMies)


                ' Add the RowUpdated event handler.
                AddHandler sqlDataAdapterObrotyMies.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

                '------------------------------------------
                'wypelnij DATASET
                Try
                    sqlConnectionObrotyMies.Open()
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                End Try

                If sqlConnectionObrotyMies.State = ConnectionState.Open Then
                    Try
                        sqlDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                        sqlDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                        sqlConnectionObrotyMies.Close()

                        'zdefiniuj unikalny klucz wyszukiwania w tabeli
                        DataSetObrotyMies.Tables(0).PrimaryKey = New DataColumn() {DataSetObrotyMies.Tables(0).Columns("IDENTKLIENTA"), DataSetObrotyMies.Tables(0).Columns("MIESIAC")}
                        DataViewObrotyMies = New DataView(DataSetObrotyMies.Tables(0))

                        ObrotyMiesOK = True

                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    End Try
                End If







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

                sqlConnectionObroty = New SqlClient.SqlConnection(sConnectionObroty)
                sqlCommandSelectObroty = New SqlClient.SqlCommand(sSelectSQLObroty, sqlConnectionObroty)
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
                End Try

                If sqlConnectionObroty.State = ConnectionState.Open Then
                    Try
                        sqlDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                        sqlDataAdapterObroty.Fill(DataSetObroty)
                        sqlConnectionObroty.Close()

                        'zdefiniuj unikalny klucz wyszukiwania w tabeli
                        DataSetObroty.Tables(0).PrimaryKey = New DataColumn() {DataSetObroty.Tables(0).Columns("IDENTKLIENTA"), DataSetObroty.Tables(0).Columns("DATA")}
                        DataViewObroty = New DataView(DataSetObroty.Tables(0))

                        ObrotyOK = True

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



        Dim IleRec As Long
        Dim BiezRec As Long
        Dim AnalMies As String = TxtMiesiac.Text
        '-------------------------------
        'usuwamy zapisy z obrotow ostatniego miesiaca
        Dim dr As DataRow
        Dim Id As String
        Dim Mies As String
        Dim Dzien As String
        Dim idr As Integer = 0

        Dim foundRow As DataRow = Nothing
        Dim findObrotyMies(1) As Object
        Dim findKlienci(0) As Object

        If ObrotyMiesOK And ObrotyOK And KlienciOK Then


            '------------------------------
            'usuwamy informacje o obrotach mies
            IleRec = DataSetObrotyMies.Tables(0).Rows.Count
            ProgressBar.Maximum = IleRec
            For Each dr In DataSetObrotyMies.Tables(0).Rows
                BiezRec += 1
                ProgressBar.Value = BiezRec

                Id = dr.Item("IDENTKLIENTA")
                Mies = dr.Item("MIESIAC")
                If Mies = AnalMies Then
                    dr.Delete()
                Else
                End If
            Next
            AktualizujObrotyMies()



            '-----------------------------------------------
            'analizuj obroty i generuj obroty mies
            BiezRec = 0
            IleRec = DataSetObroty.Tables(0).Rows.Count
            ProgressBar.Maximum = IleRec
            For Each dr In DataSetObroty.Tables(0).Rows
                BiezRec += 1
                ProgressBar.Value = BiezRec

                Id = dr.Item("IDENTKLIENTA")
                Dzien = dr.Item("DATA")
                Mies = Mid(Dzien, 1, 7)

                If Mies = AnalMies Then
                    oLWartSprzedMies = dr.Item("LWARTSPRZED")
                    oLMARSprzedMies = dr.Item("LMARSPRZED")
                    oLIloSprzedMies = dr.Item("LILOSPRZED")

                    oAWartSprzedMies = dr.Item("AWARTSPRZED")
                    oAMARSprzedMies = dr.Item("AMARSPRZED")
                    oAIloSprzedMies = dr.Item("AILOSPRZED")

                    oLOWartSprzedMies = dr.Item("LOWARTSPRZED")
                    oLOMARSprzedMies = dr.Item("LOMARSPRZED")
                    oLOIloSprzedMies = dr.Item("LOILOSPRZED")

                    oAOWartSprzedMies = dr.Item("AOWARTSPRZED")
                    oAOMARSprzedMies = dr.Item("AOMARSPRZED")
                    oAOIloSprzedMies = dr.Item("AOILOSPRZED")

                    oNAJEMWartSprzedMies = dr.Item("NAJEMWARTSPRZED")
                    oNAJEMMARSprzedMies = dr.Item("NAJEMMARSPRZED")
                    oNAJEMIloSprzedMies = dr.Item("NAJEMILOSPRZED")

                    oSTRONYWartSprzedMies = dr.Item("STRONYWARTSPRZED")
                    oSTRONYMARSprzedMies = dr.Item("STRONYMARSPRZED")
                    oSTRONYIloSprzedMies = dr.Item("STRONYILOSPRZED")

                    oDRUKARKIWartSprzedMies = dr.Item("DRUKARKIWARTSPRZED")
                    oDRUKARKIMARSprzedMies = dr.Item("DRUKARKIMARSPRZED")
                    oDRUKARKIIloSprzedMies = dr.Item("DRUKARKIILOSPRZED")

                    oSKUPWartSprzedMies = dr.Item("SKUPWARTSPRZED")
                    oSKUPMARSprzedMies = dr.Item("SKUPMARSPRZED")
                    oSKUPIloSprzedMies = dr.Item("SKUPILOSPRZED")

                    oWersjaMies = dr.Item("WERSJA")


                    'aktualizujemy dane syntetyczne OBROTY
                    findObrotyMies(0) = Id
                    findObrotyMies(1) = Mies
                    foundRow = DataSetObrotyMies.Tables(0).Rows.Find(findObrotyMies)
                    ' sprawdz czy jest...
                    If Not (foundRow Is Nothing) Then

                        oWersjaMies = foundRow.Item("WERSJA")


                        'foundRow("IDENTKLIENTA") = Id
                        'foundRow("MIESIAC") = Mies

                        foundRow("LILOSPRZED") += oLIloSprzedMies
                        foundRow("LWARTSPRZED") += oLWartSprzedMies
                        foundRow("LMARSPRZED") += oLMARSprzedMies

                        foundRow("AILOSPRZED") += oAIloSprzedMies
                        foundRow("AWARTSPRZED") += oAWartSprzedMies
                        foundRow("AMARSPRZED") += oAMARSprzedMies

                        foundRow("LOILOSPRZED") += oLOIloSprzedMies
                        foundRow("LOWARTSPRZED") += oLOWartSprzedMies
                        foundRow("LOMARSPRZED") += oLOMARSprzedMies

                        foundRow("AOILOSPRZED") += oAOIloSprzedMies
                        foundRow("AOWARTSPRZED") += oAOWartSprzedMies
                        foundRow("AOMARSPRZED") += oAOMARSprzedMies

                        foundRow("NAJEMILOSPRZED") += oNAJEMIloSprzedMies
                        foundRow("NAJEMWARTSPRZED") += oNAJEMWartSprzedMies
                        foundRow("NAJEMMARSPRZED") += oNAJEMMARSprzedMies

                        foundRow("STRONYILOSPRZED") += oSTRONYIloSprzedMies
                        foundRow("STRONYWARTSPRZED") += oSTRONYWartSprzedMies
                        foundRow("STRONYMARSPRZED") += oSTRONYMARSprzedMies

                        foundRow("DRUKARKIILOSPRZED") += oDRUKARKIIloSprzedMies
                        foundRow("DRUKARKIWARTSPRZED") += oDRUKARKIWartSprzedMies
                        foundRow("DRUKARKIMARSPRZED") += oDRUKARKIMARSprzedMies

                        foundRow("SKUPILOSPRZED") += oSKUPIloSprzedMies
                        foundRow("SKUPWARTSPRZED") += oSKUPWartSprzedMies
                        foundRow("SKUPMARSPRZED") += oSKUPMARSprzedMies

                        foundRow("WERSJA") = ZmienWersje(oWersjaMies) 'Wersja         Integer, 2
                        'aktualizuj DataGrid
                        AktualizujObrotyMies()
                    Else
                        'stworz nowy rekord
                        foundRow = DataSetObrotyMies.Tables(0).NewRow()

                        foundRow("IDENTKLIENTA") = Id
                        foundRow("MIESIAC") = Mies

                        foundRow("LILOSPRZED") = oLIloSprzedMies
                        foundRow("LWARTSPRZED") = oLWartSprzedMies
                        foundRow("LMARSPRZED") = oLMARSprzedMies

                        foundRow("AILOSPRZED") = oAIloSprzedMies
                        foundRow("AWARTSPRZED") = oAWartSprzedMies
                        foundRow("AMARSPRZED") = oAMARSprzedMies

                        foundRow("LOILOSPRZED") = oLOIloSprzedMies
                        foundRow("LOWARTSPRZED") = oLOWartSprzedMies
                        foundRow("LOMARSPRZED") = oLOMARSprzedMies

                        foundRow("AOILOSPRZED") = oAOIloSprzedMies
                        foundRow("AOWARTSPRZED") = oAOWartSprzedMies
                        foundRow("AOMARSPRZED") = oAOMARSprzedMies

                        foundRow("NAJEMILOSPRZED") = oNAJEMIloSprzedMies
                        foundRow("NAJEMWARTSPRZED") = oNAJEMWartSprzedMies
                        foundRow("NAJEMMARSPRZED") = oNAJEMMARSprzedMies

                        foundRow("STRONYILOSPRZED") = oSTRONYIloSprzedMies
                        foundRow("STRONYWARTSPRZED") = oSTRONYWartSprzedMies
                        foundRow("STRONYMARSPRZED") = oSTRONYMARSprzedMies

                        foundRow("DRUKARKIILOSPRZED") = oDRUKARKIIloSprzedMies
                        foundRow("DRUKARKIWARTSPRZED") = oDRUKARKIWartSprzedMies
                        foundRow("DRUKARKIMARSPRZED") = oDRUKARKIMARSprzedMies

                        foundRow("SKUPILOSPRZED") = oSKUPIloSprzedMies
                        foundRow("SKUPWARTSPRZED") = oSKUPWartSprzedMies
                        foundRow("SKUPMARSPRZED") = oSKUPMARSprzedMies

                        foundRow("WERSJA") = 1     'Wersja         Integer, 2
                        'dodaj rekord do DataSet
                        DataSetObrotyMies.Tables(0).Rows.Add(foundRow)
                        AktualizujObrotyMies()
                    End If










                    If chbUsun.Checked Then
                        dr.Delete()
                    End If
                Else
                End If
            Next
            AktualizujObroty()
            'AktualizujObrotyMies()






            '-------------------------------
            MessageBox.Show("Pobra³em informacje o obrotach" & vbCrLf &
            "miesi¹ca " & AnalMies,
            "OK, skoñczy³em",
            System.Windows.Forms.MessageBoxButtons.OK,
            MessageBoxIcon.Information,
            MessageBoxDefaultButton.Button1)
            Me.Close()
        End If
    End Sub



    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujObrotyMies()
        If ParL_DataBaseType = "MS ACCESS" Then
            Try
                dbConnectionObrotyMies.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If dbConnectionObrotyMies.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    dbDataAdapterObrotyMies.Update(DataSetObrotyMies, "TABELA_ObrotyMies")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    dbDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                End If
                dbConnectionObrotyMies.Close()
            End If
        Else
            Try
                sqlConnectionObrotyMies.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionObrotyMies.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterObrotyMies.Update(DataSetObrotyMies, "TABELA_ObrotyMies")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                End If
                sqlConnectionObrotyMies.Close()
            End If
        End If
    End Sub




    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujObroty()
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


End Class
