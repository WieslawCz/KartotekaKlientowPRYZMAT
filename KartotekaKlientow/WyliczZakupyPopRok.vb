Public Class WyliczZakupyPopRok
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
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents cmdAnalizuj As System.Windows.Forms.Button
    Friend WithEvents Panel2 As Panel
    Friend WithEvents ProgressBar As ProgressBar
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Label5 As Label
    Friend WithEvents rbtWGPobierzP As RadioButton
    Friend WithEvents rbtWGPobierzW As RadioButton
    Friend WithEvents rbtWGWylicz As RadioButton
    Friend WithEvents lblBiezRok As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents cbxRok As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents chkInicjujPrognozy As CheckBox
    Friend WithEvents chkWyliczZakupy As CheckBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(WyliczZakupyPopRok))
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.cmdAnalizuj = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.chkInicjujPrognozy = New System.Windows.Forms.CheckBox()
        Me.chkWyliczZakupy = New System.Windows.Forms.CheckBox()
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.rbtWGPobierzP = New System.Windows.Forms.RadioButton()
        Me.rbtWGPobierzW = New System.Windows.Forms.RadioButton()
        Me.rbtWGWylicz = New System.Windows.Forms.RadioButton()
        Me.lblBiezRok = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cbxRok = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(488, 319)
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
        Me.cmdAnalizuj.Location = New System.Drawing.Point(400, 319)
        Me.cmdAnalizuj.Name = "cmdAnalizuj"
        Me.cmdAnalizuj.Size = New System.Drawing.Size(80, 24)
        Me.cmdAnalizuj.TabIndex = 13
        Me.cmdAnalizuj.Text = "Analizuj"
        Me.cmdAnalizuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 319)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(88, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 27
        Me.PictureBox2.TabStop = False
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.chkInicjujPrognozy)
        Me.Panel2.Controls.Add(Me.chkWyliczZakupy)
        Me.Panel2.Controls.Add(Me.ProgressBar)
        Me.Panel2.Controls.Add(Me.Panel3)
        Me.Panel2.Controls.Add(Me.lblBiezRok)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.cbxRok)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Location = New System.Drawing.Point(8, 8)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(560, 305)
        Me.Panel2.TabIndex = 43
        '
        'chkInicjujPrognozy
        '
        Me.chkInicjujPrognozy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkInicjujPrognozy.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.chkInicjujPrognozy.ForeColor = System.Drawing.Color.Navy
        Me.chkInicjujPrognozy.Location = New System.Drawing.Point(15, 237)
        Me.chkInicjujPrognozy.Name = "chkInicjujPrognozy"
        Me.chkInicjujPrognozy.Size = New System.Drawing.Size(538, 18)
        Me.chkInicjujPrognozy.TabIndex = 381
        Me.chkInicjujPrognozy.Text = "Inicjuj (wyzeruj) informacje o prognozach na nowy rok."
        Me.chkInicjujPrognozy.UseVisualStyleBackColor = True
        '
        'chkWyliczZakupy
        '
        Me.chkWyliczZakupy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkWyliczZakupy.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.chkWyliczZakupy.ForeColor = System.Drawing.Color.Navy
        Me.chkWyliczZakupy.Location = New System.Drawing.Point(15, 117)
        Me.chkWyliczZakupy.Name = "chkWyliczZakupy"
        Me.chkWyliczZakupy.Size = New System.Drawing.Size(538, 18)
        Me.chkWyliczZakupy.TabIndex = 380
        Me.chkWyliczZakupy.Text = "Inicjuj (wyzeruj) pola oceny klientów na nowy rok."
        Me.chkWyliczZakupy.UseVisualStyleBackColor = True
        '
        'ProgressBar
        '
        Me.ProgressBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar.Location = New System.Drawing.Point(15, 280)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(520, 8)
        Me.ProgressBar.TabIndex = 42
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Controls.Add(Me.rbtWGPobierzP)
        Me.Panel3.Controls.Add(Me.rbtWGPobierzW)
        Me.Panel3.Controls.Add(Me.rbtWGWylicz)
        Me.Panel3.Location = New System.Drawing.Point(0, 140)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(558, 80)
        Me.Panel3.TabIndex = 379
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(12, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(508, 16)
        Me.Label5.TabIndex = 381
        Me.Label5.Text = "Sposób wyliczania wartoœci granicznej : "
        '
        'rbtWGPobierzP
        '
        Me.rbtWGPobierzP.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtWGPobierzP.ForeColor = System.Drawing.Color.Navy
        Me.rbtWGPobierzP.Location = New System.Drawing.Point(39, 61)
        Me.rbtWGPobierzP.Name = "rbtWGPobierzP"
        Me.rbtWGPobierzP.Size = New System.Drawing.Size(481, 18)
        Me.rbtWGPobierzP.TabIndex = 384
        Me.rbtWGPobierzP.TabStop = True
        Me.rbtWGPobierzP.Text = "Pobierz % sumarycznej wartoœci sprzeda¿y w pop.roku z parametrów i wylicz Wart.Gr" &
    "an"
        Me.rbtWGPobierzP.UseVisualStyleBackColor = True
        '
        'rbtWGPobierzW
        '
        Me.rbtWGPobierzW.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtWGPobierzW.ForeColor = System.Drawing.Color.Navy
        Me.rbtWGPobierzW.Location = New System.Drawing.Point(39, 43)
        Me.rbtWGPobierzW.Name = "rbtWGPobierzW"
        Me.rbtWGPobierzW.Size = New System.Drawing.Size(481, 18)
        Me.rbtWGPobierzW.TabIndex = 383
        Me.rbtWGPobierzW.TabStop = True
        Me.rbtWGPobierzW.Text = "Pobierz Wart.Gran z Parametrów"
        Me.rbtWGPobierzW.UseVisualStyleBackColor = True
        '
        'rbtWGWylicz
        '
        Me.rbtWGWylicz.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtWGWylicz.ForeColor = System.Drawing.Color.Navy
        Me.rbtWGWylicz.Location = New System.Drawing.Point(39, 25)
        Me.rbtWGWylicz.Name = "rbtWGWylicz"
        Me.rbtWGWylicz.Size = New System.Drawing.Size(481, 18)
        Me.rbtWGWylicz.TabIndex = 382
        Me.rbtWGWylicz.TabStop = True
        Me.rbtWGWylicz.Text = "Wylicz Wart.Gran jako 80 % sumarycznej wart. sprzedazy w poprzednim roku"
        Me.rbtWGWylicz.UseVisualStyleBackColor = True
        '
        'lblBiezRok
        '
        Me.lblBiezRok.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblBiezRok.ForeColor = System.Drawing.Color.Navy
        Me.lblBiezRok.Location = New System.Drawing.Point(257, 55)
        Me.lblBiezRok.Name = "lblBiezRok"
        Me.lblBiezRok.Size = New System.Drawing.Size(53, 16)
        Me.lblBiezRok.TabIndex = 57
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(0, -1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(558, 57)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = resources.GetString("Label1.Text")
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(12, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(281, 16)
        Me.Label4.TabIndex = 56
        Me.Label4.Text = "Bie¿¹cy rok kalendarzowy . . . . . . . . . . . . . . . . . . . ."
        '
        'cbxRok
        '
        Me.cbxRok.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cbxRok.ForeColor = System.Drawing.Color.Purple
        Me.cbxRok.Location = New System.Drawing.Point(257, 75)
        Me.cbxRok.Name = "cbxRok"
        Me.cbxRok.Size = New System.Drawing.Size(53, 22)
        Me.cbxRok.TabIndex = 55
        Me.cbxRok.Text = "2007"
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(12, 78)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(281, 16)
        Me.Label3.TabIndex = 38
        Me.Label3.Text = "Rok do analizy wartoœci zakupów (poprzedni) . . . . . . . . ."
        '
        'WyliczZakupyPopRok
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(576, 348)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.cmdAnalizuj)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Name = "WyliczZakupyPopRok"
        Me.ShowInTaskbar = False
        Me.Text = "Inicjowanie nowego roku rozliczeniowego"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim ObrotyMiesOK As Boolean = False
    Dim ObrotyOK As Boolean = False
    Dim KlienciOK As Boolean = False
    '----------------------------------
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
    '-------------------------
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
    '--------------------------
    Dim dbSelectSQLKlienci As String = ""       'sSelectSQLKlienci
    Dim dbUpdateSQLKlienci As String = ""       'sSelectSQLKlienci

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
    '--------------------------
    Dim ConnectionState As System.Data.ConnectionState





    Private Sub WyliczZakupyPopRok_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        lblBiezRok.Text = Microsoft.VisualBasic.Right("0000" & Trim(Str(Year(Now))), 4)
        InicjujListeLat(cbxRok)
        cbxRok.Text = Microsoft.VisualBasic.Right("0000" & Trim(Str(Year(Now) - 1)), 4)

        chkWyliczZakupy.Checked = True
        chkInicjujPrognozy.Checked = True
    End Sub

    Private Sub WyliczZakupyPopRok_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub cmdAnalizuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAnalizuj.Click
        If chkWyliczZakupy.Checked Then
            WyliczZakupy()
        End If

        If chkInicjujPrognozy.Checked Then
            InicjujPrognozy()
        End If

        '-------------------------------
        MessageBox.Show(IIf(chkWyliczZakupy.Checked, "Pobra³em informacje o wartoœci zakupów klientów w roku " & cbxRok.Text & vbCrLf, "") &
                   IIf(chkInicjujPrognozy.Checked, "Wyzerowa³em prognozy na Nowy Rok.", ""),
                    "OK, skoñczy³em",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
        Me.Close()
    End Sub


    '--------------------------------------------------
    ' inicjuje prognozy zakupu
    '---------------------------------------------------

    Private Sub inicjujprognozy()
        Dim idr As Integer = 0

        'przejdz przez katalog klienci - usuwaj prognoze z miesi¹ca analizy
        DataSetKlienci = New DataSet
        DataViewKlienci = New DataView
        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
        Else
            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
            sqlCommandSelectKlienci = New SqlClient.SqlCommand(sSelectSQLKlienciAktywni, sqlConnectionKlienci)
            'sqlCommandDeleteKlienci = New SqlClient.SqlCommand(sDeleteSQLKlienci, sqlConnectionKlienci)
            sqlCommandUpdateKlienci = New SqlClient.SqlCommand(sUpdateSQLKlienci, sqlConnectionKlienci)
            'sqlCommandInsertKlienci = New SqlClient.SqlCommand(sInsertSQLKlienci, sqlConnectionKlienci)
            sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)


            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliKlienci As System.Data.Common.DataTableMapping
            MapowanieTabeliKlienci = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

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

                If DataSetKlienci.Tables(0).Rows.Count > 0 Then
                    ProgressBar.Maximum = DataSetKlienci.Tables(0).Rows.Count
                    For idr = 0 To DataSetKlienci.Tables(0).Rows.Count - 1
                        ProgressBar.Value = idr

                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA01") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA02") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA03") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA04") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA05") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA06") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA07") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA08") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA09") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA10") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA11") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZA12") = 0

                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZAP1") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZAP2") = 0

                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZAK1") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZAK2") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZAK3") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZAK4") = 0

                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZARR") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROGNOZARRPLAN") = 0

                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGI01") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGI02") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGI03") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGI04") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGI05") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGI06") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGI07") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGI08") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGI09") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGI10") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGI11") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGI12") = 0

                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGIP1") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGIP2") = 0

                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGIK1") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGIK2") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGIK3") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGIK4") = 0

                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGIRR") = 0
                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROUSLUGIRRPLAN") = 0

                        DataSetKlienci.Tables(0).Rows.Item(idr).Item("PROSKUPRRPLAN") = 0
                    Next
                    AktualizujBazeKlienci()
                End If

            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try
            '---------------------------------
        End If
    End Sub


    Private Sub AktualizujBazeKlienci()
        If ParL_DataBaseType = "MS ACCESS" Then
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
















    '--------------------------------------------------
    ' wylicza zakupy poprzedniego roku
    '---------------------------------------------------

    Private Sub WyliczZakupy()
        ObrotyMiesOK = False
        ObrotyOK = False
        KlienciOK = False

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

        DataSetObroty = New DataSet
        DataViewObroty = New DataView
        DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")

        'Public sSelectSQLObroty As String = "SELECT " & _
        '                                             "IDENTKLIENTA," & _
        '                                             "DATA," & _
        '                                             "AILOSPRZED," & _
        '                                             "AWARTSPRZED," & _
        '                                             "AMARSPRZED," & _
        '                                             "LILOSPRZED," & _
        '                                             "LWARTSPRZED," & _
        '                                             "LMARSPRZED," & _
        '                                               "AOILOSPRZED," & _
        '                                               "AOWARTSPRZED," & _
        '                                               "AOMARSPRZED," & _
        '                                               "LOILOSPRZED," & _
        '                                               "LOWARTSPRZED," & _
        '                                               "LOMARSPRZED," & _
        '                                             "WERSJA " & _
        '                                    "FROM Obroty "

        dbSelectSQLObroty = sSelectSQLObroty &
                                " WHERE (SUBSTRING(DATA,1,4)='" & cbxRok.Text & "') "







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

        DataSetObrotyMies = New DataSet
        DataViewObrotyMies = New DataView
        DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")

        'Public sSelectSQLObrotyMies As String = "SELECT " & _
        '                                             "IDENTKLIENTA," & _
        '                                             "MIESIAC," & _
        '                                             "AILOSPRZED," & _
        '                                             "AWARTSPRZED," & _
        '                                             "AMARSPRZED," & _
        '                                             "LILOSPRZED," & _
        '                                             "LWARTSPRZED," & _
        '                                             "LMARSPRZED," & _
        '                                               "AOILOSPRZED," & _
        '                                               "AOWARTSPRZED," & _
        '                                               "AOMARSPRZED," & _
        '                                               "LOILOSPRZED," & _
        '                                               "LOWARTSPRZED," & _
        '                                               "LOMARSPRZED," & _
        '                                             "WERSJA " & _
        '                                        "FROM ObrotyMies "

        dbSelectSQLObrotyMies = sSelectSQLObrotyMies &
                                " WHERE (SUBSTRING(MIESIAC,1,4)='" & cbxRok.Text & "') "




        DataSetKlienci = New DataSet
        DataViewKlienci = New DataView
        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")
        dbSelectSQLKlienci = "SELECT " &
                                               "IDENTKLIENTA AS NRKARTY, " &
                                               "ZAKUPPOPROK, " &
                                               "USLUGIDOD, " &
                                               "POTENCJALBR, " &
                                               "POTENCJALPR, " &
                                               "POTENCJALPPR, " &
                                               "ROKOWANIABR, " &
                                               "ROKOWANIAPR, " &
                                               "ROKOWANIAPPR, " &
                                               "WERSJA " &
                                            "FROM Klienci "


        dbUpdateSQLKlienci = "UPDATE Klienci SET " &
                                       "ZAKUPPOPROK=@oZAKUPPOPROKKli, " &
                                       "USLUGIDOD=@oUSLUGIDODKli, " &
                                             "POTENCJALBR=@oPOTENCJALBR, " &
                                             "POTENCJALPR=@oPOTENCJALPR, " &
                                             "POTENCJALPPR=@oPOTENCJALPPR, " &
                                             "ROKOWANIABR=@oROKOWANIABR, " &
                                             "ROKOWANIAPR=@oROKOWANIAPR, " &
                                             "ROKOWANIAPPR=@oROKOWANIAPPR, " &
                                       "WERSJA=@oWersjaKli " &
                                    "WHERE (IDENTKLIENTA=@orygSymbol) AND " &
                                          "(WERSJA=@orygWersja)"



        If ParL_DataBaseType = "MS ACCESS" Then
        Else
            sqlConnectionObrotyMies = New SqlClient.SqlConnection(sConnectionObrotyMies)
            sqlCommandSelectObrotyMies = New SqlClient.SqlCommand(dbSelectSQLObrotyMies, sqlConnectionObrotyMies)
            'sqlCommandDeleteObrotyMies = New SqlClient.SqlCommand(sDeleteSQLObrotyMies, sqlConnectionObrotyMies)
            'sqlCommandUpdateObrotyMies = New SqlClient.SqlCommand(sUpdateSQLObrotyMies, sqlConnectionObrotyMies)
            'sqlCommandInsertObrotyMies = New SqlClient.SqlCommand(sInsertSQLObrotyMies, sqlConnectionObrotyMies)
            sqlDataAdapterObrotyMies = New SqlClient.SqlDataAdapter(sqlCommandSelectObrotyMies)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")

            ''komendy DataAdapter
            'Dim objParam As SqlClient.SqlParameter
            ''------------------------------------------
            ''komenda DELETE
            ''parametry filtrowania
            'objParam = sqlCommandDeleteObrotyMies.Parameters.Add("@orygSymbol", SqlDbType.Char, 6, "IDENTKLIENTA")
            'objParam.SourceVersion = DataRowVersion.Original
            'objParam = sqlCommandDeleteObrotyMies.Parameters.Add("@orygMies", SqlDbType.Char, 7, "MIESIAC")
            'objParam.SourceVersion = DataRowVersion.Original
            'objParam = sqlCommandDeleteObrotyMies.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            'objParam.SourceVersion = DataRowVersion.Original
            'sqlDataAdapterObrotyMies.DeleteCommand = sqlCommandDeleteObrotyMies

            ''------------------------------------------
            ''komenda UPDATE
            ''parametry aktualizacji
            ''sqlCommandUpdateObrotyMies.Parameters.Add("@oIdentMies", sqlDbType.Char, 6, "IDENTKLIENTA")
            ''sqlCommandUpdateObrotyMies.Parameters.Add("@oMiesMies", sqlDbType.Char, 7, "MIESIAC")

            'sqlCommandUpdateObrotyMies.Parameters.Add("@oAIloSprzedMies", SqlDbType.Float, Nothing, "AILOSPRZED")
            'sqlCommandUpdateObrotyMies.Parameters.Add("@oAWartSprzedMies", SqlDbType.Float, Nothing, "AWARTSPRZED")
            'sqlCommandUpdateObrotyMies.Parameters.Add("@oLIloSprzedMies", SqlDbType.Float, Nothing, "LILOSPRZED")
            'sqlCommandUpdateObrotyMies.Parameters.Add("@oLWartSprzedMies", SqlDbType.Float, Nothing, "LWARTSPRZED")

            'sqlCommandUpdateObrotyMies.Parameters.Add("@oAoIloSprzedMies", SqlDbType.Float, Nothing, "AOILOSPRZED")
            'sqlCommandUpdateObrotyMies.Parameters.Add("@oAoWartSprzedMies", SqlDbType.Float, Nothing, "AOWARTSPRZED")
            'sqlCommandUpdateObrotyMies.Parameters.Add("@oLoIloSprzedMies", SqlDbType.Float, Nothing, "LOILOSPRZED")
            'sqlCommandUpdateObrotyMies.Parameters.Add("@oLoWartSprzedMies", SqlDbType.Float, Nothing, "LOWARTSPRZED")

            'sqlCommandUpdateObrotyMies.Parameters.Add("@oWersjaMies", SqlDbType.Int, Nothing, "WERSJA")

            ''parametry filtrowania
            'objParam = sqlCommandUpdateObrotyMies.Parameters.Add("@orygSymbol", SqlDbType.Char, 6, "IDENTKLIENTA")
            'objParam.SourceVersion = DataRowVersion.Original
            'objParam = sqlCommandUpdateObrotyMies.Parameters.Add("@orygMies", SqlDbType.Char, 10, "MIESIAC")
            'objParam.SourceVersion = DataRowVersion.Original
            'objParam = sqlCommandUpdateObrotyMies.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            'objParam.SourceVersion = DataRowVersion.Original
            'sqlDataAdapterObrotyMies.UpdateCommand = sqlCommandUpdateObrotyMies

            ''------------------------------------------
            ''komenda INSERT
            'objParam = sqlCommandInsertObrotyMies.Parameters.Add("@oIdentMies", SqlDbType.Char, 6, "IDENTKLIENTA")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertObrotyMies.Parameters.Add("@oMiesMies", SqlDbType.Char, 10, "MIESIAC")
            'objParam.SourceVersion = DataRowVersion.Current

            'objParam = sqlCommandInsertObrotyMies.Parameters.Add("@oAIloSprzedMies", SqlDbType.Float, Nothing, "AILOSPRZED")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertObrotyMies.Parameters.Add("@oAWartSprzedMies", SqlDbType.Float, Nothing, "AWARTSPRZED")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertObrotyMies.Parameters.Add("@oLIloSprzedMies", SqlDbType.Float, Nothing, "LILOSPRZED")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertObrotyMies.Parameters.Add("@oLWartSprzedMies", SqlDbType.Float, Nothing, "LWARTSPRZED")
            'objParam.SourceVersion = DataRowVersion.Current

            'objParam = sqlCommandInsertObrotyMies.Parameters.Add("@oAoIloSprzedMies", SqlDbType.Float, Nothing, "AOILOSPRZED")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertObrotyMies.Parameters.Add("@oAoWartSprzedMies", SqlDbType.Float, Nothing, "AOWARTSPRZED")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertObrotyMies.Parameters.Add("@oLoIloSprzedMies", SqlDbType.Float, Nothing, "LOILOSPRZED")
            'objParam.SourceVersion = DataRowVersion.Current
            'objParam = sqlCommandInsertObrotyMies.Parameters.Add("@oLoWartSprzedMies", SqlDbType.Float, Nothing, "LOWARTSPRZED")
            'objParam.SourceVersion = DataRowVersion.Current

            'objParam = sqlCommandInsertObrotyMies.Parameters.Add("@oWersjaMies", SqlDbType.Int, Nothing, "WERSJA")
            'objParam.SourceVersion = DataRowVersion.Current
            'sqlDataAdapterObrotyMies.InsertCommand = sqlCommandInsertObrotyMies

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
                    DataSetObrotyMies.Tables(0).PrimaryKey = New DataColumn() {DataSetObrotyMies.Tables(0).Columns("IDENTKLIENTA"),
                                                                               DataSetObrotyMies.Tables(0).Columns("MIESIAC")}
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







            If ObrotyMiesOK Then
                sqlConnectionObroty = New SqlClient.SqlConnection(sConnectionObroty)
                sqlCommandSelectObroty = New SqlClient.SqlCommand(dbSelectSQLObroty, sqlConnectionObroty)
                'sqlCommandDeleteObroty = New SqlClient.SqlCommand(sDeleteSQLObroty, sqlConnectionObroty)
                'sqlCommandUpdateObroty = New SqlClient.SqlCommand(sUpdateSQLObroty, sqlConnectionObroty)
                'sqlCommandInsertObroty = New SqlClient.SqlCommand(sInsertSQLObroty, sqlConnectionObroty)
                sqlDataAdapterObroty = New SqlClient.SqlDataAdapter(sqlCommandSelectObroty)

                'mapujemy domyslna nazwe tabeli
                Dim MapowanieTabeliObroty As System.Data.Common.DataTableMapping
                MapowanieTabeliObroty = sqlDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")

                ''komendy DataAdapter
                'Dim objParamObroty As SqlClient.SqlParameter
                ''------------------------------------------
                ''komenda DELETE
                ''parametry filtrowania
                'objParamObroty = sqlCommandDeleteObroty.Parameters.Add("@orygSymbol", SqlDbType.Char, 6, "IDENTKLIENTA")
                'objParamObroty.SourceVersion = DataRowVersion.Original
                'objParamObroty = sqlCommandDeleteObroty.Parameters.Add("@orygData", SqlDbType.Char, 10, "DATA")
                'objParamObroty.SourceVersion = DataRowVersion.Original
                'objParamObroty = sqlCommandDeleteObroty.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
                'objParamObroty.SourceVersion = DataRowVersion.Original
                'sqlDataAdapterObroty.DeleteCommand = sqlCommandDeleteObroty


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
                        DataSetObroty.Tables(0).PrimaryKey = New DataColumn() {DataSetObroty.Tables(0).Columns("IDENTKLIENTA"),
                                                                               DataSetObroty.Tables(0).Columns("DATA")}
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





            If ObrotyMiesOK And ObrotyOK Then
                sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
                sqlCommandSelectKlienci = New SqlClient.SqlCommand(dbSelectSQLKlienci, sqlConnectionKlienci)
                'sqlCommandDeleteKlienci = New SqlClient.SqlCommand(sDeleteSQLKlienci, sqlConnectionKlienci)
                sqlCommandUpdateKlienci = New SqlClient.SqlCommand(dbUpdateSQLKlienci, sqlConnectionKlienci)
                'sqlCommandInsertKlienci = New SqlClient.SqlCommand(sInsertSQLKlienci, sqlConnectionKlienci)
                sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

                'mapujemy domyslna nazwe tabeli
                Dim MapowanieTabeliKlienci As System.Data.Common.DataTableMapping
                MapowanieTabeliKlienci = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

                Dim objParamSQL As SqlClient.SqlParameter

                '------------------------------------------
                'komenda UPDATE
                'parametry aktualizacji
                'sqlcommandUpdateKlienci.Parameters.Add("@oIdentKli", sqldbtype.varchar, 6, "NRKARTY")
                sqlCommandUpdateKlienci.Parameters.Add("@oZAKUPPOPROKKli", SqlDbType.Float, Nothing, "ZAKUPPOPROK")
                sqlCommandUpdateKlienci.Parameters.Add("@oUSLUGIDODKli", SqlDbType.VarChar, 200, "USLUGIDOD")
                sqlCommandUpdateKlienci.Parameters.Add("@oPOTENCJALBR", SqlDbType.VarChar, 4, "POTENCJALBR")
                sqlCommandUpdateKlienci.Parameters.Add("@oPOTENCJALPR", SqlDbType.VarChar, 4, "POTENCJALPR")
                sqlCommandUpdateKlienci.Parameters.Add("@oPOTENCJALPPR", SqlDbType.VarChar, 4, "POTENCJALPPR")
                sqlCommandUpdateKlienci.Parameters.Add("@oROKOWANIABR", SqlDbType.VarChar, 4, "ROKOWANIABR")
                sqlCommandUpdateKlienci.Parameters.Add("@oROKOWANIAPR", SqlDbType.VarChar, 4, "ROKOWANIAPR")
                sqlCommandUpdateKlienci.Parameters.Add("@oROKOWANIAPPR", SqlDbType.VarChar, 4, "ROKOWANIAPPR")
                sqlCommandUpdateKlienci.Parameters.Add("@oWersjaKli", SqlDbType.Int, Nothing, "WERSJA")

                'parametry filtrowania
                objParamSQL = sqlCommandUpdateKlienci.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 10, "NRKARTY")
                objParamSQL.SourceVersion = DataRowVersion.Original
                objParamSQL = sqlCommandUpdateKlienci.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
                objParamSQL.SourceVersion = DataRowVersion.Original
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
                End Try

                If sqlConnectionKlienci.State = ConnectionState.Open Then
                    Try
                        sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                        sqlDataAdapterKlienci.Fill(DataSetKlienci)
                        sqlConnectionKlienci.Close()

                        'zdefiniuj unikalny klucz wyszukiwania w tabeli
                        DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("NRKARTY")}
                        DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
                        KlienciOK = True

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



        Dim IleRec As Long = 0
        Dim BiezRec As Long = 0
        Dim AnalRok As String = cbxRok.Text
        Dim dr As DataRow = Nothing
        Dim Id As String = ""
        Dim Rok As String = ""
        Dim WartoscO As Double = 0
        Dim i As Integer = 0
        Dim popRok As String = WyliczRok(Mid(SysData, 1, 4), -1)
        Dim WartGran As Double = 0

        If rbtWGWylicz.Checked Then
            WartGran = WyliczWartoscGraniczna(popRok, 80)
        ElseIf rbtWGPobierzP.Checked Then
            WartGran = WyliczWartoscGraniczna(popRok, Par_WartGranProcent)
        ElseIf rbtWGPobierzW.Checked Then
            WartGran = Par_WartGranWartosc
        Else
            WartGran = WyliczWartoscGraniczna(popRok, 80)
        End If

        '-------------------------------
        If ObrotyMiesOK And ObrotyOK And KlienciOK Then

            'usuwamy zapisy biezace
            'analizujemy wartosc obrotow
            BiezRec = 0
            IleRec = DataSetKlienci.Tables(0).Rows.Count
            ProgressBar.Maximum = IleRec
            For Each dr In DataSetKlienci.Tables(0).Rows
                BiezRec += 1
                ProgressBar.Value = BiezRec
                Application.DoEvents()
                Id = dr.Item("NRKARTY")

                'analizujemy wartosc obrotow
                WartoscO = 0
                DataViewObroty.RowFilter = "IDENTKLIENTA='" & Id & "'"
                If DataViewObroty.Count > 0 Then
                    For i = 0 To DataViewObroty.Count - 1
                        WartoscO += GetDblDRVField(DataViewObroty.Item(i), "LWARTSPRZED") +
                                    GetDblDRVField(DataViewObroty.Item(i), "AWARTSPRZED") +
                                    GetDblDRVField(DataViewObroty.Item(i), "LOWARTSPRZED") +
                                    GetDblDRVField(DataViewObroty.Item(i), "AOWARTSPRZED")
                    Next
                End If
                DataViewObroty.RowFilter = ""


                'analizujemy wartosc obrotow Mies
                DataViewObrotyMies.RowFilter = "IDENTKLIENTA='" & Id & "'"
                If DataViewObrotyMies.Count > 0 Then
                    For i = 0 To DataViewObrotyMies.Count - 1
                        WartoscO += GetDblDRVField(DataViewObrotyMies.Item(i), "LWARTSPRZED") +
                                    GetDblDRVField(DataViewObrotyMies.Item(i), "AWARTSPRZED") +
                                    GetDblDRVField(DataViewObrotyMies.Item(i), "LOWARTSPRZED") +
                                    GetDblDRVField(DataViewObrotyMies.Item(i), "AOWARTSPRZED")
                    Next
                End If
                DataViewObrotyMies.RowFilter = ""

                'Aktualizujemy zakup do RZU
                dr.Item("ZAKUPPOPROK") = WartoscO







                '2018.01.19 - POPROK=BIEZROK, BIEZROK=puste

                ''Aktualizujemy Potencja³ i rokowania
                'dr.Item("POTENCJALPPR") = dr.Item("POTENCJALPR")    'PrzedPoprzedni=Poprzedni
                dr.Item("POTENCJALPR") = dr.Item("POTENCJALBR")     'Poprzedni=Bie¿¹cy
                dr.Item("POTENCJALBR") = ""

                'dr.Item("ROKOWANIAPPR") = dr.Item("ROKOWANIAPR")    'PrzedPoprzedni=Poprzedni
                dr.Item("ROKOWANIAPR") = dr.Item("ROKOWANIABR")     'Poprzedni=Bie¿¹cy
                dr.Item("ROKOWANIABR") = ""

                ''Aktualizujemy bie¿¹cy rok
                'If WartoscO < WartGran Then
                '    dr.Item("POTENCJALPR") = ""
                'ElseIf WartoscO < 1000 Then
                '    dr.Item("POTENCJALPR") = "PG"
                'ElseIf WartoscO < 3000 Then
                '    dr.Item("POTENCJALPR") = "P1"
                'ElseIf WartoscO < 6000 Then
                '    dr.Item("POTENCJALPR") = "P3"
                'Else
                '    dr.Item("POTENCJALPR") = "P6"
                'End If


            Next
            AktualizujKlienci()
        End If
    End Sub



    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujKlienci()
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

    Private Sub chkWyliczZakupy_CheckedChanged(sender As Object, e As EventArgs) Handles chkWyliczZakupy.CheckedChanged
        If Not chkWyliczZakupy.Checked Then
            rbtWGPobierzP.Enabled = False
            rbtWGPobierzW.Enabled = False
            rbtWGWylicz.Enabled = False
        Else
            rbtWGPobierzP.Enabled = True
            rbtWGPobierzW.Enabled = True
            rbtWGWylicz.Enabled = True
        End If
    End Sub
End Class
