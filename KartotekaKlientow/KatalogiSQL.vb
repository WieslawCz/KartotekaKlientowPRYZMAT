Public Class KatalogiSQL
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
    Friend WithEvents lblPlik As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents lbxTabele As System.Windows.Forms.ListBox
    Friend WithEvents ListView As System.Windows.Forms.ListView
    Friend WithEvents NazwaPola As System.Windows.Forms.ColumnHeader
    Friend WithEvents TypPola As System.Windows.Forms.ColumnHeader
    Friend WithEvents DlugoscPola As System.Windows.Forms.ColumnHeader
    Friend WithEvents UnikalnoscPola As System.Windows.Forms.ColumnHeader
    Friend WithEvents ROPola As System.Windows.Forms.ColumnHeader
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdPokazZawartosc As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogiSQL))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ListView = New System.Windows.Forms.ListView()
        Me.NazwaPola = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.TypPola = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.DlugoscPola = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.UnikalnoscPola = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ROPola = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.lbxTabele = New System.Windows.Forms.ListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblPlik = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.cmdPokazZawartosc = New System.Windows.Forms.Button()
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
        Me.Panel1.Controls.Add(Me.ListView)
        Me.Panel1.Controls.Add(Me.lbxTabele)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.lblPlik)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(616, 304)
        Me.Panel1.TabIndex = 0
        '
        'ListView
        '
        Me.ListView.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListView.BackColor = System.Drawing.SystemColors.Control
        Me.ListView.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.NazwaPola, Me.TypPola, Me.DlugoscPola, Me.UnikalnoscPola, Me.ROPola})
        Me.ListView.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.ListView.ForeColor = System.Drawing.Color.Purple
        Me.ListView.HideSelection = False
        Me.ListView.Location = New System.Drawing.Point(216, 32)
        Me.ListView.Name = "ListView"
        Me.ListView.Size = New System.Drawing.Size(384, 258)
        Me.ListView.TabIndex = 5
        Me.ListView.UseCompatibleStateImageBehavior = False
        Me.ListView.View = System.Windows.Forms.View.Details
        '
        'NazwaPola
        '
        Me.NazwaPola.Text = "Nazwa Pola"
        Me.NazwaPola.Width = 120
        '
        'TypPola
        '
        Me.TypPola.Text = "Typ"
        '
        'DlugoscPola
        '
        Me.DlugoscPola.Text = "D³ugoœæ"
        '
        'UnikalnoscPola
        '
        Me.UnikalnoscPola.Text = "Uni"
        Me.UnikalnoscPola.Width = 30
        '
        'ROPola
        '
        Me.ROPola.Text = "R/O"
        Me.ROPola.Width = 30
        '
        'lbxTabele
        '
        Me.lbxTabele.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lbxTabele.ForeColor = System.Drawing.Color.Navy
        Me.lbxTabele.Location = New System.Drawing.Point(64, 32)
        Me.lbxTabele.Name = "lbxTabele"
        Me.lbxTabele.Size = New System.Drawing.Size(144, 238)
        Me.lbxTabele.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(16, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Tabele :"
        '
        'lblPlik
        '
        Me.lblPlik.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblPlik.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPlik.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblPlik.ForeColor = System.Drawing.Color.Navy
        Me.lblPlik.Location = New System.Drawing.Point(88, 8)
        Me.lblPlik.Name = "lblPlik"
        Me.lblPlik.Size = New System.Drawing.Size(512, 16)
        Me.lblPlik.TabIndex = 1
        Me.lblPlik.Text = "Plik na dysku"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(16, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Baza danych . . . . . ."
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(544, 320)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 23)
        Me.cmdPowrot.TabIndex = 14
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(16, 320)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 16
        Me.PictureBox2.TabStop = False
        '
        'cmdPokazZawartosc
        '
        Me.cmdPokazZawartosc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPokazZawartosc.Image = CType(resources.GetObject("cmdPokazZawartosc.Image"), System.Drawing.Image)
        Me.cmdPokazZawartosc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPokazZawartosc.Location = New System.Drawing.Point(400, 320)
        Me.cmdPokazZawartosc.Name = "cmdPokazZawartosc"
        Me.cmdPokazZawartosc.Size = New System.Drawing.Size(128, 23)
        Me.cmdPokazZawartosc.TabIndex = 17
        Me.cmdPokazZawartosc.Text = "Poka¿ zawartoœæ"
        Me.cmdPokazZawartosc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'KatalogiSQL
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(632, 348)
        Me.Controls.Add(Me.cmdPokazZawartosc)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.Panel1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "KatalogiSQL"
        Me.ShowInTaskbar = False
        Me.Text = "Baza Danych"
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim NazwaTabeli As String

    Dim sConnection As String = ConnectionStrings()
    Dim sSelectSQL As String = "SELECT * FROM sysobjects " & _
                               "WHERE xtype='U'" & _
                               "ORDER BY name"
    Dim dbConnection As New SqlClient.SqlConnection(sConnection)
    Dim dbCommandSelect As New SqlClient.SqlCommand(sSelectSQL, dbConnection)
    Dim dbDataAdapter As New SqlClient.SqlDataAdapter(dbCommandSelect)
    Dim DataSetTabela As New DataSet
    Dim DataViewTabela As New DataView

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub KatalogiSQL_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        'Me.Text = oKomunikat
        lblPlik.Text = ParL_DataBase

        Dim MapowanieTabeli As System.Data.Common.DataTableMapping
        MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA")
        DataSetTabela.Locale = New System.Globalization.CultureInfo("pl-PL")
        '------------------------------------------
        'wypelnij DATASET
        Try
            dbConnection.Open()
        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        End Try

        If dbConnection.State = ConnectionState.Open Then
            Try
                dbDataAdapter.FillSchema(DataSetTabela, SchemaType.Mapped, "TABELA")
                dbDataAdapter.Fill(DataSetTabela)
                dbConnection.Close()

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                'DataSetTabela.Tables(0).PrimaryKey = New DataColumn() {DataSetTabela.Tables(0).Columns("DATANOTOW"), _
                '---------------------------------

                'definiuj DataView
                DataViewTabela = New DataView(DataSetTabela.Tables(0))
                DataViewTabela.AllowNew = False
                DataViewTabela.AllowDelete = False

                Dim ir As Long
                If DataViewTabela.Count > 0 Then
                    For ir = 0 To DataViewTabela.Count - 1
                        lbxTabele.Items.Add(DataViewTabela.Item(ir).Item("NAME"))
                    Next
                End If

            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            End Try
        End If

        'inicjuje tabele
        'lbxTabele.Items.Add("PARAMETRY")
        'lbxTabele.Items.Add("UZYTKOWNICY")
        'lbxTabele.Items.Add("UPRAWNIENIA")
        'lbxTabele.Items.Add("DOSTAWCY")
        'lbxTabele.Items.Add("ODBIORCY")
        'lbxTabele.Items.Add("PRODUCENCI")
        'lbxTabele.Items.Add("MAGAZYNY")
        'lbxTabele.Items.Add("TOWARY")
        'lbxTabele.Items.Add("ZAMIENNIKI")
        'lbxTabele.Items.Add("GRUPYTOWAROWE")
        'lbxTabele.Items.Add("SKLADOWENAZW")
        'lbxTabele.Items.Add("BANKI")
        'lbxTabele.Items.Add("KURSYWALUT")
        'lbxTabele.Items.Add("NUMERYFAKTUR")
        'lbxTabele.Items.Add("SPOSOBYDOSTAWY")
        'lbxTabele.Items.Add("TRASY")
        'lbxTabele.Items.Add("REGIONY")
        'lbxTabele.Items.Add("KOMUNIKATY")
        'lbxTabele.Items.Add("TRANSAKCJEOPS")
        'lbxTabele.Items.Add("TRANSAKCJESPC")
        'lbxTabele.Items.Add("BUFOROPS")
        'lbxTabele.Items.Add("BUFORSPC")
        'lbxTabele.Items.Add("DOKUMENTYOPS")
        'lbxTabele.Items.Add("DOKUMENTYSPC")
        'lbxTabele.Items.Add("WDRODZEOPS")
        'lbxTabele.Items.Add("WDRODZESPC")
        'lbxTabele.Items.Add("STANYMAGAZYNOWE")

        lbxTabele.Select()
    End Sub

    Private Sub lbxTabele_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbxTabele.Click
        WyswietlStrukture()
    End Sub

    Private Sub lbxTabele_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbxTabele.SelectedIndexChanged
        WyswietlStrukture()
    End Sub

    Private Sub cmdPokazZawartosc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPokazZawartosc.Click
        If Len(Trim(NazwaTabeli)) > 0 Then
            oTabela = NazwaTabeli
            Dim Form As New KatalogiSQLZawartosc
            Form.ShowDialog()
        End If
    End Sub

    Private Sub KatalogiSQL_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub WyswietlStrukture()
        If lbxTabele.Items.Count > 0 Then
            If lbxTabele.SelectedIndex >= 0 Then
                NazwaTabeli = lbxTabele.SelectedItem.ToString

                Dim sConnection As String = ConnectionStrings()
                Dim sSelectSQL As String = "SELECT * FROM " & NazwaTabeli
                Dim dbConnection As New SqlClient.SqlConnection(sConnection)
                Dim dbCommandSelect As New SqlClient.SqlCommand(sSelectSQL, dbConnection)
                Dim dbDataAdapter As New SqlClient.SqlDataAdapter(dbCommandSelect)
                Dim DataSetTabela As New DataSet
                Dim DataTableTabela As DataTable

                '------------------------------------------
                'wypelnij DATASET
                Try
                    dbConnection.Open()
                Catch Ex As System.Exception
                    MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                        System.Windows.Forms.MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, _
                        MessageBoxDefaultButton.Button1)
                End Try

                If dbConnection.State = ConnectionState.Open Then
                    Try
                        dbDataAdapter.FillSchema(DataSetTabela, SchemaType.Mapped, "TABELA")
                        dbDataAdapter.Fill(DataSetTabela)
                        dbConnection.Close()

                        DataTableTabela = DataSetTabela.Tables("TABELA")
                        Dim objColumn As DataColumn
                        Dim i As Integer = 0
                        ListView.Items.Clear()
                        ListView.ForeColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(192, Byte))
                        For Each objColumn In DataTableTabela.Columns
                            ListView.Items.Add(objColumn.ColumnName, i)
                            ListView.Items(i).SubItems.Add(objColumn.DataType.Name)
                            ListView.Items(i).SubItems.Add(IIf(objColumn.MaxLength < 0, "", objColumn.MaxLength.ToString))
                            ListView.Items(i).SubItems.Add(IIf(objColumn.Unique, "T", "N"))
                            ListView.Items(i).SubItems.Add(IIf(objColumn.ReadOnly, "T", "N"))
                            i += 1
                        Next

                    Catch Ex As System.Exception
                        MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                            System.Windows.Forms.MessageBoxButtons.OK, _
                            MessageBoxIcon.Information, _
                            MessageBoxDefaultButton.Button1)
                    End Try

                End If
            End If
        End If
    End Sub
End Class
