Imports System.Drawing.Printing

Public Class KatalogDaneDoRaportu
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
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents cmdEdytuj As System.Windows.Forms.Button
    Friend WithEvents cmdDodaj As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents stbDaneDoRaportu As System.Windows.Forms.StatusBar
    Friend WithEvents cmdWycofajsie As System.Windows.Forms.Button
    Friend WithEvents cmdUsuñ As System.Windows.Forms.Button
    Friend WithEvents cmdWszystko As System.Windows.Forms.Button
    Friend WithEvents cmdFi As System.Windows.Forms.Button
    Friend WithEvents cmdHarmonogramWizyt As System.Windows.Forms.Button
    Friend WithEvents dagDaneDoRaportu As DataGridView
    Friend WithEvents HorizScrollTimer As Timer
    Friend WithEvents menuWybieraj As ContextMenuStrip
    Friend WithEvents menuWEdytuj As ToolStripMenuItem
    Friend WithEvents menuWDodaj As ToolStripMenuItem
    Friend WithEvents menuWUsun As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents menuWWybierz As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents menuWSingleL As ToolStripMenuItem
    Friend WithEvents menuWMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents menuWOdswiez As ToolStripMenuItem
    Friend WithEvents menuEdytuj As ContextMenuStrip
    Friend WithEvents menuEEdytuj As ToolStripMenuItem
    Friend WithEvents menuEDodaj As ToolStripMenuItem
    Friend WithEvents menuEUsun As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents menuESingleL As ToolStripMenuItem
    Friend WithEvents menuEMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents menuEOdswiez As ToolStripMenuItem
    Friend WithEvents stbFiltry As System.Windows.Forms.StatusBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogDaneDoRaportu))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.stbDaneDoRaportu = New System.Windows.Forms.StatusBar()
        Me.cmdEdytuj = New System.Windows.Forms.Button()
        Me.cmdUsuñ = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdWycofajsie = New System.Windows.Forms.Button()
        Me.cmdWszystko = New System.Windows.Forms.Button()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.cmdHarmonogramWizyt = New System.Windows.Forms.Button()
        Me.dagDaneDoRaportu = New System.Windows.Forms.DataGridView()
        Me.HorizScrollTimer = New System.Windows.Forms.Timer(Me.components)
        Me.menuWybieraj = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuWEdytuj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuWDodaj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuWUsun = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuWWybierz = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuWSingleL = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuWMultiL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuWOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEdytuj = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuEEdytuj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEDodaj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEUsun = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuESingleL = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEMultiL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuEOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dagDaneDoRaportu, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuWybieraj.SuspendLayout()
        Me.menuEdytuj.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(690, 333)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(88, 23)
        Me.cmdStop.TabIndex = 37
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'stbDaneDoRaportu
        '
        Me.stbDaneDoRaportu.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbDaneDoRaportu.Dock = System.Windows.Forms.DockStyle.None
        Me.stbDaneDoRaportu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbDaneDoRaportu.Location = New System.Drawing.Point(8, 303)
        Me.stbDaneDoRaportu.Name = "stbDaneDoRaportu"
        Me.stbDaneDoRaportu.ShowPanels = True
        Me.stbDaneDoRaportu.Size = New System.Drawing.Size(762, 22)
        Me.stbDaneDoRaportu.TabIndex = 42
        Me.stbDaneDoRaportu.Text = "StatusBar1"
        '
        'cmdEdytuj
        '
        Me.cmdEdytuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdytuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdEdytuj.Image = CType(resources.GetObject("cmdEdytuj.Image"), System.Drawing.Image)
        Me.cmdEdytuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEdytuj.Location = New System.Drawing.Point(490, 333)
        Me.cmdEdytuj.Name = "cmdEdytuj"
        Me.cmdEdytuj.Size = New System.Drawing.Size(80, 23)
        Me.cmdEdytuj.TabIndex = 41
        Me.cmdEdytuj.Text = "Edytuj"
        Me.cmdEdytuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdUsuñ
        '
        Me.cmdUsuñ.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUsuñ.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdUsuñ.Image = CType(resources.GetObject("cmdUsuñ.Image"), System.Drawing.Image)
        Me.cmdUsuñ.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUsuñ.Location = New System.Drawing.Point(402, 333)
        Me.cmdUsuñ.Name = "cmdUsuñ"
        Me.cmdUsuñ.Size = New System.Drawing.Size(80, 23)
        Me.cmdUsuñ.TabIndex = 40
        Me.cmdUsuñ.Text = "Usuñ"
        Me.cmdUsuñ.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdDodaj
        '
        Me.cmdDodaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDodaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDodaj.Image = CType(resources.GetObject("cmdDodaj.Image"), System.Drawing.Image)
        Me.cmdDodaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDodaj.Location = New System.Drawing.Point(314, 333)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(80, 23)
        Me.cmdDodaj.TabIndex = 39
        Me.cmdDodaj.Text = "Dodaj"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(16, 333)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 38
        Me.PictureBox1.TabStop = False
        '
        'cmdWycofajsie
        '
        Me.cmdWycofajsie.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWycofajsie.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWycofajsie.Image = CType(resources.GetObject("cmdWycofajsie.Image"), System.Drawing.Image)
        Me.cmdWycofajsie.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWycofajsie.Location = New System.Drawing.Point(578, 333)
        Me.cmdWycofajsie.Name = "cmdWycofajsie"
        Me.cmdWycofajsie.Size = New System.Drawing.Size(104, 23)
        Me.cmdWycofajsie.TabIndex = 45
        Me.cmdWycofajsie.Text = "Wycofaj siê"
        Me.cmdWycofajsie.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdWszystko
        '
        Me.cmdWszystko.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystko.Location = New System.Drawing.Point(552, 7)
        Me.cmdWszystko.Name = "cmdWszystko"
        Me.cmdWszystko.Size = New System.Drawing.Size(26, 20)
        Me.cmdWszystko.TabIndex = 66
        '
        'stbFiltry
        '
        Me.stbFiltry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltry.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltry.Location = New System.Drawing.Point(8, 8)
        Me.stbFiltry.Name = "stbFiltry"
        Me.stbFiltry.ShowPanels = True
        Me.stbFiltry.Size = New System.Drawing.Size(770, 22)
        Me.stbFiltry.TabIndex = 85
        Me.stbFiltry.Text = "StatusBar1"
        '
        'cmdFi
        '
        Me.cmdFi.Image = CType(resources.GetObject("cmdFi.Image"), System.Drawing.Image)
        Me.cmdFi.Location = New System.Drawing.Point(528, 8)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(16, 20)
        Me.cmdFi.TabIndex = 175
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'cmdHarmonogramWizyt
        '
        Me.cmdHarmonogramWizyt.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdHarmonogramWizyt.ForeColor = System.Drawing.Color.Black
        Me.cmdHarmonogramWizyt.Image = CType(resources.GetObject("cmdHarmonogramWizyt.Image"), System.Drawing.Image)
        Me.cmdHarmonogramWizyt.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdHarmonogramWizyt.Location = New System.Drawing.Point(126, 332)
        Me.cmdHarmonogramWizyt.Name = "cmdHarmonogramWizyt"
        Me.cmdHarmonogramWizyt.Size = New System.Drawing.Size(24, 23)
        Me.cmdHarmonogramWizyt.TabIndex = 205
        Me.cmdHarmonogramWizyt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dagDaneDoRaportu
        '
        Me.dagDaneDoRaportu.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagDaneDoRaportu.Location = New System.Drawing.Point(8, 32)
        Me.dagDaneDoRaportu.Name = "dagDaneDoRaportu"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagDaneDoRaportu.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagDaneDoRaportu.Size = New System.Drawing.Size(770, 269)
        Me.dagDaneDoRaportu.TabIndex = 206
        '
        'HorizScrollTimer
        '
        '
        'menuWybieraj
        '
        Me.menuWybieraj.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuWEdytuj, Me.menuWDodaj, Me.menuWUsun, Me.ToolStripSeparator1, Me.menuWWybierz, Me.ToolStripSeparator4, Me.menuWSingleL, Me.menuWMultiL, Me.ToolStripSeparator5, Me.menuWOdswiez})
        Me.menuWybieraj.Name = "menuWybieraj"
        Me.menuWybieraj.Size = New System.Drawing.Size(187, 176)
        '
        'menuWEdytuj
        '
        Me.menuWEdytuj.Name = "menuWEdytuj"
        Me.menuWEdytuj.Size = New System.Drawing.Size(186, 22)
        Me.menuWEdytuj.Text = "Edytuj"
        '
        'menuWDodaj
        '
        Me.menuWDodaj.Name = "menuWDodaj"
        Me.menuWDodaj.Size = New System.Drawing.Size(186, 22)
        Me.menuWDodaj.Text = "Dodaj"
        '
        'menuWUsun
        '
        Me.menuWUsun.Name = "menuWUsun"
        Me.menuWUsun.Size = New System.Drawing.Size(186, 22)
        Me.menuWUsun.Text = "Usuñ"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(183, 6)
        '
        'menuWWybierz
        '
        Me.menuWWybierz.Name = "menuWWybierz"
        Me.menuWWybierz.Size = New System.Drawing.Size(186, 22)
        Me.menuWWybierz.Text = "Wybierz"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(183, 6)
        '
        'menuWSingleL
        '
        Me.menuWSingleL.Name = "menuWSingleL"
        Me.menuWSingleL.Size = New System.Drawing.Size(186, 22)
        Me.menuWSingleL.Text = "Poka¿ w jednej linii"
        '
        'menuWMultiL
        '
        Me.menuWMultiL.Name = "menuWMultiL"
        Me.menuWMultiL.Size = New System.Drawing.Size(186, 22)
        Me.menuWMultiL.Text = "Poka¿ w wielu liniach"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(183, 6)
        '
        'menuWOdswiez
        '
        Me.menuWOdswiez.Name = "menuWOdswiez"
        Me.menuWOdswiez.Size = New System.Drawing.Size(186, 22)
        Me.menuWOdswiez.Text = "Odswie¿ Tabelê"
        '
        'menuEdytuj
        '
        Me.menuEdytuj.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuEEdytuj, Me.menuEDodaj, Me.menuEUsun, Me.ToolStripSeparator2, Me.menuESingleL, Me.menuEMultiL, Me.ToolStripSeparator3, Me.menuEOdswiez})
        Me.menuEdytuj.Name = "menuEdytuj"
        Me.menuEdytuj.Size = New System.Drawing.Size(187, 148)
        '
        'menuEEdytuj
        '
        Me.menuEEdytuj.Name = "menuEEdytuj"
        Me.menuEEdytuj.Size = New System.Drawing.Size(186, 22)
        Me.menuEEdytuj.Text = "Edytuj"
        '
        'menuEDodaj
        '
        Me.menuEDodaj.Name = "menuEDodaj"
        Me.menuEDodaj.Size = New System.Drawing.Size(186, 22)
        Me.menuEDodaj.Text = "Dodaj"
        '
        'menuEUsun
        '
        Me.menuEUsun.Name = "menuEUsun"
        Me.menuEUsun.Size = New System.Drawing.Size(186, 22)
        Me.menuEUsun.Text = "Usuñ"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(183, 6)
        '
        'menuESingleL
        '
        Me.menuESingleL.Name = "menuESingleL"
        Me.menuESingleL.Size = New System.Drawing.Size(186, 22)
        Me.menuESingleL.Text = "Poka¿ w jednej linii"
        '
        'menuEMultiL
        '
        Me.menuEMultiL.Name = "menuEMultiL"
        Me.menuEMultiL.Size = New System.Drawing.Size(186, 22)
        Me.menuEMultiL.Text = "Poka¿ w wielu liniach"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(183, 6)
        '
        'menuEOdswiez
        '
        Me.menuEOdswiez.Name = "menuEOdswiez"
        Me.menuEOdswiez.Size = New System.Drawing.Size(186, 22)
        Me.menuEOdswiez.Text = "Odœwie¿ Tabelê"
        '
        'KatalogDaneDoRaportu
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(784, 361)
        Me.Controls.Add(Me.dagDaneDoRaportu)
        Me.Controls.Add(Me.cmdHarmonogramWizyt)
        Me.Controls.Add(Me.cmdFi)
        Me.Controls.Add(Me.cmdWszystko)
        Me.Controls.Add(Me.stbFiltry)
        Me.Controls.Add(Me.cmdWycofajsie)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.stbDaneDoRaportu)
        Me.Controls.Add(Me.cmdEdytuj)
        Me.Controls.Add(Me.cmdUsuñ)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.PictureBox1)
        Me.Name = "KatalogDaneDoRaportu"
        Me.Text = "Dane uzupe³niajace do Raportu Aktywnoœci Pracowników"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dagDaneDoRaportu, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menuWybieraj.ResumeLayout(False)
        Me.menuEdytuj.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim StartFormy As Boolean = True
    Dim _OCoMamRobic As String = oCoMamRobic

    '---------------------------------------------
    'SCROLL-parametry obslugi scrollingu poziomego
    Dim StartPointInCell00 As System.Drawing.Point             'pocz¹tek ektanu przed scrollingiem
    Dim ScrollXOffset As Integer = 0            'bie¿¹cy punkt ekranu wskazywany przez kursor podczas scrollingu
    Dim HorizScrollLastTime As String = ""      'Chwila rozpoczecia scrollingu - filtry wyswietlaner bêd¹ po 1 sekundzie...
    ' HorizScrollTimer - zegar odliczaj¹cy czas od poczatku scrollingu
    '---------------------------------------------



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

    Dim DataSetDaneDoRaportu As New DataSet
    Dim DataViewDaneDoRaportu As New DataView
    Dim ConnectionState As System.Data.ConnectionState


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

    'Public ddrKolExcel01 As String            'KolExCEL01           String
    'Public ddrKolExcel02 As String            'KolExCEL02           String
    'Public ddrKolExcel03 As String            'KolExCEL03           String
    'Public ddrKolExcel04 As String            'KolExCEL04           String
    'Public ddrKolExcel05 As String            'KolExCEL05           String
    'Public ddrKolExcel06 As String            'KolExCEL06           String
    'Public ddrKolExcel07 As String            'KolExCEL07           String
    'Public ddrKolExcel08 As String            'KolExCEL08           String
    'Public ddrKolExcel09 As String            'KolExCEL09           String
    'Public ddrKolExcel10 As String            'KolExCEL10           String

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

    'Public ddrWersja As Integer          'WERSJA         Integer



    Private Sub KatalogDaneDoRaportu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        DataSetDaneDoRaportu.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnection = New OleDb.OleDbConnection(sConnectionDaneDoRaportu)
            dbCommandSelect = New OleDb.OleDbCommand(sSelectSQLDaneDoRaportu, dbConnection)
            dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLDaneDoRaportu, dbConnection)
            dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLDaneDoRaportu, dbConnection)
            dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLDaneDoRaportu, dbConnection)
            dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim DBMapowanieTabeli As System.Data.Common.DataTableMapping
            DBMapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_DaneDoRaportu")

            DBDeleteDaneDoRaportu(dbCommandDelete, dbDataAdapter)
            DBUpdateDaneDoRaportu(dbCommandUpdate, dbDataAdapter)
            DBInsertDaneDoRaportu(dbCommandInsert, dbDataAdapter)

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
            sqlConnection = New SqlClient.SqlConnection(sConnectionDaneDoRaportu)
            sqlCommandSelect = New SqlClient.SqlCommand(sSelectSQLDaneDoRaportu, sqlConnection)
            sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLDaneDoRaportu, sqlConnection)
            sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLDaneDoRaportu, sqlConnection)
            sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLDaneDoRaportu, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeli As System.Data.Common.DataTableMapping
            sqlMapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_DaneDoRaportu")

            SQLDeleteDaneDoRaportu(sqlCommandDelete, sqlDataAdapter)
            SQLUpdateDaneDoRaportu(sqlCommandUpdate, sqlDataAdapter)
            SQLInsertDaneDoRaportu(sqlCommandInsert, sqlDataAdapter)

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
                    dbDataAdapter.FillSchema(DataSetDaneDoRaportu, SchemaType.Mapped, "TABELA_DaneDoRaportu")
                    dbDataAdapter.Fill(DataSetDaneDoRaportu)
                    dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(DataSetDaneDoRaportu, SchemaType.Mapped, "TABELA_DaneDoRaportu")
                    sqlDataAdapter.Fill(DataSetDaneDoRaportu)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetDaneDoRaportu.Tables(0).PrimaryKey = New DataColumn() {DataSetDaneDoRaportu.Tables(0).Columns("PRACOWNIK"),
                                                                              DataSetDaneDoRaportu.Tables(0).Columns("DATARAPORTU")}

                'definiuj DataView
                DataViewDaneDoRaportu = New DataView(DataSetDaneDoRaportu.Tables(0))
                DataViewDaneDoRaportu.AllowDelete = False
                DataViewDaneDoRaportu.AllowNew = False

                '-------------------------------------------------
                'parametry DataGridView
                '-------------------------------------------------
                ' Initialize basic DataGridView properties.
                dagDaneDoRaportu.BackgroundColor = System.Drawing.SystemColors.Control
                dagDaneDoRaportu.GridColor = System.Drawing.SystemColors.ControlDark
                dagDaneDoRaportu.ForeColor = System.Drawing.Color.Purple
                dagDaneDoRaportu.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagDaneDoRaportu.BorderStyle = BorderStyle.Fixed3D
                'dagDaneDoRaportu.Dock = DockStyle.Fill
                dagDaneDoRaportu.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                ' Set property values appropriate for read-only display and limited interactivity. 
                dagDaneDoRaportu.AllowUserToAddRows = False
                dagDaneDoRaportu.AllowUserToDeleteRows = False
                dagDaneDoRaportu.AllowUserToOrderColumns = True
                dagDaneDoRaportu.AllowUserToResizeColumns = True
                dagDaneDoRaportu.AllowUserToResizeRows = True
                dagDaneDoRaportu.ReadOnly = True
                dagDaneDoRaportu.SelectionMode = DataGridViewSelectionMode.CellSelect
                dagDaneDoRaportu.MultiSelect = False

                'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                'dagDaneDoRaportu.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                dagDaneDoRaportu.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                dagDaneDoRaportu.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                dagDaneDoRaportu.ColumnHeadersVisible = True
                dagDaneDoRaportu.ColumnHeadersHeight = 40
                dagDaneDoRaportu.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                dagDaneDoRaportu.RowHeadersVisible = True
                dagDaneDoRaportu.RowHeadersWidth = 50
                dagDaneDoRaportu.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                ' Set the selection background color for all the cells.
                dagDaneDoRaportu.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                dagDaneDoRaportu.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                dagDaneDoRaportu.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagDaneDoRaportu.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                dagDaneDoRaportu.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagDaneDoRaportu.DefaultCellStyle.NullValue = ""
                dagDaneDoRaportu.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                dagDaneDoRaportu.DefaultCellStyle.WrapMode = DataGridViewTriState.False         'single...


                ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                dagDaneDoRaportu.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                dagDaneDoRaportu.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                ' Set the background color for all rows and for alternating rows.  
                ' The value for alternating rows overrides the value for all rows. 
                dagDaneDoRaportu.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagDaneDoRaportu.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                ' Set the row and column header styles.
                dagDaneDoRaportu.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagDaneDoRaportu.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                dagDaneDoRaportu.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagDaneDoRaportu.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                '-----------------------------------
                'wypelnienie DataGridView
                dagDaneDoRaportu.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                dagDaneDoRaportu.DataMember = DataSetDaneDoRaportu.Tables(0).TableName
                'dagDaneDoRaportu.DataSource = DataViewDaneDoRaportu
                '-----------------------------------

                ' Add a GridColumnStyle and set the MappingName 
                ' to the name of a DataColumn in the DataTable. 
                ' Set the HeaderText and Width properties. 

                'Public ddrPracownik As String        'PRACOWNIK   Text, 10
                'Public ddrDataRaportu As String      'DATARAPORTU Text, 7   RRRR-MM
                'Public ddrKlBezDrukarki As Integer   'KLBEZDRUKARKI  Integer
                'Public ddrOferty As Integer          'OFERTY         Integer
                'Public ddrTestyWstawionw As Integer  'TESTYWSTAWIONE Integer
                'Public ddrCzyszczenie As Integer     'CZYSZCZENIE    Integer
                'Public ddrDostawy As Integer         'DOSTAWY        Integer
                'Public ddrSkup As Integer            'SKUP           Integer

                'Public ddrKolExcel01 As String            'KolExCEL01           String
                'Public ddrKolExcel02 As String            'KolExCEL02           String
                'Public ddrKolExcel03 As String            'KolExCEL03           String
                'Public ddrKolExcel04 As String            'KolExCEL04           String
                'Public ddrKolExcel05 As String            'KolExCEL05           String
                'Public ddrKolExcel06 As String            'KolExCEL06           String
                'Public ddrKolExcel07 As String            'KolExCEL07           String
                'Public ddrKolExcel08 As String            'KolExCEL08           String
                'Public ddrKolExcel09 As String            'KolExCEL09           String
                'Public ddrKolExcel10 As String            'KolExCEL10           String

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

                'Public ddrWersja As Integer          'WERSJA         Integer

                TxtColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(0).ColumnName, "Pracownik", 80, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(1).ColumnName, "Data", 80, DataGridViewContentAlignment.MiddleLeft, True)
                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(2).ColumnName, "Klient bez drukarki", 80, DataGridViewContentAlignment.MiddleRight, "N0", True)
                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(3).ColumnName, "Oferty", 80, DataGridViewContentAlignment.MiddleRight, "N0", True)
                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(4).ColumnName, "Testy wstawione", 80, DataGridViewContentAlignment.MiddleRight, "N0", True)
                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(5).ColumnName, "Czyszczenie", 80, DataGridViewContentAlignment.MiddleRight, "N0", True)
                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(6).ColumnName, "Dostawy", 80, DataGridViewContentAlignment.MiddleRight, "N0", True)
                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(7).ColumnName, "Skup", 80, DataGridViewContentAlignment.MiddleRight, "N0", True)

                TxtColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(8).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(9).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(10).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(11).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(12).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(13).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(14).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(15).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(16).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(17).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)

                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(18).ColumnName, Par_RAKolumna01, 80, DataGridViewContentAlignment.MiddleRight, "N0", True)
                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(19).ColumnName, Par_RAKolumna02, 80, DataGridViewContentAlignment.MiddleRight, "N0", True)
                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(20).ColumnName, Par_RAKolumna03, 80, DataGridViewContentAlignment.MiddleRight, "N0", True)
                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(21).ColumnName, Par_RAKolumna04, 80, DataGridViewContentAlignment.MiddleRight, "N0", True)
                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(22).ColumnName, Par_RAKolumna05, 80, DataGridViewContentAlignment.MiddleRight, "N0", True)
                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(23).ColumnName, Par_RAKolumna06, 80, DataGridViewContentAlignment.MiddleRight, "N0", True)
                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(24).ColumnName, Par_RAKolumna07, 80, DataGridViewContentAlignment.MiddleRight, "N0", True)
                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(25).ColumnName, Par_RAKolumna08, 80, DataGridViewContentAlignment.MiddleRight, "N0", True)
                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(26).ColumnName, Par_RAKolumna09, 80, DataGridViewContentAlignment.MiddleRight, "N0", True)
                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(27).ColumnName, Par_RAKolumna10, 80, DataGridViewContentAlignment.MiddleRight, "N0", True)

                NumColumnView(dagDaneDoRaportu, DataSetDaneDoRaportu.Tables(0).Columns(28).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N0", False)

                Me.Cursor = Cursors.WaitCursor
                dagDaneDoRaportu.DataSource = DataViewDaneDoRaportu
                Me.Cursor = Cursors.Default

                'ustaw sie na pierwszym zapisie
                If DataSetDaneDoRaportu.Tables(0).Rows.Count > 0 Then
                    dagDaneDoRaportu.CurrentCell = dagDaneDoRaportu(0, 0)
                    dagDaneDoRaportu.CurrentCell.Selected = True
                End If


            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            '------------------------------------
            'dodaj StatusBar na koncu
            stbDaneDoRaportu.BackColor = System.Drawing.SystemColors.Control
            stbDaneDoRaportu.ForeColor = System.Drawing.Color.DarkGreen
            stbDaneDoRaportu.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            stbDaneDoRaportu.Panels.Add("stbPanelIleRec")
            stbDaneDoRaportu.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbDaneDoRaportu.Panels(0).Width = 80
            stbDaneDoRaportu.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbDaneDoRaportu.Panels(0).Text = "Iloœæ rec.: " & DataSetDaneDoRaportu.Tables(0).Rows.Count.ToString

            stbDaneDoRaportu.Panels.Add("stbPanelWiersz")
            stbDaneDoRaportu.Panels(1).AutoSize = StatusBarPanelAutoSize.None
            stbDaneDoRaportu.Panels(1).Width = 80
            stbDaneDoRaportu.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagDaneDoRaportu.CurrentCell Is Nothing Then
                stbDaneDoRaportu.Panels(1).Text = "Wiersz : "
            Else
                stbDaneDoRaportu.Panels(1).Text = "Wiersz : " & dagDaneDoRaportu.CurrentCell.RowIndex.ToString
            End If

            stbDaneDoRaportu.Panels.Add("stbPanelKolumna")
            stbDaneDoRaportu.Panels(2).AutoSize = StatusBarPanelAutoSize.None
            stbDaneDoRaportu.Panels(2).Width = 80
            stbDaneDoRaportu.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagDaneDoRaportu.CurrentCell Is Nothing Then
                stbDaneDoRaportu.Panels(2).Text = "Kolumna : "
            Else
                stbDaneDoRaportu.Panels(2).Text = "Kolumna : " & dagDaneDoRaportu.CurrentCell.ColumnIndex.ToString
            End If

            stbDaneDoRaportu.Panels.Add("stbPanelSort")
            stbDaneDoRaportu.Panels(3).AutoSize = StatusBarPanelAutoSize.None
            stbDaneDoRaportu.Panels(3).Width = 150
            stbDaneDoRaportu.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbDaneDoRaportu.Panels(3).Text = "Sort : "

            stbDaneDoRaportu.Panels.Add("stbPanelFiltr")
            stbDaneDoRaportu.Panels(4).AutoSize = StatusBarPanelAutoSize.None
            stbDaneDoRaportu.Panels(4).Width = 150
            stbDaneDoRaportu.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbDaneDoRaportu.Panels(4).Text = "Filtr : "

            stbDaneDoRaportu.Panels.Add("stbPanelReszta")
            stbDaneDoRaportu.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
            'stbDaneDoRaportu.Panels(5).Width = 150
            stbDaneDoRaportu.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbDaneDoRaportu.Panels(5).Text = ""

            stbDaneDoRaportu.ShowPanels = True
            '---------------------------------
            'dodaj StatusBar na koncu
            stbFiltry.BackColor = System.Drawing.SystemColors.Control
            stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
            stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            stbFiltry.Panels.Add("stbFiltryRowHead")
            stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbFiltry.Panels(0).Width = 50  'dagDaneDoRaportu.TableStyles(0).RowHeaderWidth
            stbFiltry.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbFiltry.Panels(0).Text = "Filtry"

            stbFiltry.Panels.Add("stbFiltryReszta")
            stbFiltry.Panels(1).AutoSize = StatusBarPanelAutoSize.Spring
            'stbFiltry.Panels(1).Width = 150
            stbFiltry.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbFiltry.Panels(1).Text = ""

            stbFiltry.ShowPanels = True

            'dodaj ctrl do pobierania szablonu
            'cmdWszystko.Size = New Size(20, 20)
            'cmdWszystko.Location = New system.drawing.Point(stbFiltry.Location.X + _
            '                                 stbFiltry.Size.Width - _
            '                                 cmdWszystko.Size.Width, _
            '                                 stbFiltry.Location.Y + 2)
            'cmdWszystko.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or _
            '                            System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            cmdWszystko.Size = New Size(20, 22)
            cmdWszystko.Location = New System.Drawing.Point(stbFiltry.Location.X + 50 - 20,
                                             stbFiltry.Location.Y)
            cmdWszystko.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or
                                        System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)

        End If
        InitDaneDoRaportu() 'inicjuj zmienne
        '-----------------------------------------------------------------
        If _OCoMamRobic = "WYBIERAJ" Then
            dagDaneDoRaportu.ContextMenuStrip = Me.menuWybieraj
            cmdStop.Text = "Wybierz"
            menuWybieraj.Visible = False
            menuWybieraj.Enabled = True
            menuEdytuj.Visible = False
            menuEdytuj.Enabled = False
        Else
            dagDaneDoRaportu.ContextMenuStrip = Me.menuEdytuj
            cmdStop.Text = "Powrót"
            menuWybieraj.Visible = False
            menuWybieraj.Enabled = False
            menuEdytuj.Visible = False
            menuEdytuj.Enabled = True
        End If
        '--------------------------------------------------------------------
        'set the header width to something apporpriate
        dagDaneDoRaportu.RowHeadersWidth = 50       'use if tablestyle
        'Me.dagDaneDoRaportu.RowHeaderWidth = 80 'use if no tablestyle
        '--------------------------------------------------------------------
        'set topleft corner point so we can locate the toprow
        If DataSetDaneDoRaportu.Tables(0).Rows.Count > 0 Then
            StartPointInCell00 = New System.Drawing.Point((dagDaneDoRaportu.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (dagDaneDoRaportu.GetCellDisplayRectangle(0, 0, True).Y + 4))
            ScrollXOffset = 10000 * dagDaneDoRaportu.FirstDisplayedScrollingColumnIndex +
                        dagDaneDoRaportu.GetCellDisplayRectangle(dagDaneDoRaportu.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
        Else
            'brak zapisow - poczatek DataGrid
            StartPointInCell00 = New System.Drawing.Point((dagDaneDoRaportu.Left + 4), (dagDaneDoRaportu.Top + 4))
            ScrollXOffset = 0
        End If
        '--------------------------------------------------
        'inicjowanie zegara dla scrollingu poziomego
        HorizScrollLastTime = ""
        HorizScrollTimer.Interval = 2 * 1000
        HorizScrollTimer.Enabled = True
        '--------------------------------------------------
        'Stworz zmienne do filtrowania
        Dim ii As Integer = 0
        For ii = 0 To DataViewDaneDoRaportu.Table.Columns.Count - 1
            'stworz zmienn¹ TXTBOX
            Dim CtrlT As New System.Windows.Forms.TextBox
            CtrlT.Name = "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
            CtrlT.Visible = False
            Me.Controls.Add(CtrlT)
            AddHandler CtrlT.TextChanged, New EventHandler(AddressOf txtFiXX_TextChanged)
            AddHandler CtrlT.GotFocus, New EventHandler(AddressOf txtFiXX_GotFocus)
            AddHandler CtrlT.LostFocus, New EventHandler(AddressOf txtFiXX_LostFocus)
            AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXX_KeyDown)

            'stworz zmienn¹ BUTTON
            Dim CtrlB As New System.Windows.Forms.Button
            CtrlB.Name = "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
            CtrlB.Image = cmdFi.Image
            CtrlB.ImageAlign = ContentAlignment.MiddleCenter
            CtrlB.Visible = False
            Me.Controls.Add(CtrlB)
            AddHandler CtrlB.Click, New EventHandler(AddressOf cmdFiXX_Click)
            AddHandler CtrlB.GotFocus, New EventHandler(AddressOf cmdFiXX_GotFocus)
            AddHandler CtrlB.LostFocus, New EventHandler(AddressOf cmdFiXX_LostFocus)
        Next
        Me.Refresh()
        '--------------------------------------------------
        StartFormy = False
        dagDaneDoRaportu_CurrentCellChanged(sender, e)
        InicjujFiltryColumnDGV(Me, dagDaneDoRaportu, DataViewDaneDoRaportu.Table.Columns.Count, stbFiltry, StartFormy)
    End Sub


    Private Sub KatalogDaneDoRaportu_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub




















    Private Sub dagDaneDoRaportu_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagDaneDoRaportu.CurrentCellChanged
        If Not StartFormy Then
            If dagDaneDoRaportu.CurrentCell Is Nothing Then
                stbDaneDoRaportu.Panels(1).Text = "Wi: "
                stbDaneDoRaportu.Panels(2).Text = "Ko: "
            Else
                stbDaneDoRaportu.Panels(1).Text = "Wi: " & dagDaneDoRaportu.CurrentCell.RowIndex.ToString
                stbDaneDoRaportu.Panels(2).Text = "Ko: " & dagDaneDoRaportu.CurrentCell.ColumnIndex.ToString
            End If
        End If
    End Sub

    '-=------------------------------------------
    'obsluga zegara scrollingu poziomego
    '-=------------------------------------------
    Private Sub HorizScrollTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HorizScrollTimer.Tick
        HorizScrollTimer.Enabled = False
        If Len(HorizScrollLastTime) = 0 Then
            'nie zainicjowano scrollingu
        Else
            'zainicjowano scrolling - sprawdz czy juz trwa 2 sec
            If IleSekundMinelo(HorizScrollLastTime, SysGodz()) >= 0 Then
                If dagDaneDoRaportu.RowCount > 0 Then
                    If ScrollXOffset <> 10000 * dagDaneDoRaportu.FirstDisplayedScrollingColumnIndex +
                                    dagDaneDoRaportu.GetCellDisplayRectangle(dagDaneDoRaportu.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X Then

                        PokazFiltryColumnDGV(Me, dagDaneDoRaportu, DataViewDaneDoRaportu.Table.Columns.Count, stbFiltry, StartFormy)
                        '-----------------------------
                        ScrollXOffset = 10000 * dagDaneDoRaportu.FirstDisplayedScrollingColumnIndex +
                                    dagDaneDoRaportu.GetCellDisplayRectangle(dagDaneDoRaportu.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X
                        HorizScrollLastTime = ""
                    End If
                Else
                    PokazFiltryColumnDGV(Me, dagDaneDoRaportu, DataViewDaneDoRaportu.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Else
            End If
        End If
        HorizScrollTimer.Enabled = True
    End Sub


    Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Dim rc As DataGridViewSelectedRowCollection = dagDaneDoRaportu.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                dagDaneDoRaportu.Rows(nextrow).Selected = True
            End If
        End If
    End Sub



    '*********************************************************
    '***** Filtrowanie wg poszczególnych pól
    '*********************************************************

    'AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXX_KeyDown)

    Private Sub txtFiXX_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Not StartFormy Then
            Select Case e.KeyCode
                'Case Keys.Enter
                'Case Keys.Insert, Keys.Add
                'Case Keys.Delete, Keys.Subtract

                Case Keys.Down
                    PrzejdzDoDGV(dagDaneDoRaportu, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            FiltrujDataViewDGV(Me, dagDaneDoRaportu, DataViewDaneDoRaportu.Table.Columns.Count, DataViewDaneDoRaportu, stbDaneDoRaportu)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me, dagDaneDoRaportu, Mid(sender.name, 6, 3), "DaneDoRaportu")
    End Sub
    Private Sub cmdFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiFiltraCmd(Me, sender)
    End Sub
    Private Sub cmdFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraCmd(Me, sender)
    End Sub


    '-------------------------------------------------
    ' filtrowanie 
    '-------------------------------------------------

    Private Sub cmdWszystko_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWszystko.Click
        Me.Cursor = Cursors.WaitCursor
        oLokalnyFiltr = ""
        Try
            StartFormy = True
            CzyscFiltryDGV(Me, dagDaneDoRaportu, DataViewDaneDoRaportu.Table.Columns.Count, DataViewDaneDoRaportu, stbDaneDoRaportu)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbDaneDoRaportu.Panels(0).Text = "Il.zap.: " & DataViewDaneDoRaportu.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagDaneDoRaportu_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagDaneDoRaportu.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagDaneDoRaportu.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagDaneDoRaportu.DefaultCellStyle.Font,
                      b,
                      e.RowBounds.Location.X + 15,
                      e.RowBounds.Location.Y + 4)
        End Using
    End Sub



    '*************************************************
    '*** Odœwiezanie filtrow kolumnowych 
    '*************************************************

    Private Sub stbFiltry_PanelClick(sender As Object, e As StatusBarPanelClickEventArgs) Handles stbFiltry.PanelClick
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagDaneDoRaportu, DataViewDaneDoRaportu.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagDaneDoRaportu_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagDaneDoRaportu.CellFormatting
    '    If Not StartFormy Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagDaneDoRaportu, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagDaneDoRaportu, e.RowIndex, 4)

    '        Select Case UserName = Program_IdUzytkownika
    '            Case True
    '                e.CellStyle.ForeColor = Color.Red
    '            Case Else
    '                Select Case Mid(Upr, _Uprawnnienia_NrZnaku, 1)
    '                    Case "A"
    '                        e.CellStyle.ForeColor = Color.Green
    '                    Case "S"
    '                        e.CellStyle.ForeColor = Color.Purple
    '                    Case "U"
    '                        e.CellStyle.ForeColor = Color.Navy
    '                    Case Else
    '                        e.CellStyle.ForeColor = Color.Black
    '                End Select
    '        End Select
    '        '-----------------------
    '        ''zmieniamy BackColor
    '        'Dim Wal As String = GetTxtField(dagKursyWalut, e.RowIndex, 0)
    '        'Select Case Wal
    '        '    Case "EUR"
    '        '        e.CellStyle.BackColor = Color.White
    '        '    Case Else
    '        'End Select
    '        '-----------------------
    '    End If
    'End Sub




    '*********************************************************
    '***** obsluga dag
    '*********************************************************

    Private Sub dagDaneDoRaportu_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagDaneDoRaportu.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagDaneDoRaportu, DataViewDaneDoRaportu.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagDaneDoRaportu_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagDaneDoRaportu.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagDaneDoRaportu, DataViewDaneDoRaportu.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagDaneDoRaportu_Scroll(sender As Object, e As ScrollEventArgs) Handles dagDaneDoRaportu.Scroll
        'If (Not StartFormy) And (DataViewDaneDoRaportu.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me, dagDaneDoRaportu, DataViewDaneDoRaportu.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me, dagDaneDoRaportu, DataViewDaneDoRaportu.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If

        If (Not StartFormy) And (DataViewDaneDoRaportu.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagDaneDoRaportu_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagDaneDoRaportu.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagDaneDoRaportu, DataViewDaneDoRaportu.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagDaneDoRaportu_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagDaneDoRaportu.MouseMove
        'przypisz biezacy DataGrid do myGrid
        Dim myGridView As DataGridView = CType(sender, DataGridView)
        Dim hti As System.Windows.Forms.DataGridView.HitTestInfo
        hti = myGridView.HitTest(e.X, e.Y)

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None     '"the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                'If Not StartFormy Then
                '    PokazFiltryColumnDGV(Me, dagDaneDoRaportu, DataViewDaneDoRaportu.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me, dagDaneDoRaportu, DataViewDaneDoRaportu.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagDaneDoRaportu_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagDaneDoRaportu.MouseUp
        'przypisz biezacy DataGrid do myGrid
        Dim myGridView As DataGridView = CType(sender, DataGridView)
        Dim hti As System.Windows.Forms.DataGridView.HitTestInfo
        hti = myGridView.HitTest(e.X, e.Y)

        Dim message As String = "You clicked "

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None
                message &= "the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                message &= "cell at row " & hti.RowIndex & ", col " & hti.ColumnIndex
                stbDaneDoRaportu.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbDaneDoRaportu.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagDaneDoRaportu(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbDaneDoRaportu.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbDaneDoRaportu.Panels(3).Text = "Sort: " & DataSetDaneDoRaportu.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbDaneDoRaportu.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagDaneDoRaportu(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbDaneDoRaportu.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbDaneDoRaportu.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub


    Private Sub dagDaneDoRaportu_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagDaneDoRaportu.DoubleClick
        If _OCoMamRobic = "WYBIERAJ" Then
            cmdStop_Click(sender, e)
        Else
            cmdEdytuj_Click(sender, e)
        End If
    End Sub

    Private Sub dagDaneDoRaportu_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagDaneDoRaportu.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                If _OCoMamRobic = "WYBIERAJ" Then
                    cmdStop_Click(sender, e)
                Else
                    cmdEdytuj_Click(sender, e)
                End If
            Case Keys.Insert
                cmdDodaj_Click(sender, e)
            Case Keys.Delete
                cmdUsuñ_Click(sender, e)
            Case Else
        End Select
    End Sub











    '*************************************************
    '*** przenoszenie danych miêdzy rekordem i zmiennymi
    '*************************************************
    'Public ddrPracownik As String        'PRACOWNIK   Text, 10
    'Public ddrDataRaportu As String      'DATARAPORTU Text, 7   RRRR-MM
    'Public ddrKlBezDrukarki As Integer   'KLBEZDRUKARKI  Integer
    'Public ddrOferty As Integer          'OFERTY         Integer
    'Public ddrTestyWstawionw As Integer  'TESTYWSTAWIONE Integer
    'Public ddrCzyszczenie As Integer     'CZYSZCZENIE    Integer
    'Public ddrDostawy As Integer         'DOSTAWY        Integer
    'Public ddrSkup As Integer            'SKUP           Integer

    'Public ddrKolExcel01 As String            'KolExCEL01           String
    'Public ddrKolExcel02 As String            'KolExCEL02           String
    'Public ddrKolExcel03 As String            'KolExCEL03           String
    'Public ddrKolExcel04 As String            'KolExCEL04           String
    'Public ddrKolExcel05 As String            'KolExCEL05           String
    'Public ddrKolExcel06 As String            'KolExCEL06           String
    'Public ddrKolExcel07 As String            'KolExCEL07           String
    'Public ddrKolExcel08 As String            'KolExCEL08           String
    'Public ddrKolExcel09 As String            'KolExCEL09           String
    'Public ddrKolExcel10 As String            'KolExCEL10           String

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

    'Public ddrWersja As Integer          'WERSJA         Integer

    Private Sub InitDaneDoRaportu()
        ddrPracownik = ""
        ddrDataRaportu = SysData()
        ddrKlBezDrukarki = 0
        ddrOferty = 0
        ddrTestyWstawionw = 0
        ddrCzyszczenie = 0
        ddrDostawy = 0
        ddrSkup = 0

        ddrKolExcel01 = ""
        ddrKolExcel02 = ""
        ddrKolExcel03 = ""
        ddrKolExcel04 = ""
        ddrKolExcel05 = ""
        ddrKolExcel06 = ""
        ddrKolExcel07 = ""
        ddrKolExcel08 = ""
        ddrKolExcel09 = ""
        ddrKolExcel10 = ""

        ddrExcel01 = 0
        ddrExcel02 = 0
        ddrExcel03 = 0
        ddrExcel04 = 0
        ddrExcel05 = 0
        ddrExcel06 = 0
        ddrExcel07 = 0
        ddrExcel08 = 0
        ddrExcel09 = 0
        ddrExcel10 = 0

        ddrWersja = 1
    End Sub

    Private Sub PobierzDaneDoRaportu()
        'pobierz wartosci przed aktualizacja
        ddrPracownik = GetTxtField(dagDaneDoRaportu, 0)
        ddrDataRaportu = GetTxtField(dagDaneDoRaportu, 1)
        ddrKlBezDrukarki = GetIntField(dagDaneDoRaportu, 2)
        ddrOferty = GetIntField(dagDaneDoRaportu, 3)
        ddrTestyWstawionw = GetIntField(dagDaneDoRaportu, 4)
        ddrCzyszczenie = GetIntField(dagDaneDoRaportu, 5)
        ddrDostawy = GetIntField(dagDaneDoRaportu, 6)
        ddrSkup = GetIntField(dagDaneDoRaportu, 7)

        ddrKolExcel01 = GetTxtField(dagDaneDoRaportu, 8)
        ddrKolExcel02 = GetTxtField(dagDaneDoRaportu, 9)
        ddrKolExcel03 = GetTxtField(dagDaneDoRaportu, 10)
        ddrKolExcel04 = GetTxtField(dagDaneDoRaportu, 11)
        ddrKolExcel05 = GetTxtField(dagDaneDoRaportu, 12)
        ddrKolExcel06 = GetTxtField(dagDaneDoRaportu, 13)
        ddrKolExcel07 = GetTxtField(dagDaneDoRaportu, 14)
        ddrKolExcel08 = GetTxtField(dagDaneDoRaportu, 15)
        ddrKolExcel09 = GetTxtField(dagDaneDoRaportu, 16)
        ddrKolExcel10 = GetTxtField(dagDaneDoRaportu, 17)

        ddrExcel01 = GetIntField(dagDaneDoRaportu, 18)
        ddrExcel02 = GetIntField(dagDaneDoRaportu, 19)
        ddrExcel03 = GetIntField(dagDaneDoRaportu, 20)
        ddrExcel04 = GetIntField(dagDaneDoRaportu, 21)
        ddrExcel05 = GetIntField(dagDaneDoRaportu, 22)
        ddrExcel06 = GetIntField(dagDaneDoRaportu, 23)
        ddrExcel07 = GetIntField(dagDaneDoRaportu, 24)
        ddrExcel08 = GetIntField(dagDaneDoRaportu, 25)
        ddrExcel09 = GetIntField(dagDaneDoRaportu, 26)
        ddrExcel10 = GetIntField(dagDaneDoRaportu, 27)

        ddrWersja = GetIntField(dagDaneDoRaportu, 28)
    End Sub

    '**************************************************************
    '* obsluga komend
    '**************************************************************

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        If DataSetDaneDoRaportu.Tables(0).Rows.Count > 0 Then
        Else
        End If
        oAktualizuj = True
        Me.Close()
    End Sub

    Private Sub cmdWycofajsie_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWycofajsie.Click
        oAktualizuj = False
        Me.Close()
    End Sub

    Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytuj.Click
        If DataSetDaneDoRaportu.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            oOperacja = "EDYTUJ"
            PobierzDaneDoRaportu()
            Dim Form As New EdytujKatalogDaneDoRaportu
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
            Else
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = ddrPracownik
                findTheseVals(1) = ddrDataRaportu
                foundRow = DataSetDaneDoRaportu.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    'foundRow("PRACOWNIK") = ddrPracownik
                    'foundRow("DATARAPORTU") = ddrDataRaportu
                    foundRow("KLBEZDRUKARKI") = ddrKlBezDrukarki
                    foundRow("OFERTY") = ddrOferty
                    foundRow("TESTYWSTAWIONE") = ddrTestyWstawionw
                    foundRow("CZYSZCZENIE") = ddrCzyszczenie
                    foundRow("DOSTAWY") = ddrDostawy
                    foundRow("SKUP") = ddrSkup

                    foundRow("KOLEXCEL01") = ddrKolExcel01
                    foundRow("KOLEXCEL02") = ddrKolExcel02
                    foundRow("KOLEXCEL03") = ddrKolExcel03
                    foundRow("KOLEXCEL04") = ddrKolExcel04
                    foundRow("KOLEXCEL05") = ddrKolExcel05
                    foundRow("KOLEXCEL06") = ddrKolExcel06
                    foundRow("KOLEXCEL07") = ddrKolExcel07
                    foundRow("KOLEXCEL08") = ddrKolExcel08
                    foundRow("KOLEXCEL09") = ddrKolExcel09
                    foundRow("KOLEXCEL10") = ddrKolExcel10

                    foundRow("EXCEL01") = ddrExcel01
                    foundRow("EXCEL02") = ddrExcel02
                    foundRow("EXCEL03") = ddrExcel03
                    foundRow("EXCEL04") = ddrExcel04
                    foundRow("EXCEL05") = ddrExcel05
                    foundRow("EXCEL06") = ddrExcel06
                    foundRow("EXCEL07") = ddrExcel07
                    foundRow("EXCEL08") = ddrExcel08
                    foundRow("EXCEL09") = ddrExcel09
                    foundRow("EXCEL10") = ddrExcel10

                    foundRow("WERSJA") = ZmienWersje(ddrWersja)    'Wersja         Integer, 2

                    'wyswietl ilosc rekordow
                    stbDaneDoRaportu.Panels(0).Text = "Iloœæ rec.: " & DataSetDaneDoRaportu.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    dagDaneDoRaportu.Update()
                End If
            End If
        End If
    End Sub


    Private Sub cmdUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñ.Click
        If DataSetDaneDoRaportu.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie usun¹æ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            If MessageBox.Show("Czy na pewno mam usun¹æ ten zapis ?",
                        "Proszê o potwierdzenie :",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                oOperacja = "USUN"
                PobierzDaneDoRaportu()
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = ddrPracownik
                findTheseVals(1) = ddrDataRaportu
                foundRow = DataSetDaneDoRaportu.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    foundRow.Delete()
                    'wyswietl ilosc rekordow
                    stbDaneDoRaportu.Panels(0).Text = "Iloœæ rec.: " & DataSetDaneDoRaportu.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                Else
                End If
            End If
        End If
    End Sub

    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        oOperacja = "DODAJ"
        InitDaneDoRaportu()
        Dim Form As New EdytujKatalogDaneDoRaportu
        Do While True
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
                Exit Do
            Else
                Dim foundRow As DataRow
                Dim NewRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = ddrPracownik
                findTheseVals(1) = ddrDataRaportu
                foundRow = DataSetDaneDoRaportu.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                        "Pracownik = " & ddrPracownik &
                        "Data Raportu = " & ddrDataRaportu,
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    'stworz nowy rekord
                    NewRow = DataSetDaneDoRaportu.Tables(0).NewRow()

                    NewRow("PRACOWNIK") = ddrPracownik
                    NewRow("DATARAPORTU") = ddrDataRaportu
                    NewRow("KLBEZDRUKARKI") = ddrKlBezDrukarki
                    NewRow("OFERTY") = ddrOferty
                    NewRow("TESTYWSTAWIONE") = ddrTestyWstawionw
                    NewRow("CZYSZCZENIE") = ddrCzyszczenie
                    NewRow("DOSTAWY") = ddrDostawy
                    NewRow("SKUP") = ddrSkup

                    NewRow("KOLEXCEL01") = ddrKolExcel01
                    NewRow("KOLEXCEL02") = ddrKolExcel02
                    NewRow("KOLEXCEL03") = ddrKolExcel03
                    NewRow("KOLEXCEL04") = ddrKolExcel04
                    NewRow("KOLEXCEL05") = ddrKolExcel05
                    NewRow("KOLEXCEL06") = ddrKolExcel06
                    NewRow("KOLEXCEL07") = ddrKolExcel07
                    NewRow("KOLEXCEL08") = ddrKolExcel08
                    NewRow("KOLEXCEL09") = ddrKolExcel09
                    NewRow("KOLEXCEL10") = ddrKolExcel10

                    NewRow("EXCEL01") = ddrExcel01
                    NewRow("EXCEL02") = ddrExcel02
                    NewRow("EXCEL03") = ddrExcel03
                    NewRow("EXCEL04") = ddrExcel04
                    NewRow("EXCEL05") = ddrExcel05
                    NewRow("EXCEL06") = ddrExcel06
                    NewRow("EXCEL07") = ddrExcel07
                    NewRow("EXCEL08") = ddrExcel08
                    NewRow("EXCEL09") = ddrExcel09
                    NewRow("EXCEL10") = ddrExcel10

                    NewRow("WERSJA") = 1                     'Wersja         Integer, 2
                    'dodaj rekord do DataSet
                    DataSetDaneDoRaportu.Tables(0).Rows.Add(NewRow)

                    'wyswietl ilosc rekordow
                    stbDaneDoRaportu.Panels(0).Text = "Iloœæ rec.: " & DataSetDaneDoRaportu.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    dagDaneDoRaportu.Update()
                    Exit Do
                End If
            End If
        Loop
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
                    dbDataAdapter.Update(DataSetDaneDoRaportu, "TABELA_DaneDoRaportu")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    dbDataAdapter.Fill(DataSetDaneDoRaportu)
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
                    sqlDataAdapter.Update(DataSetDaneDoRaportu, "TABELA_DaneDoRaportu")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapter.Fill(DataSetDaneDoRaportu)
                End If
                sqlConnection.Close()
            End If
        End If
    End Sub


    Private Sub OdswiezBaze()
        DataSetDaneDoRaportu.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            Try
                dbConnection.Open()
                If dbConnection.State = ConnectionState.Open Then
                    dbDataAdapter.Fill(DataSetDaneDoRaportu)
                    dbConnection.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        Else
            Try
                sqlConnection.Open()
                If sqlConnection.State = ConnectionState.Open Then
                    sqlDataAdapter.Fill(DataSetDaneDoRaportu)
                    sqlConnection.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub




    Private Sub cmdHarmonogramWizyt_Click(sender As Object, e As EventArgs) Handles cmdHarmonogramWizyt.Click
        Dim Form As New KatalogHarmonogramWizyt
        Form.Show()
    End Sub


    '****************************************
    '** obsluga Menu Kontekstowego
    '****************************************

    Private Sub menuWEdytuj_Click(sender As Object, e As EventArgs) Handles menuWEdytuj.Click
        cmdEdytuj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuWDodaj_Click(sender As Object, e As EventArgs) Handles menuWDodaj.Click
        cmdDodaj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuWUsun_Click(sender As Object, e As EventArgs) Handles menuWUsun.Click
        cmdUsuñ_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuWWybierz_Click(sender As Object, e As EventArgs) Handles menuWWybierz.Click
        cmdStop_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuWSingleL_Click(sender As Object, e As EventArgs) Handles menuWSingleL.Click
        dagDaneDoRaportu.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagDaneDoRaportu.Refresh()
    End Sub

    Private Sub menuWMultiL_Click(sender As Object, e As EventArgs) Handles menuWMultiL.Click
        dagDaneDoRaportu.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagDaneDoRaportu.Refresh()
    End Sub

    Private Sub menuWOdswiez_Click(sender As Object, e As EventArgs) Handles menuWOdswiez.Click
        OdswiezBaze()
    End Sub




    Private Sub menuEEdytuj_Click(sender As Object, e As EventArgs) Handles menuEEdytuj.Click
        cmdEdytuj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuEDodaj_Click(sender As Object, e As EventArgs) Handles menuEDodaj.Click
        cmdDodaj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuEUsun_Click(sender As Object, e As EventArgs) Handles menuEUsun.Click
        cmdUsuñ_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuESingleL_Click(sender As Object, e As EventArgs) Handles menuESingleL.Click
        dagDaneDoRaportu.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagDaneDoRaportu.Refresh()
    End Sub

    Private Sub menuEMultiL_Click(sender As Object, e As EventArgs) Handles menuEMultiL.Click
        dagDaneDoRaportu.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagDaneDoRaportu.Refresh()
    End Sub

    Private Sub menuEOdswiez_Click(sender As Object, e As EventArgs) Handles menuEOdswiez.Click
        OdswiezBaze()
    End Sub




End Class
