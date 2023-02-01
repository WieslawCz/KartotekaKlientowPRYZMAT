Public Class EdytujKatalogAkcjiOpis
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

#Region " Windows Form Designer generated code "


    Public Sub New()
        MyBase.New()        'This call is required by the Windows Form Designer.
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
    Friend WithEvents cmdWycofajSie As System.Windows.Forms.Button
    Friend WithEvents cmdPrzywroc As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents txtIdent As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtOpis As System.Windows.Forms.TextBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents Akcja As System.Windows.Forms.TabPage
    Friend WithEvents txtUwagi As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpDataAkcji As DateTimePicker
    Friend WithEvents HorizScrollTimer As Timer
    Friend WithEvents menuEdytuj As ContextMenuStrip
    Friend WithEvents MenuEEdytuj As ToolStripMenuItem
    Friend WithEvents MenuEDodaj As ToolStripMenuItem
    Friend WithEvents MenuEUsun As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents menuESingleL As ToolStripMenuItem
    Friend WithEvents menuEMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents menuEOdswiez As ToolStripMenuItem
    Friend WithEvents Specyfikacja As TabPage
    Friend WithEvents txtDA As TextBox
    Friend WithEvents cmdEdytuj As Button
    Friend WithEvents cmdFi As Button
    Friend WithEvents cmdWszystko As Button
    Friend WithEvents stbFiltry As StatusBar
    Friend WithEvents dagAkcjeSpec As DataGridView
    Friend WithEvents stbAkcjeSpec As StatusBar
    Friend WithEvents Panel1 As Panel
    Friend WithEvents cmdUsuñ As Button
    Friend WithEvents cmdDodaj As Button
    Friend WithEvents txtUW As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txtID As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents txtOP As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EdytujKatalogAkcjiOpis))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmdWycofajSie = New System.Windows.Forms.Button()
        Me.cmdPrzywroc = New System.Windows.Forms.Button()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.txtOpis = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtIdent = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.Akcja = New System.Windows.Forms.TabPage()
        Me.dtpDataAkcji = New System.Windows.Forms.DateTimePicker()
        Me.txtUwagi = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Specyfikacja = New System.Windows.Forms.TabPage()
        Me.txtDA = New System.Windows.Forms.TextBox()
        Me.cmdEdytuj = New System.Windows.Forms.Button()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.cmdWszystko = New System.Windows.Forms.Button()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
        Me.dagAkcjeSpec = New System.Windows.Forms.DataGridView()
        Me.stbAkcjeSpec = New System.Windows.Forms.StatusBar()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdUsuñ = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.txtUW = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtOP = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.HorizScrollTimer = New System.Windows.Forms.Timer(Me.components)
        Me.menuEdytuj = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.MenuEEdytuj = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuEDodaj = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuEUsun = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuESingleL = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEMultiL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuEOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.Akcja.SuspendLayout()
        Me.Specyfikacja.SuspendLayout()
        CType(Me.dagAkcjeSpec, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuEdytuj.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdWycofajSie
        '
        Me.cmdWycofajSie.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWycofajSie.Image = CType(resources.GetObject("cmdWycofajSie.Image"), System.Drawing.Image)
        Me.cmdWycofajSie.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWycofajSie.Location = New System.Drawing.Point(672, 634)
        Me.cmdWycofajSie.Name = "cmdWycofajSie"
        Me.cmdWycofajSie.Size = New System.Drawing.Size(104, 22)
        Me.cmdWycofajSie.TabIndex = 14
        Me.cmdWycofajSie.Text = "Wycofaj siê"
        Me.cmdWycofajSie.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPrzywroc
        '
        Me.cmdPrzywroc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPrzywroc.Image = CType(resources.GetObject("cmdPrzywroc.Image"), System.Drawing.Image)
        Me.cmdPrzywroc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPrzywroc.Location = New System.Drawing.Point(580, 634)
        Me.cmdPrzywroc.Name = "cmdPrzywroc"
        Me.cmdPrzywroc.Size = New System.Drawing.Size(86, 22)
        Me.cmdPrzywroc.TabIndex = 15
        Me.cmdPrzywroc.Text = "Przywróæ"
        Me.cmdPrzywroc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(782, 634)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 22)
        Me.cmdPowrot.TabIndex = 13
        Me.cmdPowrot.Text = "Zapisz"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 634)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 16
        Me.PictureBox2.TabStop = False
        '
        'txtOpis
        '
        Me.txtOpis.ForeColor = System.Drawing.Color.Purple
        Me.txtOpis.Location = New System.Drawing.Point(80, 31)
        Me.txtOpis.Name = "txtOpis"
        Me.txtOpis.Size = New System.Drawing.Size(752, 20)
        Me.txtOpis.TabIndex = 9
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(16, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 16)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Opis akcji . . . . . . . . ."
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(298, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 16)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Data akcji . . . . . . . . ."
        '
        'txtIdent
        '
        Me.txtIdent.ForeColor = System.Drawing.Color.Purple
        Me.txtIdent.Location = New System.Drawing.Point(80, 8)
        Me.txtIdent.Name = "txtIdent"
        Me.txtIdent.Size = New System.Drawing.Size(144, 20)
        Me.txtIdent.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(16, 11)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Ident. akcji . . . . . . . . ."
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.Akcja)
        Me.TabControl1.Controls.Add(Me.Specyfikacja)
        Me.TabControl1.Location = New System.Drawing.Point(8, 8)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(854, 621)
        Me.TabControl1.TabIndex = 18
        '
        'Akcja
        '
        Me.Akcja.Controls.Add(Me.dtpDataAkcji)
        Me.Akcja.Controls.Add(Me.txtUwagi)
        Me.Akcja.Controls.Add(Me.Label3)
        Me.Akcja.Controls.Add(Me.txtIdent)
        Me.Akcja.Controls.Add(Me.Label4)
        Me.Akcja.Controls.Add(Me.txtOpis)
        Me.Akcja.Controls.Add(Me.Label1)
        Me.Akcja.Controls.Add(Me.Label2)
        Me.Akcja.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Akcja.ForeColor = System.Drawing.Color.Navy
        Me.Akcja.Location = New System.Drawing.Point(4, 22)
        Me.Akcja.Name = "Akcja"
        Me.Akcja.Size = New System.Drawing.Size(846, 595)
        Me.Akcja.TabIndex = 0
        Me.Akcja.Text = "Akcja"
        '
        'dtpDataAkcji
        '
        Me.dtpDataAkcji.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dtpDataAkcji.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpDataAkcji.CustomFormat = "yyyy-MM-dd"
        Me.dtpDataAkcji.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDataAkcji.Location = New System.Drawing.Point(362, 8)
        Me.dtpDataAkcji.Name = "dtpDataAkcji"
        Me.dtpDataAkcji.Size = New System.Drawing.Size(100, 20)
        Me.dtpDataAkcji.TabIndex = 185
        Me.dtpDataAkcji.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'txtUwagi
        '
        Me.txtUwagi.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUwagi.ForeColor = System.Drawing.Color.Purple
        Me.txtUwagi.Location = New System.Drawing.Point(80, 55)
        Me.txtUwagi.Multiline = True
        Me.txtUwagi.Name = "txtUwagi"
        Me.txtUwagi.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtUwagi.Size = New System.Drawing.Size(752, 529)
        Me.txtUwagi.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(16, 58)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 16)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Uwagi . . . . . . . . ."
        '
        'Specyfikacja
        '
        Me.Specyfikacja.BackColor = System.Drawing.SystemColors.Control
        Me.Specyfikacja.Controls.Add(Me.txtDA)
        Me.Specyfikacja.Controls.Add(Me.cmdEdytuj)
        Me.Specyfikacja.Controls.Add(Me.cmdFi)
        Me.Specyfikacja.Controls.Add(Me.cmdWszystko)
        Me.Specyfikacja.Controls.Add(Me.stbFiltry)
        Me.Specyfikacja.Controls.Add(Me.dagAkcjeSpec)
        Me.Specyfikacja.Controls.Add(Me.stbAkcjeSpec)
        Me.Specyfikacja.Controls.Add(Me.Panel1)
        Me.Specyfikacja.Controls.Add(Me.cmdUsuñ)
        Me.Specyfikacja.Controls.Add(Me.cmdDodaj)
        Me.Specyfikacja.Controls.Add(Me.txtUW)
        Me.Specyfikacja.Controls.Add(Me.Label5)
        Me.Specyfikacja.Controls.Add(Me.txtID)
        Me.Specyfikacja.Controls.Add(Me.Label6)
        Me.Specyfikacja.Controls.Add(Me.txtOP)
        Me.Specyfikacja.Controls.Add(Me.Label7)
        Me.Specyfikacja.Controls.Add(Me.Label8)
        Me.Specyfikacja.Location = New System.Drawing.Point(4, 22)
        Me.Specyfikacja.Name = "Specyfikacja"
        Me.Specyfikacja.Size = New System.Drawing.Size(846, 595)
        Me.Specyfikacja.TabIndex = 1
        Me.Specyfikacja.Text = "Specyfikacja"
        '
        'txtDA
        '
        Me.txtDA.ForeColor = System.Drawing.Color.Purple
        Me.txtDA.Location = New System.Drawing.Point(80, 32)
        Me.txtDA.Name = "txtDA"
        Me.txtDA.ReadOnly = True
        Me.txtDA.Size = New System.Drawing.Size(144, 20)
        Me.txtDA.TabIndex = 204
        '
        'cmdEdytuj
        '
        Me.cmdEdytuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdytuj.Enabled = False
        Me.cmdEdytuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdEdytuj.Image = CType(resources.GetObject("cmdEdytuj.Image"), System.Drawing.Image)
        Me.cmdEdytuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEdytuj.Location = New System.Drawing.Point(754, 564)
        Me.cmdEdytuj.Name = "cmdEdytuj"
        Me.cmdEdytuj.Size = New System.Drawing.Size(80, 23)
        Me.cmdEdytuj.TabIndex = 203
        Me.cmdEdytuj.Text = "Edytuj"
        Me.cmdEdytuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdFi
        '
        Me.cmdFi.Image = CType(resources.GetObject("cmdFi.Image"), System.Drawing.Image)
        Me.cmdFi.Location = New System.Drawing.Point(598, 96)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(16, 20)
        Me.cmdFi.TabIndex = 201
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'cmdWszystko
        '
        Me.cmdWszystko.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystko.Location = New System.Drawing.Point(619, 94)
        Me.cmdWszystko.Name = "cmdWszystko"
        Me.cmdWszystko.Size = New System.Drawing.Size(26, 22)
        Me.cmdWszystko.TabIndex = 199
        '
        'stbFiltry
        '
        Me.stbFiltry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltry.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltry.Location = New System.Drawing.Point(12, 94)
        Me.stbFiltry.Name = "stbFiltry"
        Me.stbFiltry.ShowPanels = True
        Me.stbFiltry.Size = New System.Drawing.Size(822, 22)
        Me.stbFiltry.TabIndex = 200
        Me.stbFiltry.Text = "StatusBar1"
        '
        'dagAkcjeSpec
        '
        Me.dagAkcjeSpec.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagAkcjeSpec.Location = New System.Drawing.Point(12, 120)
        Me.dagAkcjeSpec.Name = "dagAkcjeSpec"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagAkcjeSpec.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagAkcjeSpec.Size = New System.Drawing.Size(822, 410)
        Me.dagAkcjeSpec.TabIndex = 198
        '
        'stbAkcjeSpec
        '
        Me.stbAkcjeSpec.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbAkcjeSpec.Dock = System.Windows.Forms.DockStyle.None
        Me.stbAkcjeSpec.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbAkcjeSpec.Location = New System.Drawing.Point(12, 533)
        Me.stbAkcjeSpec.Name = "stbAkcjeSpec"
        Me.stbAkcjeSpec.ShowPanels = True
        Me.stbAkcjeSpec.Size = New System.Drawing.Size(822, 22)
        Me.stbAkcjeSpec.TabIndex = 197
        Me.stbAkcjeSpec.Text = "StatusBar1"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Location = New System.Drawing.Point(1, 84)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(845, 1)
        Me.Panel1.TabIndex = 196
        '
        'cmdUsuñ
        '
        Me.cmdUsuñ.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUsuñ.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdUsuñ.Image = CType(resources.GetObject("cmdUsuñ.Image"), System.Drawing.Image)
        Me.cmdUsuñ.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUsuñ.Location = New System.Drawing.Point(668, 564)
        Me.cmdUsuñ.Name = "cmdUsuñ"
        Me.cmdUsuñ.Size = New System.Drawing.Size(80, 23)
        Me.cmdUsuñ.TabIndex = 195
        Me.cmdUsuñ.Text = "Usuñ"
        Me.cmdUsuñ.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdDodaj
        '
        Me.cmdDodaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDodaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDodaj.Image = CType(resources.GetObject("cmdDodaj.Image"), System.Drawing.Image)
        Me.cmdDodaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDodaj.Location = New System.Drawing.Point(582, 564)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(80, 23)
        Me.cmdDodaj.TabIndex = 194
        Me.cmdDodaj.Text = "Dodaj"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtUW
        '
        Me.txtUW.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUW.ForeColor = System.Drawing.Color.Purple
        Me.txtUW.Location = New System.Drawing.Point(378, 8)
        Me.txtUW.Multiline = True
        Me.txtUW.Name = "txtUW"
        Me.txtUW.ReadOnly = True
        Me.txtUW.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtUW.Size = New System.Drawing.Size(460, 68)
        Me.txtUW.TabIndex = 192
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(318, 11)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 16)
        Me.Label5.TabIndex = 193
        Me.Label5.Text = "Uwagi . . . . . . . . ."
        '
        'txtID
        '
        Me.txtID.ForeColor = System.Drawing.Color.Purple
        Me.txtID.Location = New System.Drawing.Point(80, 8)
        Me.txtID.Name = "txtID"
        Me.txtID.ReadOnly = True
        Me.txtID.Size = New System.Drawing.Size(144, 20)
        Me.txtID.TabIndex = 187
        '
        'Label6
        '
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(16, 11)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(112, 16)
        Me.Label6.TabIndex = 188
        Me.Label6.Text = "Ident. akcji . . . . . . . . ."
        '
        'txtOP
        '
        Me.txtOP.ForeColor = System.Drawing.Color.Purple
        Me.txtOP.Location = New System.Drawing.Point(80, 56)
        Me.txtOP.Name = "txtOP"
        Me.txtOP.ReadOnly = True
        Me.txtOP.Size = New System.Drawing.Size(293, 20)
        Me.txtOP.TabIndex = 190
        '
        'Label7
        '
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(16, 35)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 16)
        Me.Label7.TabIndex = 189
        Me.Label7.Text = "Data akcji . . . . . . . . ."
        '
        'Label8
        '
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(16, 59)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 16)
        Me.Label8.TabIndex = 191
        Me.Label8.Text = "Opis akcji . . . . . . . . ."
        '
        'HorizScrollTimer
        '
        '
        'menuEdytuj
        '
        Me.menuEdytuj.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuEEdytuj, Me.MenuEDodaj, Me.MenuEUsun, Me.ToolStripSeparator4, Me.menuESingleL, Me.menuEMultiL, Me.ToolStripSeparator5, Me.menuEOdswiez})
        Me.menuEdytuj.Name = "menuEdytuj"
        Me.menuEdytuj.Size = New System.Drawing.Size(187, 148)
        '
        'MenuEEdytuj
        '
        Me.MenuEEdytuj.Enabled = False
        Me.MenuEEdytuj.Name = "MenuEEdytuj"
        Me.MenuEEdytuj.Size = New System.Drawing.Size(186, 22)
        Me.MenuEEdytuj.Text = "Edytuj"
        '
        'MenuEDodaj
        '
        Me.MenuEDodaj.Name = "MenuEDodaj"
        Me.MenuEDodaj.Size = New System.Drawing.Size(186, 22)
        Me.MenuEDodaj.Text = "Dodaj"
        '
        'MenuEUsun
        '
        Me.MenuEUsun.Name = "MenuEUsun"
        Me.MenuEUsun.Size = New System.Drawing.Size(186, 22)
        Me.MenuEUsun.Text = "Usuñ"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(183, 6)
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
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(183, 6)
        '
        'menuEOdswiez
        '
        Me.menuEOdswiez.Name = "menuEOdswiez"
        Me.menuEOdswiez.Size = New System.Drawing.Size(186, 22)
        Me.menuEOdswiez.Text = "Odœwie¿ Tabele"
        '
        'EdytujKatalogAkcjiOpis
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(868, 662)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.cmdWycofajSie)
        Me.Controls.Add(Me.cmdPrzywroc)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Location = New System.Drawing.Point(8, 8)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "EdytujKatalogAkcjiOpis"
        Me.ShowInTaskbar = False
        Me.Text = "Akcja marketingowa"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.Akcja.ResumeLayout(False)
        Me.Akcja.PerformLayout()
        Me.Specyfikacja.ResumeLayout(False)
        Me.Specyfikacja.PerformLayout()
        CType(Me.dagAkcjeSpec, System.ComponentModel.ISupportInitialize).EndInit()
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

    Dim dbSelectSQLAkcjeSpec As String = sSelectSQLAkcjeSpec

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

    Dim DataSetAkcjeSpec As DataSet
    Dim DataViewAkcjeSpec As DataView
    Dim ConnectionState As System.Data.ConnectionState


    Private Sub EdytujKatalogAkcjiOpis_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        PobierzDane()
        If oOperacja = "EDYTUJ" Then
            'nie pozwalaj na edycja PrimaryKey
            txtIdent.Enabled = False
        Else
            txtIdent.Enabled = True
            txtIdent.Focus()
        End If

    End Sub

    Private Sub EdytujKatalogAkcjiOpis_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If oOperacja = "EDYTUJ" Then
            dtpDataAkcji.Focus()
        Else
            txtIdent.Focus()
        End If
    End Sub

    Private Sub EdytujKatalogAkcjiOpis_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        ZapiszDane()
        oAktualizuj = True
        Me.Close()
    End Sub


    Private Sub cmdWycofajSie_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWycofajSie.Click
        oAktualizuj = False
        Me.Close()
    End Sub

    Private Sub cmdPrzywroc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrzywroc.Click
        PobierzDane()
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Select Case TabControl1.SelectedIndex
            Case 0
                cmdPowrot.Focus()
            Case 1
                '----------------------------



                dbSelectSQLAkcjeSpec = "SELECT AkcjeSpec.IDENTAKCJI," &
                                              "AkcjeSpec.IDENTKLI," &
                                              "AkcjeSpec.WERSJA," &
                                           "Klienci.NAZWA1 AS NAZWAKLIENTA," &
                                           "Klienci.MIEJSCOWOSC," &
                                           "Klienci.KODPOCZTOWY," &
                                           "Klienci.ULICA," &
                                           "Klienci.NRDOMU," &
                                           "Klienci.REJKLIENTA AS REJONKLIENTA," &
                                           "Klienci.TELEFON," &
                                           "Klienci.FAX," &
                                           "Klienci.EMAIL " &
                                       "FROM AkcjeSpec INNER JOIN Klienci " &
                                       "ON AkcjeSpec.IDENTKLI=Klienci.IDENTKLIENTA " &
                                       "WHERE AkcjeSpec.IDENTAKCJI='" & Trim(txtIdent.Text) & "'" &
                                       "ORDER BY AkcjeSpec.IDENTAKCJI"

                DataSetAkcjeSpec = New DataSet
                DataViewAkcjeSpec = New DataView
                DataSetAkcjeSpec.Locale = New System.Globalization.CultureInfo("pl-PL")

                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbConnection = New OleDb.OleDbConnection(sConnectionAkcjeSpec)
                    'dbCommandSelect = New OleDb.OleDbCommand(dbSelectSQLAkcjeSpec, dbConnection)
                    'dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLAkcjeSpec, dbConnection)
                    'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLAkcjeSpec, dbConnection)
                    'dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLAkcjeSpec, dbConnection)
                    'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

                    'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
                    'MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_AkcjeSpec")

                    'DBDeleteAkcjeSpec(dbCommandDelete, dbDataAdapter)
                    'DBUpdateAkcjeSpec(dbCommandUpdate, dbDataAdapter)
                    'DBInsertAkcjeSpec(dbCommandInsert, dbDataAdapter)

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
                    sqlConnection = New SqlClient.SqlConnection(sConnectionAkcjeSpec)
                    sqlCommandSelect = New SqlClient.SqlCommand(dbSelectSQLAkcjeSpec, sqlConnection)
                    sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLAkcjeSpec, sqlConnection)
                    sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLAkcjeSpec, sqlConnection)
                    sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLAkcjeSpec, sqlConnection)
                    sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

                    Dim MapowanieTabeli As System.Data.Common.DataTableMapping
                    MapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_AkcjeSpec")

                    SQLDeleteAkcjeSpec(sqlCommandDelete, sqlDataAdapter)
                    SQLUpdateAkcjeSpec(sqlCommandUpdate, sqlDataAdapter)
                    SQLInsertAkcjeSpec(sqlCommandInsert, sqlDataAdapter)

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
                            'dbDataAdapter.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_AkcjeSpec")
                            'dbDataAdapter.Fill(DataSetAkcjeSpec)
                            'dbConnection.Close()
                        Else
                            sqlDataAdapter.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_AkcjeSpec")
                            sqlDataAdapter.Fill(DataSetAkcjeSpec)
                            sqlConnection.Close()
                        End If

                        'zdefiniuj unikalny klucz wyszukiwania w tabeli
                        DataSetAkcjeSpec.Tables(0).PrimaryKey = New DataColumn() {DataSetAkcjeSpec.Tables(0).Columns("IDENTAKCJI"), DataSetAkcjeSpec.Tables(0).Columns("IDENTKLI")}

                        'definiuj DataView
                        DataViewAkcjeSpec = New DataView(DataSetAkcjeSpec.Tables(0))
                        DataViewAkcjeSpec.AllowDelete = False
                        DataViewAkcjeSpec.AllowNew = False

                        '-------------------------------------------------
                        'parametry DataGridView
                        '-------------------------------------------------
                        ' Initialize basic DataGridView properties.
                        dagAkcjeSpec.BackgroundColor = System.Drawing.SystemColors.Control
                        dagAkcjeSpec.GridColor = System.Drawing.SystemColors.ControlDark
                        dagAkcjeSpec.ForeColor = System.Drawing.Color.Purple
                        dagAkcjeSpec.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagAkcjeSpec.BorderStyle = BorderStyle.Fixed3D
                        'dagAkcjeSpec.Dock = DockStyle.Fill
                        dagAkcjeSpec.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                        ' Set property values appropriate for read-only display and limited interactivity. 
                        dagAkcjeSpec.AllowUserToAddRows = False
                        dagAkcjeSpec.AllowUserToDeleteRows = False
                        dagAkcjeSpec.AllowUserToOrderColumns = True
                        dagAkcjeSpec.AllowUserToResizeColumns = True
                        dagAkcjeSpec.AllowUserToResizeRows = True
                        dagAkcjeSpec.ReadOnly = True
                        dagAkcjeSpec.SelectionMode = DataGridViewSelectionMode.CellSelect
                        dagAkcjeSpec.MultiSelect = False

                        'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                        'dagAkcjeSpec.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                        dagAkcjeSpec.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                        dagAkcjeSpec.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                        dagAkcjeSpec.ColumnHeadersVisible = True
                        dagAkcjeSpec.ColumnHeadersHeight = 40
                        dagAkcjeSpec.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                        dagAkcjeSpec.RowHeadersVisible = True
                        dagAkcjeSpec.RowHeadersWidth = 50
                        dagAkcjeSpec.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                        ' Set the selection background color for all the cells.
                        dagAkcjeSpec.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                        dagAkcjeSpec.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                        dagAkcjeSpec.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagAkcjeSpec.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                        dagAkcjeSpec.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagAkcjeSpec.DefaultCellStyle.NullValue = ""
                        dagAkcjeSpec.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        dagAkcjeSpec.DefaultCellStyle.WrapMode = DataGridViewTriState.True


                        ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                        ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                        dagAkcjeSpec.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                        dagAkcjeSpec.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                        ' Set the background color for all rows and for alternating rows.  
                        ' The value for alternating rows overrides the value for all rows. 
                        dagAkcjeSpec.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagAkcjeSpec.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                        ' Set the row and column header styles.
                        dagAkcjeSpec.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagAkcjeSpec.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        dagAkcjeSpec.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagAkcjeSpec.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        '-----------------------------------
                        'wypelnienie DataGridView
                        dagAkcjeSpec.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                        dagAkcjeSpec.DataMember = DataSetAkcjeSpec.Tables(0).TableName
                        'dagAkcjeSpec.DataSource = DataViewAkcjeSpec
                        '-----------------------------------

                        'dbSelectSQLAkcjeSpec = "SELECT AkcjeSpec.IDENTAKCJI," &
                        '                      "AkcjeSpec.IDENTKLI," &
                        '                      "AkcjeSpec.WERSJA," &
                        '                   "Klienci.NAZWA1 AS NAZWAKLIENTA," &
                        '                   "Klienci.MIEJSCOWOSC," &
                        '                   "Klienci.KODPOCZTOWY," &
                        '                   "Klienci.ULICA," &
                        '                   "Klienci.NRDOMU," &
                        '                   "Klienci.REJKLIENTA AS REJONKLIENTA," &
                        '                   "Klienci.TELEFON," &
                        '                   "Klienci.FAX," &
                        '                   "Klienci.EMAIL," &
                        '                   "Klienci.WWW as STRONAWWW," &
                        '                   "Klienci.OSOBAKONTAKTOWA " &
                        '               "FROM AkcjeSpec INNER JOIN Klienci " &
                        '               "ON AkcjeSpec.IDENTKLI=Klienci.IDENTKLIENTA " &
                        '               "WHERE AkcjeSpec.IDENTAKCJI='" & Trim(txtIdent.Text) & "'" &
                        '               "ORDER BY AkcjeSpec.IDENTAKCJI"



                        TxtColumnView(dagAkcjeSpec, DataSetAkcjeSpec.Tables(0).Columns(0).ColumnName, "Ident.Akcji", 150, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagAkcjeSpec, DataSetAkcjeSpec.Tables(0).Columns(1).ColumnName, "Klient", 80, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagAkcjeSpec, DataSetAkcjeSpec.Tables(0).Columns(2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)

                        TxtColumnView(dagAkcjeSpec, DataSetAkcjeSpec.Tables(0).Columns(3).ColumnName, "Nazwa klienta", 300, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagAkcjeSpec, DataSetAkcjeSpec.Tables(0).Columns(4).ColumnName, "Miejscowoœæ", 200, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagAkcjeSpec, DataSetAkcjeSpec.Tables(0).Columns(5).ColumnName, "Kod Poczt.", 50, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagAkcjeSpec, DataSetAkcjeSpec.Tables(0).Columns(6).ColumnName, "Ulica", 200, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagAkcjeSpec, DataSetAkcjeSpec.Tables(0).Columns(7).ColumnName, "Nr domu", 50, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagAkcjeSpec, DataSetAkcjeSpec.Tables(0).Columns(8).ColumnName, "Rejon", 200, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagAkcjeSpec, DataSetAkcjeSpec.Tables(0).Columns(9).ColumnName, "Telefon", 200, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagAkcjeSpec, DataSetAkcjeSpec.Tables(0).Columns(10).ColumnName, "Fax", 200, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagAkcjeSpec, DataSetAkcjeSpec.Tables(0).Columns(11).ColumnName, "eMail", 200, DataGridViewContentAlignment.MiddleLeft, True)


                        Me.Cursor = Cursors.WaitCursor
                        dagAkcjeSpec.DataSource = DataViewAkcjeSpec
                        Me.Cursor = Cursors.Default

                        'ustaw sie na pierwszym zapisie
                        If DataSetAkcjeSpec.Tables(0).Rows.Count > 0 Then
                            dagAkcjeSpec.CurrentCell = dagAkcjeSpec(0, 0)
                            dagAkcjeSpec.CurrentCell.Selected = True
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
                    stbAkcjeSpec.BackColor = System.Drawing.SystemColors.Control
                    stbAkcjeSpec.ForeColor = System.Drawing.Color.DarkGreen
                    stbAkcjeSpec.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                    stbAkcjeSpec.Panels.Add("stbPanelIleRec")
                    stbAkcjeSpec.Panels(0).AutoSize = StatusBarPanelAutoSize.None
                    stbAkcjeSpec.Panels(0).Width = 80
                    stbAkcjeSpec.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbAkcjeSpec.Panels(0).Text = "Iloœæ rec.: " & DataSetAkcjeSpec.Tables(0).Rows.Count.ToString

                    stbAkcjeSpec.Panels.Add("stbPanelWiersz")
                    stbAkcjeSpec.Panels(1).AutoSize = StatusBarPanelAutoSize.None
                    stbAkcjeSpec.Panels(1).Width = 80
                    stbAkcjeSpec.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    If dagAkcjeSpec.CurrentCell Is Nothing Then
                        stbAkcjeSpec.Panels(1).Text = "Wi: "
                    Else
                        stbAkcjeSpec.Panels(1).Text = "Wi: " & dagAkcjeSpec.CurrentCell.RowIndex.ToString
                    End If

                    stbAkcjeSpec.Panels.Add("stbPanelKolumna")
                    stbAkcjeSpec.Panels(2).AutoSize = StatusBarPanelAutoSize.None
                    stbAkcjeSpec.Panels(2).Width = 80
                    stbAkcjeSpec.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    If dagAkcjeSpec.CurrentCell Is Nothing Then
                        stbAkcjeSpec.Panels(2).Text = "Ko: "
                    Else
                        stbAkcjeSpec.Panels(2).Text = "Ko: " & dagAkcjeSpec.CurrentCell.ColumnIndex.ToString
                    End If

                    stbAkcjeSpec.Panels.Add("stbPanelSort")
                    stbAkcjeSpec.Panels(3).AutoSize = StatusBarPanelAutoSize.None
                    stbAkcjeSpec.Panels(3).Width = 150
                    stbAkcjeSpec.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbAkcjeSpec.Panels(3).Text = "Sort : "

                    stbAkcjeSpec.Panels.Add("stbPanelFiltr")
                    stbAkcjeSpec.Panels(4).AutoSize = StatusBarPanelAutoSize.None
                    stbAkcjeSpec.Panels(4).Width = 150
                    stbAkcjeSpec.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbAkcjeSpec.Panels(4).Text = "Filtr : "

                    stbAkcjeSpec.Panels.Add("stbPanelReszta")
                    stbAkcjeSpec.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
                    'stbAkcjeSpec.Panels(5).Width = 150
                    stbAkcjeSpec.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbAkcjeSpec.Panels(5).Text = ""

                    stbAkcjeSpec.ShowPanels = True
                    '---------------------------------
                    'dodaj StatusBar na koncu
                    stbFiltry.BackColor = System.Drawing.SystemColors.Control
                    stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
                    stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                    stbFiltry.Panels.Add("stbFiltryRowHead")
                    stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
                    stbFiltry.Panels(0).Width = 50  'dagAkcjeSpec.TableStyles(0).RowHeaderWidth
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

                InitAkcjeSpec() 'inicjuj zmienne
                '-----------------------------------------------------------------
                dagAkcjeSpec.ContextMenuStrip = Me.menuEdytuj
                '--------------------------------------------------------------------
                'set the header width to something apporpriate
                dagAkcjeSpec.RowHeadersWidth = 50       'use if tablestyle
                'Me.dagAkcjeSpec.RowHeaderWidth = 80 'use if no tablestyle
                '--------------------------------------------------------------------
                'set topleft corner point so we can locate the toprow
                If DataSetAkcjeSpec.Tables(0).Rows.Count > 0 Then
                    StartPointInCell00 = New System.Drawing.Point((dagAkcjeSpec.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (dagAkcjeSpec.GetCellDisplayRectangle(0, 0, True).Y + 4))
                    ScrollXOffset = 10000 * dagAkcjeSpec.FirstDisplayedScrollingColumnIndex +
                        dagAkcjeSpec.GetCellDisplayRectangle(dagAkcjeSpec.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
                Else
                    'brak zapisow - poczatek DataGrid
                    StartPointInCell00 = New System.Drawing.Point((dagAkcjeSpec.Left + 4), (dagAkcjeSpec.Top + 4))
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
                For ii = 0 To DataViewAkcjeSpec.Table.Columns.Count - 1
                    'stworz zmienn¹ TXTBOX
                    Dim CtrlT As New System.Windows.Forms.TextBox
                    CtrlT.Name = "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
                    CtrlT.Visible = False
                    Me.Specyfikacja.Controls.Add(CtrlT)
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
                    Me.Specyfikacja.Controls.Add(CtrlB)
                    AddHandler CtrlB.Click, New EventHandler(AddressOf cmdFiXX_Click)
                    AddHandler CtrlB.GotFocus, New EventHandler(AddressOf cmdFiXX_GotFocus)
                    AddHandler CtrlB.LostFocus, New EventHandler(AddressOf cmdFiXX_LostFocus)
                Next
                Me.Refresh()
                '--------------------------------------------------
                StartFormy = False
                dagAkcjeSpec_CurrentCellChanged(sender, e)
                InicjujFiltryColumnDGV(Me.Specyfikacja, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltry, StartFormy)
                '----------------------------
                cmdPowrot.Focus()

        End Select
    End Sub





    Private Sub dagAkcjeSpec_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAkcjeSpec.CurrentCellChanged
        If Not StartFormy Then
            If dagAkcjeSpec.CurrentCell Is Nothing Then
                stbAkcjeSpec.Panels(1).Text = "Wi: "
                stbAkcjeSpec.Panels(2).Text = "Ko: "
            Else
                stbAkcjeSpec.Panels(1).Text = "Wi: " & dagAkcjeSpec.CurrentCell.RowIndex.ToString
                stbAkcjeSpec.Panels(2).Text = "Ko: " & dagAkcjeSpec.CurrentCell.ColumnIndex.ToString
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
                If dagAkcjeSpec.RowCount > 0 Then
                    If ScrollXOffset <> 10000 * dagAkcjeSpec.FirstDisplayedScrollingColumnIndex +
                                    dagAkcjeSpec.GetCellDisplayRectangle(dagAkcjeSpec.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X Then

                        PokazFiltryColumnDGV(Me.Specyfikacja, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltry, StartFormy)
                        '-----------------------------
                        ScrollXOffset = 10000 * dagAkcjeSpec.FirstDisplayedScrollingColumnIndex +
                                    dagAkcjeSpec.GetCellDisplayRectangle(dagAkcjeSpec.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X
                        HorizScrollLastTime = ""
                    End If
                Else
                    PokazFiltryColumnDGV(Me.Specyfikacja, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Else
            End If
        End If
        HorizScrollTimer.Enabled = True
    End Sub


    Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Dim rc As DataGridViewSelectedRowCollection = dagAkcjeSpec.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                dagAkcjeSpec.Rows(nextrow).Selected = True
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
                    PrzejdzDoDGV(dagAkcjeSpec, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            FiltrujDataViewDGV(Me.Specyfikacja, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, DataViewAkcjeSpec, stbAkcjeSpec)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me.Specyfikacja, dagAkcjeSpec, Mid(sender.name, 6, 3), "AkcjeSpec")
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
            CzyscFiltryDGV(Me.Specyfikacja, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, DataViewAkcjeSpec, stbAkcjeSpec)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbAkcjeSpec.Panels(0).Text = "Il.zap.: " & DataViewAkcjeSpec.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagAkcjeSpec_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagAkcjeSpec.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagAkcjeSpec.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagAkcjeSpec.DefaultCellStyle.Font,
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
            PokazFiltryColumnDGV(Me.Specyfikacja, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagAkcjeSpec_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagAkcjeSpec.CellFormatting
    '    If Not StartFormy Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagAkcjeSpec, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagAkcjeSpec, e.RowIndex, 4)

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

    Private Sub dagAkcjeSpec_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagAkcjeSpec.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me.Specyfikacja, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagAkcjeSpec_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAkcjeSpec.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me.Specyfikacja, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagAkcjeSpec_Scroll(sender As Object, e As ScrollEventArgs) Handles dagAkcjeSpec.Scroll
        'If (Not StartFormy) And (DataViewAkcjeSpec.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(me.specyfikacja, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(me.specyfikacja, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If

        If (Not StartFormy) And (DataViewAkcjeSpec.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagAkcjeSpec_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAkcjeSpec.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me.Specyfikacja, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagAkcjeSpec_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagAkcjeSpec.MouseMove
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
                '    PokazFiltryColumnDGV(me.specyfikacja, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me.Specyfikacja, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagAkcjeSpec_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagAkcjeSpec.MouseUp
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
                stbAkcjeSpec.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbAkcjeSpec.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagAkcjeSpec(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbAkcjeSpec.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbAkcjeSpec.Panels(3).Text = "Sort: " & DataSetAkcjeSpec.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbAkcjeSpec.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagAkcjeSpec(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbAkcjeSpec.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbAkcjeSpec.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub




    Private Sub dagAkcjeSpec_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAkcjeSpec.DoubleClick
        'If _CoMamRobic = "WYBIERAJ" Then
        '    cmdStop_Click(sender, e)
        'Else
        '    cmdEdytuj_Click(sender, e)
        'End If
    End Sub

    Private Sub dagAkcjeSpec_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagAkcjeSpec.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                'If _CoMamRobic = "WYBIERAJ" Then
                '    cmdStop_Click(sender, e)
                'Else
                '    cmdEdytuj_Click(sender, e)
                'End If
            Case Keys.Insert
                cmdDodaj_Click(sender, e)
            Case Keys.Delete
                cmdUsuñ_Click(sender, e)
            Case Else
        End Select
    End Sub

















    '***************************************************
    '* procedury pobierania/zapisywania danych
    '***************************************************

    Private Sub PobierzDane()
        txtIdent.Text = aoIdentAkcji
        dtpDataAkcji.Value = aoDataAkcji
        txtOpis.Text = aoOpisAkcji
        txtUwagi.Text = aoUwagiAkcji
    End Sub

    Private Sub ZapiszDane()
        aoIdentAkcji = txtIdent.Text
        aoDataAkcji = dtpDataAkcji.Value.ToString("yyyy-MM-dd")
        aoOpisAkcji = txtOpis.Text
        aoUwagiAkcji = txtUwagi.Text
    End Sub

    '**************************************************************
    '** obsluga obiektu w trakcie edycji
    '**************************************************************

    Private Sub txtIdent_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIdent.TextChanged
        TestLen(txtIdent, "IDENTYFIKATOR", 20)
    End Sub
    Private Sub txtIdent_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIdent.GotFocus
        StartEdycjiTxt(txtIdent)
    End Sub
    Private Sub txtIdent_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIdent.LostFocus
        KoniecEdycjiTxt(txtIdent)
        '----------------------------
        aoIdentAkcji = txtIdent.Text
    End Sub

    Private Sub txtOpis_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOpis.TextChanged
        TestLen(txtOpis, "OPIS", 100)
    End Sub
    Private Sub txtOpis_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOpis.GotFocus
        StartEdycjiTxt(txtOpis)
    End Sub
    Private Sub txtOpis_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOpis.LostFocus
        KoniecEdycjiTxt(txtOpis)
    End Sub

    'Private Sub txtData_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    TestLen(txtData, "DATA", 10)
    'End Sub
    'Private Sub txtData_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    StartEdycjiTxt(txtData)
    'End Sub
    'Private Sub txtData_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    KoniecEdycjiTxt(txtData)
    'End Sub

    Private Sub txtUwagi_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUwagi.GotFocus
        StartEdycjiTxt(txtUwagi)
    End Sub
    Private Sub txtUwagi_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUwagi.LostFocus
        KoniecEdycjiTxt(txtUwagi)
    End Sub



    '*************************************************
    '*** przenoszenie danych miêdzy rekordem i zmiennymi
    '*************************************************

    Private Sub InitAkcjeSpec()
        asIdentAkcji = aoIdentAkcji        'IDENT       Text, 10
        asIdentKlienta = ""
        asWersjaAkcji = 1        'WERSJA      Integer, 2
    End Sub

    Private Sub PobierzAkcjeSpec()
        asIdentAkcji = GetTxtField(dagAkcjeSpec, 0)
        asIdentKlienta = GetTxtField(dagAkcjeSpec, 1)
        asWersjaAkcji = GetNumField(dagAkcjeSpec, 2)
    End Sub

    '**************************************************************
    '* obsluga komend
    '**************************************************************





    Private Sub cmdEdytuj_Click(sender As Object, e As EventArgs) Handles cmdEdytuj.Click

    End Sub

    'Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If DataSetAkcjeSpec.Tables(0).Rows.Count <= 0 Then
    '        'Raiseevent(Dodaj,Click )
    '        MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
    '            "Proszê najpierw dopisaæ rekord do tabeli...",
    '            "Uwaga :",
    '            System.Windows.Forms.MessageBoxButtons.OK,
    '            MessageBoxIcon.Information,
    '            MessageBoxDefaultButton.Button1)
    '    Else
    '        oOperacja = "EDYTUJ"
    '        PobierzAkcjeSpec()
    '        Dim Form As New EdytujKatalogAkcjeSpec
    '        oAktualizuj = False
    '        Form.ShowDialog()
    '        If Not oAktualizuj Then
    '        Else
    '            Dim foundRow As DataRow
    '            ' Create an array for the key values to find.
    '            Dim findTheseVals(1) As Object
    '            findTheseVals(0) = asIdentAkcji
    '            findTheseVals(1) = asIdentKlienta
    '            foundRow = DataSetAkcjeSpec.Tables(0).Rows.Find(findTheseVals)
    '            ' sprawdz czy jest...
    '            If Not (foundRow Is Nothing) Then
    '                'foundRow("IDENTAKCJI") = asIdentAkcji
    '                'foundRow("IDENTKLI") = asIdentKlienta
    '                foundRow("WERSJA") = ZmienWersje(asWersjaAkcji)    'Wersja         Integer, 2

    '                dagAkcjeSpec.Update()
    '                AktualizujBaze()
    '            End If
    '        End If
    '    End If
    'End Sub

    Private Sub cmdUsuñ_Click(sender As Object, e As EventArgs) Handles cmdUsuñ.Click
        If DataSetAkcjeSpec.Tables(0).Rows.Count <= 0 Then
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
                        MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then
                oOperacja = "USUN"
                PobierzAkcjeSpec()
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = asIdentAkcji
                findTheseVals(1) = asIdentKlienta
                foundRow = DataSetAkcjeSpec.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest..
                If Not (foundRow Is Nothing) Then
                    foundRow.Delete()
                    AktualizujBaze()
                Else
                End If
            End If
        End If
    End Sub

    Private Sub cmdDodaj_Click(sender As Object, e As EventArgs) Handles cmdDodaj.Click
        oOperacja = "DODAJ"
        InitAkcjeSpec()
        If Len(txtIdent.Text) = 0 Then
            MessageBox.Show("Nie potrafie dopisaæ specyfikacji" & vbCrLf &
                "Proszê najpierw zdefiniowaæ Ident.Akcji...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            Dim Form As New EdytujKatalogAkcjeSpec
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
                    findTheseVals(0) = asIdentAkcji
                    findTheseVals(1) = asIdentKlienta
                    foundRow = DataSetAkcjeSpec.Tables(0).Rows.Find(findTheseVals)
                    ' sprawdz czy jest...
                    If Not (foundRow Is Nothing) Then
                        MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                            "Ident.akcji = " & asIdentAkcji & vbCrLf &
                            "Ident.klienta = " & asIdentKlienta,
                            "Uwaga :",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1)
                    Else
                        'dbSelectSQLAkcjeSpec = "SELECT AkcjeSpec.IDENTAKCJI," &
                        '                      "AkcjeSpec.IDENTKLI," &
                        '                      "AkcjeSpec.WERSJA," &
                        '                   "Klienci.NAZWA1 AS NAZWAKLIENTA," &
                        '                   "Klienci.MIEJSCOWOSC," &
                        '                   "Klienci.KODPOCZTOWY," &
                        '                   "Klienci.ULICA," &
                        '                   "Klienci.NRDOMU," &
                        '                   "Klienci.REJKLIENTA AS REJONKLIENTA," &
                        '                   "Klienci.TELEFON," &
                        '                   "Klienci.FAX," &
                        '                   "Klienci.EMAIL " &
                        '               "FROM AkcjeSpec INNER JOIN Klienci " &
                        '               "ON AkcjeSpec.IDENTKLI=Klienci.IDENTKLIENTA " &
                        '               "WHERE AkcjeSpec.IDENTAKCJI='" & Trim(txtIdent.Text) & "'" &
                        '               "ORDER BY AkcjeSpec.IDENTAKCJI"


                        'stworz nowy rekord
                        NewRow = DataSetAkcjeSpec.Tables(0).NewRow()

                        NewRow("IDENTAKCJI") = asIdentAkcji
                        NewRow("IDENTKLI") = asIdentKlienta
                        NewRow("WERSJA") = ZmienWersje(asWersjaAkcji)    'Wersja         Integer, 2

                        NewRow("NAZWAKLIENTA") = ""
                        NewRow("MIEJSCOWOSC") = ""
                        NewRow("KODPOCZTOWY") = ""
                        NewRow("ULICA") = ""
                        NewRow("NRDOMU") = ""
                        NewRow("REJONKLIENTA") = ""
                        NewRow("TELEFON") = ""
                        NewRow("FAx") = ""
                        NewRow("EMAIL") = ""

                        'dodaj rekord do DataSet
                        DataSetAkcjeSpec.Tables(0).Rows.Add(NewRow)
                        AktualizujBaze()
                        OdswiezBaze()
                        dagAkcjeSpec.Update()
                        Exit Do
                    End If
                End If
            Loop
        End If
    End Sub

    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujBaze()
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
            '        dbDataAdapter.Update(DataSetAkcjeSpec, "TABELA_AkcjeSpec")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapter.Fill(DataSetAkcjeSpec)
            '    End If
            '    dbConnection.Close()
            'End If
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
                    sqlDataAdapter.Update(DataSetAkcjeSpec, "TABELA_AkcjeSpec")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapter.Fill(DataSetAkcjeSpec)
                End If
                sqlConnection.Close()
            End If
        End If
    End Sub



    Private Sub OdswiezBaze()
        DataSetAkcjeSpec.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnection.Open()
            '    If dbConnection.State = ConnectionState.Open Then
            '        dbDataAdapter.Fill(DataSetAkcjeSpec)
            '        dbConnection.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnection.Open()
                If sqlConnection.State = ConnectionState.Open Then
                    sqlDataAdapter.Fill(DataSetAkcjeSpec)
                    sqlConnection.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub



    '****************************************
    '** obsluga Menu Kontekstowego
    '****************************************

    Private Sub MenuUEDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEDodaj.Click
        cmdDodaj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub MenuUEEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEEdytuj.Click
        'cmdEdytuj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub MenuUEUsun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEUsun.Click
        cmdUsuñ_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuESingleL_Click(sender As Object, e As EventArgs) Handles menuESingleL.Click
        dagAkcjeSpec.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagAkcjeSpec.Refresh()
    End Sub

    Private Sub menuEMultiL_Click(sender As Object, e As EventArgs) Handles menuEMultiL.Click
        dagAkcjeSpec.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagAkcjeSpec.Refresh()
    End Sub

    Private Sub menuEOdswiez_Click(sender As Object, e As EventArgs) Handles menuEOdswiez.Click
        OdswiezBaze()
    End Sub





End Class
