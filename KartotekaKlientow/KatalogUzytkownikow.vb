Imports System.Drawing.Printing

Public Class KatalogUzytkownikow
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
    Friend WithEvents stbUzytkownicy As System.Windows.Forms.StatusBar
    Friend WithEvents menuWybieraj As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents cmdUsuñ As System.Windows.Forms.Button
    Friend WithEvents MenuWEdytuj As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuWDodaj As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuWUsuñ As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmdFi As Button
    Friend WithEvents cmdWszystko As Button
    Friend WithEvents stbFiltry As StatusBar
    Friend WithEvents lblUwagi As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents lblFunkcja As Label
    Friend WithEvents lblNazwa As Label
    Friend WithEvents lblIdent As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents dagUzytkownicy As DataGridView
    Friend WithEvents MenuEEdytuj As ToolStripMenuItem
    Friend WithEvents MenuEDodaj As ToolStripMenuItem
    Friend WithEvents MenuEUsun As ToolStripMenuItem
    Friend WithEvents menuEdytuj As ContextMenuStrip
    Friend WithEvents HorizScrollTimer As Timer
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents menuWSingleL As ToolStripMenuItem
    Friend WithEvents menuWMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents menuWOdswiez As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents menuESingleL As ToolStripMenuItem
    Friend WithEvents menuEMultiL As ToolStripMenuItem
    Friend WithEvents menuEOdswiez As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents MenuWWybierz As System.Windows.Forms.ToolStripMenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogUzytkownikow))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.stbUzytkownicy = New System.Windows.Forms.StatusBar()
        Me.cmdEdytuj = New System.Windows.Forms.Button()
        Me.cmdUsuñ = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.menuWybieraj = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.MenuWEdytuj = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuWDodaj = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuWUsuñ = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.MenuWWybierz = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuWSingleL = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuWMultiL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuWOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.cmdWszystko = New System.Windows.Forms.Button()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
        Me.lblUwagi = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblFunkcja = New System.Windows.Forms.Label()
        Me.lblNazwa = New System.Windows.Forms.Label()
        Me.lblIdent = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.dagUzytkownicy = New System.Windows.Forms.DataGridView()
        Me.MenuEEdytuj = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuEDodaj = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuEUsun = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEdytuj = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuESingleL = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEMultiL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuEOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        Me.HorizScrollTimer = New System.Windows.Forms.Timer(Me.components)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuWybieraj.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.dagUzytkownicy, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuEdytuj.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(723, 357)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(88, 23)
        Me.cmdStop.TabIndex = 37
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'stbUzytkownicy
        '
        Me.stbUzytkownicy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbUzytkownicy.Dock = System.Windows.Forms.DockStyle.None
        Me.stbUzytkownicy.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbUzytkownicy.Location = New System.Drawing.Point(8, 329)
        Me.stbUzytkownicy.Name = "stbUzytkownicy"
        Me.stbUzytkownicy.ShowPanels = True
        Me.stbUzytkownicy.Size = New System.Drawing.Size(803, 22)
        Me.stbUzytkownicy.TabIndex = 42
        Me.stbUzytkownicy.Text = "StatusBar1"
        '
        'cmdEdytuj
        '
        Me.cmdEdytuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdytuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdEdytuj.Image = CType(resources.GetObject("cmdEdytuj.Image"), System.Drawing.Image)
        Me.cmdEdytuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEdytuj.Location = New System.Drawing.Point(637, 357)
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
        Me.cmdUsuñ.Location = New System.Drawing.Point(551, 357)
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
        Me.cmdDodaj.Location = New System.Drawing.Point(465, 357)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(80, 23)
        Me.cmdDodaj.TabIndex = 39
        Me.cmdDodaj.Text = "Dodaj"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 357)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 38
        Me.PictureBox1.TabStop = False
        '
        'menuWybieraj
        '
        Me.menuWybieraj.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuWEdytuj, Me.MenuWDodaj, Me.MenuWUsuñ, Me.ToolStripSeparator1, Me.MenuWWybierz, Me.ToolStripSeparator2, Me.menuWSingleL, Me.menuWMultiL, Me.ToolStripSeparator3, Me.menuWOdswiez})
        Me.menuWybieraj.Name = "menuWybieraj"
        Me.menuWybieraj.Size = New System.Drawing.Size(187, 176)
        '
        'MenuWEdytuj
        '
        Me.MenuWEdytuj.Name = "MenuWEdytuj"
        Me.MenuWEdytuj.Size = New System.Drawing.Size(186, 22)
        Me.MenuWEdytuj.Text = "Edytuj"
        '
        'MenuWDodaj
        '
        Me.MenuWDodaj.Name = "MenuWDodaj"
        Me.MenuWDodaj.Size = New System.Drawing.Size(186, 22)
        Me.MenuWDodaj.Text = "Dodaj"
        '
        'MenuWUsuñ
        '
        Me.MenuWUsuñ.Name = "MenuWUsuñ"
        Me.MenuWUsuñ.Size = New System.Drawing.Size(186, 22)
        Me.MenuWUsuñ.Text = "Usuñ"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(183, 6)
        '
        'MenuWWybierz
        '
        Me.MenuWWybierz.Name = "MenuWWybierz"
        Me.MenuWWybierz.Size = New System.Drawing.Size(186, 22)
        Me.MenuWWybierz.Text = "Wybierz"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(183, 6)
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
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(183, 6)
        '
        'menuWOdswiez
        '
        Me.menuWOdswiez.Name = "menuWOdswiez"
        Me.menuWOdswiez.Size = New System.Drawing.Size(186, 22)
        Me.menuWOdswiez.Text = "Odœwie¿ Tabele"
        '
        'cmdFi
        '
        Me.cmdFi.Image = CType(resources.GetObject("cmdFi.Image"), System.Drawing.Image)
        Me.cmdFi.Location = New System.Drawing.Point(587, 80)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(16, 20)
        Me.cmdFi.TabIndex = 179
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'cmdWszystko
        '
        Me.cmdWszystko.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystko.Location = New System.Drawing.Point(608, 78)
        Me.cmdWszystko.Name = "cmdWszystko"
        Me.cmdWszystko.Size = New System.Drawing.Size(26, 22)
        Me.cmdWszystko.TabIndex = 175
        '
        'stbFiltry
        '
        Me.stbFiltry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltry.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltry.Location = New System.Drawing.Point(8, 78)
        Me.stbFiltry.Name = "stbFiltry"
        Me.stbFiltry.ShowPanels = True
        Me.stbFiltry.Size = New System.Drawing.Size(803, 22)
        Me.stbFiltry.TabIndex = 178
        Me.stbFiltry.Text = "StatusBar1"
        '
        'lblUwagi
        '
        Me.lblUwagi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUwagi.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblUwagi.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblUwagi.ForeColor = System.Drawing.Color.Purple
        Me.lblUwagi.Location = New System.Drawing.Point(378, 8)
        Me.lblUwagi.Name = "lblUwagi"
        Me.lblUwagi.Size = New System.Drawing.Size(433, 64)
        Me.lblUwagi.TabIndex = 177
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.lblFunkcja)
        Me.Panel1.Controls.Add(Me.lblNazwa)
        Me.Panel1.Controls.Add(Me.lblIdent)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(364, 64)
        Me.Panel1.TabIndex = 176
        '
        'lblFunkcja
        '
        Me.lblFunkcja.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFunkcja.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblFunkcja.ForeColor = System.Drawing.Color.Purple
        Me.lblFunkcja.Location = New System.Drawing.Point(79, 38)
        Me.lblFunkcja.Name = "lblFunkcja"
        Me.lblFunkcja.Size = New System.Drawing.Size(271, 16)
        Me.lblFunkcja.TabIndex = 35
        '
        'lblNazwa
        '
        Me.lblNazwa.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa.Location = New System.Drawing.Point(79, 22)
        Me.lblNazwa.Name = "lblNazwa"
        Me.lblNazwa.Size = New System.Drawing.Size(271, 16)
        Me.lblNazwa.TabIndex = 31
        '
        'lblIdent
        '
        Me.lblIdent.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdent.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIdent.ForeColor = System.Drawing.Color.Purple
        Me.lblIdent.Location = New System.Drawing.Point(79, 6)
        Me.lblIdent.Name = "lblIdent"
        Me.lblIdent.Size = New System.Drawing.Size(88, 16)
        Me.lblIdent.TabIndex = 34
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 16)
        Me.Label3.TabIndex = 36
        Me.Label3.Text = "Funkcja . . . . . . . . "
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(92, 16)
        Me.Label1.TabIndex = 33
        Me.Label1.Text = "Ident. . . . . . . . . ."
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(8, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(104, 16)
        Me.Label4.TabIndex = 32
        Me.Label4.Text = "Nazwisko . . . . . . . . "
        '
        'dagUzytkownicy
        '
        Me.dagUzytkownicy.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagUzytkownicy.Location = New System.Drawing.Point(8, 100)
        Me.dagUzytkownicy.Name = "dagUzytkownicy"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagUzytkownicy.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagUzytkownicy.Size = New System.Drawing.Size(803, 229)
        Me.dagUzytkownicy.TabIndex = 174
        '
        'MenuEEdytuj
        '
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
        'menuEdytuj
        '
        Me.menuEdytuj.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuEEdytuj, Me.MenuEDodaj, Me.MenuEUsun, Me.ToolStripSeparator4, Me.menuESingleL, Me.menuEMultiL, Me.ToolStripSeparator5, Me.menuEOdswiez})
        Me.menuEdytuj.Name = "menuEdytuj"
        Me.menuEdytuj.Size = New System.Drawing.Size(187, 148)
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
        'HorizScrollTimer
        '
        '
        'KatalogUzytkownikow
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(817, 389)
        Me.Controls.Add(Me.cmdFi)
        Me.Controls.Add(Me.cmdWszystko)
        Me.Controls.Add(Me.stbFiltry)
        Me.Controls.Add(Me.lblUwagi)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.dagUzytkownicy)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.stbUzytkownicy)
        Me.Controls.Add(Me.cmdEdytuj)
        Me.Controls.Add(Me.cmdUsuñ)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.PictureBox1)
        Me.Name = "KatalogUzytkownikow"
        Me.ShowInTaskbar = False
        Me.Text = "Katalog U¿ytkowników"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menuWybieraj.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        CType(Me.dagUzytkownicy, System.ComponentModel.ISupportInitialize).EndInit()
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


    Dim DataSetUzytkownicy As New DataSet
    Dim DataViewUzytkownicy As New DataView
    Dim ConnectionState As System.Data.ConnectionState

    Dim _CoMamRobic As String = ""

    '----------------------------------------------------------
    'Katalog Uzytkownicy
    '----------------------------------------------------------
    'Public oIdentUzytkownika As String             'IDENT  Text    10
    'Public oNazwaUzytkownika As String             'NAZWA             Text    100
    'Public oFunkcjaUzytkownika As String           'FUNKCJA           Text    100
    'Public oHasloUzytkownika As String             'HASLO             Text    100
    'Public oUprawnieniaUzytkownika As String       'UPRAWNIENIA       Text    100
    'Public oUwagiUzytkownika As String   'UWAGI          Text
    'Public oWersjaUzytkownika As Integer 'WERSJA         Integer


    Private Sub KatalogUzytkownikow_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        _CoMamRobic = oCoMamRobic
        '--------------------------------
        DataSetUzytkownicy.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionUzytkownicy)
            'dbCommandSelect = New OleDb.OleDbCommand(sSelectSQLUzytkownicy, dbConnection)
            'dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLUzytkownicy, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLUzytkownicy, dbConnection)
            'dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLUzytkownicy, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Uzytkownicy")

            'DBDeleteUzytkownicy(dbCommandDelete, dbDataAdapter)
            'DBUpdateUzytkownicy(dbCommandUpdate, dbDataAdapter)
            'DBInsertUzytkownicy(dbCommandInsert, dbDataAdapter)

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
            sqlConnection = New SqlClient.SqlConnection(sConnectionUzytkownicy)
            sqlCommandSelect = New SqlClient.SqlCommand(sSelectSQLUzytkownicy, sqlConnection)
            sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLUzytkownicy, sqlConnection)
            sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLUzytkownicy, sqlConnection)
            sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLUzytkownicy, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_Uzytkownicy")

            SQLDeleteUzytkownicy(sqlCommandDelete, sqlDataAdapter)
            SQLUpdateUzytkownicy(sqlCommandUpdate, sqlDataAdapter)
            SQLInsertUzytkownicy(sqlCommandInsert, sqlDataAdapter)

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
                    'dbDataAdapter.FillSchema(DataSetUzytkownicy, SchemaType.Mapped, "TABELA_Uzytkownicy")
                    'dbDataAdapter.Fill(DataSetUzytkownicy)
                    'dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(DataSetUzytkownicy, SchemaType.Mapped, "TABELA_Uzytkownicy")
                    sqlDataAdapter.Fill(DataSetUzytkownicy)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetUzytkownicy.Tables(0).PrimaryKey = New DataColumn() {DataSetUzytkownicy.Tables(0).Columns("IDENT")}


                'definiuj DataView
                DataViewUzytkownicy = New DataView(DataSetUzytkownicy.Tables(0))
                DataViewUzytkownicy.AllowDelete = False
                DataViewUzytkownicy.AllowNew = False

                '-------------------------------------------------
                'parametry DataGridView
                '-------------------------------------------------
                ' Initialize basic DataGridView properties.
                dagUzytkownicy.BackgroundColor = System.Drawing.SystemColors.Control
                dagUzytkownicy.GridColor = System.Drawing.SystemColors.ControlDark
                dagUzytkownicy.ForeColor = System.Drawing.Color.Purple
                dagUzytkownicy.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagUzytkownicy.BorderStyle = BorderStyle.Fixed3D
                'dagUzytkownicy.Dock = DockStyle.Fill
                dagUzytkownicy.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                ' Set property values appropriate for read-only display and limited interactivity. 
                dagUzytkownicy.AllowUserToAddRows = False
                dagUzytkownicy.AllowUserToDeleteRows = False
                dagUzytkownicy.AllowUserToOrderColumns = True
                dagUzytkownicy.AllowUserToResizeColumns = True
                dagUzytkownicy.AllowUserToResizeRows = True
                dagUzytkownicy.ReadOnly = True
                dagUzytkownicy.SelectionMode = DataGridViewSelectionMode.CellSelect
                dagUzytkownicy.MultiSelect = False

                'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                'dagUzytkownicy.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                dagUzytkownicy.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                dagUzytkownicy.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                dagUzytkownicy.ColumnHeadersVisible = True
                dagUzytkownicy.ColumnHeadersHeight = 40
                dagUzytkownicy.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                dagUzytkownicy.RowHeadersVisible = True
                dagUzytkownicy.RowHeadersWidth = 50
                dagUzytkownicy.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                ' Set the selection background color for all the cells.
                dagUzytkownicy.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                dagUzytkownicy.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                dagUzytkownicy.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagUzytkownicy.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                dagUzytkownicy.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagUzytkownicy.DefaultCellStyle.NullValue = ""
                dagUzytkownicy.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                dagUzytkownicy.DefaultCellStyle.WrapMode = DataGridViewTriState.True


                ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                dagUzytkownicy.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                dagUzytkownicy.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                ' Set the background color for all rows and for alternating rows.  
                ' The value for alternating rows overrides the value for all rows. 
                dagUzytkownicy.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagUzytkownicy.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                ' Set the row and column header styles.
                dagUzytkownicy.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagUzytkownicy.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                dagUzytkownicy.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagUzytkownicy.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                '-----------------------------------
                'wypelnienie DataGridView
                dagUzytkownicy.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                dagUzytkownicy.DataMember = DataSetUzytkownicy.Tables(0).TableName
                'dagUzytkownicy.DataSource = DataViewUzytkownicy
                '-----------------------------------

                'Public oIdentUzytkownika As String             'IDENT  Text    10
                'Public oNazwaUzytkownika As String             'NAZWA             Text    100
                'Public oFunkcjaUzytkownika As String           'FUNKCJA           Text    100
                'Public oHasloUzytkownika As String             'HASLO             Text    100
                'Public oUprawnieniaUzytkownika As String       'UPRAWNIENIA       Text    100
                'Public oUwagiUzytkownika As String   'UWAGI          Text
                'Public oWersjaUzytkownika As Integer 'WERSJA         Integer

                TxtColumnView(dagUzytkownicy, DataSetUzytkownicy.Tables(0).Columns(0).ColumnName, "Ident", 100, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagUzytkownicy, DataSetUzytkownicy.Tables(0).Columns(1).ColumnName, "Imie i nazwisko", 300, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagUzytkownicy, DataSetUzytkownicy.Tables(0).Columns(2).ColumnName, "Funkcja", 300, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagUzytkownicy, DataSetUzytkownicy.Tables(0).Columns(3).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagUzytkownicy, DataSetUzytkownicy.Tables(0).Columns(4).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagUzytkownicy, DataSetUzytkownicy.Tables(0).Columns(5).ColumnName, "Uwagi", 300, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagUzytkownicy, DataSetUzytkownicy.Tables(0).Columns(6).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)

                Me.Cursor = Cursors.WaitCursor
                dagUzytkownicy.DataSource = DataViewUzytkownicy
                Me.Cursor = Cursors.Default

                'ustaw sie na pierwszym zapisie
                If DataSetUzytkownicy.Tables(0).Rows.Count > 0 Then
                    dagUzytkownicy.CurrentCell = dagUzytkownicy(0, 0)
                    dagUzytkownicy.CurrentCell.Selected = True
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
            stbUzytkownicy.BackColor = System.Drawing.SystemColors.Control
            stbUzytkownicy.ForeColor = System.Drawing.Color.DarkGreen
            stbUzytkownicy.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            stbUzytkownicy.Panels.Add("stbPanelIleRec")
            stbUzytkownicy.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbUzytkownicy.Panels(0).Width = 80
            stbUzytkownicy.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbUzytkownicy.Panels(0).Text = "Iloœæ rec.: " & DataSetUzytkownicy.Tables(0).Rows.Count.ToString

            stbUzytkownicy.Panels.Add("stbPanelWiersz")
            stbUzytkownicy.Panels(1).AutoSize = StatusBarPanelAutoSize.None
            stbUzytkownicy.Panels(1).Width = 80
            stbUzytkownicy.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagUzytkownicy.CurrentCell Is Nothing Then
                stbUzytkownicy.Panels(1).Text = "Wi: "
            Else
                stbUzytkownicy.Panels(1).Text = "Wi: " & dagUzytkownicy.CurrentCell.RowIndex.ToString
            End If

            stbUzytkownicy.Panels.Add("stbPanelKolumna")
            stbUzytkownicy.Panels(2).AutoSize = StatusBarPanelAutoSize.None
            stbUzytkownicy.Panels(2).Width = 80
            stbUzytkownicy.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagUzytkownicy.CurrentCell Is Nothing Then
                stbUzytkownicy.Panels(2).Text = "Ko: "
            Else
                stbUzytkownicy.Panels(2).Text = "Ko: " & dagUzytkownicy.CurrentCell.ColumnIndex.ToString
            End If

            stbUzytkownicy.Panels.Add("stbPanelSort")
            stbUzytkownicy.Panels(3).AutoSize = StatusBarPanelAutoSize.None
            stbUzytkownicy.Panels(3).Width = 150
            stbUzytkownicy.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbUzytkownicy.Panels(3).Text = "Sort : "

            stbUzytkownicy.Panels.Add("stbPanelFiltr")
            stbUzytkownicy.Panels(4).AutoSize = StatusBarPanelAutoSize.None
            stbUzytkownicy.Panels(4).Width = 150
            stbUzytkownicy.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbUzytkownicy.Panels(4).Text = "Filtr : "

            stbUzytkownicy.Panels.Add("stbPanelReszta")
            stbUzytkownicy.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
            'stbUzytkownicy.Panels(5).Width = 150
            stbUzytkownicy.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbUzytkownicy.Panels(5).Text = ""

            stbUzytkownicy.ShowPanels = True
            '---------------------------------
            'dodaj StatusBar na koncu
            stbFiltry.BackColor = System.Drawing.SystemColors.Control
            stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
            stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            stbFiltry.Panels.Add("stbFiltryRowHead")
            stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbFiltry.Panels(0).Width = 50  'dagUzytkownicy.TableStyles(0).RowHeaderWidth
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

        InitInfoUzytkownika() 'inicjuj zmienne
        '-----------------------------------------------------------------
        If _OCoMamRobic = "WYBIERAJ" Then
            dagUzytkownicy.ContextMenuStrip = Me.menuWybieraj
            cmdStop.Text = "Wybierz"
            menuWybieraj.Visible = False
            menuWybieraj.Enabled = True
            menuEdytuj.Visible = False
            menuEdytuj.Enabled = False
        Else
            dagUzytkownicy.ContextMenuStrip = Me.menuEdytuj
            cmdStop.Text = "Powrót"
            menuWybieraj.Visible = False
            menuWybieraj.Enabled = False
            menuEdytuj.Visible = False
            menuEdytuj.Enabled = True
        End If
        '--------------------------------------------------------------------
        'set the header width to something apporpriate
        dagUzytkownicy.RowHeadersWidth = 50       'use if tablestyle
        'Me.dagUzytkownicy.RowHeaderWidth = 80 'use if no tablestyle
        '--------------------------------------------------------------------
        'set topleft corner point so we can locate the toprow
        If DataSetUzytkownicy.Tables(0).Rows.Count > 0 Then
            StartPointInCell00 = New System.Drawing.Point((dagUzytkownicy.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (dagUzytkownicy.GetCellDisplayRectangle(0, 0, True).Y + 4))
            ScrollXOffset = 10000 * dagUzytkownicy.FirstDisplayedScrollingColumnIndex +
                        dagUzytkownicy.GetCellDisplayRectangle(dagUzytkownicy.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
        Else
            'brak zapisow - poczatek DataGrid
            StartPointInCell00 = New System.Drawing.Point((dagUzytkownicy.Left + 4), (dagUzytkownicy.Top + 4))
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
        For ii = 0 To DataViewUzytkownicy.Table.Columns.Count - 1
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
        dagUzytkownicy_CurrentCellChanged(sender, e)
        InicjujFiltryColumnDGV(Me, dagUzytkownicy, DataViewUzytkownicy.Table.Columns.Count, stbFiltry, StartFormy)
    End Sub

    Private Sub KatalogUzytkownikow_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub


    Private Sub dagUzytkownicy_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagUzytkownicy.CurrentCellChanged
        If Not StartFormy Then
            If dagUzytkownicy.CurrentCell Is Nothing Then
                stbUzytkownicy.Panels(1).Text = "Wi: "
                stbUzytkownicy.Panels(2).Text = "Ko: "

                lblIdent.Text = ""
                lblNazwa.Text = ""
                lblFunkcja.Text = ""
                lblUwagi.Text = ""
            Else
                stbUzytkownicy.Panels(1).Text = "Wi: " & dagUzytkownicy.CurrentCell.RowIndex.ToString
                stbUzytkownicy.Panels(2).Text = "Ko: " & dagUzytkownicy.CurrentCell.ColumnIndex.ToString

                If DataViewUzytkownicy.Count = 0 Then
                    lblIdent.Text = ""
                    lblNazwa.Text = ""
                    lblFunkcja.Text = ""
                    lblUwagi.Text = ""
                Else
                    lblIdent.Text = GetTxtField(dagUzytkownicy, 0)
                    lblNazwa.Text = GetTxtField(dagUzytkownicy, 1)
                    lblFunkcja.Text = GetTxtField(dagUzytkownicy, 2)
                    lblUwagi.Text = GetTxtField(dagUzytkownicy, 5)

                    If Program_IdUzytkownika = Program_AdministratorLogin Or
                       Program_IdUzytkownika = lblIdent.Text Then
                        'mo¿e sam siebie edytowaæ....
                        cmdEdytuj.Enabled = True
                        menuEdytuj.Enabled = True
                    Else
                        cmdEdytuj.Enabled = False
                        menuEdytuj.Enabled = False
                    End If
                End If
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
                If dagUzytkownicy.RowCount > 0 Then
                    If ScrollXOffset <> 10000 * dagUzytkownicy.FirstDisplayedScrollingColumnIndex +
                                    dagUzytkownicy.GetCellDisplayRectangle(dagUzytkownicy.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X Then

                        PokazFiltryColumnDGV(Me, dagUzytkownicy, DataViewUzytkownicy.Table.Columns.Count, stbFiltry, StartFormy)
                        '-----------------------------
                        ScrollXOffset = 10000 * dagUzytkownicy.FirstDisplayedScrollingColumnIndex +
                                    dagUzytkownicy.GetCellDisplayRectangle(dagUzytkownicy.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X
                        HorizScrollLastTime = ""
                    End If
                Else
                    PokazFiltryColumnDGV(Me, dagUzytkownicy, DataViewUzytkownicy.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Else
            End If
        End If
        HorizScrollTimer.Enabled = True
    End Sub


    Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Dim rc As DataGridViewSelectedRowCollection = dagUzytkownicy.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                dagUzytkownicy.Rows(nextrow).Selected = True
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
                    PrzejdzDoDGV(dagUzytkownicy, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            FiltrujDataViewDGV(Me, dagUzytkownicy, DataViewUzytkownicy.Table.Columns.Count, DataViewUzytkownicy, stbUzytkownicy)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me, dagUzytkownicy, Mid(sender.name, 6, 3), "Uzytkownicy")
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
            CzyscFiltryDGV(Me, dagUzytkownicy, DataViewUzytkownicy.Table.Columns.Count, DataViewUzytkownicy, stbUzytkownicy)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbUzytkownicy.Panels(0).Text = "Il.zap.: " & DataViewUzytkownicy.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagUzytkownicy_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagUzytkownicy.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagUzytkownicy.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagUzytkownicy.DefaultCellStyle.Font,
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
            PokazFiltryColumnDGV(Me, dagUzytkownicy, DataViewUzytkownicy.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagUzytkownicy_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagUzytkownicy.CellFormatting
    '    If Not StartFormy Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagUzytkownicy, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagUzytkownicy, e.RowIndex, 4)

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

    Private Sub dagUzytkownicy_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagUzytkownicy.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagUzytkownicy, DataViewUzytkownicy.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagUzytkownicy_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagUzytkownicy.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagUzytkownicy, DataViewUzytkownicy.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagUzytkownicy_Scroll(sender As Object, e As ScrollEventArgs) Handles dagUzytkownicy.Scroll
        'If (Not StartFormy) And (DataViewUzytkownicy.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me, dagUzytkownicy, DataViewUzytkownicy.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me, dagUzytkownicy, DataViewUzytkownicy.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If

        If (Not StartFormy) And (DataViewUzytkownicy.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagUzytkownicy_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagUzytkownicy.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagUzytkownicy, DataViewUzytkownicy.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagUzytkownicy_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagUzytkownicy.MouseMove
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
                '    PokazFiltryColumnDGV(Me, dagUzytkownicy, DataViewUzytkownicy.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me, dagUzytkownicy, DataViewUzytkownicy.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagUzytkownicy_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagUzytkownicy.MouseUp
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
                stbUzytkownicy.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbUzytkownicy.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagUzytkownicy(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbUzytkownicy.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbUzytkownicy.Panels(3).Text = "Sort: " & DataSetUzytkownicy.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbUzytkownicy.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagUzytkownicy(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbUzytkownicy.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbUzytkownicy.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub




    Private Sub dagUzytkownicy_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagUzytkownicy.DoubleClick
        If _CoMamRobic = "WYBIERAJ" Then
            cmdStop_Click(sender, e)
        Else
            cmdEdytuj_Click(sender, e)
        End If
    End Sub

    Private Sub dagUzytkownicy_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagUzytkownicy.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                If _CoMamRobic = "WYBIERAJ" Then
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
    '----------------------------------------------------------
    'Katalog Uzytkownicy
    '----------------------------------------------------------
    'Public oIdentUzytkownika As String             'IDENT  Text    10
    'Public oNazwaUzytkownika As String             'NAZWA             Text    100
    'Public oFunkcjaUzytkownika As String           'FUNKCJA           Text    100
    'Public oHasloUzytkownika As String             'HASLO             Text    100
    'Public oUprawnieniaUzytkownika As String       'UPRAWNIENIA       Text    100
    'Public oUwagiUzytkownika As String   'UWAGI          Text
    'Public oWersjaUzytkownika As Integer 'WERSJA         Integer


    'Private Sub InitUzytkownicy()
    '    oIdentUzytkownika = ""        'IDENT       Text, 10
    '    oNazwaUzytkownika = ""        'NAZWA       Text, 50
    '    oFunkcjaUzytkownika = ""         'OPIS        Text, 50
    '    oWersjaUzytkownika = 1        'WERSJA      Integer, 2
    'End Sub

    Private Sub PobierzUzytkownicy()
        'pobierz wartosci przed aktualizacja
        oIdentUzytkownika = GetTxtField(dagUzytkownicy, 0)
        oNazwaUzytkownika = GetTxtField(dagUzytkownicy, 1)
        oFunkcjaUzytkownika = GetTxtField(dagUzytkownicy, 2)
        oHasloUzytkownika = GetTxtField(dagUzytkownicy, 3)
        oUprawnieniaUzytkownika = GetTxtField(dagUzytkownicy, 4)
        oUwagiUzytkownika = GetTxtField(dagUzytkownicy, 5)
        oWersjaUzytkownika = GetIntField(dagUzytkownicy, 6)
    End Sub

    '**************************************************************
    '* obsluga komend
    '**************************************************************

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        If DataSetUzytkownicy.Tables(0).Rows.Count > 0 Then
            oPracownik = GetTxtField(dagUzytkownicy, 0)
        Else
            oPracownik = ""        'IDENT       Text, 10
        End If
        Me.Close()
    End Sub

    Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytuj.Click
        If DataSetUzytkownicy.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            oOperacja = "EDYTUJ"
            PobierzUzytkownicy()
            Dim Form As New frmEdytujKatalogUzytkownikow
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
            Else
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = oIdentUzytkownika
                foundRow = DataSetUzytkownicy.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then

                    'Public oIdentUzytkownika As String             'IDENT  Text    10
                    'Public oNazwaUzytkownika As String             'NAZWA             Text    100
                    'Public oFunkcjaUzytkownika As String           'FUNKCJA           Text    100
                    'Public oHasloUzytkownika As String             'HASLO             Text    100
                    'Public oUprawnieniaUzytkownika As String       'UPRAWNIENIA       Text    100
                    'Public oUwagiUzytkownika As String   'UWAGI          Text
                    'Public oWersjaUzytkownika As Integer 'WERSJA         Integer

                    'foundRow("IDENT") = oIdentUzytkownika
                    foundRow("NAZWA") = oNazwaUzytkownika         'NAZWA     Text 50
                    foundRow("FUNKCJA") = oFunkcjaUzytkownika           'OPIS      Text 50
                    foundRow("HASLO") = oHasloUzytkownika         'NAZWA     Text 50
                    foundRow("UPRAWNIENIA") = oUprawnieniaUzytkownika         'NAZWA     Text 50
                    foundRow("UWAGI") = oUwagiUzytkownika         'NAZWA     Text 50
                    foundRow("WERSJA") = ZmienWersje(oWersjaUzytkownika)    'Wersja         Integer, 2

                    'wyswietl ilosc rekordow
                    stbUzytkownicy.Panels(0).Text = "Iloœæ rec.: " & DataSetUzytkownicy.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    dagUzytkownicy.Update()
                End If
            End If
        End If
    End Sub


    Private Sub cmdUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñ.Click
        If DataSetUzytkownicy.Tables(0).Rows.Count <= 0 Then
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
                PobierzUzytkownicy()
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = oIdentUzytkownika
                foundRow = DataSetUzytkownicy.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    foundRow.Delete()
                    'wyswietl ilosc rekordow
                    stbUzytkownicy.Panels(0).Text = "Iloœæ rec.: " & DataSetUzytkownicy.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                Else
                End If
            End If
        End If
    End Sub

    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        oOperacja = "DODAJ"
        InitInfoUzytkownika()
        Dim Form As New frmEdytujKatalogUzytkownikow
        Do While True
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
                Exit Do
            Else
                Dim foundRow As DataRow
                Dim NewRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = oIdentUzytkownika
                foundRow = DataSetUzytkownicy.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                        "Ident.u¿ytkownika = " & oIdentUzytkownika,
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    'stworz nowy rekord
                    NewRow = DataSetUzytkownicy.Tables(0).NewRow()

                    NewRow("IDENT") = oIdentUzytkownika
                    NewRow("NAZWA") = oNazwaUzytkownika         'NAZWA     Text 50
                    NewRow("FUNKCJA") = oFunkcjaUzytkownika           'OPIS      Text 50
                    NewRow("HASLO") = oHasloUzytkownika         'NAZWA     Text 50
                    NewRow("UPRAWNIENIA") = oUprawnieniaUzytkownika         'NAZWA     Text 50
                    NewRow("UWAGI") = oUwagiUzytkownika         'NAZWA     Text 50
                    NewRow("WERSJA") = 1                     'Wersja         Integer, 2
                    'dodaj rekord do DataSet
                    DataSetUzytkownicy.Tables(0).Rows.Add(NewRow)

                    'wyswietl ilosc rekordow
                    stbUzytkownicy.Panels(0).Text = "Iloœæ rec.: " & DataSetUzytkownicy.Tables(0).Rows.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    dagUzytkownicy.Update()
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
            'Try
            '    dbConnection.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnection.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapter.Update(DataSetUzytkownicy, "TABELA_Uzytkownicy")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapter.Fill(DataSetUzytkownicy)
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
                    sqlDataAdapter.Update(DataSetUzytkownicy, "TABELA_Uzytkownicy")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapter.Fill(DataSetUzytkownicy)
                End If
                sqlConnection.Close()
            End If
        End If
    End Sub

    Private Sub OdswiezBaze()
        DataSetUzytkownicy.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnection.Open()
            '    If dbConnection.State = ConnectionState.Open Then
            '        dbDataAdapter.Fill(DataSetUzytkownicy)
            '        dbConnection.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnection.Open()
                If sqlConnection.State = ConnectionState.Open Then
                    sqlDataAdapter.Fill(DataSetUzytkownicy)
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
        cmdEdytuj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub MenuUEUsun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuEUsun.Click
        cmdUsuñ_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub MenuUWEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuWEdytuj.Click
        cmdEdytuj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuESingleL_Click(sender As Object, e As EventArgs) Handles menuESingleL.Click
        dagUzytkownicy.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagUzytkownicy.Refresh()
    End Sub

    Private Sub menuEMultiL_Click(sender As Object, e As EventArgs) Handles menuEMultiL.Click
        dagUzytkownicy.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagUzytkownicy.Refresh()
    End Sub

    Private Sub menuEOdswiez_Click(sender As Object, e As EventArgs) Handles menuEOdswiez.Click
        OdswiezBaze()
    End Sub




    Private Sub MenuUWDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuWDodaj.Click
        cmdDodaj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub MenuUWUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuWUsuñ.Click
        cmdUsuñ_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub MenuUWWybierz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuWWybierz.Click
        cmdStop_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuWSingleL_Click(sender As Object, e As EventArgs) Handles menuWSingleL.Click
        dagUzytkownicy.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagUzytkownicy.Refresh()
    End Sub

    Private Sub menuWMultiL_Click(sender As Object, e As EventArgs) Handles menuWMultiL.Click
        dagUzytkownicy.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagUzytkownicy.Refresh()
    End Sub

    Private Sub menuWOdswiez_Click(sender As Object, e As EventArgs) Handles menuWOdswiez.Click
        OdswiezBaze()
    End Sub

End Class
