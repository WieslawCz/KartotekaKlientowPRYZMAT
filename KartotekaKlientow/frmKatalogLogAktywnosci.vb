Imports System.Drawing.Printing

Public Class frmKatalogLogAktywnosci
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents cmdDodaj As System.Windows.Forms.Button
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents dagLogAktywnosci As System.Windows.Forms.DataGridView
    Friend WithEvents stbLogAktywnosci As System.Windows.Forms.StatusBar
    Friend WithEvents cmdUsuñ As System.Windows.Forms.Button
    Friend WithEvents cmdWszystko As System.Windows.Forms.Button
    Friend WithEvents cmdFi As System.Windows.Forms.Button
    Friend WithEvents cmdEdytuj As System.Windows.Forms.Button
    Friend WithEvents menuEdytuj As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents menuEDodaj As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuEUsuñ As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuEEdytuj As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents menuEPoka¿WJednejLinii As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuEPoka¿WWieluLiniach As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents HorizScrollTimer As System.Windows.Forms.Timer
    Friend WithEvents cmdDoExcela As System.Windows.Forms.Button
    Friend WithEvents menuEOdswiez As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblUwagi As System.Windows.Forms.Label
    Friend WithEvents lblDataZapisu As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblKatalog As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblPracownik As System.Windows.Forms.Label
    Friend WithEvents lblTemat As System.Windows.Forms.Label
    Friend WithEvents lblOperacja As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As Panel
    Friend WithEvents stbFiltry As System.Windows.Forms.StatusBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.stbLogAktywnosci = New System.Windows.Forms.StatusBar()
        Me.cmdUsuñ = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.dagLogAktywnosci = New System.Windows.Forms.DataGridView()
        Me.menuEdytuj = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuEDodaj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEUsuñ = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEEdytuj = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuEPoka¿WJednejLinii = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEPoka¿WWieluLiniach = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuEOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        Me.cmdWszystko = New System.Windows.Forms.Button()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.cmdEdytuj = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.HorizScrollTimer = New System.Windows.Forms.Timer(Me.components)
        Me.cmdDoExcela = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblPracownik = New System.Windows.Forms.Label()
        Me.lblOperacja = New System.Windows.Forms.Label()
        Me.lblTemat = New System.Windows.Forms.Label()
        Me.lblKatalog = New System.Windows.Forms.Label()
        Me.lblDataZapisu = New System.Windows.Forms.Label()
        Me.lblUwagi = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dagLogAktywnosci, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuEdytuj.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(664, 376)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(88, 22)
        Me.cmdStop.TabIndex = 37
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'stbLogAktywnosci
        '
        Me.stbLogAktywnosci.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbLogAktywnosci.Dock = System.Windows.Forms.DockStyle.None
        Me.stbLogAktywnosci.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbLogAktywnosci.Location = New System.Drawing.Point(0, 237)
        Me.stbLogAktywnosci.Name = "stbLogAktywnosci"
        Me.stbLogAktywnosci.ShowPanels = True
        Me.stbLogAktywnosci.Size = New System.Drawing.Size(744, 21)
        Me.stbLogAktywnosci.TabIndex = 42
        Me.stbLogAktywnosci.Text = "StatusBar1"
        '
        'cmdUsuñ
        '
        Me.cmdUsuñ.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUsuñ.Enabled = False
        Me.cmdUsuñ.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdUsuñ.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUsuñ.Location = New System.Drawing.Point(492, 376)
        Me.cmdUsuñ.Name = "cmdUsuñ"
        Me.cmdUsuñ.Size = New System.Drawing.Size(80, 22)
        Me.cmdUsuñ.TabIndex = 40
        Me.cmdUsuñ.Text = "Usuñ"
        Me.cmdUsuñ.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdDodaj
        '
        Me.cmdDodaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDodaj.Enabled = False
        Me.cmdDodaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDodaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDodaj.Location = New System.Drawing.Point(406, 376)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(80, 22)
        Me.cmdDodaj.TabIndex = 39
        Me.cmdDodaj.Text = "Dodaj"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'picLogo
        '
        Me.picLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.picLogo.ErrorImage = Nothing
        Me.picLogo.Location = New System.Drawing.Point(8, 376)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(104, 24)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 38
        Me.picLogo.TabStop = False
        '
        'dagLogAktywnosci
        '
        Me.dagLogAktywnosci.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagLogAktywnosci.ContextMenuStrip = Me.menuEdytuj
        Me.dagLogAktywnosci.Location = New System.Drawing.Point(0, 22)
        Me.dagLogAktywnosci.Name = "dagLogAktywnosci"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagLogAktywnosci.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagLogAktywnosci.Size = New System.Drawing.Size(744, 213)
        Me.dagLogAktywnosci.TabIndex = 44
        '
        'menuEdytuj
        '
        Me.menuEdytuj.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuEDodaj, Me.menuEUsuñ, Me.menuEEdytuj, Me.ToolStripSeparator1, Me.menuEPoka¿WJednejLinii, Me.menuEPoka¿WWieluLiniach, Me.ToolStripSeparator4, Me.menuEOdswiez})
        Me.menuEdytuj.Name = "menuEdytuj"
        Me.menuEdytuj.Size = New System.Drawing.Size(280, 148)
        '
        'menuEDodaj
        '
        Me.menuEDodaj.Enabled = False
        Me.menuEDodaj.Name = "menuEDodaj"
        Me.menuEDodaj.Size = New System.Drawing.Size(279, 22)
        Me.menuEDodaj.Text = "Dodaj"
        '
        'menuEUsuñ
        '
        Me.menuEUsuñ.Enabled = False
        Me.menuEUsuñ.Name = "menuEUsuñ"
        Me.menuEUsuñ.Size = New System.Drawing.Size(279, 22)
        Me.menuEUsuñ.Text = "Usuñ"
        '
        'menuEEdytuj
        '
        Me.menuEEdytuj.Enabled = False
        Me.menuEEdytuj.Name = "menuEEdytuj"
        Me.menuEEdytuj.Size = New System.Drawing.Size(279, 22)
        Me.menuEEdytuj.Text = "Edytuj"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(276, 6)
        '
        'menuEPoka¿WJednejLinii
        '
        Me.menuEPoka¿WJednejLinii.Name = "menuEPoka¿WJednejLinii"
        Me.menuEPoka¿WJednejLinii.Size = New System.Drawing.Size(279, 22)
        Me.menuEPoka¿WJednejLinii.Text = "Poka¿ w jednej linii"
        '
        'menuEPoka¿WWieluLiniach
        '
        Me.menuEPoka¿WWieluLiniach.Name = "menuEPoka¿WWieluLiniach"
        Me.menuEPoka¿WWieluLiniach.Size = New System.Drawing.Size(279, 22)
        Me.menuEPoka¿WWieluLiniach.Text = "Poka¿ w wielu liniach"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(276, 6)
        '
        'menuEOdswiez
        '
        Me.menuEOdswiez.Name = "menuEOdswiez"
        Me.menuEOdswiez.Size = New System.Drawing.Size(279, 22)
        Me.menuEOdswiez.Text = "Odœwie¿ Log Aktywnoœci Pracowników"
        '
        'cmdWszystko
        '
        Me.cmdWszystko.Location = New System.Drawing.Point(549, 0)
        Me.cmdWszystko.Name = "cmdWszystko"
        Me.cmdWszystko.Size = New System.Drawing.Size(32, 23)
        Me.cmdWszystko.TabIndex = 66
        '
        'stbFiltry
        '
        Me.stbFiltry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltry.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltry.Location = New System.Drawing.Point(0, 0)
        Me.stbFiltry.Name = "stbFiltry"
        Me.stbFiltry.ShowPanels = True
        Me.stbFiltry.Size = New System.Drawing.Size(744, 22)
        Me.stbFiltry.TabIndex = 85
        Me.stbFiltry.Text = "StatusBar1"
        '
        'cmdFi
        '
        Me.cmdFi.Location = New System.Drawing.Point(461, 0)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(16, 20)
        Me.cmdFi.TabIndex = 175
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'cmdEdytuj
        '
        Me.cmdEdytuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdytuj.Enabled = False
        Me.cmdEdytuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdEdytuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEdytuj.Location = New System.Drawing.Point(578, 376)
        Me.cmdEdytuj.Name = "cmdEdytuj"
        Me.cmdEdytuj.Size = New System.Drawing.Size(80, 22)
        Me.cmdEdytuj.TabIndex = 41
        Me.cmdEdytuj.Text = "Edytuj"
        Me.cmdEdytuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ToolTip1
        '
        Me.ToolTip1.AutoPopDelay = 5000
        Me.ToolTip1.InitialDelay = 100
        Me.ToolTip1.IsBalloon = True
        Me.ToolTip1.ReshowDelay = 100
        '
        'HorizScrollTimer
        '
        '
        'cmdDoExcela
        '
        Me.cmdDoExcela.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdDoExcela.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDoExcela.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDoExcela.Location = New System.Drawing.Point(118, 376)
        Me.cmdDoExcela.Name = "cmdDoExcela"
        Me.cmdDoExcela.Size = New System.Drawing.Size(84, 22)
        Me.cmdDoExcela.TabIndex = 184
        Me.cmdDoExcela.Text = "Do Excela"
        Me.cmdDoExcela.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.lblPracownik)
        Me.Panel1.Controls.Add(Me.lblOperacja)
        Me.Panel1.Controls.Add(Me.lblTemat)
        Me.Panel1.Controls.Add(Me.lblKatalog)
        Me.Panel1.Controls.Add(Me.lblDataZapisu)
        Me.Panel1.Controls.Add(Me.lblUwagi)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(743, 96)
        Me.Panel1.TabIndex = 186
        '
        'lblPracownik
        '
        Me.lblPracownik.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPracownik.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblPracownik.ForeColor = System.Drawing.Color.Purple
        Me.lblPracownik.Location = New System.Drawing.Point(293, 6)
        Me.lblPracownik.Name = "lblPracownik"
        Me.lblPracownik.Size = New System.Drawing.Size(88, 16)
        Me.lblPracownik.TabIndex = 10
        '
        'lblOperacja
        '
        Me.lblOperacja.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOperacja.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblOperacja.ForeColor = System.Drawing.Color.Purple
        Me.lblOperacja.Location = New System.Drawing.Point(104, 38)
        Me.lblOperacja.Name = "lblOperacja"
        Me.lblOperacja.Size = New System.Drawing.Size(277, 16)
        Me.lblOperacja.TabIndex = 7
        '
        'lblTemat
        '
        Me.lblTemat.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTemat.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTemat.ForeColor = System.Drawing.Color.Purple
        Me.lblTemat.Location = New System.Drawing.Point(104, 54)
        Me.lblTemat.Name = "lblTemat"
        Me.lblTemat.Size = New System.Drawing.Size(277, 16)
        Me.lblTemat.TabIndex = 9
        Me.lblTemat.Visible = False
        '
        'lblKatalog
        '
        Me.lblKatalog.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKatalog.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblKatalog.ForeColor = System.Drawing.Color.Purple
        Me.lblKatalog.Location = New System.Drawing.Point(104, 22)
        Me.lblKatalog.Name = "lblKatalog"
        Me.lblKatalog.Size = New System.Drawing.Size(277, 16)
        Me.lblKatalog.TabIndex = 2
        '
        'lblDataZapisu
        '
        Me.lblDataZapisu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDataZapisu.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblDataZapisu.ForeColor = System.Drawing.Color.Purple
        Me.lblDataZapisu.Location = New System.Drawing.Point(104, 6)
        Me.lblDataZapisu.Name = "lblDataZapisu"
        Me.lblDataZapisu.Size = New System.Drawing.Size(108, 16)
        Me.lblDataZapisu.TabIndex = 30
        '
        'lblUwagi
        '
        Me.lblUwagi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUwagi.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblUwagi.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblUwagi.ForeColor = System.Drawing.Color.Purple
        Me.lblUwagi.Location = New System.Drawing.Point(396, 8)
        Me.lblUwagi.Name = "lblUwagi"
        Me.lblUwagi.Size = New System.Drawing.Size(333, 80)
        Me.lblUwagi.TabIndex = 52
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(128, 16)
        Me.Label1.TabIndex = 29
        Me.Label1.Text = "Data zapisu . . . . . . . ."
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(128, 16)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Operacja . . . . . . . . . ."
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(8, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(128, 16)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "Katalog . . . . . . . . . . . ."
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(8, 56)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(128, 16)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Temat . . . . . . . . . ."
        Me.Label7.Visible = False
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(226, 6)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(120, 16)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Pracownik . . . . ."
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.Controls.Add(Me.stbLogAktywnosci)
        Me.Panel2.Controls.Add(Me.dagLogAktywnosci)
        Me.Panel2.Controls.Add(Me.cmdFi)
        Me.Panel2.Controls.Add(Me.cmdWszystko)
        Me.Panel2.Controls.Add(Me.stbFiltry)
        Me.Panel2.Location = New System.Drawing.Point(8, 110)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(744, 259)
        Me.Panel2.TabIndex = 187
        '
        'frmKatalogLogAktywnosci
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(758, 403)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdDoExcela)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.cmdEdytuj)
        Me.Controls.Add(Me.cmdUsuñ)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.picLogo)
        Me.Name = "frmKatalogLogAktywnosci"
        Me.Text = "Katalog Osób Kontaktowych"
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dagLogAktywnosci, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menuEdytuj.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
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

    Dim dbSelect As String = sSelectSQLLogAktywnosci

    'Dim dbConnection As OleDb.OleDbConnection
    'Dim dbCommandSelect As OleDb.OleDbCommand
    'Dim dbCommandDelete As OleDb.OleDbCommand
    'Dim dbCommandUpdate As OleDb.OleDbCommand
    'Dim dbCommandInsert As OleDb.OleDbCommand
    'Dim dbDataAdapter As OleDb.OleDbDataAdapter

    Dim sqlConnection As SqlClient.SqlConnection
    Dim sqlCommandSelect As SqlClient.SqlCommand
    Dim sqlCommandDelete As SqlClient.SqlCommand
    Dim sqlCommandUpdate As SqlClient.SqlCommand
    Dim sqlCommandInsert As SqlClient.SqlCommand
    Dim sqlDataAdapter As SqlClient.SqlDataAdapter

    Dim DataSetLogAktywnosci As New DataSet
    Dim DataViewLogAktywnosci As New DataView
    Dim ConnectionState As System.Data.ConnectionState

    'Public LG_UniqID As String       'UNIQID     STRING 40
    'Public LG_Temat As String        'TEMAT    STRING 100
    'Public LG_Katalog As String      'KATALOG    STRING 100
    'Public LG_DataZapisu As String   'DATAZAPISU STRING 30
    'Public LG_Pracownik As String    'PRACOWNIK  STRING 10
    'Public LG_Operacja As String     'OPERACJA   STRING 20
    'Public LG_Uwagi As String        'UWAGI      Text
    'Public LG_Wersja As Integer      'WERSJA     Integer

    Private Sub frmKatalogLogAktywnosci_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Icon = My.Resources.MrJones
        Me.picLogo.Image = My.Resources.logomini
        Me.cmdStop.Image = My.Resources._EXIT.ToBitmap

        Me.cmdFi.Image = My.Resources.ROZWIN3.ToBitmap
        Me.cmdWszystko.Image = My.Resources.close_16

        '-------------------------------
        dbSelect = sSelectSQLLogAktywnosci & " ORDER BY DATAZAPISU DESC"
        Me.Text = "Log aktywnoœci pracowników"

        DataSetLogAktywnosci.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionLogAktywnosci)
            'dbCommandSelect = New OleDb.OleDbCommand(sSelectSQLLogAktywnosci, dbConnection)
            'dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLLogAktywnosci, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLLogAktywnosci, dbConnection)
            'dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLLogAktywnosci, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            ''mapujemy domyslna nazwe tabeli
            'Dim DBMapowanieTabeli As System.Data.Common.DataTableMapping
            'DBMapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_LogAktywnosci")

            'DBDeleteLogAktywnosci(dbCommandDelete, dbDataAdapter)
            'DBUpdateLogAktywnosci(dbCommandUpdate, dbDataAdapter)
            'DBInsertLogAktywnosci(dbCommandInsert, dbDataAdapter)

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapter.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf dbOnRowUpdated)

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
            sqlConnection = New SqlClient.SqlConnection(sConnectionLogAktywnosci)
            sqlCommandSelect = New SqlClient.SqlCommand(dbSelect, sqlConnection)
            sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLLogAktywnosci, sqlConnection)
            sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLLogAktywnosci, sqlConnection)
            sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLLogAktywnosci, sqlConnection)
            sqlDataAdapter = New SqlClient.SqlDataAdapter(sqlCommandSelect)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeli As System.Data.Common.DataTableMapping
            sqlMapowanieTabeli = sqlDataAdapter.TableMappings.Add("Table", "TABELA_LogAktywnosci")

            SQLDeleteLogAktywnosci(sqlCommandDelete, sqlDataAdapter)
            SQLUpdateLogAktywnosci(sqlCommandUpdate, sqlDataAdapter)
            SQLInsertLogAktywnosci(sqlCommandInsert, sqlDataAdapter)

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
                    'dbDataAdapter.FillSchema(DataSetLogAktywnosci, SchemaType.Mapped, "TABELA_LogAktywnosci")
                    'dbDataAdapter.Fill(DataSetLogAktywnosci)
                    'dbConnection.Close()
                Else
                    sqlDataAdapter.FillSchema(DataSetLogAktywnosci, SchemaType.Mapped, "TABELA_LogAktywnosci")
                    sqlDataAdapter.Fill(DataSetLogAktywnosci)
                    sqlConnection.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetLogAktywnosci.Tables(0).PrimaryKey = New DataColumn() {DataSetLogAktywnosci.Tables(0).Columns("UNIQID")}

                'definiuj DataView
                DataViewLogAktywnosci = New DataView(DataSetLogAktywnosci.Tables(0))
                DataViewLogAktywnosci.AllowDelete = False
                DataViewLogAktywnosci.AllowNew = False

                '-------------------------------------------------
                'parametry DataGridView
                '-------------------------------------------------
                ' Initialize basic DataGridView properties.
                dagLogAktywnosci.BackgroundColor = System.Drawing.SystemColors.Control
                dagLogAktywnosci.GridColor = System.Drawing.SystemColors.ControlDark
                dagLogAktywnosci.ForeColor = System.Drawing.Color.Purple
                dagLogAktywnosci.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagLogAktywnosci.BorderStyle = BorderStyle.Fixed3D
                'dagLogAktywnosci.Dock = DockStyle.Fill
                dagLogAktywnosci.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                ' Set property values appropriate for read-only display and limited interactivity. 
                dagLogAktywnosci.AllowUserToAddRows = False
                dagLogAktywnosci.AllowUserToDeleteRows = False
                dagLogAktywnosci.AllowUserToOrderColumns = True
                dagLogAktywnosci.AllowUserToResizeColumns = True
                dagLogAktywnosci.AllowUserToResizeRows = True
                dagLogAktywnosci.ReadOnly = True
                dagLogAktywnosci.SelectionMode = DataGridViewSelectionMode.CellSelect
                dagLogAktywnosci.MultiSelect = False

                'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                'dagLogAktywnosci.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                dagLogAktywnosci.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                dagLogAktywnosci.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                dagLogAktywnosci.ColumnHeadersVisible = True
                dagLogAktywnosci.ColumnHeadersHeight = 40
                dagLogAktywnosci.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                dagLogAktywnosci.RowHeadersVisible = True
                dagLogAktywnosci.RowHeadersWidth = 50
                dagLogAktywnosci.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                ' Set the selection background color for all the cells.
                dagLogAktywnosci.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                dagLogAktywnosci.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                dagLogAktywnosci.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagLogAktywnosci.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                dagLogAktywnosci.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagLogAktywnosci.DefaultCellStyle.NullValue = ""
                dagLogAktywnosci.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                dagLogAktywnosci.DefaultCellStyle.WrapMode = DataGridViewTriState.False


                ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                dagLogAktywnosci.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                dagLogAktywnosci.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                ' Set the background color for all rows and for alternating rows.  
                ' The value for alternating rows overrides the value for all rows. 
                dagLogAktywnosci.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagLogAktywnosci.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                ' Set the row and column header styles.
                dagLogAktywnosci.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagLogAktywnosci.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                dagLogAktywnosci.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagLogAktywnosci.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                '-----------------------------------
                'wypelnienie DataGridView
                dagLogAktywnosci.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                dagLogAktywnosci.DataMember = DataSetLogAktywnosci.Tables(0).TableName
                'dagLogAktywnosci.DataSource = DataViewLogAktywnosci
                '-----------------------------------


                'Public LG_UniqID As String       'UNIQID     STRING 40
                'Public LG_Temat As String        'TEMAT      STRING 100
                'Public LG_Katalog As String      'KATALOG    STRING 100
                'Public LG_DataZapisu As String   'DATAZAPISU STRING 30
                'Public LG_Pracownik As String    'PRACOWNIK  STRING 10
                'Public LG_Operacja As String     'OPERACJA   STRING 20
                'Public LG_Uwagi As String        'UWAGI      Text
                'Public LG_Wersja As Integer      'WERSJA     Integer

                TxtColumnView(dagLogAktywnosci, DataSetLogAktywnosci.Tables(0).Columns(0).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagLogAktywnosci, DataSetLogAktywnosci.Tables(0).Columns(1).ColumnName, "Temat", 200, DataGridViewContentAlignment.MiddleLeft, False)
                TxtColumnView(dagLogAktywnosci, DataSetLogAktywnosci.Tables(0).Columns(2).ColumnName, "Katalog", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagLogAktywnosci, DataSetLogAktywnosci.Tables(0).Columns(3).ColumnName, "Data zapisu", 150, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagLogAktywnosci, DataSetLogAktywnosci.Tables(0).Columns(4).ColumnName, "Pracownik", 100, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagLogAktywnosci, DataSetLogAktywnosci.Tables(0).Columns(5).ColumnName, "Operacja", 120, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagLogAktywnosci, DataSetLogAktywnosci.Tables(0).Columns(6).ColumnName, "Uwagi", 500, DataGridViewContentAlignment.MiddleLeft, True)
                NumColumnView(dagLogAktywnosci, DataSetLogAktywnosci.Tables(0).Columns(7).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "F0", False)

                Me.Cursor = Cursors.WaitCursor
                dagLogAktywnosci.DataSource = DataViewLogAktywnosci
                Me.Cursor = Cursors.Default

                ''ustaw sie na pierwszym zapisie
                If DataSetLogAktywnosci.Tables(0).Rows.Count > 0 Then
                    dagLogAktywnosci.CurrentCell = dagLogAktywnosci(2, 0)
                    dagLogAktywnosci.CurrentCell.Selected = True
                End If

            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            'dodaj StatusBar na koncu
            stbLogAktywnosci.BackColor = System.Drawing.SystemColors.Control
            stbLogAktywnosci.ForeColor = System.Drawing.Color.DarkGreen
            stbLogAktywnosci.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            stbLogAktywnosci.Panels.Add("stbPanelIleRec")
            stbLogAktywnosci.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbLogAktywnosci.Panels(0).Width = 120
            stbLogAktywnosci.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbLogAktywnosci.Panels(0).Text = "Il.zap.: " & DataViewLogAktywnosci.Count.ToString

            stbLogAktywnosci.Panels.Add("stbPanelWiersz")
            stbLogAktywnosci.Panels(1).AutoSize = StatusBarPanelAutoSize.None
            stbLogAktywnosci.Panels(1).Width = 120
            stbLogAktywnosci.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagLogAktywnosci.CurrentCell Is Nothing Then
                stbLogAktywnosci.Panels(1).Text = "Wi: "
            Else
                stbLogAktywnosci.Panels(1).Text = "Wi: " & dagLogAktywnosci.CurrentCell.RowIndex.ToString
            End If

            stbLogAktywnosci.Panels.Add("stbPanelKolumna")
            stbLogAktywnosci.Panels(2).AutoSize = StatusBarPanelAutoSize.None
            stbLogAktywnosci.Panels(2).Width = 120
            stbLogAktywnosci.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagLogAktywnosci.CurrentCell Is Nothing Then
                stbLogAktywnosci.Panels(2).Text = "Ko: "
            Else
                stbLogAktywnosci.Panels(2).Text = "Ko: " & dagLogAktywnosci.CurrentCell.ColumnIndex.ToString
            End If

            stbLogAktywnosci.Panels.Add("stbPanelSort")
            stbLogAktywnosci.Panels(3).AutoSize = StatusBarPanelAutoSize.None
            stbLogAktywnosci.Panels(3).Width = 150
            stbLogAktywnosci.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbLogAktywnosci.Panels(3).Text = "Sort: " & DataSetLogAktywnosci.Tables(0).Columns(0).ColumnName

            stbLogAktywnosci.Panels.Add("stbPanelReszta")
            stbLogAktywnosci.Panels(4).AutoSize = StatusBarPanelAutoSize.Spring
            'stbLogAktywnosci.Panels(4).Width = 150
            stbLogAktywnosci.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbLogAktywnosci.Panels(4).Text = ""

            stbLogAktywnosci.ShowPanels = True
            '---------------------------------
            'dodaj StatusBar na koncu
            stbFiltry.BackColor = System.Drawing.SystemColors.Control
            stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
            stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            stbFiltry.Panels.Add("stbFiltryRowHead")
            stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbFiltry.Panels(0).Width = 50  'dagLogAktywnosci.TableStyles(0).RowHeaderWidth
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
        InitLogAktywnosci() 'inicjuj zmienne
        '--------------------------------------------------------------------
        'set the header width to something apporpriate
        dagLogAktywnosci.RowHeadersWidth = 50       'use if tablestyle
        'Me.dagLogAktywnosci.RowHeaderWidth = 80 'use if no tablestyle
        '--------------------------------------------------------------------
        'set topleft corner point so we can locate the toprow
        If DataSetLogAktywnosci.Tables(0).Rows.Count > 0 Then
            StartPointInCell00 = New System.Drawing.Point((dagLogAktywnosci.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (dagLogAktywnosci.GetCellDisplayRectangle(0, 0, True).Y + 4))
            ScrollXOffset = 10000 * dagLogAktywnosci.FirstDisplayedScrollingColumnIndex +
                        dagLogAktywnosci.GetCellDisplayRectangle(dagLogAktywnosci.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
        Else
            'brak zapisow - poczatek DataGrid
            StartPointInCell00 = New System.Drawing.Point((dagLogAktywnosci.Left + 4), (dagLogAktywnosci.Top + 4))
            ScrollXOffset = 0
        End If
        '--------------------------------------------------
        'inicjowanie zegara dla scrollingu poziomego
        HorizScrollLastTime = ""
        HorizScrollTimer.Interval = 2 * 1000
        HorizScrollTimer.Enabled = True
        '--------------------------------------------------
        dagLogAktywnosci.ContextMenuStrip = Me.menuEdytuj
        cmdStop.Text = "Powrót"
        menuEdytuj.Visible = True
        menuEdytuj.Enabled = True
        '--------------------------------------------------
        'Stworz zmienne do filtrowania
        Dim ii As Integer = 0
        For ii = 0 To DataViewLogAktywnosci.Table.Columns.Count - 1
            'stworz zmienn¹ TXTBOX
            Dim CtrlT As New System.Windows.Forms.TextBox
            CtrlT.Name = "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
            CtrlT.Visible = False
            Me.Panel2.Controls.Add(CtrlT)
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
            Me.Panel2.Controls.Add(CtrlB)
            AddHandler CtrlB.Click, New EventHandler(AddressOf cmdFiXX_Click)
            AddHandler CtrlB.GotFocus, New EventHandler(AddressOf cmdFiXX_GotFocus)
            AddHandler CtrlB.LostFocus, New EventHandler(AddressOf cmdFiXX_LostFocus)
        Next
        Me.Refresh()
        '--------------------------------------------------
        StartFormy = False
        InicjujFiltryColumnDGV(Me.Panel2, dagLogAktywnosci, DataViewLogAktywnosci.Table.Columns.Count, stbFiltry, StartFormy)
        dagLogAktywnosci_CurrentCellChanged(sender, e)
    End Sub

    Private Sub frmKatalogBankow_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdDoExcela_Click(sender As Object, e As EventArgs) Handles cmdDoExcela.Click
        'Me.cmdDoExcela.Image = My.Resources.Excel_16
        Me.Cursor = Cursors.WaitCursor
        ExportDataGrid2Excel(dagLogAktywnosci, DataViewLogAktywnosci, Me.Text)
        Me.Cursor = Cursors.Default
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
                If dagLogAktywnosci.RowCount > 0 Then
                    If ScrollXOffset <> 10000 * dagLogAktywnosci.FirstDisplayedScrollingColumnIndex +
                                    dagLogAktywnosci.GetCellDisplayRectangle(dagLogAktywnosci.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X Then

                        PokazFiltryColumnDGV(Me.Panel2, dagLogAktywnosci, DataViewLogAktywnosci.Table.Columns.Count, stbFiltry, StartFormy)
                        '-----------------------------
                        ScrollXOffset = 10000 * dagLogAktywnosci.FirstDisplayedScrollingColumnIndex +
                                    dagLogAktywnosci.GetCellDisplayRectangle(dagLogAktywnosci.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X
                        HorizScrollLastTime = ""
                    End If
                Else
                    PokazFiltryColumnDGV(Me.Panel2, dagLogAktywnosci, DataViewLogAktywnosci.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Else
            End If
        End If
        HorizScrollTimer.Enabled = True
    End Sub


    '***************************************************
    ' zmiana komorki...
    '***************************************************

    Private Sub dagLogAktywnosci_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagLogAktywnosci.CurrentCellChanged
        'If Not StartFormy Then
        '    If dagLogAktywnosci.CurrentCell Is Nothing Then
        '        stbLogAktywnosci.Panels(1).Text = "Wi: "
        '        stbLogAktywnosci.Panels(2).Text = "Ko: "
        '    Else
        '        stbLogAktywnosci.Panels(1).Text = "Wi: " & dagLogAktywnosci.CurrentCell.RowIndex.ToString
        '        stbLogAktywnosci.Panels(2).Text = "Ko: " & dagLogAktywnosci.CurrentCell.ColumnIndex.ToString
        '        ''-----------------------------------------------
        '        If DataViewLogAktywnosci.Count = 0 Then
        '        Else
        '        End If
        '    End If
        'End If

        If Not StartFormy Then
            If dagLogAktywnosci.CurrentCell Is Nothing Then
                stbLogAktywnosci.Panels(1).Text = "Wi: "
                stbLogAktywnosci.Panels(2).Text = "Ko: "

                lblDataZapisu.Text = ""
                lblKatalog.Text = ""
                lblTemat.Text = ""
                lblOperacja.Text = ""
                lblPracownik.Text = ""
                lblUwagi.Text = ""
            Else
                stbLogAktywnosci.Panels(1).Text = "Wi: " & dagLogAktywnosci.CurrentCell.RowIndex.ToString
                stbLogAktywnosci.Panels(2).Text = "Ko: " & dagLogAktywnosci.CurrentCell.ColumnIndex.ToString
                ''-----------------------------------------------
                If DataViewLogAktywnosci.Count = 0 Then
                    lblDataZapisu.Text = ""
                    lblKatalog.Text = ""
                    lblTemat.Text = ""
                    lblOperacja.Text = ""
                    lblPracownik.Text = ""
                    lblUwagi.Text = ""
                Else
                    'Public LG_UniqID As String       'UNIQID     STRING 40
                    'Public LG_Temat As String        'TEMAT    STRING 100
                    'Public LG_Katalog As String      'KATALOG    STRING 100
                    'Public LG_DataZapisu As String   'DATAZAPISU STRING 30
                    'Public LG_Pracownik As String    'PRACOWNIK  STRING 10
                    'Public LG_Operacja As String     'OPERACJA   STRING 20
                    'Public LG_Uwagi As String        'UWAGI      Text
                    'Public LG_Wersja As Integer      'WERSJA     Integer

                    lblDataZapisu.Text = GetTxtField(dagLogAktywnosci, 3)
                    lblKatalog.Text = GetTxtField(dagLogAktywnosci, 2)
                    lblTemat.Text = GetTxtField(dagLogAktywnosci, 1)
                    lblOperacja.Text = GetTxtField(dagLogAktywnosci, 5)
                    lblPracownik.Text = GetTxtField(dagLogAktywnosci, 4)
                    lblUwagi.Text = GetTxtField(dagLogAktywnosci, 6)
                End If
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
                    PrzejdzDoDGV(dagLogAktywnosci, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub

    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            FiltrujDataViewDGV(Me.Panel2, dagLogAktywnosci, DataViewLogAktywnosci.Table.Columns.Count, DataViewLogAktywnosci, stbLogAktywnosci)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me.Panel2, dagLogAktywnosci, Mid(sender.name, 6, 3), "LogAktywnosci")
    End Sub
    Private Sub cmdFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiFiltraCmd(Me.Panel2, sender)
    End Sub
    Private Sub cmdFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraCmd(Me.Panel2, sender)
    End Sub


    '*********************************************************
    '***** obsluga dag
    '*********************************************************

    Private Sub dagLogAktywnosci_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagLogAktywnosci.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me.Panel2, dagLogAktywnosci, DataViewLogAktywnosci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagLogAktywnosci_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagLogAktywnosci.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me.Panel2, dagLogAktywnosci, DataViewLogAktywnosci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub


    Private Sub dagLogAktywnosci_Scroll(sender As Object, e As ScrollEventArgs) Handles dagLogAktywnosci.Scroll
        'If (Not StartFormy) And (DataViewLogAktywnosci.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me.Panel2, dagLogAktywnosci, DataViewLogAktywnosci.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me.Panel2, dagLogAktywnosci, DataViewLogAktywnosci.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If


        If (Not StartFormy) And (DataViewLogAktywnosci.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If
    End Sub

    Private Sub frmKatalogBankow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Dim rc As DataGridViewSelectedRowCollection = dagLogAktywnosci.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                dagLogAktywnosci.Rows(nextrow).Selected = True
            End If
        End If
    End Sub

    Private Sub dagLogAktywnosci_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagLogAktywnosci.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me.Panel2, dagLogAktywnosci, DataViewLogAktywnosci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagLogAktywnosci_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagLogAktywnosci.MouseMove
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
                '    PokazFiltryColumnDGV(Me.Panel2, dagLogAktywnosci, DataViewLogAktywnosci.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me.Panel2, dagLogAktywnosci, DataViewLogAktywnosci.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagLogAktywnosci_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagLogAktywnosci.MouseUp
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
                stbLogAktywnosci.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbLogAktywnosci.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagLogAktywnosci(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbLogAktywnosci.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbLogAktywnosci.Panels(3).Text = "Sort: " & DataSetLogAktywnosci.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbLogAktywnosci.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagLogAktywnosci(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbLogAktywnosci.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbLogAktywnosci.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagLogAktywnosci_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagLogAktywnosci.DoubleClick
        If _OCoMamRobic = "WYBIERAJ" Then
            cmdStop_Click(sender, e)
        Else
            cmdEdytuj_Click(sender, e)
        End If
    End Sub

    Private Sub dagLogAktywnosci_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagLogAktywnosci.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                If _OCoMamRobic = "WYBIERAJ" Then
                    cmdStop_Click(sender, e)
                Else
                    cmdEdytuj_Click(sender, e)
                End If
            Case Keys.Insert, Keys.Add
                cmdDodaj_Click(sender, e)
            Case Keys.Delete, Keys.Subtract
                cmdUsuñ_Click(sender, e)
            Case Else
        End Select
    End Sub




    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagLogAktywnosci_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagLogAktywnosci.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagLogAktywnosci.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagLogAktywnosci.DefaultCellStyle.Font,
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
            PokazFiltryColumnDGV(Me.Panel2, dagLogAktywnosci, DataViewLogAktywnosci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub




    '*************************************************
    '*** przenoszenie danych miêdzy rekordem i zmiennymi
    '*************************************************

    'Public LG_UniqID As String       'UNIQID     STRING 40
    'Public LG_Temat As String        'TEMAT    STRING 100
    'Public LG_Katalog As String      'KATALOG    STRING 100
    'Public LG_DataZapisu As String   'DATAZAPISU STRING 30
    'Public LG_Pracownik As String    'PRACOWNIK  STRING 10
    'Public LG_Operacja As String     'OPERACJA   STRING 20
    'Public LG_Uwagi As String        'UWAGI      Text
    'Public LG_Wersja As Integer      'WERSJA     Integer

    Private Sub PobierzLogAktywnosci()
        'pobierz wartosci przed aktualizacja
        LG_UniqID = GetTxtField(dagLogAktywnosci, 0)       'UNIQID     STRING 40
        LG_Temat = GetTxtField(dagLogAktywnosci, 1)        'TEMAT      STRING 100
        LG_Katalog = GetTxtField(dagLogAktywnosci, 2)      'KATALOG    STRING 100
        LG_DataZapisu = GetTxtField(dagLogAktywnosci, 3)   'DATAZAPISU STRING 30
        LG_Pracownik = GetTxtField(dagLogAktywnosci, 4)    'PRACOWNIK  STRING 10
        LG_Operacja = GetTxtField(dagLogAktywnosci, 5)     'OPERACJA   STRING 20
        LG_Uwagi = GetTxtField(dagLogAktywnosci, 6)        'UWAGI      Text
        LG_Wersja = GetIntField(dagLogAktywnosci, 11)      'WERSJA     Integer
    End Sub

    '**************************************************************
    '* obsluga komend
    '**************************************************************

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        If DataSetLogAktywnosci.Tables(0).Rows.Count > 0 Then
            PobierzLogAktywnosci()
            'oOsobaKontaktowa = Trim(GetTxtField(dagLogAktywnosci, 0))
        Else
            'oOsobaKontaktowa = ""
        End If
        Me.Close()
    End Sub


    Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytuj.Click
        'If DataSetLogAktywnosci.Tables(0).Rows.Count <= 0 Then
        '    'Raiseevent(Dodaj,Click )
        '    MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
        '        "Proszê najpierw dopisaæ rekord do tabeli...",
        '        "Uwaga :",
        '        System.Windows.Forms.MessageBoxButtons.OK,
        '        MessageBoxIcon.Information,
        '        MessageBoxDefaultButton.Button1)
        'Else
        '    If dagLogAktywnosci.CurrentCell Is Nothing Then
        '        MessageBox.Show("Nie potrafie edytowaæ niezdefiniowanego zapisu" & vbCrLf &
        '            "Proszê najpierw wybraæ rekord z tabeli...",
        '            "Uwaga :",
        '            System.Windows.Forms.MessageBoxButtons.OK,
        '            MessageBoxIcon.Information,
        '            MessageBoxDefaultButton.Button1)
        '    Else
        '        oOperacja = "EDYTUJ"
        '        PobierzLogAktywnosci()
        '        Dim Form As New frmEdytujKatalogLogAktywnosci
        '        oAktualizuj = False
        '        Form.ShowDialog()
        '        If Not oAktualizuj Then
        '        Else
        '            Dim foundRow As DataRow
        '            ' Create an array for the key values to find.
        '            Dim findTheseVals(0) As Object
        '            findTheseVals(0) = LG_UniqID
        '            foundRow = DataSetLogAktywnosci.Tables(0).Rows.Find(findTheseVals)
        '            ' sprawdz czy jest...
        '            If Not (foundRow Is Nothing) Then
        '                'Public LG_UniqID As String       'UNIQID     STRING 40
        '                'Public LG_Temat As String        'TEMAT    STRING 100
        '                'Public LG_Katalog As String      'KATALOG    STRING 100
        '                'Public LG_DataZapisu As String   'DATAZAPISU STRING 30
        '                'Public LG_Pracownik As String    'PRACOWNIK  STRING 10
        '                'Public LG_Operacja As String     'OPERACJA   STRING 20
        '                'Public LG_Uwagi As String        'UWAGI      Text
        '                'Public LG_Wersja As Integer      'WERSJA     Integer

        '                'foundRow("UNIQID") = lg_UniqID
        '                foundRow("TEMAT") = LG_Temat
        '                foundRow("KATALOG") = LG_Katalog
        '                foundRow("DATAZAPISU") = LG_DataZapisu
        '                foundRow("PRACOWNIK") = LG_Pracownik
        '                foundRow("OPERACJA") = LG_Operacja
        '                foundRow("UWAGI") = LG_Uwagi
        '                foundRow("WERSJA") = ZmienWersje(LG_Wersja)    'Wersja         Integer, 2

        '                'wyswietl ilosc rekordow
        '                stbLogAktywnosci.Panels(0).Text = "Iloœæ rec.: " & DataSetLogAktywnosci.Tables(0).Rows.Count.ToString
        '                AktualizujBaze()
        '                'aktualizuj DataGrid
        '                dagLogAktywnosci.Update()
        '            End If
        '        End If
        '    End If
        'End If
    End Sub



    Private Sub cmdUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñ.Click
        'If DataSetLogAktywnosci.Tables(0).Rows.Count <= 0 Then
        '    'Raiseevent(Dodaj,Click )
        '    MessageBox.Show("Nie potrafie usun¹æ nieistniej¹cego zapisu" & vbCrLf &
        '        "Proszê najpierw dopisaæ rekord do tabeli...",
        '        "Uwaga :",
        '        System.Windows.Forms.MessageBoxButtons.OK,
        '        MessageBoxIcon.Information,
        '        MessageBoxDefaultButton.Button1)
        'Else
        '    If dagLogAktywnosci.CurrentCell Is Nothing Then
        '        MessageBox.Show("Nie potrafie usun¹æ niezdefiniowanego zapisu" & vbCrLf &
        '            "Proszê najpierw wybraæ rekord z tabeli...",
        '            "Uwaga :",
        '            System.Windows.Forms.MessageBoxButtons.OK,
        '            MessageBoxIcon.Information,
        '            MessageBoxDefaultButton.Button1)
        '    Else
        '        If MessageBox.Show("Czy na pewno mam usun¹æ ten zapis ?",
        '                    "Proszê o potwierdzenie :",
        '                    System.Windows.Forms.MessageBoxButtons.YesNo,
        '                    MessageBoxIcon.Question,
        '                    MessageBoxDefaultButton.Button1) = DialogResult.Yes Then
        '            oOperacja = "USUN"
        '            PobierzLogAktywnosci()

        '            Dim foundRow As DataRow
        '            ' Create an array for the key values to find.
        '            Dim findTheseVals(0) As Object
        '            findTheseVals(0) = LG_UniqID
        '            foundRow = DataSetLogAktywnosci.Tables(0).Rows.Find(findTheseVals)
        '            ' sprawdz czy jest...
        '            If Not (foundRow Is Nothing) Then
        '                foundRow.Delete()
        '                'wyswietl ilosc rekordow
        '                stbLogAktywnosci.Panels(0).Text = "Iloœæ rec.: " & DataSetLogAktywnosci.Tables(0).Rows.Count.ToString
        '                AktualizujBaze()
        '            Else
        '            End If
        '        End If
        '    End If
        'End If
    End Sub


    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        'oOperacja = "DODAJ"
        'InitLogAktywnosci()
        'Dim Form As New frmEdytujKatalogLogAktywnosci
        'Do While True
        '    oAktualizuj = False
        '    Form.ShowDialog()
        '    If Not oAktualizuj Then
        '        Exit Do
        '    Else
        '        Dim foundRow As DataRow
        '        Dim NewRow As DataRow
        '        ' Create an array for the key values to find.
        '        Dim findTheseVals(0) As Object
        '        findTheseVals(0) = LG_UniqID
        '        foundRow = DataSetLogAktywnosci.Tables(0).Rows.Find(findTheseVals)
        '        ' sprawdz czy jest...
        '        If Not (foundRow Is Nothing) Then
        '            MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
        '                "Ident.Unikalny = " & LG_UniqID,
        '                "Uwaga :",
        '                System.Windows.Forms.MessageBoxButtons.OK,
        '                MessageBoxIcon.Information,
        '                MessageBoxDefaultButton.Button1)
        '        Else
        '            'stworz nowy rekord
        '            NewRow = DataSetLogAktywnosci.Tables(0).NewRow()

        '            NewRow("UNIQID") = LG_UniqID
        '            NewRow("TEMAT") = LG_Temat
        '            NewRow("KATALOG") = LG_Katalog
        '            NewRow("DATAZAPISU") = LG_DataZapisu
        '            NewRow("PRACOWNIK") = LG_Pracownik
        '            NewRow("OPERACJA") = LG_Operacja
        '            NewRow("UWAGI") = LG_Uwagi
        '            NewRow("WERSJA") = 1                     'Wersja         Integer, 2

        '            'dodaj rekord do DataSet
        '            DataSetLogAktywnosci.Tables(0).Rows.Add(NewRow)

        '            'wyswietl ilosc rekordow
        '            stbLogAktywnosci.Panels(0).Text = "Iloœæ rec.: " & DataSetLogAktywnosci.Tables(0).Rows.Count.ToString
        '            AktualizujBaze()
        '            'aktualizuj DataGrid
        '            dagLogAktywnosci.Update()
        '            '---------------------------------
        '            'ustaw sie na ostatnim zapisie
        '            If DataViewLogAktywnosci.Count > 0 Then
        '                dagLogAktywnosci.FirstDisplayedScrollingRowIndex = dagLogAktywnosci.RowCount - 1
        '                dagLogAktywnosci.PerformLayout()

        '                dagLogAktywnosci.CurrentCell = dagLogAktywnosci(1, DataViewLogAktywnosci.Count - 1)
        '                'dagLogAktywnosci.CurrentCell.Selected = True
        '                dagLogAktywnosci.Select()
        '            End If
        '            Exit Do
        '        End If
        '    End If
        'Loop
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
            '        dbDataAdapter.Update(DataSetLogAktywnosci, "TABELA_LogAktywnosci")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapter.Fill(DataSetLogAktywnosci)
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
                    sqlDataAdapter.Update(DataSetLogAktywnosci, "TABELA_LogAktywnosci")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapter.Fill(DataSetLogAktywnosci)
                End If
                sqlConnection.Close()
            End If
        End If
    End Sub



    Private Sub OdswiezBaze()
        DataSetLogAktywnosci.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnection.Open()
            '    If dbConnection.State = ConnectionState.Open Then
            '        dbDataAdapter.Fill(DataSetLogAktywnosci)
            '        dbConnection.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnection.Open()
                If sqlConnection.State = ConnectionState.Open Then
                    sqlDataAdapter.Fill(DataSetLogAktywnosci)
                    sqlConnection.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub




    '****************************************
    ' filtrowanie 
    '****************************************

    Private Sub cmdWszystko_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWszystko.Click
        Me.Cursor = Cursors.WaitCursor
        oLokalnyFiltr = ""
        Try
            StartFormy = True
            CzyscFiltryDGV(Me.Panel2, dagLogAktywnosci, DataViewLogAktywnosci.Table.Columns.Count, DataViewLogAktywnosci, stbLogAktywnosci)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbLogAktywnosci.Panels(0).Text = "Il.zap.: " & DataViewLogAktywnosci.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub



    '****************************************
    '** Kolorowanie
    '****************************************

    Private Sub dagLogAktywnosci_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagLogAktywnosci.CellFormatting
        If Not StartFormy Then
            'Public LG_UniqID As String       'UNIQID     STRING 40
            'Public LG_Temat As String        'TEMAT    STRING 100
            'Public LG_Katalog As String      'KATALOG    STRING 100
            'Public LG_DataZapisu As String   'DATAZAPISU STRING 30
            'Public LG_Pracownik As String    'PRACOWNIK  STRING 10
            'Public LG_Operacja As String     'OPERACJA   STRING 20
            'Public LG_Uwagi As String        'UWAGI      Text
            'Public LG_Wersja As Integer      'WERSJA     Integer

            Dim Oper As String = GetTxtField(dagLogAktywnosci, e.RowIndex, 5)
            Select Case Oper
                Case LOG_OperacjaDodaj
                    e.CellStyle.ForeColor = Color.Green
                Case LOG_OperacjaEdytuj
                    e.CellStyle.ForeColor = Color.Navy
                Case LOG_OperacjaUsun
                    e.CellStyle.ForeColor = Color.Red
                Case Else
                    e.CellStyle.ForeColor = Color.Navy
            End Select

            'Dim Wal As String = GetTxtField(dagLogAktywnosci, e.RowIndex, 0)
            'Select Case Wal
            '    Case "EUR"
            '        e.CellStyle.BackColor = Color.White
            '    Case Else
            'End Select
        End If
    End Sub







    '****************************************
    '** obsluga Menu Kontekstowego
    '****************************************

    Private Sub menuEDodaj_Click(sender As Object, e As EventArgs) Handles menuEDodaj.Click
        cmdDodaj_Click(sender, e)
    End Sub

    Private Sub menuEUsuñ_Click(sender As Object, e As EventArgs) Handles menuEUsuñ.Click
        cmdUsuñ_Click(sender, e)
    End Sub

    Private Sub menuEEdytuj_Click(sender As Object, e As EventArgs) Handles menuEEdytuj.Click
        cmdEdytuj_Click(sender, e)
    End Sub

    Private Sub MnuEdytuj_Poka¿WJednejLinii_Click(sender As Object, e As EventArgs) Handles menuEPoka¿WJednejLinii.Click
        'dagLogAktywnosci.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        dagLogAktywnosci.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagLogAktywnosci.Refresh()
    End Sub

    Private Sub mnuEdytuj_Poka¿WWieluLiniach_Click(sender As Object, e As EventArgs) Handles menuEPoka¿WWieluLiniach.Click
        'dagLogAktywnosci.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        dagLogAktywnosci.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagLogAktywnosci.Refresh()
    End Sub


    Private Sub menuWOdswiez_Click(sender As Object, e As EventArgs)
        OdswiezBaze()
    End Sub

    Private Sub menuEOdswiez_Click(sender As Object, e As EventArgs) Handles menuEOdswiez.Click
        OdswiezBaze()
    End Sub



    ''------------------------
    '' obs³uga selekcji panelu...
    ''--------------------------------

    'Dim BufColor As System.Drawing.Color = Nothing

    'Private Sub dagLogAktywnosci_LostFocus(sender As Object, e As EventArgs) Handles dagLogAktywnosci.LostFocus
    '    dagLogAktywnosci.EnableHeadersVisualStyles = False
    '    dagLogAktywnosci.ColumnHeadersDefaultCellStyle.BackColor = SoftartTabHeaderUnSelectedColor
    'End Sub

    'Private Sub dagLogAktywnosci_GotFocus(sender As Object, e As EventArgs) Handles dagLogAktywnosci.GotFocus
    '    dagLogAktywnosci.EnableHeadersVisualStyles = False
    '    BufColor = dagLogAktywnosci.ColumnHeadersDefaultCellStyle.BackColor
    '    dagLogAktywnosci.ColumnHeadersDefaultCellStyle.BackColor = SoftartTabHeaderSelectedColor
    'End Sub


End Class
