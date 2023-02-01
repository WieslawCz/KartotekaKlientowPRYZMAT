<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class GeneratorWydrukow
    Inherits System.Windows.Forms.Form

    Dim _DataSet As DataSet
    Dim _DataView As DataView

    Public Sub New(ByRef ParDS As DataSet, ByRef ParDV As DataView)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _DataSet = ParDS
        _DataView = ParDV
    End Sub

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(GeneratorWydrukow))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdEdytuj = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lblSzerokoscWydruku = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblSzerokoscArkusza = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtNazwa = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdDol = New System.Windows.Forms.Button()
        Me.cmdGora = New System.Windows.Forms.Button()
        Me.cmdUsun = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.ListaWydruku = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ListaPol = New System.Windows.Forms.ListView()
        Me.Nazwa = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Typ = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdZrezygnuj = New System.Windows.Forms.Button()
        Me.cmdZapisz = New System.Windows.Forms.Button()
        Me.cmdPobierz = New System.Windows.Forms.Button()
        Me.cmdCzysc = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.lblPath = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtSchemat = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.cmdEdytuj)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.lblSzerokoscWydruku)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.lblSzerokoscArkusza)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.txtNazwa)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.cmdDol)
        Me.Panel1.Controls.Add(Me.cmdGora)
        Me.Panel1.Controls.Add(Me.cmdUsun)
        Me.Panel1.Controls.Add(Me.cmdDodaj)
        Me.Panel1.Controls.Add(Me.ListaWydruku)
        Me.Panel1.Controls.Add(Me.ListaPol)
        Me.Panel1.Location = New System.Drawing.Point(8, 65)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(739, 406)
        Me.Panel1.TabIndex = 0
        '
        'cmdEdytuj
        '
        Me.cmdEdytuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdEdytuj.Image = CType(resources.GetObject("cmdEdytuj.Image"), System.Drawing.Image)
        Me.cmdEdytuj.Location = New System.Drawing.Point(311, 360)
        Me.cmdEdytuj.Name = "cmdEdytuj"
        Me.cmdEdytuj.Size = New System.Drawing.Size(43, 39)
        Me.cmdEdytuj.TabIndex = 71
        Me.cmdEdytuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label5.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(360, 73)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(363, 16)
        Me.Label5.TabIndex = 70
        Me.Label5.Text = "Kolejne pola na wydruku"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 73)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(297, 16)
        Me.Label3.TabIndex = 69
        Me.Label3.Text = "Poszczególne pola Katalogu Klientów (do wyboru)"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(0, 62)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(740, 1)
        Me.Panel2.TabIndex = 68
        '
        'lblSzerokoscWydruku
        '
        Me.lblSzerokoscWydruku.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSzerokoscWydruku.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblSzerokoscWydruku.ForeColor = System.Drawing.Color.Purple
        Me.lblSzerokoscWydruku.Location = New System.Drawing.Point(533, 36)
        Me.lblSzerokoscWydruku.Name = "lblSzerokoscWydruku"
        Me.lblSzerokoscWydruku.Size = New System.Drawing.Size(95, 16)
        Me.lblSzerokoscWydruku.TabIndex = 67
        Me.lblSzerokoscWydruku.Text = "0"
        Me.lblSzerokoscWydruku.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(357, 36)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(179, 16)
        Me.Label4.TabIndex = 66
        Me.Label4.Text = "Szerokoœæ wydruku  [mm]. . . . . . . . "
        '
        'lblSzerokoscArkusza
        '
        Me.lblSzerokoscArkusza.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSzerokoscArkusza.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblSzerokoscArkusza.ForeColor = System.Drawing.Color.Purple
        Me.lblSzerokoscArkusza.Location = New System.Drawing.Point(181, 35)
        Me.lblSzerokoscArkusza.Name = "lblSzerokoscArkusza"
        Me.lblSzerokoscArkusza.Size = New System.Drawing.Size(95, 16)
        Me.lblSzerokoscArkusza.TabIndex = 65
        Me.lblSzerokoscArkusza.Text = "0"
        Me.lblSzerokoscArkusza.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(5, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(179, 16)
        Me.Label2.TabIndex = 64
        Me.Label2.Text = "Szerokoœæ arkusza papieru [mm]. . . . . . . . "
        '
        'txtNazwa
        '
        Me.txtNazwa.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNazwa.ForeColor = System.Drawing.Color.Purple
        Me.txtNazwa.Location = New System.Drawing.Point(181, 12)
        Me.txtNazwa.Name = "txtNazwa"
        Me.txtNazwa.Size = New System.Drawing.Size(542, 20)
        Me.txtNazwa.TabIndex = 62
        Me.txtNazwa.Text = "Data"
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(5, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(130, 16)
        Me.Label1.TabIndex = 63
        Me.Label1.Text = "Nag³ówek zestawienia . . . . . . . . ."
        '
        'cmdDol
        '
        Me.cmdDol.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDol.Image = CType(resources.GetObject("cmdDol.Image"), System.Drawing.Image)
        Me.cmdDol.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDol.Location = New System.Drawing.Point(311, 246)
        Me.cmdDol.Name = "cmdDol"
        Me.cmdDol.Size = New System.Drawing.Size(43, 39)
        Me.cmdDol.TabIndex = 61
        Me.cmdDol.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdGora
        '
        Me.cmdGora.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdGora.Image = CType(resources.GetObject("cmdGora.Image"), System.Drawing.Image)
        Me.cmdGora.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdGora.Location = New System.Drawing.Point(311, 201)
        Me.cmdGora.Name = "cmdGora"
        Me.cmdGora.Size = New System.Drawing.Size(43, 39)
        Me.cmdGora.TabIndex = 60
        Me.cmdGora.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdUsun
        '
        Me.cmdUsun.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdUsun.Image = CType(resources.GetObject("cmdUsun.Image"), System.Drawing.Image)
        Me.cmdUsun.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUsun.Location = New System.Drawing.Point(311, 156)
        Me.cmdUsun.Name = "cmdUsun"
        Me.cmdUsun.Size = New System.Drawing.Size(43, 39)
        Me.cmdUsun.TabIndex = 59
        Me.cmdUsun.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdDodaj
        '
        Me.cmdDodaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDodaj.Image = CType(resources.GetObject("cmdDodaj.Image"), System.Drawing.Image)
        Me.cmdDodaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDodaj.Location = New System.Drawing.Point(311, 111)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(43, 39)
        Me.cmdDodaj.TabIndex = 58
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ListaWydruku
        '
        Me.ListaWydruku.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListaWydruku.BackColor = System.Drawing.SystemColors.Control
        Me.ListaWydruku.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6})
        Me.ListaWydruku.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.ListaWydruku.ForeColor = System.Drawing.Color.Purple
        Me.ListaWydruku.FullRowSelect = True
        Me.ListaWydruku.GridLines = True
        Me.ListaWydruku.HideSelection = False
        Me.ListaWydruku.Location = New System.Drawing.Point(361, 90)
        Me.ListaWydruku.MultiSelect = False
        Me.ListaWydruku.Name = "ListaWydruku"
        Me.ListaWydruku.Size = New System.Drawing.Size(363, 310)
        Me.ListaWydruku.TabIndex = 3
        Me.ListaWydruku.UseCompatibleStateImageBehavior = False
        Me.ListaWydruku.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Nazwa pola"
        Me.ColumnHeader1.Width = 200
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Typ"
        Me.ColumnHeader2.Width = 80
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Rozmiar"
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Nag³ówek"
        Me.ColumnHeader4.Width = 400
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Wyrównanie"
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Max iloœæ wierszy"
        '
        'ListaPol
        '
        Me.ListaPol.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ListaPol.BackColor = System.Drawing.SystemColors.Control
        Me.ListaPol.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Nazwa, Me.Typ})
        Me.ListaPol.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.ListaPol.ForeColor = System.Drawing.Color.Purple
        Me.ListaPol.FullRowSelect = True
        Me.ListaPol.GridLines = True
        Me.ListaPol.HideSelection = False
        Me.ListaPol.Location = New System.Drawing.Point(8, 90)
        Me.ListaPol.MultiSelect = False
        Me.ListaPol.Name = "ListaPol"
        Me.ListaPol.Size = New System.Drawing.Size(297, 310)
        Me.ListaPol.TabIndex = 2
        Me.ListaPol.UseCompatibleStateImageBehavior = False
        Me.ListaPol.View = System.Windows.Forms.View.Details
        '
        'Nazwa
        '
        Me.Nazwa.Text = "Nazwa pola"
        Me.Nazwa.Width = 200
        '
        'Typ
        '
        Me.Typ.Text = "Typ"
        Me.Typ.Width = 80
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(679, 479)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(68, 23)
        Me.cmdStop.TabIndex = 57
        Me.cmdStop.Text = "Drukuj"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 479)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(80, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 58
        Me.PictureBox1.TabStop = False
        '
        'cmdZrezygnuj
        '
        Me.cmdZrezygnuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdZrezygnuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdZrezygnuj.Image = CType(resources.GetObject("cmdZrezygnuj.Image"), System.Drawing.Image)
        Me.cmdZrezygnuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdZrezygnuj.Location = New System.Drawing.Point(577, 479)
        Me.cmdZrezygnuj.Name = "cmdZrezygnuj"
        Me.cmdZrezygnuj.Size = New System.Drawing.Size(96, 23)
        Me.cmdZrezygnuj.TabIndex = 59
        Me.cmdZrezygnuj.Text = "Zrezygnuj"
        Me.cmdZrezygnuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdZapisz
        '
        Me.cmdZapisz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdZapisz.Image = CType(resources.GetObject("cmdZapisz.Image"), System.Drawing.Image)
        Me.cmdZapisz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdZapisz.Location = New System.Drawing.Point(488, 479)
        Me.cmdZapisz.Name = "cmdZapisz"
        Me.cmdZapisz.Size = New System.Drawing.Size(83, 23)
        Me.cmdZapisz.TabIndex = 60
        Me.cmdZapisz.Text = "Zapisz"
        Me.cmdZapisz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPobierz
        '
        Me.cmdPobierz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPobierz.Image = CType(resources.GetObject("cmdPobierz.Image"), System.Drawing.Image)
        Me.cmdPobierz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPobierz.Location = New System.Drawing.Point(399, 479)
        Me.cmdPobierz.Name = "cmdPobierz"
        Me.cmdPobierz.Size = New System.Drawing.Size(83, 23)
        Me.cmdPobierz.TabIndex = 61
        Me.cmdPobierz.Text = "Pobierz"
        Me.cmdPobierz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdCzysc
        '
        Me.cmdCzysc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCzysc.Image = CType(resources.GetObject("cmdCzysc.Image"), System.Drawing.Image)
        Me.cmdCzysc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCzysc.Location = New System.Drawing.Point(306, 479)
        Me.cmdCzysc.Name = "cmdCzysc"
        Me.cmdCzysc.Size = New System.Drawing.Size(87, 23)
        Me.cmdCzysc.TabIndex = 62
        Me.cmdCzysc.Text = "Wyczyœæ"
        Me.cmdCzysc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.Button1)
        Me.Panel3.Controls.Add(Me.lblPath)
        Me.Panel3.Controls.Add(Me.Label11)
        Me.Panel3.Controls.Add(Me.txtSchemat)
        Me.Panel3.Controls.Add(Me.Label6)
        Me.Panel3.Location = New System.Drawing.Point(8, 12)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(739, 47)
        Me.Panel3.TabIndex = 64
        '
        'Button1
        '
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Location = New System.Drawing.Point(693, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(32, 20)
        Me.Button1.TabIndex = 30
        '
        'lblPath
        '
        Me.lblPath.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPath.ForeColor = System.Drawing.Color.Navy
        Me.lblPath.Location = New System.Drawing.Point(248, 12)
        Me.lblPath.Name = "lblPath"
        Me.lblPath.Size = New System.Drawing.Size(451, 20)
        Me.lblPath.TabIndex = 28
        '
        'Label11
        '
        Me.Label11.ForeColor = System.Drawing.Color.Navy
        Me.Label11.Location = New System.Drawing.Point(225, 15)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(24, 16)
        Me.Label11.TabIndex = 29
        Me.Label11.Text = "Plik z szablonem . . . . . . . . ."
        '
        'txtSchemat
        '
        Me.txtSchemat.ForeColor = System.Drawing.Color.Purple
        Me.txtSchemat.Location = New System.Drawing.Point(99, 12)
        Me.txtSchemat.Name = "txtSchemat"
        Me.txtSchemat.Size = New System.Drawing.Size(96, 20)
        Me.txtSchemat.TabIndex = 26
        Me.txtSchemat.Text = "Data"
        '
        'Label6
        '
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(7, 15)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(112, 16)
        Me.Label6.TabIndex = 27
        Me.Label6.Text = "Nazwa szablonu . . . . . . . . ."
        '
        'GeneratorWydrukow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(759, 506)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.cmdCzysc)
        Me.Controls.Add(Me.cmdPobierz)
        Me.Controls.Add(Me.cmdZapisz)
        Me.Controls.Add(Me.cmdZrezygnuj)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "GeneratorWydrukow"
        Me.ShowInTaskbar = False
        Me.Text = "Generator wydruków"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents ListaPol As System.Windows.Forms.ListView
    Friend WithEvents Nazwa As System.Windows.Forms.ColumnHeader
    Friend WithEvents Typ As System.Windows.Forms.ColumnHeader
    Friend WithEvents ListaWydruku As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdDol As System.Windows.Forms.Button
    Friend WithEvents cmdGora As System.Windows.Forms.Button
    Friend WithEvents cmdUsun As System.Windows.Forms.Button
    Friend WithEvents cmdDodaj As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents txtNazwa As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblSzerokoscArkusza As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents lblSzerokoscWydruku As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmdZrezygnuj As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdZapisz As System.Windows.Forms.Button
    Friend WithEvents cmdPobierz As System.Windows.Forms.Button
    Friend WithEvents cmdCzysc As System.Windows.Forms.Button
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents lblPath As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtSchemat As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cmdEdytuj As System.Windows.Forms.Button
End Class
