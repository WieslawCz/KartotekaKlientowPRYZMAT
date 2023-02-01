'Public Delegate Sub delegateAktualizuj()

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AkcjaMarketingowaWybierzKilka
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

    Dim _dsKlenci As DataSet
    Dim _AktualizujKartKlientow As delegateAktualizuj

    Public Sub New(ByVal ds As DataSet, _
                   ByVal Aktualizuj As delegateAktualizuj)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _dsKlenci = ds
        _AktualizujKartKlientow = Aktualizuj
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AkcjaMarketingowaWybierzKilka))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdWybierzPodbranze = New System.Windows.Forms.Button()
        Me.cmdWybierzBranze = New System.Windows.Forms.Button()
        Me.txtPodbranza = New System.Windows.Forms.TextBox()
        Me.LabelPodBranza = New System.Windows.Forms.Label()
        Me.txtBranza = New System.Windows.Forms.TextBox()
        Me.LabelBranza = New System.Windows.Forms.Label()
        Me.rbtDefiniujBranze = New System.Windows.Forms.RadioButton()
        Me.dtpOdDaty = New System.Windows.Forms.DateTimePicker()
        Me.dtpDoDaty = New System.Windows.Forms.DateTimePicker()
        Me.lblOdDaty = New System.Windows.Forms.Label()
        Me.lblDoDaty = New System.Windows.Forms.Label()
        Me.rbtNieObsluzeni = New System.Windows.Forms.RadioButton()
        Me.rbtObsluzeni = New System.Windows.Forms.RadioButton()
        Me.rbtWszyscy = New System.Windows.Forms.RadioButton()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.cmdZapamietaj = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.rbtWybierzWspolnych = New System.Windows.Forms.RadioButton()
        Me.rbtWybierzWszystkich = New System.Windows.Forms.RadioButton()
        Me.dagAkcjeOpis = New System.Windows.Forms.DataGrid()
        Me.stbAkcjeOpis = New System.Windows.Forms.StatusBar()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        CType(Me.dagAkcjeOpis, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.cmdWybierzPodbranze)
        Me.Panel1.Controls.Add(Me.cmdWybierzBranze)
        Me.Panel1.Controls.Add(Me.txtPodbranza)
        Me.Panel1.Controls.Add(Me.LabelPodBranza)
        Me.Panel1.Controls.Add(Me.txtBranza)
        Me.Panel1.Controls.Add(Me.LabelBranza)
        Me.Panel1.Controls.Add(Me.rbtDefiniujBranze)
        Me.Panel1.Controls.Add(Me.dtpOdDaty)
        Me.Panel1.Controls.Add(Me.dtpDoDaty)
        Me.Panel1.Controls.Add(Me.lblOdDaty)
        Me.Panel1.Controls.Add(Me.lblDoDaty)
        Me.Panel1.Controls.Add(Me.rbtNieObsluzeni)
        Me.Panel1.Controls.Add(Me.rbtObsluzeni)
        Me.Panel1.Controls.Add(Me.rbtWszyscy)
        Me.Panel1.Location = New System.Drawing.Point(8, 286)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(620, 174)
        Me.Panel1.TabIndex = 0
        '
        'cmdWybierzPodbranze
        '
        Me.cmdWybierzPodbranze.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzPodbranze.Image = CType(resources.GetObject("cmdWybierzPodbranze.Image"), System.Drawing.Image)
        Me.cmdWybierzPodbranze.Location = New System.Drawing.Point(554, 141)
        Me.cmdWybierzPodbranze.Name = "cmdWybierzPodbranze"
        Me.cmdWybierzPodbranze.Size = New System.Drawing.Size(32, 23)
        Me.cmdWybierzPodbranze.TabIndex = 297
        Me.cmdWybierzPodbranze.Text = "2"
        '
        'cmdWybierzBranze
        '
        Me.cmdWybierzBranze.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzBranze.Image = CType(resources.GetObject("cmdWybierzBranze.Image"), System.Drawing.Image)
        Me.cmdWybierzBranze.Location = New System.Drawing.Point(554, 119)
        Me.cmdWybierzBranze.Name = "cmdWybierzBranze"
        Me.cmdWybierzBranze.Size = New System.Drawing.Size(32, 23)
        Me.cmdWybierzBranze.TabIndex = 296
        Me.cmdWybierzBranze.Text = "2"
        '
        'txtPodbranza
        '
        Me.txtPodbranza.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPodbranza.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtPodbranza.ForeColor = System.Drawing.Color.Purple
        Me.txtPodbranza.Location = New System.Drawing.Point(123, 143)
        Me.txtPodbranza.Name = "txtPodbranza"
        Me.txtPodbranza.Size = New System.Drawing.Size(463, 20)
        Me.txtPodbranza.TabIndex = 295
        '
        'LabelPodBranza
        '
        Me.LabelPodBranza.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelPodBranza.ForeColor = System.Drawing.Color.Navy
        Me.LabelPodBranza.Location = New System.Drawing.Point(25, 146)
        Me.LabelPodBranza.Name = "LabelPodBranza"
        Me.LabelPodBranza.Size = New System.Drawing.Size(183, 16)
        Me.LabelPodBranza.TabIndex = 294
        Me.LabelPodBranza.Text = "Podbran¿a . . . . . . ."
        '
        'txtBranza
        '
        Me.txtBranza.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBranza.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtBranza.ForeColor = System.Drawing.Color.Purple
        Me.txtBranza.Location = New System.Drawing.Point(123, 121)
        Me.txtBranza.Name = "txtBranza"
        Me.txtBranza.Size = New System.Drawing.Size(463, 20)
        Me.txtBranza.TabIndex = 293
        '
        'LabelBranza
        '
        Me.LabelBranza.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelBranza.ForeColor = System.Drawing.Color.Navy
        Me.LabelBranza.Location = New System.Drawing.Point(25, 124)
        Me.LabelBranza.Name = "LabelBranza"
        Me.LabelBranza.Size = New System.Drawing.Size(183, 16)
        Me.LabelBranza.TabIndex = 292
        Me.LabelBranza.Text = "Bran¿a . . . . . . . . . . ."
        '
        'rbtDefiniujBranze
        '
        Me.rbtDefiniujBranze.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rbtDefiniujBranze.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtDefiniujBranze.ForeColor = System.Drawing.Color.Navy
        Me.rbtDefiniujBranze.Location = New System.Drawing.Point(13, 99)
        Me.rbtDefiniujBranze.Name = "rbtDefiniujBranze"
        Me.rbtDefiniujBranze.Size = New System.Drawing.Size(485, 16)
        Me.rbtDefiniujBranze.TabIndex = 291
        Me.rbtDefiniujBranze.Text = "Wybierz klientów z Akcji i zdefiniuj dla nich Bran¿ê / Podbran¿ê"
        '
        'dtpOdDaty
        '
        Me.dtpOdDaty.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dtpOdDaty.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpOdDaty.CustomFormat = "yyyy-MM-dd"
        Me.dtpOdDaty.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpOdDaty.Location = New System.Drawing.Point(265, 76)
        Me.dtpOdDaty.Name = "dtpOdDaty"
        Me.dtpOdDaty.Size = New System.Drawing.Size(91, 20)
        Me.dtpOdDaty.TabIndex = 179
        Me.dtpOdDaty.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'dtpDoDaty
        '
        Me.dtpDoDaty.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dtpDoDaty.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpDoDaty.CustomFormat = "yyyy-MM-dd"
        Me.dtpDoDaty.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDoDaty.Location = New System.Drawing.Point(415, 76)
        Me.dtpDoDaty.Name = "dtpDoDaty"
        Me.dtpDoDaty.Size = New System.Drawing.Size(91, 20)
        Me.dtpDoDaty.TabIndex = 181
        Me.dtpDoDaty.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'lblOdDaty
        '
        Me.lblOdDaty.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblOdDaty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblOdDaty.ForeColor = System.Drawing.Color.Navy
        Me.lblOdDaty.Location = New System.Drawing.Point(29, 80)
        Me.lblOdDaty.Name = "lblOdDaty"
        Me.lblOdDaty.Size = New System.Drawing.Size(261, 16)
        Me.lblOdDaty.TabIndex = 180
        Me.lblOdDaty.Text = "Analizuj Kontakty Handlowe w okresie Od daty . . . . . . . . . "
        '
        'lblDoDaty
        '
        Me.lblDoDaty.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblDoDaty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblDoDaty.ForeColor = System.Drawing.Color.Navy
        Me.lblDoDaty.Location = New System.Drawing.Point(366, 80)
        Me.lblDoDaty.Name = "lblDoDaty"
        Me.lblDoDaty.Size = New System.Drawing.Size(56, 16)
        Me.lblDoDaty.TabIndex = 182
        Me.lblDoDaty.Text = "Do daty . . . . . . . . . "
        '
        'rbtNieObsluzeni
        '
        Me.rbtNieObsluzeni.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rbtNieObsluzeni.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtNieObsluzeni.ForeColor = System.Drawing.Color.Navy
        Me.rbtNieObsluzeni.Location = New System.Drawing.Point(13, 56)
        Me.rbtNieObsluzeni.Name = "rbtNieObsluzeni"
        Me.rbtNieObsluzeni.Size = New System.Drawing.Size(601, 16)
        Me.rbtNieObsluzeni.TabIndex = 111
        Me.rbtNieObsluzeni.Text = "Analizuj wybranych klientów którym NIE ODNOTOWANO obs³ugê w kontaktach handlowych" &
    " (nieobs³u¿eni)"
        '
        'rbtObsluzeni
        '
        Me.rbtObsluzeni.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rbtObsluzeni.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtObsluzeni.ForeColor = System.Drawing.Color.Navy
        Me.rbtObsluzeni.Location = New System.Drawing.Point(13, 34)
        Me.rbtObsluzeni.Name = "rbtObsluzeni"
        Me.rbtObsluzeni.Size = New System.Drawing.Size(601, 16)
        Me.rbtObsluzeni.TabIndex = 110
        Me.rbtObsluzeni.Text = "Analizuj wybranych klientów którym odnotowano obs³ugê w kontaktach handlowych (ob" &
    "s³u¿eni)"
        '
        'rbtWszyscy
        '
        Me.rbtWszyscy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rbtWszyscy.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtWszyscy.ForeColor = System.Drawing.Color.Navy
        Me.rbtWszyscy.Location = New System.Drawing.Point(13, 12)
        Me.rbtWszyscy.Name = "rbtWszyscy"
        Me.rbtWszyscy.Size = New System.Drawing.Size(601, 16)
        Me.rbtWszyscy.TabIndex = 109
        Me.rbtWszyscy.Text = "Analizuj wszystkich wybranych klientów"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 466)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 40
        Me.PictureBox1.TabStop = False
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.ForeColor = System.Drawing.Color.Black
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(546, 467)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(80, 23)
        Me.cmdStop.TabIndex = 41
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdZapamietaj
        '
        Me.cmdZapamietaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdZapamietaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdZapamietaj.ForeColor = System.Drawing.Color.Black
        Me.cmdZapamietaj.Image = CType(resources.GetObject("cmdZapamietaj.Image"), System.Drawing.Image)
        Me.cmdZapamietaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdZapamietaj.Location = New System.Drawing.Point(281, 467)
        Me.cmdZapamietaj.Name = "cmdZapamietaj"
        Me.cmdZapamietaj.Size = New System.Drawing.Size(131, 23)
        Me.cmdZapamietaj.TabIndex = 42
        Me.cmdZapamietaj.Text = "Pobierz Klientów"
        Me.cmdZapamietaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdDodaj
        '
        Me.cmdDodaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDodaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDodaj.ForeColor = System.Drawing.Color.Black
        Me.cmdDodaj.Image = CType(resources.GetObject("cmdDodaj.Image"), System.Drawing.Image)
        Me.cmdDodaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDodaj.Location = New System.Drawing.Point(418, 467)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(122, 23)
        Me.cmdDodaj.TabIndex = 43
        Me.cmdDodaj.Text = "Dodaj Klientów"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.rbtWybierzWspolnych)
        Me.Panel2.Controls.Add(Me.rbtWybierzWszystkich)
        Me.Panel2.Location = New System.Drawing.Point(8, 226)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(618, 54)
        Me.Panel2.TabIndex = 44
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(10, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(576, 16)
        Me.Label1.TabIndex = 181
        Me.Label1.Text = "Okreœl których klientów wybraæ i których z wybranych analizowaæ :"
        '
        'rbtWybierzWspolnych
        '
        Me.rbtWybierzWspolnych.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtWybierzWspolnych.ForeColor = System.Drawing.Color.Navy
        Me.rbtWybierzWspolnych.Location = New System.Drawing.Point(316, 27)
        Me.rbtWybierzWspolnych.Name = "rbtWybierzWspolnych"
        Me.rbtWybierzWspolnych.Size = New System.Drawing.Size(286, 18)
        Me.rbtWybierzWspolnych.TabIndex = 110
        Me.rbtWybierzWspolnych.Text = "Wybierz klientów WSPÓLNYCH z oznaczonych Akcji"
        '
        'rbtWybierzWszystkich
        '
        Me.rbtWybierzWszystkich.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtWybierzWszystkich.ForeColor = System.Drawing.Color.Navy
        Me.rbtWybierzWszystkich.Location = New System.Drawing.Point(10, 27)
        Me.rbtWybierzWszystkich.Name = "rbtWybierzWszystkich"
        Me.rbtWybierzWszystkich.Size = New System.Drawing.Size(300, 18)
        Me.rbtWybierzWszystkich.TabIndex = 109
        Me.rbtWybierzWszystkich.Text = "Wybierz WSZYSTKICH klientów z oznaczonych Akcji"
        '
        'dagAkcjeOpis
        '
        Me.dagAkcjeOpis.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagAkcjeOpis.DataMember = ""
        Me.dagAkcjeOpis.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dagAkcjeOpis.Location = New System.Drawing.Point(8, 12)
        Me.dagAkcjeOpis.Name = "dagAkcjeOpis"
        Me.dagAkcjeOpis.Size = New System.Drawing.Size(618, 182)
        Me.dagAkcjeOpis.TabIndex = 46
        '
        'stbAkcjeOpis
        '
        Me.stbAkcjeOpis.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbAkcjeOpis.Dock = System.Windows.Forms.DockStyle.None
        Me.stbAkcjeOpis.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbAkcjeOpis.Location = New System.Drawing.Point(8, 198)
        Me.stbAkcjeOpis.Name = "stbAkcjeOpis"
        Me.stbAkcjeOpis.ShowPanels = True
        Me.stbAkcjeOpis.Size = New System.Drawing.Size(618, 22)
        Me.stbAkcjeOpis.TabIndex = 45
        Me.stbAkcjeOpis.Text = "StatusBar1"
        '
        'AkcjaMarketingowaWybierzKilka
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(638, 496)
        Me.Controls.Add(Me.dagAkcjeOpis)
        Me.Controls.Add(Me.stbAkcjeOpis)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.cmdZapamietaj)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "AkcjaMarketingowaWybierzKilka"
        Me.ShowInTaskbar = False
        Me.Text = "Wybór akcji marketingowej"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        CType(Me.dagAkcjeOpis, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents cmdZapamietaj As System.Windows.Forms.Button
    Friend WithEvents cmdDodaj As System.Windows.Forms.Button
    Friend WithEvents rbtNieObsluzeni As System.Windows.Forms.RadioButton
    Friend WithEvents rbtObsluzeni As System.Windows.Forms.RadioButton
    Friend WithEvents rbtWszyscy As System.Windows.Forms.RadioButton
    Friend WithEvents dtpOdDaty As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpDoDaty As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblOdDaty As System.Windows.Forms.Label
    Friend WithEvents lblDoDaty As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents rbtWybierzWspolnych As System.Windows.Forms.RadioButton
    Friend WithEvents rbtWybierzWszystkich As System.Windows.Forms.RadioButton
    Friend WithEvents dagAkcjeOpis As System.Windows.Forms.DataGrid
    Friend WithEvents stbAkcjeOpis As System.Windows.Forms.StatusBar
    Friend WithEvents cmdWybierzPodbranze As Button
    Friend WithEvents cmdWybierzBranze As Button
    Friend WithEvents txtPodbranza As TextBox
    Friend WithEvents LabelPodBranza As Label
    Friend WithEvents txtBranza As TextBox
    Friend WithEvents LabelBranza As Label
    Friend WithEvents rbtDefiniujBranze As RadioButton
End Class
