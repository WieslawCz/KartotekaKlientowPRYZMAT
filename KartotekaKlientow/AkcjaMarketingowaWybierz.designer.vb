Public Delegate Sub delegateAktualizuj()

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AkcjaMarketingowaWybierz
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AkcjaMarketingowaWybierz))
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
        Me.cmdAkcja = New System.Windows.Forms.Button()
        Me.txtIdent = New System.Windows.Forms.TextBox()
        Me.txtUwagi = New System.Windows.Forms.TextBox()
        Me.txtIleZap = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtData = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtOpis = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.cmdZapamietaj = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
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
        Me.Panel1.Controls.Add(Me.cmdAkcja)
        Me.Panel1.Controls.Add(Me.txtIdent)
        Me.Panel1.Controls.Add(Me.txtUwagi)
        Me.Panel1.Controls.Add(Me.txtIleZap)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.txtData)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.txtOpis)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(504, 459)
        Me.Panel1.TabIndex = 0
        '
        'cmdWybierzPodbranze
        '
        Me.cmdWybierzPodbranze.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzPodbranze.Image = CType(resources.GetObject("cmdWybierzPodbranze.Image"), System.Drawing.Image)
        Me.cmdWybierzPodbranze.Location = New System.Drawing.Point(455, 425)
        Me.cmdWybierzPodbranze.Name = "cmdWybierzPodbranze"
        Me.cmdWybierzPodbranze.Size = New System.Drawing.Size(32, 23)
        Me.cmdWybierzPodbranze.TabIndex = 290
        Me.cmdWybierzPodbranze.Text = "2"
        '
        'cmdWybierzBranze
        '
        Me.cmdWybierzBranze.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzBranze.Image = CType(resources.GetObject("cmdWybierzBranze.Image"), System.Drawing.Image)
        Me.cmdWybierzBranze.Location = New System.Drawing.Point(455, 403)
        Me.cmdWybierzBranze.Name = "cmdWybierzBranze"
        Me.cmdWybierzBranze.Size = New System.Drawing.Size(32, 23)
        Me.cmdWybierzBranze.TabIndex = 289
        Me.cmdWybierzBranze.Text = "2"
        '
        'txtPodbranza
        '
        Me.txtPodbranza.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPodbranza.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtPodbranza.ForeColor = System.Drawing.Color.Purple
        Me.txtPodbranza.Location = New System.Drawing.Point(120, 427)
        Me.txtPodbranza.Name = "txtPodbranza"
        Me.txtPodbranza.Size = New System.Drawing.Size(367, 20)
        Me.txtPodbranza.TabIndex = 288
        '
        'LabelPodBranza
        '
        Me.LabelPodBranza.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelPodBranza.ForeColor = System.Drawing.Color.Navy
        Me.LabelPodBranza.Location = New System.Drawing.Point(22, 430)
        Me.LabelPodBranza.Name = "LabelPodBranza"
        Me.LabelPodBranza.Size = New System.Drawing.Size(183, 16)
        Me.LabelPodBranza.TabIndex = 287
        Me.LabelPodBranza.Text = "Podbran¿a . . . . . . ."
        '
        'txtBranza
        '
        Me.txtBranza.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBranza.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtBranza.ForeColor = System.Drawing.Color.Purple
        Me.txtBranza.Location = New System.Drawing.Point(120, 405)
        Me.txtBranza.Name = "txtBranza"
        Me.txtBranza.Size = New System.Drawing.Size(367, 20)
        Me.txtBranza.TabIndex = 286
        '
        'LabelBranza
        '
        Me.LabelBranza.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelBranza.ForeColor = System.Drawing.Color.Navy
        Me.LabelBranza.Location = New System.Drawing.Point(22, 408)
        Me.LabelBranza.Name = "LabelBranza"
        Me.LabelBranza.Size = New System.Drawing.Size(183, 16)
        Me.LabelBranza.TabIndex = 285
        Me.LabelBranza.Text = "Bran¿a . . . . . . . . . . ."
        '
        'rbtDefiniujBranze
        '
        Me.rbtDefiniujBranze.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rbtDefiniujBranze.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtDefiniujBranze.ForeColor = System.Drawing.Color.Navy
        Me.rbtDefiniujBranze.Location = New System.Drawing.Point(10, 383)
        Me.rbtDefiniujBranze.Name = "rbtDefiniujBranze"
        Me.rbtDefiniujBranze.Size = New System.Drawing.Size(485, 16)
        Me.rbtDefiniujBranze.TabIndex = 183
        Me.rbtDefiniujBranze.Text = "Wybierz klientów z Akcji i zdefiniuj dla nich Bran¿ê / Podbran¿ê"
        '
        'dtpOdDaty
        '
        Me.dtpOdDaty.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dtpOdDaty.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpOdDaty.CustomFormat = "yyyy-MM-dd"
        Me.dtpOdDaty.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpOdDaty.Location = New System.Drawing.Point(246, 357)
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
        Me.dtpDoDaty.Location = New System.Drawing.Point(396, 357)
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
        Me.lblOdDaty.Location = New System.Drawing.Point(25, 361)
        Me.lblOdDaty.Name = "lblOdDaty"
        Me.lblOdDaty.Size = New System.Drawing.Size(228, 16)
        Me.lblOdDaty.TabIndex = 180
        Me.lblOdDaty.Text = "Analizuj Kontakty Handlowe Od daty . . . . . . . . . "
        '
        'lblDoDaty
        '
        Me.lblDoDaty.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblDoDaty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblDoDaty.ForeColor = System.Drawing.Color.Navy
        Me.lblDoDaty.Location = New System.Drawing.Point(347, 361)
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
        Me.rbtNieObsluzeni.Location = New System.Drawing.Point(10, 337)
        Me.rbtNieObsluzeni.Name = "rbtNieObsluzeni"
        Me.rbtNieObsluzeni.Size = New System.Drawing.Size(485, 16)
        Me.rbtNieObsluzeni.TabIndex = 111
        Me.rbtNieObsluzeni.Text = "Wybierz klientów z Akcji z którymi NIE ODNOTOWANO kontaktów handlowych (nieobs³u¿" &
    "eni)"
        '
        'rbtObsluzeni
        '
        Me.rbtObsluzeni.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rbtObsluzeni.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtObsluzeni.ForeColor = System.Drawing.Color.Navy
        Me.rbtObsluzeni.Location = New System.Drawing.Point(10, 315)
        Me.rbtObsluzeni.Name = "rbtObsluzeni"
        Me.rbtObsluzeni.Size = New System.Drawing.Size(485, 16)
        Me.rbtObsluzeni.TabIndex = 110
        Me.rbtObsluzeni.Text = "Wybierz klientów z Akcji z którymi odnotowano kontakty handlowe (obs³u¿eni)"
        '
        'rbtWszyscy
        '
        Me.rbtWszyscy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rbtWszyscy.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtWszyscy.ForeColor = System.Drawing.Color.Navy
        Me.rbtWszyscy.Location = New System.Drawing.Point(10, 293)
        Me.rbtWszyscy.Name = "rbtWszyscy"
        Me.rbtWszyscy.Size = New System.Drawing.Size(485, 16)
        Me.rbtWszyscy.TabIndex = 109
        Me.rbtWszyscy.Text = "Wybierz wszystkich klientów z Akcji"
        '
        'cmdAkcja
        '
        Me.cmdAkcja.Image = CType(resources.GetObject("cmdAkcja.Image"), System.Drawing.Image)
        Me.cmdAkcja.Location = New System.Drawing.Point(216, 6)
        Me.cmdAkcja.Name = "cmdAkcja"
        Me.cmdAkcja.Size = New System.Drawing.Size(32, 20)
        Me.cmdAkcja.TabIndex = 89
        '
        'txtIdent
        '
        Me.txtIdent.ForeColor = System.Drawing.Color.Purple
        Me.txtIdent.Location = New System.Drawing.Point(80, 6)
        Me.txtIdent.Name = "txtIdent"
        Me.txtIdent.Size = New System.Drawing.Size(144, 20)
        Me.txtIdent.TabIndex = 13
        '
        'txtUwagi
        '
        Me.txtUwagi.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUwagi.ForeColor = System.Drawing.Color.Purple
        Me.txtUwagi.Location = New System.Drawing.Point(80, 84)
        Me.txtUwagi.Multiline = True
        Me.txtUwagi.Name = "txtUwagi"
        Me.txtUwagi.ReadOnly = True
        Me.txtUwagi.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtUwagi.Size = New System.Drawing.Size(415, 203)
        Me.txtUwagi.TabIndex = 22
        '
        'txtIleZap
        '
        Me.txtIleZap.ForeColor = System.Drawing.Color.Purple
        Me.txtIleZap.Location = New System.Drawing.Point(80, 32)
        Me.txtIleZap.Name = "txtIleZap"
        Me.txtIleZap.ReadOnly = True
        Me.txtIleZap.Size = New System.Drawing.Size(80, 20)
        Me.txtIleZap.TabIndex = 20
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(7, 9)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 18)
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "Ident. akcji . . . . . . . . ."
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(7, 87)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 16)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "Uwagi . . . . . . . . ."
        '
        'txtData
        '
        Me.txtData.ForeColor = System.Drawing.Color.Purple
        Me.txtData.Location = New System.Drawing.Point(274, 32)
        Me.txtData.Name = "txtData"
        Me.txtData.ReadOnly = True
        Me.txtData.Size = New System.Drawing.Size(80, 20)
        Me.txtData.TabIndex = 15
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(7, 35)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 16)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Iloœæ zapisów . . . . . . . . ."
        '
        'txtOpis
        '
        Me.txtOpis.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOpis.ForeColor = System.Drawing.Color.Purple
        Me.txtOpis.Location = New System.Drawing.Point(80, 58)
        Me.txtOpis.Name = "txtOpis"
        Me.txtOpis.ReadOnly = True
        Me.txtOpis.Size = New System.Drawing.Size(415, 20)
        Me.txtOpis.TabIndex = 17
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(201, 35)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 16)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Data akcji . . . . . . . . ."
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(7, 61)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 16)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Opis akcji . . . . . . . . ."
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 473)
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
        Me.cmdStop.Location = New System.Drawing.Point(430, 474)
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
        Me.cmdZapamietaj.Location = New System.Drawing.Point(197, 474)
        Me.cmdZapamietaj.Name = "cmdZapamietaj"
        Me.cmdZapamietaj.Size = New System.Drawing.Size(115, 23)
        Me.cmdZapamietaj.TabIndex = 42
        Me.cmdZapamietaj.Text = "Pobierz akcje"
        Me.cmdZapamietaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdDodaj
        '
        Me.cmdDodaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDodaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDodaj.ForeColor = System.Drawing.Color.Black
        Me.cmdDodaj.Image = CType(resources.GetObject("cmdDodaj.Image"), System.Drawing.Image)
        Me.cmdDodaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDodaj.Location = New System.Drawing.Point(318, 474)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(106, 23)
        Me.cmdDodaj.TabIndex = 43
        Me.cmdDodaj.Text = "Dodaj akcje"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'AkcjaMarketingowaWybierz
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(522, 503)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.cmdZapamietaj)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "AkcjaMarketingowaWybierz"
        Me.ShowInTaskbar = False
        Me.Text = "Wybór akcji marketingowej"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents cmdAkcja As System.Windows.Forms.Button
    Friend WithEvents txtIdent As System.Windows.Forms.TextBox
    Friend WithEvents txtUwagi As System.Windows.Forms.TextBox
    Friend WithEvents txtIleZap As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtData As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtOpis As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents rbtDefiniujBranze As RadioButton
    Friend WithEvents cmdWybierzPodbranze As Button
    Friend WithEvents cmdWybierzBranze As Button
    Friend WithEvents txtPodbranza As TextBox
    Friend WithEvents LabelPodBranza As Label
    Friend WithEvents txtBranza As TextBox
    Friend WithEvents LabelBranza As Label
End Class
