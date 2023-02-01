'Public Delegate Sub delegateAktualizuj()

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AkcjaMarketingowaWybierzWgObrotow
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

    Dim _dsKlenci As DataSet
    Dim _CoAnalizowac As String
    Dim _AktualizujKartKlientow As delegateAktualizuj

    Public Sub New(ByVal ds As DataSet,
                   ByVal parCoAnalizowac As String,
                   ByVal Aktualizuj As delegateAktualizuj)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _dsKlenci = ds
        _CoAnalizowac = parCoAnalizowac
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AkcjaMarketingowaWybierzWgObrotow))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtKwotaMin = New System.Windows.Forms.TextBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.dtpOdDaty2 = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dtpDoDaty2 = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.dtpOdDaty1 = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.cbxJakSieZmienilo = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dtpDoDaty1 = New System.Windows.Forms.DateTimePicker()
        Me.lblOdDaty = New System.Windows.Forms.Label()
        Me.lblDoDaty = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.cmdZapamietaj = New System.Windows.Forms.Button()
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
        Me.Panel1.Controls.Add(Me.txtKwotaMin)
        Me.Panel1.Controls.Add(Me.Label26)
        Me.Panel1.Controls.Add(Me.dtpOdDaty2)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.dtpDoDaty2)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.dtpOdDaty1)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.ProgressBar1)
        Me.Panel1.Controls.Add(Me.cbxJakSieZmienilo)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.dtpDoDaty1)
        Me.Panel1.Controls.Add(Me.lblOdDaty)
        Me.Panel1.Controls.Add(Me.lblDoDaty)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(504, 130)
        Me.Panel1.TabIndex = 0
        '
        'txtKwotaMin
        '
        Me.txtKwotaMin.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtKwotaMin.ForeColor = System.Drawing.Color.Purple
        Me.txtKwotaMin.Location = New System.Drawing.Point(243, 79)
        Me.txtKwotaMin.Name = "txtKwotaMin"
        Me.txtKwotaMin.Size = New System.Drawing.Size(91, 20)
        Me.txtKwotaMin.TabIndex = 192
        '
        'Label26
        '
        Me.Label26.ForeColor = System.Drawing.Color.Navy
        Me.Label26.Location = New System.Drawing.Point(7, 82)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(262, 16)
        Me.Label26.TabIndex = 193
        Me.Label26.Text = "Minimalna kwota ró¿nicy wartoœci sprzeda¿y . . . . . . . . ."
        '
        'dtpOdDaty2
        '
        Me.dtpOdDaty2.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpOdDaty2.CustomFormat = "yyyy-MM-dd"
        Me.dtpOdDaty2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpOdDaty2.Location = New System.Drawing.Point(243, 33)
        Me.dtpOdDaty2.Name = "dtpOdDaty2"
        Me.dtpOdDaty2.Size = New System.Drawing.Size(91, 20)
        Me.dtpOdDaty2.TabIndex = 187
        Me.dtpOdDaty2.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(196, 37)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 16)
        Me.Label3.TabIndex = 191
        Me.Label3.Text = "Od daty . . . . . . . . . "
        '
        'dtpDoDaty2
        '
        Me.dtpDoDaty2.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpDoDaty2.CustomFormat = "yyyy-MM-dd"
        Me.dtpDoDaty2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDoDaty2.Location = New System.Drawing.Point(393, 33)
        Me.dtpDoDaty2.Name = "dtpDoDaty2"
        Me.dtpDoDaty2.Size = New System.Drawing.Size(91, 20)
        Me.dtpDoDaty2.TabIndex = 189
        Me.dtpDoDaty2.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(7, 37)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(183, 16)
        Me.Label4.TabIndex = 188
        Me.Label4.Text = "Porównaj do sprzeda¿y w okresie"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(344, 37)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 16)
        Me.Label5.TabIndex = 190
        Me.Label5.Text = "Do daty . . . . . . . . . "
        '
        'dtpOdDaty1
        '
        Me.dtpOdDaty1.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpOdDaty1.CustomFormat = "yyyy-MM-dd"
        Me.dtpOdDaty1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpOdDaty1.Location = New System.Drawing.Point(243, 10)
        Me.dtpOdDaty1.Name = "dtpOdDaty1"
        Me.dtpOdDaty1.Size = New System.Drawing.Size(91, 20)
        Me.dtpOdDaty1.TabIndex = 179
        Me.dtpOdDaty1.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(196, 14)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 16)
        Me.Label2.TabIndex = 186
        Me.Label2.Text = "Od daty . . . . . . . . . "
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(10, 109)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(474, 4)
        Me.ProgressBar1.TabIndex = 185
        '
        'cbxJakSieZmienilo
        '
        Me.cbxJakSieZmienilo.ForeColor = System.Drawing.Color.Purple
        Me.cbxJakSieZmienilo.Location = New System.Drawing.Point(243, 56)
        Me.cbxJakSieZmienilo.Name = "cbxJakSieZmienilo"
        Me.cbxJakSieZmienilo.Size = New System.Drawing.Size(241, 21)
        Me.cbxJakSieZmienilo.TabIndex = 184
        Me.cbxJakSieZmienilo.Text = "ComboBox1"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(7, 59)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(327, 16)
        Me.Label1.TabIndex = 183
        Me.Label1.Text = "Wybierz klientów którzy  . . . . . . . . . . . . . . . . . . . . ."
        '
        'dtpDoDaty1
        '
        Me.dtpDoDaty1.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpDoDaty1.CustomFormat = "yyyy-MM-dd"
        Me.dtpDoDaty1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDoDaty1.Location = New System.Drawing.Point(393, 10)
        Me.dtpDoDaty1.Name = "dtpDoDaty1"
        Me.dtpDoDaty1.Size = New System.Drawing.Size(91, 20)
        Me.dtpDoDaty1.TabIndex = 181
        Me.dtpDoDaty1.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'lblOdDaty
        '
        Me.lblOdDaty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblOdDaty.ForeColor = System.Drawing.Color.Navy
        Me.lblOdDaty.Location = New System.Drawing.Point(7, 14)
        Me.lblOdDaty.Name = "lblOdDaty"
        Me.lblOdDaty.Size = New System.Drawing.Size(183, 16)
        Me.lblOdDaty.TabIndex = 180
        Me.lblOdDaty.Text = "Analizuj sprzeda¿ w okresie"
        '
        'lblDoDaty
        '
        Me.lblDoDaty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblDoDaty.ForeColor = System.Drawing.Color.Navy
        Me.lblDoDaty.Location = New System.Drawing.Point(344, 14)
        Me.lblDoDaty.Name = "lblDoDaty"
        Me.lblDoDaty.Size = New System.Drawing.Size(56, 16)
        Me.lblDoDaty.TabIndex = 182
        Me.lblDoDaty.Text = "Do daty . . . . . . . . . "
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 144)
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
        Me.cmdStop.Location = New System.Drawing.Point(430, 145)
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
        Me.cmdZapamietaj.Location = New System.Drawing.Point(294, 145)
        Me.cmdZapamietaj.Name = "cmdZapamietaj"
        Me.cmdZapamietaj.Size = New System.Drawing.Size(130, 23)
        Me.cmdZapamietaj.TabIndex = 42
        Me.cmdZapamietaj.Text = "Pobierz klientów"
        Me.cmdZapamietaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'AkcjaMarketingowaWybierzWgObrotow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(522, 174)
        Me.Controls.Add(Me.cmdZapamietaj)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "AkcjaMarketingowaWybierzWgObrotow"
        Me.ShowInTaskbar = False
        Me.Text = "Wybór klientów wg zmiennoœci obrotów"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents cmdZapamietaj As System.Windows.Forms.Button
    Friend WithEvents dtpOdDaty1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpDoDaty1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblOdDaty As System.Windows.Forms.Label
    Friend WithEvents lblDoDaty As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents cbxJakSieZmienilo As System.Windows.Forms.ComboBox
    Friend WithEvents dtpOdDaty2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpDoDaty2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtKwotaMin As TextBox
    Friend WithEvents Label26 As Label
End Class
