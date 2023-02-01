'Public Delegate Sub delegateAktualizuj()

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AkcjaMarketingowaWybierzWgUslug
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AkcjaMarketingowaWybierzWgUslug))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.chkU9 = New System.Windows.Forms.CheckBox()
        Me.chkU8 = New System.Windows.Forms.CheckBox()
        Me.chkU7 = New System.Windows.Forms.CheckBox()
        Me.chkU6 = New System.Windows.Forms.CheckBox()
        Me.chkU5 = New System.Windows.Forms.CheckBox()
        Me.chkU4 = New System.Windows.Forms.CheckBox()
        Me.chkU3 = New System.Windows.Forms.CheckBox()
        Me.chkU2 = New System.Windows.Forms.CheckBox()
        Me.chkU1 = New System.Windows.Forms.CheckBox()
        Me.dtpOdDaty = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dtpDoDaty = New System.Windows.Forms.DateTimePicker()
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
        Me.Panel1.Controls.Add(Me.chkU9)
        Me.Panel1.Controls.Add(Me.chkU8)
        Me.Panel1.Controls.Add(Me.chkU7)
        Me.Panel1.Controls.Add(Me.chkU6)
        Me.Panel1.Controls.Add(Me.chkU5)
        Me.Panel1.Controls.Add(Me.chkU4)
        Me.Panel1.Controls.Add(Me.chkU3)
        Me.Panel1.Controls.Add(Me.chkU2)
        Me.Panel1.Controls.Add(Me.chkU1)
        Me.Panel1.Controls.Add(Me.dtpOdDaty)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.ProgressBar1)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.dtpDoDaty)
        Me.Panel1.Controls.Add(Me.lblOdDaty)
        Me.Panel1.Controls.Add(Me.lblDoDaty)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(504, 284)
        Me.Panel1.TabIndex = 0
        '
        'chkU9
        '
        Me.chkU9.ForeColor = System.Drawing.Color.Navy
        Me.chkU9.Location = New System.Drawing.Point(25, 238)
        Me.chkU9.Name = "chkU9"
        Me.chkU9.Size = New System.Drawing.Size(432, 18)
        Me.chkU9.TabIndex = 195
        '
        'chkU8
        '
        Me.chkU8.ForeColor = System.Drawing.Color.Navy
        Me.chkU8.Location = New System.Drawing.Point(25, 216)
        Me.chkU8.Name = "chkU8"
        Me.chkU8.Size = New System.Drawing.Size(432, 18)
        Me.chkU8.TabIndex = 194
        '
        'chkU7
        '
        Me.chkU7.ForeColor = System.Drawing.Color.Navy
        Me.chkU7.Location = New System.Drawing.Point(25, 194)
        Me.chkU7.Name = "chkU7"
        Me.chkU7.Size = New System.Drawing.Size(432, 18)
        Me.chkU7.TabIndex = 193
        Me.chkU7.Text = "NAPRAWA DRUKAREK LASEROWYCH"
        '
        'chkU6
        '
        Me.chkU6.ForeColor = System.Drawing.Color.Navy
        Me.chkU6.Location = New System.Drawing.Point(25, 172)
        Me.chkU6.Name = "chkU6"
        Me.chkU6.Size = New System.Drawing.Size(432, 18)
        Me.chkU6.TabIndex = 192
        Me.chkU6.Text = "CZYSZCZENIE DRUKAREK LASEROWYCH"
        '
        'chkU5
        '
        Me.chkU5.ForeColor = System.Drawing.Color.Navy
        Me.chkU5.Location = New System.Drawing.Point(25, 150)
        Me.chkU5.Name = "chkU5"
        Me.chkU5.Size = New System.Drawing.Size(432, 18)
        Me.chkU5.TabIndex = 191
        Me.chkU5.Text = "PRZEGL¥D TECHNICZNY"
        '
        'chkU4
        '
        Me.chkU4.ForeColor = System.Drawing.Color.Navy
        Me.chkU4.Location = New System.Drawing.Point(25, 128)
        Me.chkU4.Name = "chkU4"
        Me.chkU4.Size = New System.Drawing.Size(432, 18)
        Me.chkU4.TabIndex = 190
        Me.chkU4.Text = "DOBOR URZ¥DZENIA DRUKUJ¥CEGO"
        '
        'chkU3
        '
        Me.chkU3.ForeColor = System.Drawing.Color.Navy
        Me.chkU3.Location = New System.Drawing.Point(25, 106)
        Me.chkU3.Name = "chkU3"
        Me.chkU3.Size = New System.Drawing.Size(432, 18)
        Me.chkU3.TabIndex = 189
        Me.chkU3.Text = "AUDYT DRUKU"
        '
        'chkU2
        '
        Me.chkU2.ForeColor = System.Drawing.Color.Navy
        Me.chkU2.Location = New System.Drawing.Point(25, 84)
        Me.chkU2.Name = "chkU2"
        Me.chkU2.Size = New System.Drawing.Size(432, 18)
        Me.chkU2.TabIndex = 188
        Me.chkU2.Text = "WSPARCIE INFORMATYCZNE"
        '
        'chkU1
        '
        Me.chkU1.ForeColor = System.Drawing.Color.Navy
        Me.chkU1.Location = New System.Drawing.Point(25, 62)
        Me.chkU1.Name = "chkU1"
        Me.chkU1.Size = New System.Drawing.Size(432, 18)
        Me.chkU1.TabIndex = 187
        Me.chkU1.Text = "PLATFORMA DRUKU"
        '
        'dtpOdDaty
        '
        Me.dtpOdDaty.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpOdDaty.CustomFormat = "yyyy-MM-dd"
        Me.dtpOdDaty.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpOdDaty.Location = New System.Drawing.Point(243, 10)
        Me.dtpOdDaty.Name = "dtpOdDaty"
        Me.dtpOdDaty.Size = New System.Drawing.Size(91, 20)
        Me.dtpOdDaty.TabIndex = 179
        Me.dtpOdDaty.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(191, 14)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 16)
        Me.Label2.TabIndex = 186
        Me.Label2.Text = "Od daty . . . . . . . . . "
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(10, 262)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(474, 4)
        Me.ProgressBar1.TabIndex = 185
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(7, 43)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(490, 16)
        Me.Label1.TabIndex = 183
        Me.Label1.Text = "Wybierz klientów którzy wybrali nastêpuj¹ce Us³ugi w ramach Rozszerzonego Zakresu" &
    " Us³ug :"
        '
        'dtpDoDaty
        '
        Me.dtpDoDaty.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpDoDaty.CustomFormat = "yyyy-MM-dd"
        Me.dtpDoDaty.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDoDaty.Location = New System.Drawing.Point(393, 10)
        Me.dtpDoDaty.Name = "dtpDoDaty"
        Me.dtpDoDaty.Size = New System.Drawing.Size(91, 20)
        Me.dtpDoDaty.TabIndex = 181
        Me.dtpDoDaty.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'lblOdDaty
        '
        Me.lblOdDaty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblOdDaty.ForeColor = System.Drawing.Color.Navy
        Me.lblOdDaty.Location = New System.Drawing.Point(7, 14)
        Me.lblOdDaty.Name = "lblOdDaty"
        Me.lblOdDaty.Size = New System.Drawing.Size(112, 16)
        Me.lblOdDaty.TabIndex = 180
        Me.lblOdDaty.Text = "Analizuj okres"
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
        Me.PictureBox1.Location = New System.Drawing.Point(8, 298)
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
        Me.cmdStop.Location = New System.Drawing.Point(430, 299)
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
        Me.cmdZapamietaj.Location = New System.Drawing.Point(294, 299)
        Me.cmdZapamietaj.Name = "cmdZapamietaj"
        Me.cmdZapamietaj.Size = New System.Drawing.Size(130, 23)
        Me.cmdZapamietaj.TabIndex = 42
        Me.cmdZapamietaj.Text = "Pobierz klientów"
        Me.cmdZapamietaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'AkcjaMarketingowaWybierzWgUslug
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(522, 328)
        Me.Controls.Add(Me.cmdZapamietaj)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "AkcjaMarketingowaWybierzWgUslug"
        Me.ShowInTaskbar = False
        Me.Text = "Wybór klientów wg wyboru us³ug"
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents cmdZapamietaj As System.Windows.Forms.Button
    Friend WithEvents dtpOdDaty As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpDoDaty As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblOdDaty As System.Windows.Forms.Label
    Friend WithEvents lblDoDaty As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkU9 As System.Windows.Forms.CheckBox
    Friend WithEvents chkU8 As System.Windows.Forms.CheckBox
    Friend WithEvents chkU7 As System.Windows.Forms.CheckBox
    Friend WithEvents chkU6 As System.Windows.Forms.CheckBox
    Friend WithEvents chkU5 As System.Windows.Forms.CheckBox
    Friend WithEvents chkU4 As System.Windows.Forms.CheckBox
    Friend WithEvents chkU3 As System.Windows.Forms.CheckBox
    Friend WithEvents chkU2 As System.Windows.Forms.CheckBox
    Friend WithEvents chkU1 As System.Windows.Forms.CheckBox
End Class
