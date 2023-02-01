'Public Delegate Sub delegateAktualizuj()

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AkcjaMarketingowaWybierzWgOsobKontaktowych
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AkcjaMarketingowaWybierzWgOsobKontaktowych))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtImieINazwisko = New System.Windows.Forms.TextBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.lblImieINazwisko = New System.Windows.Forms.Label()
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
        Me.Panel1.Controls.Add(Me.txtImieINazwisko)
        Me.Panel1.Controls.Add(Me.ProgressBar1)
        Me.Panel1.Controls.Add(Me.lblImieINazwisko)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(641, 57)
        Me.Panel1.TabIndex = 0
        '
        'txtImieINazwisko
        '
        Me.txtImieINazwisko.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtImieINazwisko.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtImieINazwisko.ForeColor = System.Drawing.Color.Purple
        Me.txtImieINazwisko.Location = New System.Drawing.Point(195, 11)
        Me.txtImieINazwisko.Name = "txtImieINazwisko"
        Me.txtImieINazwisko.Size = New System.Drawing.Size(426, 20)
        Me.txtImieINazwisko.TabIndex = 193
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(10, 43)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(611, 4)
        Me.ProgressBar1.TabIndex = 185
        '
        'lblImieINazwisko
        '
        Me.lblImieINazwisko.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblImieINazwisko.ForeColor = System.Drawing.Color.Navy
        Me.lblImieINazwisko.Location = New System.Drawing.Point(7, 14)
        Me.lblImieINazwisko.Name = "lblImieINazwisko"
        Me.lblImieINazwisko.Size = New System.Drawing.Size(300, 16)
        Me.lblImieINazwisko.TabIndex = 180
        Me.lblImieINazwisko.Text = "Imiê i nazwisko osoby kontaktowej . . . . . . . . . . . . ."
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 71)
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
        Me.cmdStop.Location = New System.Drawing.Point(567, 72)
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
        Me.cmdZapamietaj.Location = New System.Drawing.Point(431, 72)
        Me.cmdZapamietaj.Name = "cmdZapamietaj"
        Me.cmdZapamietaj.Size = New System.Drawing.Size(130, 23)
        Me.cmdZapamietaj.TabIndex = 42
        Me.cmdZapamietaj.Text = "Pobierz klientów"
        Me.cmdZapamietaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'AkcjaMarketingowaWybierzWgOsobKontaktowych
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(659, 101)
        Me.Controls.Add(Me.cmdZapamietaj)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "AkcjaMarketingowaWybierzWgOsobKontaktowych"
        Me.ShowInTaskbar = False
        Me.Text = "Wybór klientów wg osoby kontaktowej"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents cmdZapamietaj As System.Windows.Forms.Button
    Friend WithEvents lblImieINazwisko As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents txtImieINazwisko As TextBox
End Class
