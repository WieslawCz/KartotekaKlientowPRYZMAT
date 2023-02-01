<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FiltrowanieWybierzWartoscPola
    Inherits System.Windows.Forms.Form

    Dim _NazwaPola As String = ""
    Dim _NazwaTabeli As String = ""
    Dim _Filtr As String = ""

    Public Sub New(ByVal ePole As String, _
                   ByVal eTabela As String, _
                   ByVal eFiltr As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _NazwaPola = ePole
        _NazwaTabeli = eTabela
        _Filtr = eFiltr
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FiltrowanieWybierzWartoscPola))
        Me.dagWartosc = New System.Windows.Forms.DataGrid()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdWybierz = New System.Windows.Forms.Button()
        CType(Me.dagWartosc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dagWartosc
        '
        Me.dagWartosc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagWartosc.DataMember = ""
        Me.dagWartosc.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dagWartosc.Location = New System.Drawing.Point(8, 8)
        Me.dagWartosc.Name = "dagWartosc"
        Me.dagWartosc.Size = New System.Drawing.Size(330, 395)
        Me.dagWartosc.TabIndex = 53
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 409)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 54
        Me.PictureBox1.TabStop = False
        '
        'cmdWybierz
        '
        Me.cmdWybierz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierz.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWybierz.ForeColor = System.Drawing.Color.Black
        Me.cmdWybierz.Image = CType(resources.GetObject("cmdWybierz.Image"), System.Drawing.Image)
        Me.cmdWybierz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWybierz.Location = New System.Drawing.Point(252, 409)
        Me.cmdWybierz.Name = "cmdWybierz"
        Me.cmdWybierz.Size = New System.Drawing.Size(80, 23)
        Me.cmdWybierz.TabIndex = 55
        Me.cmdWybierz.Text = "Powrót"
        Me.cmdWybierz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'FiltrowanieWybierzWartoscPola
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(344, 437)
        Me.Controls.Add(Me.cmdWybierz)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.dagWartosc)
        Me.Name = "FiltrowanieWybierzWartoscPola"
        Me.ShowInTaskbar = False
        Me.Text = "Wybierz Wartoœæ Pola"
        CType(Me.dagWartosc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dagWartosc As System.Windows.Forms.DataGrid
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdWybierz As System.Windows.Forms.Button
End Class
