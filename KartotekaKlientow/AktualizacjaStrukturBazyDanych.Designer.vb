<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AktualizacjaStrukturBazyDanych
    Inherits System.Windows.Forms.Form


    Private _CoAktualizujemy As String = ""

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _CoAktualizujemy = ""
    End Sub


    Public Sub New(ByVal pCoAktualizujemy As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _CoAktualizujemy = pCoAktualizujemy
    End Sub



    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AktualizacjaStrukturBazyDanych))
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.cmdAktualizuj = New System.Windows.Forms.Button()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.txtRaport = New System.Windows.Forms.TextBox()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 239)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 35
        Me.PictureBox2.TabStop = False
        '
        'cmdAktualizuj
        '
        Me.cmdAktualizuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdAktualizuj.Image = CType(resources.GetObject("cmdAktualizuj.Image"), System.Drawing.Image)
        Me.cmdAktualizuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAktualizuj.Location = New System.Drawing.Point(296, 240)
        Me.cmdAktualizuj.Name = "cmdAktualizuj"
        Me.cmdAktualizuj.Size = New System.Drawing.Size(152, 23)
        Me.cmdAktualizuj.TabIndex = 34
        Me.cmdAktualizuj.Text = "Aktualizuj bazê danych"
        Me.cmdAktualizuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(456, 240)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(136, 23)
        Me.cmdPowrot.TabIndex = 33
        Me.cmdPowrot.Text = "OK, skoñczy³em..."
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtRaport
        '
        Me.txtRaport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRaport.Location = New System.Drawing.Point(8, 8)
        Me.txtRaport.Multiline = True
        Me.txtRaport.Name = "txtRaport"
        Me.txtRaport.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtRaport.Size = New System.Drawing.Size(584, 224)
        Me.txtRaport.TabIndex = 32
        Me.txtRaport.Text = "TextBox1"
        '
        'AktualizacjaStrukturBazyDanych
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(600, 270)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.cmdAktualizuj)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.txtRaport)
        Me.Name = "AktualizacjaStrukturBazyDanych"
        Me.Text = "Aktualizacja Struktur Bazy Danych"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdAktualizuj As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents txtRaport As System.Windows.Forms.TextBox
End Class
